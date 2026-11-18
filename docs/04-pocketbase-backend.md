# 4. PocketBase backend

Everything is plain PocketBase with JavaScript hooks. No Go extension is needed.

```text
pocketbase/
├── Dockerfile                         Alpine + official PocketBase release binary
├── entrypoint.sh                      upserts the superuser, starts `serve`
├── pb_migrations/
│   └── 1763491200_entra_sharepoint.js collections and fields
└── pb_hooks/
    ├── auth.pb.js                     OAuth2 provider config, OAuth2 hook, /api/app-config, /api/entra-login
    ├── data.pb.js                     directory + SharePoint routes, mirror route, cron
    └── lib/
        ├── jwt.js                     RS256 JWT verification in plain JavaScript
        ├── entra.js                   token -> identity -> PocketBase user
        ├── graph.js                   app-only Graph token and requests
        ├── sharepoint.js              list items, Person column resolution, item creation
        └── mirror.js                  copies directory and list into collections
```

## Container

[`pocketbase/Dockerfile`](../pocketbase/Dockerfile) downloads the release for the build platform (`TARGETARCH`), so it builds on both Intel and Apple Silicon. `pb_hooks` and `pb_migrations` are mounted as volumes; PocketBase reloads changed hook files automatically. It does **not** notice deleted files, so restart the container after removing one.

Environment ([`.env.example`](../.env.example)):

| Variable | Purpose |
| --- | --- |
| `PB_SUPERUSER_EMAIL`, `PB_SUPERUSER_PASSWORD` | admin UI login at `/_/`, created on start |
| `ENTRA_TENANT_ID`, `ENTRA_CLIENT_ID`, `ENTRA_CLIENT_SECRET` | from [2. Entra](02-entra-app-registration.md) |
| `SHAREPOINT_SITE`, `SHAREPOINT_LIST`, `SHAREPOINT_OWNER_COLUMN` | from [3. SharePoint](03-sharepoint.md) |

## Collections (migration)

[`pb_migrations/1763491200_entra_sharepoint.js`](../pocketbase/pb_migrations/1763491200_entra_sharepoint.js):

| Collection | Fields | API rules |
| --- | --- | --- |
| `users` (built-in, extended) | `+ entraOid` (unique when set), `+ upn` | default |
| `directory_users` | `oid` (unique), `displayName`, `mail`, `upn`, `jobTitle`, `accountEnabled`, `account → users` | list/view: signed in |
| `sp_items` | `spItemId`, `siteId`, `listId` (unique pair `listId`+`spItemId`), `title`, `ownerEmail`, `owner → directory_users`, `modified`, `fields` (JSON) | list/view: signed in |
| `item_notes` | `item → sp_items` (cascade delete), `author → users`, `text`, `created` | list/view: signed in; create: only as yourself; update/delete: author only |

Mirrored collections have no create/update/delete rules, so only the hooks (and superusers) write them. [6. Relations](06-relations.md) explains each relation.

## Customization 1: built-in Microsoft OAuth2, tenant-only

[`auth.pb.js`](../pocketbase/pb_hooks/auth.pb.js), `onBootstrap`: configures the `microsoft` provider on `users` from the environment, with tenant-specific endpoints so only your organization can sign in:

```js
const tenantEndpoint = `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0`
unmarshal({ oauth2: { enabled: true, providers: [{
    name: "microsoft", clientId, clientSecret,
    authURL: tenantEndpoint + "/authorize",
    tokenURL: tenantEndpoint + "/token",
}] } }, users)
$app.save(users)
```

The same can be done by hand in the admin UI (**users > Options > OAuth2**). `unmarshal` applies the JSON field names (`authURL`, `tokenURL`) directly to the Go struct.

`onRecordAuthWithOAuth2Request` stores the Entra object ID (`e.oauth2User.id`, Graph's `/me` `id`) in `users.entraOid` after a built-in OAuth2 sign-in, so paths A and B end up as the same user.

## Customization 2: `/api/entra-login`

```js
routerAdd("POST", "/api/entra-login", (e) => {
    const entra = require(`${__hooks}/lib/entra.js`)
    const body = e.requestInfo().body
    const method = body.method === "obo" ? "obo" : "jwks"
    let identity
    try {
        identity = method === "obo" ? entra.identityFromOnBehalfOf(body.token) : entra.identityFromJwks(body.token)
    } catch (error) {
        throw new UnauthorizedError(`Entra token rejected (${method}): ${error}`)
    }
    const record = entra.findOrCreateUser(identity)
    return $apis.recordAuthResponse(e, record, "entra-" + method, { method })
})
```

`$apis.recordAuthResponse` returns `{ token, record, meta }`, exactly like the built-in auth endpoints, so the SDK's `authStore` works unchanged.

### Verifying the token in JavaScript (`lib/jwt.js`)

PocketBase's `$security.parseJWT` only supports HS256, but RS256 can be done in plain JavaScript inside the PocketBase JS runtime:

1. Split the JWT, base64url-decode header and payload (UTF-8 safe).
2. Reject anything but `alg: RS256`; pick the key by `kid`.
3. `signature^e mod n` with `BigInt` (square-and-multiply).
4. Compare with the expected PKCS#1 v1.5 encoding: `00 01 FF.. 00 || SHA-256 DigestInfo prefix || $security.sha256(header + "." + payload)`.
5. Check claims: `exp`/`nbf` (5 minutes leeway), `aud` (client ID or `api://<client-id>`), `iss` (v2 or v1 issuer of your tenant), `tid`.

`lib/entra.js` additionally requires the `access_as_user` scope in `scp`, and caches Entra's JWKS (`https://login.microsoftonline.com/<tenant>/discovery/v2.0/keys`) in `$app.store()` for an hour, refetching early when a token names an unknown `kid` (key rotation).

Verified: real Entra v2 tokens from MSAL and Teams NAA are accepted. Tampered payloads, wrong keys, `alg: none`, expired tokens and wrong audience, issuer or tenant are rejected. The round trip of `/api/entra-login` is about 50-60 ms.

### On-behalf-of alternative

`identityFromOnBehalfOf` exchanges the user token (`grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer`, `requested_token_use=on_behalf_of`, scope `https://graph.microsoft.com/User.Read`) and calls `/me`. No local crypto, but one or two extra network calls (0.3-1 s). Use it if you also need delegated Graph calls on the user's behalf.

### Finding or creating the user

`findOrCreateUser` looks up `users` by `entraOid`. If none exists it creates one with a random password and `verified: true`. It then updates `name`, `upn` and `email` (lower-cased) and links the matching `directory_users` entry (`account`).

## Customization 3: app-only Graph (`lib/graph.js`)

- `appToken()`: client-credentials token for `https://graph.microsoft.com/.default`, cached in `$app.store()` until 60 s before expiry.
- `graphGet(path)`, `graphGetAll(path)` (follows `@odata.nextLink`), `graphRequest(method, path, { body, headers, accessToken })`.

PocketBase JS hooks notes:

- Each handler runs in its own JS runtime from a pool, so module-level variables are **not** shared between requests. Use `$app.store()` for caches and `require()` inside each handler.
- `$http.send` returns `{ statusCode, json, body, headers }`; `body` is bytes, use `toString(response.body)` for text.

## Customization 4: data routes ([`data.pb.js`](../pocketbase/pb_hooks/data.pb.js))

All routes require a signed-in `users` record (`$apis.requireAuth("users")`).

| Route | Graph call | Returns |
| --- | --- | --- |
| `GET /api/directory/users?search=` | `/users?$filter=startswith(displayName,'…')&$top=25` | live tenant users |
| `GET /api/sharepoint/items` | list items `$expand=fields`, columns, User Information List | `{ site, list, rows[] }`, each row with `people` resolved to emails |
| `POST /api/sharepoint/items` | `POST …/lists/{id}/items { fields }` | created item; needs the `write` role (not tested) |
| `POST /api/mirror` | all of the above | `{ directory: { users }, sharePoint: { items, ownerColumn } }` |

A cron job runs the mirror every 15 minutes:

```js
cronAdd("mirror-directory-and-sharepoint", "*/15 * * * *", () => { … mirror.mirrorAll() … })
```

Mirroring several hundred users plus a small list took about 2 seconds. For large tenants, switch to Graph delta queries (`/users/delta`, list item `delta`) instead of full reads.

## Run it

```bash
docker compose up -d --build
docker compose logs -f pocketbase
```

Admin UI: <http://localhost:8080/_/> with the superuser from `.env`.

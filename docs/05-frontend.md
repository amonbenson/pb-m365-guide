# 5. Frontend

A dependency-light example without a framework: ES modules bundled with esbuild and served by nginx.

```text
web/
├── Dockerfile               node build stage -> nginx
├── nginx.conf               static files + proxy /api and /_/ to PocketBase (SSE-safe)
├── package.json             @azure/msal-browser, @microsoft/teams-js, pocketbase, esbuild
├── public/
│   ├── index.html           demo page
│   ├── redirect.html        MSAL v5 redirect bridge
│   └── tab-config.html      Teams channel tab configuration page
└── src/
    ├── auth.js              Teams detection, MSAL (nestable), token acquisition
    ├── data.js              every PocketBase call: auth, users, lists, notes
    ├── app.js               page wiring
    ├── redirect.js          broadcastResponseToMainFrame()
    └── tab-config.js        pages.config for channel tabs
```

## nginx

```nginx
location ~ ^/(api|_)/ {
    proxy_pass http://pocketbase:8090;
    proxy_http_version 1.1;
    proxy_set_header Connection "";
    proxy_buffering off;          # PocketBase realtime (SSE)
    proxy_read_timeout 360s;
    ...
}
```

## Authentication ([`src/auth.js`](../web/src/auth.js))

### 1. Detect Teams

```js
await Promise.race([teamsApp.initialize(), timeout(3000)])
const context = await teamsApp.getContext()
teamsApp.notifySuccess()
```

Outside Teams `initialize()` rejects ("No Parent window found"); the timeout covers hosts where it hangs. Initialize TeamsJS **before** MSAL. Useful context fields: `app.host.name` / `clientType`, `page.frameContext`, `channel.membershipType` (`Regular`, `Private`, `Shared`), `user.loginHint`.

### 2. Create the MSAL client

```js
const configuration = await (await fetch("/api/app-config")).json()
const msal = await createNestablePublicClientApplication({
    auth: {
        clientId: configuration.clientId,
        authority: `https://login.microsoftonline.com/${configuration.tenantId}`,
        redirectUri: window.location.origin + "/redirect.html",
    },
    cache: { cacheLocation: "localStorage" },
})
const redirectResult = await msal.handleRedirectPromise()
```

- **MSAL v5:** `createNestablePublicClientApplication` replaces the old `supportsNestedAppAuth: true` option, which Microsoft's Teams NAA docs still show and which v5 removed. In Teams it uses the NAA bridge; elsewhere it returns a standard client.
- **Redirect bridge (new in v5):** popup, silent and redirect flows all return through the `redirectUri` page, which must call:
  ```js
  import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge"
  broadcastResponseToMainFrame()
  ```
  That page must not be served with a `Cross-Origin-Opener-Policy` header (nginx serves it without one).

### 3. Get a token

```js
const request = { scopes: [configuration.apiScope], loginHint: context?.user?.loginHint }

// silent: cached account first, otherwise reuse the Entra session cookie
const account = msal.getActiveAccount() || msal.getAllAccounts()[0]
const result = account
    ? await msal.acquireTokenSilent({ ...request, account })
    : await msal.ssoSilent(request)

// interactive: call straight from the click handler (Safari blocks popups opened after an await)
button.onclick = () => msal.acquireTokenPopup(request).then(signInWithEntraToken)

// fallback without popups
msal.acquireTokenRedirect(request)
```

What happened where:

| Host | Silent result |
| --- | --- |
| Teams web, desktop, Android (personal tab); Teams desktop channel tabs (standard and private) | `ssoSilent` via NAA succeeds, no UI |
| Chrome / Edge with an Entra session | `ssoSilent` succeeds (~1 s) |
| Chrome without a session | `timed_out` after ~10 s |
| Safari | `login_required` at once (third-party cookies blocked); use popup or redirect |

### 4. Exchange for a PocketBase session ([`src/data.js`](../web/src/data.js))

```js
const response = await fetch("/api/entra-login", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ token: result.accessToken, method: "jwks" }),
})
const { token, record } = await response.json()
pocketBase.authStore.save(token, record)
```

From here on all SDK calls are authenticated. The SDK keeps the session in `localStorage`; returning users skip the whole flow until it expires.

### Built-in PocketBase OAuth2 (browser only)

```js
await pocketBase.collection("users").authWithOAuth2({ provider: "microsoft" })
```

The SDK opens a blank popup, connects to PocketBase realtime (SSE), and only then navigates the popup to Microsoft. If SSE is blocked (for example a Cloudflare quick tunnel), the popup stays blank and the call fails after a timeout.

### Sign out

```js
pocketBase.authStore.clear()
await msal.clearCache()   // local MSAL cache only; the Entra session stays
```

## Tenant users

```js
// live from Graph (app-only, via PocketBase)
pocketBase.send("/api/directory/users", { query: { search: "Ann" } })

// mirrored collection, with the PocketBase account if the person has signed in
pocketBase.collection("directory_users").getList(1, 25, {
    filter: pocketBase.filter("displayName ~ {:search} || mail ~ {:search}", { search }),
    sort: "displayName",
    expand: "account",
})
```

## SharePoint list

```js
// live: { site, list: { personColumns }, rows: [{ id, title, fields, people: { Owner: { lookupId, email } } }] }
pocketBase.send("/api/sharepoint/items", {})

// create (needs the "write" site role)
pocketBase.send("/api/sharepoint/items", { method: "POST", body: { fields: { Title: "New", Status: "Planned" } } })

// mirrored, with owner (directory user) and owner's PocketBase account
pocketBase.collection("sp_items").getFullList({ sort: "title", expand: "owner,owner.account" })

// items owned by the signed-in user
pocketBase.collection("sp_items").getFullList({
    filter: pocketBase.filter("owner.account = {:userId}", { userId: pocketBase.authStore.record.id }),
})

// refresh the mirror now
pocketBase.send("/api/mirror", { method: "POST" })
```

## PocketBase data attached to SharePoint items

```js
pocketBase.collection("item_notes").create({ item: spItemRecordId, author: pocketBase.authStore.record.id, text })

pocketBase.collection("item_notes").getFullList({
    filter: pocketBase.filter("item = {:itemId}", { itemId: spItemRecordId }),
    sort: "created",
    expand: "author",
})

// reverse direction: an item with its notes
pocketBase.collection("sp_items").getOne(spItemRecordId, { expand: "item_notes_via_item" })
```

## Build and run

The bundle is built inside the image, so rebuild it after frontend changes:

```bash
docker compose up -d --build web
```

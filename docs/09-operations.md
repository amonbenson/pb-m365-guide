# 9. Operations: running, testing, cleanup

## Run locally

```bash
cp .env.example .env      # fill in values from steps 2 and 3
docker compose up -d --build
```

- App: <http://localhost:8080>
- PocketBase admin: <http://localhost:8080/_/>

`http://localhost:8080` is enough for every browser flow, including PocketBase's built-in OAuth2.

## Public HTTPS for Teams

Teams tabs and the `brk-multihub://` redirect need a public HTTPS host.

**Quick tunnel (testing only):**

```bash
docker compose --profile tunnel up -d
docker compose logs tunnel | grep trycloudflare.com
```

- The `*.trycloudflare.com` host changes every time the tunnel container restarts. You then have to update the Entra redirect URIs and rebuild the Teams package.
- Quick tunnels don't pass Server-Sent Events: PocketBase's built-in OAuth2 popup and realtime subscriptions fail through it ([caveat 7](08-caveats.md#7-built-in-oauth2-popup-needs-server-sent-events)).
- The URL is public. Everything except `/api/app-config` and `/api/entra-login` requires a signed-in user.

**Stable alternatives:** a named Cloudflare tunnel (needs a Cloudflare account and domain), a reserved ngrok domain, or a real deployment.

## Test checklist

| # | Test | Expected |
| --- | --- | --- |
| 1 | Browser with an Entra session, open the app | signed in silently |
| 2 | Safari, **Sign in (popup)** | popup, then signed in |
| 3 | **Sign in (PocketBase OAuth2)** on localhost | signed in, same user as #1 |
| 4 | **Search live** | tenant users from Graph |
| 5 | **Refresh mirror**, then **Search mirrored** | users in `directory_users`, your entry with `account` |
| 6 | **Items live** | list rows; `people.Owner.email` filled |
| 7 | **Items mirrored** | `expand.owner` filled; `expand.owner.expand.account` only for people who signed in |
| 8 | **Items I own** | only your items |
| 9 | Add a note to a mirrored item, **Show notes** | note with `expand.author` |
| 10 | Teams personal tab, standard and private channel tab | signed in silently, host line shows the context |

## Production notes

- Keep the frontend and PocketBase on one origin (as with nginx here), or configure CORS.
- Run behind real TLS; set PocketBase's trusted proxy headers (admin UI > Settings) so rate limits and logs see client IPs.
- Store `ENTRA_CLIENT_SECRET` in a secret store; prefer a certificate credential if your policy requires it, and track the expiry.
- Restrict `POST /api/mirror` (for example to superusers) if regular users shouldn't trigger syncs.
- Back up `pb_data` (the `pb_data` Docker volume).

## Cleanup

| Where | What |
| --- | --- |
| Entra admin center | **App registrations > your app > Delete** (removes secret and consents) |
| Graph Explorer | **Entra ID > Enterprise apps > Graph Explorer > Permissions**: revoke `Sites.FullControl.All` if no longer needed |
| SharePoint | `DELETE /sites/{site-id}/permissions/{permission-id}`, or delete the test site |
| Teams admin center | **Teams apps > Manage apps**: delete the app |
| Docker | `docker compose --profile tunnel down -v` (also deletes the PocketBase data volume) |

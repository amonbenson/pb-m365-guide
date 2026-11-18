# 8. Caveats and surprises

Everything here was observed while building and testing this setup. Items marked **unexpected** contradicted documentation or the original plan.

## 1. Person columns contain no email

**Unexpected.** Graph returns a SharePoint Person column only as `<Column>LookupId`, an ID in the site's hidden User Information List, not as an email or Entra ID. Resolve it by reading that list ([6. Relations](06-relations.md#4-sharepoint-person-column--tenant-user)).

## 2. The User Information List has a localized title

**Unexpected.** `GET /sites/{id}/lists/User Information List` returns `404 itemNotFound` on a German site, where the list is called "Benutzerinformationsliste". Find it by its internal name `users`. System lists are only returned when `system` is in `$select`:

```http
GET /sites/{site-id}/lists?$select=id,name,system
```

Expect the same for other built-in lists and column display names; use internal names.

## 3. Built-in columns are Person columns too

`Author` ("Created By") and `Editor` ("Modified By") show up as Person columns. Picking "the first Person column" as owner can hit them; they are `readOnly`, so filter those out or name the column explicitly (`SHAREPOINT_OWNER_COLUMN`).

## 4. `Sites.Selected` hides contents, not existence

**Unexpected.** With a grant on one site, other sites' names and URLs are still readable (`/sites/root`, `/sites/{host}:/sites/{path}` return 200). Their lists come back empty and their drives return 403; site search returns 403. Fine for most cases, but site names are not secret from the app.

## 5. MSAL v5 differs from Microsoft's NAA documentation

**Unexpected.** The Teams NAA docs still configure `supportsNestedAppAuth: true`. MSAL Browser v5 removed that option:

- Use `await createNestablePublicClientApplication(config)`.
- All popup, silent and redirect flows now go through a **redirect bridge** page at `redirectUri` that calls `broadcastResponseToMainFrame()` from `@azure/msal-browser/redirect-bridge`. Without it, popups and `ssoSilent` never complete.
- `handleRedirectPromise()` takes an options object; `getAccountByUsername` and the like are replaced by `getAccount({ … })`.

## 6. PocketBase blocks iframes (Teams tabs)

**Unexpected.** PocketBase 0.40 adds these headers to **every** response, static files from `pb_public` included:

```text
X-Frame-Options: SAMEORIGIN
Cross-Origin-Opener-Policy: same-origin
```

A frontend served by PocketBase therefore can't be embedded as a Teams tab. Serve the frontend separately (this repository uses nginx), or unbind the `pbSecurityHeaders` middleware for the frontend routes in a hook.

## 7. Built-in OAuth2 popup needs Server-Sent Events

**Unexpected.** The PocketBase SDK opens `about:blank`, waits for its realtime (SSE) connection, and only then loads Microsoft's sign-in page. Behind a **Cloudflare quick tunnel** (`*.trycloudflare.com`) SSE is buffered, so the popup stayed blank and the call failed after a timeout. Works on localhost, on a normal host, or behind a named Cloudflare tunnel. `/api/entra-login` doesn't use SSE and is unaffected.

## 8. Built-in OAuth2 is browser-only

Microsoft's sign-in page refuses to load inside the Teams tab iframe, so `authWithOAuth2` can't be used in Teams. Use `/api/entra-login` with NAA there.

## 9. PocketBase's Microsoft provider takes email from `mail`

Since PocketBase v0.40 the Microsoft provider fills the user's email from Graph's `mail` property (optionally from ID token claims via `idTokenEmailClaim`). Accounts without a mailbox get no email. Key users by `entraOid`, never by email.

## 10. `$security.parseJWT` is HS256 only

It can't verify Entra's RS256 tokens. A ~100-line plain-JavaScript verifier using `BigInt` and `$security.sha256` works inside the PocketBase JS runtime ([4. Backend](04-pocketbase-backend.md#verifying-the-token-in-javascript-libjwtjs)); no Go extension is needed.

## 11. Silent sign-in failures look different per browser

| Situation | Result |
| --- | --- |
| Chrome without an Entra session | `ssoSilent` waits ~10 s, then `timed_out` |
| Safari | `login_required` at once (third-party cookies blocked) |
| Chrome with a session | works, ~1 s |

Don't block the UI on `ssoSilent`. Try it when a session is likely (returning user, known `loginHint`), show a sign-in button otherwise, or lower `system.iframeBridgeTimeout`.

## 12. Safari blocks popups opened after an `await`

Call `acquireTokenPopup`/`loginPopup` synchronously in the click handler. If anything is awaited first, Safari blocks the popup (`popup_window_error`). The redirect flow is the fallback.

## 13. Production tenants usually block custom app upload

Normal users often can't upload Teams apps. Publish through the Teams admin center with availability limited to test users ([7. Teams](07-teams.md#publishing)), and expect the app to take a while to appear in clients.

## 14. Private channels

- A private channel has its own SharePoint site; `Sites.Selected` grants on it work like on any other site.
- In a private channel tab, `context.team` is empty; use `context.channel.id`.
- Apps in private channels were rolled out during the first half of 2026; older tenants may not have it yet.

## 15. Graph and SharePoint limits worth knowing

- Results are paged: `/users` allows `$top` up to 999; always follow `@odata.nextLink` (`graphGetAll` does).
- The mirror does full reads; for big tenants or lists use delta queries.
- Deleted list items are not removed from the mirror by this code.

## 16. PocketBase JS hooks

- Hook files are reloaded on change but **not** on deletion; restart the container.
- Each request handler gets its own JS runtime from a pool: no shared module state; use `$app.store()` and `require()` inside handlers.
- `$http.send(...).body` is a byte array; use `toString(...)` or `.json`.

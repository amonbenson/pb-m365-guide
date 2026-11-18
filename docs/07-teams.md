# 7. Teams tab

The same page runs as a personal tab and as a channel tab. Inside Teams, MSAL's nestable client signs in through Nested App Authentication (NAA) with no prompt.

## Requirements

- A public HTTPS origin (Teams won't load `http://localhost`), see [9. Operations](09-operations.md#public-https-for-teams).
- The SPA redirect URI `brk-multihub://<host>` in the Entra app ([2. Entra](02-entra-app-registration.md#2-redirect-uris)).
- Teams context must be initialized before MSAL ([5. Frontend](05-frontend.md#1-detect-teams)).

## App package

```bash
python3 teams/build_package.py app.example.com
```

[`teams/build_package.py`](../teams/build_package.py) reads `ENTRA_CLIENT_ID` from `.env` and writes `teams/app-package.zip` with:

- `manifest.json`, schema **1.22**:
  - `staticTabs`: personal tab → `https://<host>/index.html`
  - `configurableTabs`: channel and group chat tab → `https://<host>/tab-config.html`
  - `validDomains`: `[<host>]`
  - `webApplicationInfo`: `{ id: <client-id>, resource: "api://<client-id>" }`
  - `id`: derived from the client ID, so every deployment gets its own Teams app ID
- `color.png` (192×192) and `outline.png` (32×32, white on transparent): plain placeholders; replace them.

`nestedAppAuthInfo` (token prefetch) is optional and wasn't needed.

The channel tab configuration page ([`src/tab-config.js`](../web/src/tab-config.js)) only saves the content URL:

```js
await app.initialize()
pages.config.registerOnSaveHandler((saveEvent) => {
    pages.config.setConfig({ contentUrl: origin + "/index.html", entityId: "entra-demo-channel-tab", suggestedDisplayName: "Entra Demo" })
        .then(() => saveEvent.notifySuccess())
})
pages.config.setValidityState(true)
```

## Publishing

**If users may upload custom apps:** Teams > **Apps > Manage your apps > Upload an app > Upload a custom app**.

**If custom app upload is disabled** (common in production tenants), use the Teams admin center with an admin account:

1. <https://admin.teams.microsoft.com> > **Teams apps > Manage apps > Upload new app**.
2. Open the app > **Users and groups** > **Edit availability**: limit it to specific users or groups while testing.
3. Make sure the app is **Allowed**.

Submitting "for admin review" from a normal account and approving it as admin gives the same result.

### Making it show up faster

The Teams client can take hours to show newly published org apps. To get there sooner:

- Open Teams on the web in a private window: a fresh session doesn't use the desktop app's cached catalog.
- Fully quit and restart the desktop client.
- Open it by link: `https://teams.microsoft.com/l/app/<catalog app id>` (the **App ID** shown in the admin center, not the manifest ID).
- Install it for yourself with Graph:
  ```http
  GET  https://graph.microsoft.com/v1.0/appCatalogs/teamsApps?$filter=externalId eq '<manifest id>'
  POST https://graph.microsoft.com/v1.0/me/teamwork/installedApps
  { "teamsApp@odata.bind": "https://graph.microsoft.com/v1.0/appCatalogs/teamsApps/<catalog app id>" }
  ```
  (delegated `TeamsAppInstallation.ReadWriteSelfForUser`)

## Adding the tabs

- **Personal:** open the app from **Apps**.
- **Channel:** in the channel, **+ (Add a tab)** > the app > **Save**. Works in standard and private channels.

## Results

| Context | Client | Sign-in |
| --- | --- | --- |
| Personal tab | Teams web | silent (NAA) |
| Personal tab | Teams desktop (macOS) | silent (NAA) |
| Personal tab | Teams Android | silent (NAA) |
| Channel tab, standard channel | Teams desktop | silent (NAA) |
| Channel tab, private channel | Teams desktop | silent (NAA) |

`nestedAppAuth.isNAAChannelRecommended()` returned `true` in all of them. Not tested: channel tabs on web and mobile, shared channels, meetings, Outlook and the Microsoft 365 app.

In a **private** channel, `context.team` is empty (no `groupId`); identify the location by `context.channel.id`.

## Classic Teams SSO (not used)

`authentication.getAuthToken()` also returns a token for your API, but needs `webApplicationInfo.resource` in the form `api://<host>/<client-id>` (the Application ID URI must contain the host), plus pre-authorized Teams client IDs. NAA needs none of that and works the same in the browser, so this guide uses NAA only.

# 2. Entra app registration

Admin center: <https://entra.microsoft.com>. You need an account that can create app registrations and grant admin consent (for example Cloud Application Administrator or Global Administrator).

Throughout this guide, `https://app.example.com` is the public origin of your app and `app.example.com` its host name. For local-only testing use `http://localhost:8080`. Teams needs a public HTTPS host ([9. Operations](09-operations.md#public-https-for-teams)).

## 1. Register the app

**Entra ID > App registrations > New registration**

- Name: anything, for example `Entra Demo`.
- Supported account types: **Single tenant only**.
- Redirect URI: leave empty.
- Select **Register**. Note the **Application (client) ID** and the **Directory (tenant) ID** from the Overview page.

## 2. Redirect URIs

**Manage > Authentication > Redirect URI configuration > Add Redirect URI**. Add each platform separately and select **Configure**.

| Platform | Redirect URI | Used by |
| --- | --- | --- |
| **Web** | `https://app.example.com/api/oauth2-redirect` | PocketBase built-in OAuth2 (path B) |
| **Web** | `http://localhost:8080/api/oauth2-redirect` | same, local |
| **Single-page application** | `https://app.example.com/redirect.html` | MSAL popup, silent and redirect flows |
| **Single-page application** | `http://localhost:8080/redirect.html` | same, local |
| **Single-page application** | `brk-multihub://app.example.com` | Teams Nested App Authentication |

- PocketBase's OAuth2 redirect must be on the **Web** platform, not SPA: PocketBase redeems the code server-side, and Entra only allows SPA redirect URIs to be redeemed from a browser.
- `brk-multihub://` takes the origin only (host name, no path).
- Leave both **Implicit grant** checkboxes unchecked.

## 3. Client secret

**Manage > Certificates & secrets > Client secrets > New client secret**. Copy the **Value** (not the Secret ID) into `.env` as `ENTRA_CLIENT_SECRET`. Note the expiry date; PocketBase stops reaching Graph when it expires.

## 4. Expose an API

**Manage > Expose an API**

1. Next to **Application ID URI** select **Add**, keep the default `api://<client-id>`, **Save**.
2. **Add a scope**:
   - Scope name: `access_as_user`
   - Who can consent: **Admins and users**
   - Display names and descriptions: any text
   - State: **Enabled**

The frontend requests `api://<client-id>/access_as_user`; that token's audience is this app, which is what PocketBase checks.

## 5. Access token version 2

**Manage > Manifest** (Microsoft Graph format): set `"requestedAccessTokenVersion": 2` inside `"api"` (it is `null` by default) and **Save**.

With v2 tokens `aud` is the client ID and `iss` is `https://login.microsoftonline.com/<tenant>/v2.0`. The backend also accepts v1 tokens, but v2 is cleaner.

## 6. API permissions and admin consent

**Manage > API permissions > Add a permission > Microsoft Graph**

| Type | Permission | Why |
| --- | --- | --- |
| Delegated | `User.Read` | sign-in, `/me` (PocketBase OAuth2, OBO) |
| Delegated | `openid`, `profile`, `offline_access` | requested by MSAL |
| Application | `User.Read.All` | read the tenant directory |
| Application | `Sites.Selected` | SharePoint access to granted sites only |

Then **Grant admin consent for <tenant>**. Every row must show "Granted". With admin consent, users never see a consent screen.

`Sites.Selected` alone grants no SharePoint access; [3. SharePoint](03-sharepoint.md) adds the site grant.

## 7. Fill in `.env`

```bash
cp .env.example .env
```

```ini
ENTRA_TENANT_ID=<directory (tenant) id>
ENTRA_CLIENT_ID=<application (client) id>
ENTRA_CLIENT_SECRET=<secret value>
```

Never commit `.env`; `.gitignore` already excludes it.

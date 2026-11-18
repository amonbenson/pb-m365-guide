# 1. Architecture and decisions

## Goal

A web app that

- signs users in with their Microsoft Entra ID (work) account, silently where possible,
- runs in the browser and as a Microsoft Teams tab (personal and channel tabs),
- uses PocketBase as its backend and database,
- reads and relates data from SharePoint lists and the tenant's user directory.

No special Microsoft licenses are needed (see [Licensing](#licensing)).

## Components

```text
 Browser / Teams tab                      Docker host
 ┌────────────────────────┐   HTTPS    ┌──────────────────────────────────────────┐
 │ web app                │──────────▶│ nginx (web)                               │
 │  MSAL (nestable client)│           │  /            static frontend             │
 │  TeamsJS               │           │  /api, /_/    proxy ─▶ PocketBase :8090   │
 │  PocketBase JS SDK     │           └───────────────────────┬──────────────────┘
 └──────────┬─────────────┘                                   │ app-only token
            │ user token (MSAL / NAA)                         ▼
            ▼                                        Microsoft Graph
      Microsoft Entra ID                             /users, /sites/.../lists
```

- **One origin.** nginx serves the frontend and proxies `/api` and `/_/` to PocketBase, so there is no CORS setup and one public host name for redirect URIs and the Teams manifest.
- **Frontend not served by PocketBase.** PocketBase adds `X-Frame-Options: SAMEORIGIN` to every response, including `pb_public`, which would stop the page loading as a Teams tab ([caveat 6](08-caveats.md#6-pocketbase-blocks-iframes-teams-tabs)).
- **One Entra app registration** for everything: user sign-in (Web and SPA platforms), the exposed API scope, and app-only Graph access.

## Authentication paths

There are two ways to get a PocketBase session. Both resolve to the same PocketBase user, keyed by the Entra object ID (`users.entraOid`).

| Path | Where | How |
| --- | --- | --- |
| **A. Custom route `/api/entra-login`** (recommended) | Browser and Teams | Frontend gets an access token for `api://<client-id>/access_as_user` with MSAL (Teams: Nested App Authentication). PocketBase verifies it and returns a normal session. |
| **B. PocketBase built-in OAuth2** | Browser only | `authWithOAuth2({ provider: "microsoft" })` popup. Doesn't work inside Teams (Microsoft's sign-in page refuses to load in the tab iframe). |

Path A, step by step:

1. Frontend: in Teams, `createNestablePublicClientApplication` uses the Teams host as broker (NAA) and gets a token without any UI. In a normal browser the same client falls back to standard MSAL: `ssoSilent` (reuses the Entra session cookie), else popup or redirect.
2. Frontend: `POST /api/entra-login { token }`.
3. PocketBase verifies the token. Two methods are implemented:
   - **`jwks`** (default): RS256 signature against Entra's published keys, plus `aud`, `iss`, `tid`, `exp`/`nbf` and the `access_as_user` scope. Pure JavaScript in the hooks, ~50 ms round trip.
   - **`obo`**: on-behalf-of exchange for a Graph token, then `GET /me`. Entra rejects tokens not issued for this app. 0.3-1 s round trip.
4. PocketBase finds or creates the user by `oid` and returns `$apis.recordAuthResponse(...)`, the same response as a password login.

Never accept a plain Graph token and just call `/me` with it: any other app holding a user's Graph token could then sign in as that user. Accepting only tokens whose audience is this app's own API prevents that.

## Data access

PocketBase calls Microsoft Graph with an **app-only** token (client credentials):

- `User.Read.All` (application): read the tenant directory.
- `Sites.Selected` (application): no SharePoint access at all until an admin grants the app a role on a specific site.

Two ways to use the data, both implemented (details in [6. Relations](06-relations.md)):

- **Live**: custom routes call Graph on each request.
- **Mirror**: a cron job copies users and list items into PocketBase collections, which gives real relation fields, `expand`, filters and API rules.

## Licensing

| Area | What's needed |
| --- | --- |
| Sign-in (SSO) | Entra ID Free (included with Microsoft 365) covers SSO to custom apps. P1 only for extras such as Conditional Access or assigning the app to groups. ([Entra ID Free](https://learn.microsoft.com/azure/cost-management-billing/manage/microsoft-entra-id-free)) |
| Teams tab | A normal Teams license; Teams Premium is not needed. ([Custom apps](https://learn.microsoft.com/microsoftteams/custom-app-overview)) |
| SharePoint data in the app | Users who see SharePoint data through the app must themselves be licensed for SharePoint ("multiplexing"). Normal Microsoft 365 licenses cover this. |

This is not licensing advice; confirm the multiplexing point with your Microsoft partner.

## Versions this guide was tested with

| Component | Version |
| --- | --- |
| PocketBase | 0.40.6 |
| `@azure/msal-browser` | 5.29.1 |
| `@microsoft/teams-js` | 2.59.0 |
| `pocketbase` (JS SDK) | 0.29.0 |
| Teams app manifest | 1.22 |

Tested in November 2026 against a production Microsoft 365 tenant: Chrome and Safari on macOS; Teams web, desktop (macOS) and Android.

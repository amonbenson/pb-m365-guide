# PocketBase + Microsoft Entra ID + SharePoint + Teams

A step-by-step guide and working reference app for building a web app that

- signs users in with their **Microsoft Entra ID** work account, silently in the browser and inside **Microsoft Teams** tabs (Nested App Authentication),
- uses **PocketBase** as backend and database,
- reads, mirrors and relates **SharePoint list** items and **tenant users**.

Everything runs in Docker. No Go code and no special Microsoft licenses are needed.

## What you get

- **Login:** one custom PocketBase route, `/api/entra-login`, turns an Entra token from MSAL or Teams into a normal PocketBase session. The token's RS256 signature is verified in plain JavaScript hooks. PocketBase's built-in Microsoft OAuth2 also works in the browser, and both paths sign in the same user.
- **Data access:** app-only Microsoft Graph access, limited to chosen SharePoint sites with `Sites.Selected`.
- **Relations:** mirrored collections with real relations: SharePoint item → owner (Person column) → tenant user → PocketBase account, plus app-only data attached to SharePoint items.
- **Frontend:** a framework-free demo with every call needed for sign-in, users and lists.

## Guide

Follow in order:

| Step | Doc | Who |
| --- | --- | --- |
| 1 | [Architecture and decisions](docs/01-architecture.md) | everyone, read first |
| 2 | [Entra app registration](docs/02-entra-app-registration.md) | Entra admin |
| 3 | [SharePoint site, list and site grant](docs/03-sharepoint.md) | SharePoint admin |
| 4 | [PocketBase backend](docs/04-pocketbase-backend.md) | developer |
| 5 | [Frontend](docs/05-frontend.md) | developer |
| 6 | [Relations between PocketBase, SharePoint and users](docs/06-relations.md) | developer |
| 7 | [Teams tab](docs/07-teams.md) | developer + Teams admin |
| 8 | [Caveats and surprises](docs/08-caveats.md) | everyone |
| 9 | [Operations: running, testing, cleanup](docs/09-operations.md) | developer |

## Quick start

Prerequisites: Docker with Compose, a Microsoft 365 tenant where you (or an admin) can register apps and grant admin consent.

1. Do steps 2 and 3 of the guide (about 15 minutes in the admin centers).
2. Configure and start:
   ```bash
   cp .env.example .env
   ```
   Fill in `.env`, then:
   ```bash
   docker compose up -d --build
   ```
3. Open <http://localhost:8080>, sign in, and work through the [test checklist](docs/09-operations.md#test-checklist).
4. For Teams, start the tunnel profile and build the app package ([7. Teams](docs/07-teams.md)).

## Repository layout

```text
.
├── docker-compose.yml     pocketbase, web (nginx), optional tunnel
├── .env.example           all settings
├── pocketbase/            Dockerfile, migrations, JS hooks
├── web/                   frontend (MSAL v5, TeamsJS, PocketBase SDK) + nginx
├── teams/                 Teams app package builder
└── docs/                  the guide
```

## Status

A reference implementation, tested in November 2026 against a production Microsoft 365 tenant (versions in [1. Architecture](docs/01-architecture.md#versions-this-guide-was-tested-with)). Each doc marks the parts that were not tested. Review security settings before using it in production ([9. Operations](docs/09-operations.md#production-notes)).

## License

[MIT](LICENSE)

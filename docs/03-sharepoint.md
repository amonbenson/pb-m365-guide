# 3. SharePoint site, list and site grant

## 1. Site and list

Any existing site works, including the site behind a private Teams channel. For a new one, a **Communication site** avoids creating a Microsoft 365 group.

Create a list, for example `Projects`, with:

| Column | Type | Notes |
| --- | --- | --- |
| `Title` | default | |
| `Owner` | **Person** (people only, single value) | becomes the owner relation |
| `Status` | **Choice** | any extra column; it is mirrored as JSON |

Add a few items: some owned by you, some by colleagues who will never sign in to the app, one without an owner. That covers every relation case.

Set in `.env`:

```ini
SHAREPOINT_SITE=contoso.sharepoint.com:/sites/demo
SHAREPOINT_LIST=Projects
SHAREPOINT_OWNER_COLUMN=Owner
```

`SHAREPOINT_SITE` uses Graph's path syntax: `<host>:/sites/<site path>`. The list can be given by display name or ID.

## 2. Grant the app access to this site only

`Sites.Selected` needs a per-site grant: `POST /sites/{site-id}/permissions`. The caller needs `Sites.FullControl.All`, so this is an admin action. The easiest tool is Graph Explorer.

1. Open <https://developer.microsoft.com/graph/graph-explorer> and sign in with an admin account.
2. **Modify permissions**: consent to `Sites.FullControl.All`. If the dialog offers **Consent on behalf of your organization**, leave it unchecked; that consents for your account only.
3. Get the site ID:
   ```http
   GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/demo?$select=id
   ```
   The `id` looks like `contoso.sharepoint.com,<guid>,<guid>`.
4. Grant the role:
   ```http
   POST https://graph.microsoft.com/v1.0/sites/<site-id>/permissions
   Content-Type: application/json

   {
     "roles": ["read"],
     "grantedToIdentities": [
       { "application": { "id": "<client-id>", "displayName": "Entra Demo" } }
     ]
   }
   ```
   Expected: `201 Created`. Keep the returned `id`; you need it to revoke the grant.

Roles: `read`, `write` (needed for `POST /api/sharepoint/items`), `owner`, `fullcontrol`.

Alternatives: Microsoft Graph PowerShell (`New-MgSitePermission`), or PnP PowerShell (`Grant-PnPAzureADAppSitePermission`, which now needs its own app registration).

### Narrower: a single list

`Lists.SelectedOperations.Selected` (application permission) instead of `Sites.Selected`, granted with `POST /sites/{site-id}/lists/{list-id}/permissions`. Not tested here. Resolving Person columns needs read access to the site's hidden User Information List as well ([caveat 1](08-caveats.md#1-person-columns-contain-no-email)), so test it before relying on it.

## 3. What the grant does and doesn't restrict

Measured with a `read` grant on one site:

| Request | Result |
| --- | --- |
| Granted site, its lists and items | 200 |
| Other site by path, `/sites/root` | **200, name and URL only** |
| Other site's `/lists` | 200 with an empty list |
| Other site's `/drive` | 403 |
| `/sites?search=*`, `/sites/getAllSites` | 403 |

The contents of other sites stay hidden, but their existence, names and URLs are readable.

## Revoking

```http
DELETE https://graph.microsoft.com/v1.0/sites/<site-id>/permissions/<permission-id>
```

Remove the Graph Explorer consent under **Entra ID > Enterprise apps > Graph Explorer > Permissions** if you don't need it anymore.

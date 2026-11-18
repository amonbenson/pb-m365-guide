/// <reference path="../pb_data/types.d.ts" />

// Live tenant directory search (app-only, User.Read.All).
routerAdd(
    "GET",
    "/api/directory/users",
    (e) => {
        const graph = require(`${__hooks}/lib/graph.js`)
        const search = String(e.request.url.query().get("search") || "").replace(/'/g, "''")
        const filter = search ? `&$filter=${encodeURIComponent(`startswith(displayName,'${search}')`)}` : ""
        const result = graph.graphGet(`/users?$select=id,displayName,mail,userPrincipalName,jobTitle&$top=25${filter}`)
        return e.json(200, result.value)
    },
    $apis.requireAuth("users"),
)

// Live SharePoint list items, with Person columns resolved to email addresses.
routerAdd(
    "GET",
    "/api/sharepoint/items",
    (e) => {
        const sharepoint = require(`${__hooks}/lib/sharepoint.js`)
        return e.json(200, sharepoint.readItems())
    },
    $apis.requireAuth("users"),
)

// Creates a list item. Requires the "write" site role. body: { fields: { Title, Status, OwnerLookupId } }
routerAdd(
    "POST",
    "/api/sharepoint/items",
    (e) => {
        const sharepoint = require(`${__hooks}/lib/sharepoint.js`)
        return e.json(201, sharepoint.createItem(e.requestInfo().body.fields || {}))
    },
    $apis.requireAuth("users"),
)

// Refreshes the mirrored collections on demand.
routerAdd(
    "POST",
    "/api/mirror",
    (e) => {
        const mirror = require(`${__hooks}/lib/mirror.js`)
        return e.json(200, mirror.mirrorAll())
    },
    $apis.requireAuth("users"),
)

cronAdd("mirror-directory-and-sharepoint", "*/15 * * * *", () => {
    if (!$os.getenv("ENTRA_CLIENT_SECRET")) {
        return
    }
    const mirror = require(`${__hooks}/lib/mirror.js`)
    try {
        console.log("mirror", JSON.stringify(mirror.mirrorAll()))
    } catch (error) {
        console.log("mirror failed", error)
    }
})

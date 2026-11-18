// Live access to one SharePoint list through Microsoft Graph (app-only, Sites.Selected).

const graph = require(`${__hooks}/lib/graph.js`)

function siteAndList() {
    const settings = graph.config()
    const site = graph.graphGet(`/sites/${settings.sharePointSite}?$select=id,displayName,webUrl`)
    const list = graph.graphGet(
        `/sites/${site.id}/lists/${encodeURIComponent(settings.sharePointList)}?$select=id,displayName&$expand=columns`,
    )
    return { site, list }
}

// Built-in "Created By" / "Modified By" are read-only person columns; skip them.
function personColumnNames(list) {
    return (list.columns || [])
        .filter((column) => column.personOrGroup && !column.readOnly)
        .map((column) => column.name)
}

// The hidden User Information List has a localized title but always the internal name "users".
function userInformationListId(siteId) {
    const lists = graph.graphGetAll(`/sites/${siteId}/lists?$select=id,name,system`)
    const userList = lists.find((list) => list.name === "users")
    if (!userList) {
        throw new Error("User Information List not found")
    }
    return userList.id
}

// Person columns come back as <Column>LookupId, an id in the User Information List.
function siteUserEmails(siteId) {
    const entries = graph.graphGetAll(
        `/sites/${siteId}/lists/${userInformationListId(siteId)}/items?$expand=fields($select=Title,EMail,UserName)&$top=500`,
    )
    const emailsByLookupId = {}
    entries.forEach((entry) => {
        emailsByLookupId[entry.id] = String(entry.fields.EMail || entry.fields.UserName || "").toLowerCase()
    })
    return emailsByLookupId
}

function resolvePeople(item, personColumns, emailsByLookupId) {
    const people = {}
    personColumns.forEach((name) => {
        const lookupId = item.fields[name + "LookupId"]
        people[name] = lookupId ? { lookupId, email: emailsByLookupId[String(lookupId)] || "" } : null
    })
    return people
}

function readItems() {
    const { site, list } = siteAndList()
    const items = graph.graphGetAll(`/sites/${site.id}/lists/${list.id}/items?$expand=fields&$top=200`)
    const personColumns = personColumnNames(list)
    const emailsByLookupId = personColumns.length ? siteUserEmails(site.id) : {}
    const rows = items.map((item) => ({
        id: item.id,
        title: item.fields.Title || "",
        modified: item.lastModifiedDateTime,
        fields: item.fields,
        people: resolvePeople(item, personColumns, emailsByLookupId),
    }))
    return { site, list: { id: list.id, displayName: list.displayName, personColumns }, rows }
}

// Needs the "write" role on the site. Person columns are set as <Column>LookupId.
function createItem(fields) {
    const { site, list } = siteAndList()
    const result = graph.graphRequest("POST", `/sites/${site.id}/lists/${list.id}/items`, { body: { fields } })
    if (result.status >= 400) {
        throw new BadRequestError("SharePoint rejected the item", result.json)
    }
    return result.json
}

module.exports = { readItems, createItem }

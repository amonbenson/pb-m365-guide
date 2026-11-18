// Mirrors the tenant directory and one SharePoint list into PocketBase collections.

const graph = require(`${__hooks}/lib/graph.js`)
const sharepoint = require(`${__hooks}/lib/sharepoint.js`)

const USER_FIELDS = "id,displayName,mail,userPrincipalName,jobTitle,accountEnabled"

function upsertByFields(collectionName, filter, parameters, values) {
    const existing = $app.findRecordsByFilter(collectionName, filter, "", 1, 0, parameters)
    const record = existing.length ? existing[0] : new Record($app.findCollectionByNameOrId(collectionName))
    Object.entries(values).forEach(([key, value]) => record.set(key, value))
    $app.save(record)
    return record
}

function accountIdForOid(oid) {
    const accounts = $app.findRecordsByFilter("users", "entraOid = {:oid}", "", 1, 0, { oid })
    return accounts.length ? accounts[0].id : ""
}

function directoryIdForEmail(email) {
    if (!email) {
        return ""
    }
    const matches = $app.findRecordsByFilter("directory_users", "mail = {:email} || upn = {:email}", "", 1, 0, { email })
    return matches.length ? matches[0].id : ""
}

function mirrorDirectory() {
    const users = graph.graphGetAll(`/users?$select=${USER_FIELDS}&$top=999`)
    users.forEach((user) =>
        upsertByFields("directory_users", "oid = {:oid}", { oid: user.id }, {
            oid: user.id,
            displayName: user.displayName || "",
            mail: (user.mail || "").toLowerCase(),
            upn: (user.userPrincipalName || "").toLowerCase(),
            jobTitle: user.jobTitle || "",
            accountEnabled: !!user.accountEnabled,
            account: accountIdForOid(user.id),
        }),
    )
    return { users: users.length }
}

function mirrorSharePointList() {
    const live = sharepoint.readItems()
    const ownerColumn = graph.config().sharePointOwnerColumn || live.list.personColumns[0]
    live.rows.forEach((row) => {
        const owner = ownerColumn ? row.people[ownerColumn] : null
        const ownerEmail = owner ? owner.email : ""
        upsertByFields("sp_items", "listId = {:listId} && spItemId = {:itemId}", { listId: live.list.id, itemId: row.id }, {
            spItemId: row.id,
            siteId: live.site.id,
            listId: live.list.id,
            title: row.title,
            ownerEmail,
            owner: directoryIdForEmail(ownerEmail),
            modified: row.modified,
            fields: row.fields,
        })
    })
    return { items: live.rows.length, ownerColumn: ownerColumn || null }
}

function mirrorAll() {
    return { directory: mirrorDirectory(), sharePoint: mirrorSharePointList() }
}

module.exports = { mirrorAll, mirrorDirectory, mirrorSharePointList }

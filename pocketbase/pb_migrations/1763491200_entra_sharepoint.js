/// <reference path="../pb_data/types.d.ts" />

const SIGNED_IN = "@request.auth.id != ''"

function extendUsers(app) {
    const users = app.findCollectionByNameOrId("users")
    users.fields.add(new TextField({ name: "entraOid" }))
    users.fields.add(new TextField({ name: "upn" }))
    users.addIndex("idx_users_entra_oid", true, "entraOid", "entraOid != ''")
    app.save(users)
    return users
}

function createDirectoryUsers(app, users) {
    const collection = new Collection({
        type: "base",
        name: "directory_users",
        listRule: SIGNED_IN,
        viewRule: SIGNED_IN,
        fields: [
            { type: "text", name: "oid", required: true },
            { type: "text", name: "displayName" },
            { type: "text", name: "mail" },
            { type: "text", name: "upn" },
            { type: "text", name: "jobTitle" },
            { type: "bool", name: "accountEnabled" },
            { type: "relation", name: "account", collectionId: users.id, maxSelect: 1 },
        ],
        indexes: [
            "CREATE UNIQUE INDEX idx_directory_users_oid ON directory_users (oid)",
            "CREATE INDEX idx_directory_users_mail ON directory_users (mail)",
            "CREATE INDEX idx_directory_users_upn ON directory_users (upn)",
        ],
    })
    app.save(collection)
    return collection
}

function createSharePointItems(app, directoryUsers) {
    const collection = new Collection({
        type: "base",
        name: "sp_items",
        listRule: SIGNED_IN,
        viewRule: SIGNED_IN,
        fields: [
            { type: "text", name: "spItemId", required: true },
            { type: "text", name: "siteId" },
            { type: "text", name: "listId" },
            { type: "text", name: "title" },
            { type: "text", name: "ownerEmail" },
            { type: "relation", name: "owner", collectionId: directoryUsers.id, maxSelect: 1 },
            { type: "text", name: "modified" },
            { type: "json", name: "fields" },
        ],
        indexes: ["CREATE UNIQUE INDEX idx_sp_items_item ON sp_items (listId, spItemId)"],
    })
    app.save(collection)
    return collection
}

function createItemNotes(app, users, sharePointItems) {
    const isAuthor = "@request.auth.id != '' && author = @request.auth.id"
    app.save(
        new Collection({
            type: "base",
            name: "item_notes",
            listRule: SIGNED_IN,
            viewRule: SIGNED_IN,
            createRule: "@request.auth.id != '' && @request.body.author = @request.auth.id",
            updateRule: isAuthor,
            deleteRule: isAuthor,
            fields: [
                { type: "relation", name: "item", collectionId: sharePointItems.id, maxSelect: 1, required: true, cascadeDelete: true },
                { type: "relation", name: "author", collectionId: users.id, maxSelect: 1, required: true },
                { type: "text", name: "text", required: true },
                { type: "autodate", name: "created", onCreate: true },
            ],
        }),
    )
}

migrate(
    (app) => {
        const users = extendUsers(app)
        const directoryUsers = createDirectoryUsers(app, users)
        const sharePointItems = createSharePointItems(app, directoryUsers)
        createItemNotes(app, users, sharePointItems)
    },
    (app) => {
        ;["item_notes", "sp_items", "directory_users"].forEach((name) => app.delete(app.findCollectionByNameOrId(name)))
        const users = app.findCollectionByNameOrId("users")
        users.removeIndex("idx_users_entra_oid")
        users.fields.removeByName("entraOid")
        users.fields.removeByName("upn")
        app.save(users)
    },
)

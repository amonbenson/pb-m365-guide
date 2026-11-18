import PocketBase from "pocketbase"

export const pocketBase = new PocketBase(window.location.origin)

// Authentication

export async function signInWithEntraToken(accessToken, method = "jwks") {
    const response = await fetch("/api/entra-login", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ token: accessToken, method }),
    })
    const body = await response.json()
    if (!response.ok) {
        throw new Error(body.message)
    }
    pocketBase.authStore.save(body.token, body.record)
    return body.record
}

// Built-in PocketBase OAuth2 popup. Needs the realtime (SSE) connection to reach PocketBase.
export async function signInWithPocketBaseOAuth() {
    const result = await pocketBase.collection("users").authWithOAuth2({ provider: "microsoft" })
    return result.record
}

export function signOut() {
    pocketBase.authStore.clear()
}

export function currentUser() {
    return pocketBase.authStore.isValid ? pocketBase.authStore.record : null
}

// Tenant users

export function searchDirectoryLive(search) {
    return pocketBase.send("/api/directory/users", { query: { search } })
}

export function searchDirectoryMirrored(search) {
    return pocketBase.collection("directory_users").getList(1, 25, {
        filter: pocketBase.filter("displayName ~ {:search} || mail ~ {:search}", { search }),
        sort: "displayName",
        expand: "account",
    })
}

// SharePoint list

export function listItemsLive() {
    return pocketBase.send("/api/sharepoint/items", {})
}

export function createItemLive(fields) {
    return pocketBase.send("/api/sharepoint/items", { method: "POST", body: { fields } })
}

export function listItemsMirrored() {
    return pocketBase.collection("sp_items").getFullList({ sort: "title", expand: "owner,owner.account" })
}

export function itemsOwnedByCurrentUser() {
    return pocketBase.collection("sp_items").getFullList({
        filter: pocketBase.filter("owner.account = {:userId}", { userId: currentUser().id }),
    })
}

export function refreshMirror() {
    return pocketBase.send("/api/mirror", { method: "POST" })
}

// PocketBase-owned data attached to SharePoint items

export function notesForItem(itemId) {
    return pocketBase.collection("item_notes").getFullList({
        filter: pocketBase.filter("item = {:itemId}", { itemId }),
        sort: "created",
        expand: "author",
    })
}

export function addNote(itemId, text) {
    return pocketBase.collection("item_notes").create({ item: itemId, author: currentUser().id, text })
}

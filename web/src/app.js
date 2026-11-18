import * as auth from "./auth.js"
import * as data from "./data.js"

const statusElement = document.getElementById("status")
const outputElement = document.getElementById("output")
const hostElement = document.getElementById("host")

function show(label, value) {
    outputElement.textContent = `${label}\n\n${typeof value === "string" ? value : JSON.stringify(value, null, 2)}`
}

function showError(label, error) {
    show(`${label} failed`, error?.response || error?.message || String(error))
}

function updateStatus() {
    const user = data.currentUser()
    statusElement.textContent = user ? `Signed in as ${user.name || user.email}` : "Not signed in"
}

async function completeSignIn(tokenResult) {
    await data.signInWithEntraToken(tokenResult.accessToken)
    updateStatus()
}

async function trySilentSignIn() {
    try {
        await completeSignIn(await auth.acquireTokenSilently())
    } catch (error) {
        show("Silent sign-in not possible, use a sign-in button", error.message)
    }
}

function selectedValue(id) {
    return document.getElementById(id).value.trim()
}

const actions = {
    "sign-in-popup": () =>
        auth
            .acquireTokenWithPopup()
            .then(completeSignIn)
            .catch((error) => showError("Popup sign-in", error)),
    "sign-in-redirect": () => auth.acquireTokenWithRedirect(),
    "sign-in-pocketbase": () => data.signInWithPocketBaseOAuth().then(updateStatus),
    "sign-out": async () => {
        data.signOut()
        await auth.forgetMicrosoftAccount()
        updateStatus()
    },
    "directory-live": () => data.searchDirectoryLive(selectedValue("search")),
    "directory-mirrored": () => data.searchDirectoryMirrored(selectedValue("search")),
    "items-live": () => data.listItemsLive(),
    "items-mirrored": () => data.listItemsMirrored(),
    "items-mine": () => data.itemsOwnedByCurrentUser(),
    "mirror-refresh": () => data.refreshMirror(),
    "notes-list": () => data.notesForItem(selectedValue("item-id")),
    "notes-add": () => data.addNote(selectedValue("item-id"), selectedValue("note-text")),
}

async function runAction(name) {
    try {
        const result = await actions[name]()
        if (result !== undefined) {
            show(name, result)
        }
    } catch (error) {
        showError(name, error)
    }
}

document.querySelectorAll("button[data-action]").forEach((button) =>
    button.addEventListener("click", () => runAction(button.dataset.action)),
)

async function start() {
    updateStatus()
    const context = await auth.detectTeams()
    hostElement.textContent = context
        ? `Running in ${context.app.host.name} (${context.app.host.clientType}), ${context.channel ? context.channel.membershipType + " channel" : "personal"} tab`
        : "Running in a browser"
    const redirectResult = await auth.initializeMsal()
    if (redirectResult) {
        await completeSignIn(redirectResult)
        return
    }
    if (!data.currentUser()) {
        await trySilentSignIn()
    }
}

start().catch((error) => showError("Startup", error))

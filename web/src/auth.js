import { createNestablePublicClientApplication } from "@azure/msal-browser"
import { app as teamsApp } from "@microsoft/teams-js"

const TEAMS_INITIALIZE_TIMEOUT_MILLISECONDS = 3000

let configuration
let msal
let teamsContext = null

export async function detectTeams() {
    const timeout = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("not running inside Teams")), TEAMS_INITIALIZE_TIMEOUT_MILLISECONDS),
    )
    try {
        await Promise.race([teamsApp.initialize(), timeout])
        teamsContext = await teamsApp.getContext()
        teamsApp.notifySuccess()
    } catch (_) {
        teamsContext = null
    }
    return teamsContext
}

// Inside Teams/Outlook/M365 this uses Nested App Authentication; elsewhere it falls back to a normal MSAL client.
export async function initializeMsal() {
    configuration = await (await fetch("/api/app-config")).json()
    msal = await createNestablePublicClientApplication({
        auth: {
            clientId: configuration.clientId,
            authority: `https://login.microsoftonline.com/${configuration.tenantId}`,
            redirectUri: window.location.origin + "/redirect.html",
        },
        cache: { cacheLocation: "localStorage" },
    })
    const redirectResult = await msal.handleRedirectPromise()
    if (redirectResult) {
        msal.setActiveAccount(redirectResult.account)
    }
    return redirectResult
}

function tokenRequest() {
    const request = { scopes: [configuration.apiScope] }
    const loginHint = teamsContext?.user?.loginHint
    return loginHint ? { ...request, loginHint } : request
}

export async function acquireTokenSilently() {
    const account = msal.getActiveAccount() || msal.getAllAccounts()[0]
    const result = account
        ? await msal.acquireTokenSilent({ ...tokenRequest(), account })
        : await msal.ssoSilent(tokenRequest())
    msal.setActiveAccount(result.account)
    return result
}

// Call directly from a click handler (no await before it), otherwise Safari blocks the popup.
export function acquireTokenWithPopup() {
    return msal.acquireTokenPopup(tokenRequest()).then((result) => {
        msal.setActiveAccount(result.account)
        return result
    })
}

export function acquireTokenWithRedirect() {
    return msal.acquireTokenRedirect(tokenRequest())
}

export async function forgetMicrosoftAccount() {
    await msal.clearCache()
}

/// <reference path="../pb_data/types.d.ts" />

// Configures PocketBase's built-in Microsoft OAuth2 provider from the environment,
// restricted to one tenant by using tenant-specific endpoints instead of /common.
onBootstrap((e) => {
    e.next()
    const graph = require(`${__hooks}/lib/graph.js`)
    const settings = graph.config()
    if (!settings.tenantId || !settings.clientId) {
        console.log("ENTRA_TENANT_ID / ENTRA_CLIENT_ID not set, Microsoft OAuth2 stays unconfigured")
        return
    }
    const users = $app.findCollectionByNameOrId("users")
    const tenantEndpoint = `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0`
    unmarshal(
        {
            oauth2: {
                enabled: true,
                providers: [
                    {
                        name: "microsoft",
                        clientId: settings.clientId,
                        clientSecret: settings.clientSecret,
                        authURL: tenantEndpoint + "/authorize",
                        tokenURL: tenantEndpoint + "/token",
                    },
                ],
            },
        },
        users,
    )
    $app.save(users)
})

// Built-in OAuth2 sign-in: store the Entra object id so both login paths resolve to the same user.
onRecordAuthWithOAuth2Request((e) => {
    e.next()
    if (e.providerName !== "microsoft") {
        return
    }
    const entra = require(`${__hooks}/lib/entra.js`)
    e.record.set("entraOid", e.oauth2User.id)
    e.record.set("upn", String(e.oauth2User.rawUser.userPrincipalName || "").toLowerCase())
    $app.save(e.record)
    entra.linkDirectoryEntry(e.record)
}, "users")

// Public settings the frontend needs to create its MSAL client.
routerAdd("GET", "/api/app-config", (e) => {
    const clientId = $os.getenv("ENTRA_CLIENT_ID")
    return e.json(200, {
        tenantId: $os.getenv("ENTRA_TENANT_ID"),
        clientId,
        apiScope: `api://${clientId}/access_as_user`,
    })
})

// Exchanges an Entra access token (MSAL / Teams NAA) for a normal PocketBase session.
// body: { token, method: "jwks" (default, local signature check) | "obo" (on-behalf-of exchange) }
routerAdd("POST", "/api/entra-login", (e) => {
    const entra = require(`${__hooks}/lib/entra.js`)
    const body = e.requestInfo().body
    const method = body.method === "obo" ? "obo" : "jwks"
    let identity
    try {
        identity = method === "obo" ? entra.identityFromOnBehalfOf(body.token) : entra.identityFromJwks(body.token)
    } catch (error) {
        throw new UnauthorizedError(`Entra token rejected (${method}): ${error}`)
    }
    const record = entra.findOrCreateUser(identity)
    return $apis.recordAuthResponse(e, record, "entra-" + method, { method })
})

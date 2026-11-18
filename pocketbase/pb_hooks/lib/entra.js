// Turns an Entra access token (from MSAL / Teams NAA) into a PocketBase user.

const jwt = require(`${__hooks}/lib/jwt.js`)
const graph = require(`${__hooks}/lib/graph.js`)

const KEYS_CACHE_KEY = "entraSigningKeys"
const KEYS_MAX_AGE_SECONDS = 3600
const REQUIRED_SCOPE = "access_as_user"

function fetchSigningKeys(tenantId) {
    const response = $http.send({
        url: `https://login.microsoftonline.com/${tenantId}/discovery/v2.0/keys`,
        method: "GET",
        timeout: 20,
    })
    if (response.statusCode !== 200) {
        throw new Error("could not load signing keys")
    }
    const keys = response.json.keys
    $app.store().set(KEYS_CACHE_KEY, { keys, fetchedAt: Math.floor(Date.now() / 1000) })
    return keys
}

// Cached for an hour; refetched early when a token names an unknown key (key rotation).
function signingKeys(tenantId, requiredKeyId) {
    const cached = $app.store().get(KEYS_CACHE_KEY)
    const nowSeconds = Math.floor(Date.now() / 1000)
    const isFresh = cached && cached.fetchedAt + KEYS_MAX_AGE_SECONDS > nowSeconds
    const hasKey = cached && cached.keys.some((key) => key.kid === requiredKeyId)
    return isFresh && hasKey ? cached.keys : fetchSigningKeys(tenantId)
}

function expectedClaims(settings) {
    return {
        tenantId: settings.tenantId,
        audiences: [settings.clientId, "api://" + settings.clientId],
        issuers: [
            `https://login.microsoftonline.com/${settings.tenantId}/v2.0`,
            `https://sts.windows.net/${settings.tenantId}/`,
        ],
    }
}

function identityFromJwks(token) {
    const settings = graph.config()
    const keyId = jwt.decode(token).header.kid
    const claims = jwt.verifyRs256(token, signingKeys(settings.tenantId, keyId), expectedClaims(settings))
    if (!String(claims.scp || "").split(" ").includes(REQUIRED_SCOPE)) {
        throw new Error(`token lacks the ${REQUIRED_SCOPE} scope`)
    }
    const userName = claims.preferred_username || claims.upn || ""
    return { oid: claims.oid, name: claims.name || "", email: userName, upn: userName }
}

function identityFromOnBehalfOf(token) {
    const exchanged = graph.requestToken({
        grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
        requested_token_use: "on_behalf_of",
        assertion: token,
        scope: "https://graph.microsoft.com/User.Read",
    })
    const me = graph.graphGet("/me?$select=id,displayName,mail,userPrincipalName", { accessToken: exchanged.access_token })
    return { oid: me.id, name: me.displayName || "", email: me.mail || me.userPrincipalName || "", upn: me.userPrincipalName || "" }
}

function linkDirectoryEntry(userRecord) {
    const entries = $app.findRecordsByFilter("directory_users", "oid = {:oid}", "", 1, 0, { oid: userRecord.get("entraOid") })
    if (!entries.length) {
        return
    }
    entries[0].set("account", userRecord.id)
    $app.save(entries[0])
}

function newUserRecord(oid) {
    const record = new Record($app.findCollectionByNameOrId("users"))
    record.set("entraOid", oid)
    record.set("verified", true)
    record.setRandomPassword()
    return record
}

function findOrCreateUser(identity) {
    const existing = $app.findRecordsByFilter("users", "entraOid = {:oid}", "", 1, 0, { oid: identity.oid })
    const record = existing.length ? existing[0] : newUserRecord(identity.oid)
    record.set("name", identity.name)
    record.set("upn", identity.upn.toLowerCase())
    if (identity.email) {
        record.set("email", identity.email.toLowerCase())
    }
    $app.save(record)
    linkDirectoryEntry(record)
    return record
}

module.exports = { identityFromJwks, identityFromOnBehalfOf, findOrCreateUser, linkDirectoryEntry }

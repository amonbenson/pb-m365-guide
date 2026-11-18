// App-only Microsoft Graph access (client credentials flow).

const GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const TOKEN_CACHE_KEY = "graphAppToken"

function config() {
    return {
        tenantId: $os.getenv("ENTRA_TENANT_ID"),
        clientId: $os.getenv("ENTRA_CLIENT_ID"),
        clientSecret: $os.getenv("ENTRA_CLIENT_SECRET"),
        sharePointSite: $os.getenv("SHAREPOINT_SITE"),
        sharePointList: $os.getenv("SHAREPOINT_LIST"),
        sharePointOwnerColumn: $os.getenv("SHAREPOINT_OWNER_COLUMN"),
    }
}

function formEncode(fields) {
    return Object.entries(fields)
        .map(([key, value]) => encodeURIComponent(key) + "=" + encodeURIComponent(value))
        .join("&")
}

function requestToken(fields) {
    const settings = config()
    const response = $http.send({
        url: `https://login.microsoftonline.com/${settings.tenantId}/oauth2/v2.0/token`,
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: formEncode(Object.assign({ client_id: settings.clientId, client_secret: settings.clientSecret }, fields)),
        timeout: 20,
    })
    if (response.statusCode !== 200) {
        throw new Error("token request failed: " + toString(response.body))
    }
    return response.json
}

function appToken() {
    const cached = $app.store().get(TOKEN_CACHE_KEY)
    const nowSeconds = Math.floor(Date.now() / 1000)
    if (cached && cached.expiresAt - 60 > nowSeconds) {
        return cached.accessToken
    }
    const result = requestToken({ grant_type: "client_credentials", scope: "https://graph.microsoft.com/.default" })
    $app.store().set(TOKEN_CACHE_KEY, { accessToken: result.access_token, expiresAt: nowSeconds + result.expires_in })
    return result.access_token
}

function graphRequest(method, path, options) {
    const settings = options || {}
    const url = path.startsWith("https://") ? path : GRAPH_BASE + path
    const response = $http.send({
        url,
        method,
        headers: Object.assign(
            { Authorization: "Bearer " + (settings.accessToken || appToken()), "Content-Type": "application/json" },
            settings.headers || {},
        ),
        body: settings.body ? JSON.stringify(settings.body) : "",
        timeout: 30,
    })
    return { status: response.statusCode, json: response.json }
}

function graphGet(path, options) {
    const result = graphRequest("GET", path, options)
    if (result.status >= 400) {
        const error = new Error(`Graph GET ${path} -> ${result.status}: ${JSON.stringify(result.json)}`)
        error.status = result.status
        throw error
    }
    return result.json
}

function graphGetAll(path, options) {
    const items = []
    let next = path
    while (next) {
        const page = graphGet(next, options)
        items.push(...(page.value || []))
        next = page["@odata.nextLink"]
    }
    return items
}

module.exports = { config, formEncode, requestToken, appToken, graphRequest, graphGet, graphGetAll }

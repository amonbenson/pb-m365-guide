import { app, pages } from "@microsoft/teams-js"

await app.initialize()
pages.config.registerOnSaveHandler((saveEvent) => {
    pages.config
        .setConfig({
            contentUrl: window.location.origin + "/index.html",
            websiteUrl: window.location.origin + "/index.html",
            entityId: "entra-demo-channel-tab",
            suggestedDisplayName: "Entra Demo",
        })
        .then(() => saveEvent.notifySuccess())
        .catch((error) => saveEvent.notifyFailure(String(error)))
})
pages.config.setValidityState(true)

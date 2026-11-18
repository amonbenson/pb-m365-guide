import { broadcastResponseToMainFrame } from "@azure/msal-browser/redirect-bridge"

broadcastResponseToMainFrame().catch((error) => {
    document.body.textContent = "Sign-in bridge error: " + error
})

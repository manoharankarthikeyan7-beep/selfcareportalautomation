export const msalConfig = {
    auth: {
        clientId: "ef8c8368-c7bf-4a2e-b204-070aa4100256", // From Phase 2 SPA registration
        authority: "https://login.microsoftonline.com/YOUR_TENANT_ID",
        redirectUri: "http://localhost:3000", // Or your production URL
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

// This scope must match the one you "Exposed" in the API registration
export const loginRequest = {
    scopes: ["api://2c51c622-567f-41cc-b46c-1a1ace37c0ed/Pipeline.Access"]
};
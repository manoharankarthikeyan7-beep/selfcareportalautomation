import { LogLevel } from "@azure/msal-browser";

export const msalConfig = {
    auth: {
        clientId: "ef8c8368-c7bf-4a2e-b204-070aa4100256", 
        authority: "https://login.microsoftonline.com/251c8343-663c-4ab5-996e-8bf8e88aca58",
        redirectUri: window.location.origin,
        // SECURITY FIX: Ensures user returns to the correct app page after login,
        // preventing unauthorized redirects.
        navigateToLoginRequestUrl: true, 
        // SECURITY FIX: Validates the Microsoft login endpoint to prevent spoofing.
        validateAuthority: true, 
    },
    cache: {
        // SECURITY FIX: sessionStorage is safer than localStorage as it 
        // wipes tokens when the tab is closed.
        cacheLocation: "sessionStorage", 
        storeAuthStateInCookie: false,
    },
    system: {
        // SECURITY FIX: Configures logging to hide sensitive PII (Personal Identifiable Info)
        // and only show critical errors in the browser console.
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) return; // Never log PII
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    default:
                        return;
                }
            },
            piiLoggingEnabled: false,
            logLevel: LogLevel.Error,
        }
    }
};

export const loginRequest = {
    // SECURITY FIX: Using custom scopes ensures "Principle of Least Privilege".
    // The token generated can ONLY be used for your specific Pipeline API.
    scopes: ["api://2c51c622-567f-41cc-b46c-1a1ace37c0ed/Pipeline.Access"]
};
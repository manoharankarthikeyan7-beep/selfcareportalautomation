import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { MsalProvider } from "@azure/msal-react";
import { msalConfig } from "./authConfig";

/**
 * SECURITY FIX: Initialize MSAL outside the render cycle.
 * This ensures the authentication context is stable and does not re-initialize 
 * during React re-renders, which could leak memory or invalidate sessions.
 */
const msalInstance = new PublicClientApplication(msalConfig);

// SECURITY FIX: Handle the redirect response immediately on page load.
// This clears the sensitive authentication code/token from the URL bar 
// after a successful login, preventing "Token Leakage" in browser history.
msalInstance.initialize().then(() => {
    
    // Optional: Account selection logic
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        msalInstance.setActiveAccount(accounts[0]);
    }

    // SECURITY FIX: Listen for login events to keep the app state synchronized.
    msalInstance.addEventCallback((event) => {
        if (event.eventType === EventType.LOGIN_SUCCESS && event.payload.account) {
            const account = event.payload.account;
            msalInstance.setActiveAccount(account);
        }
    });

    const root = ReactDOM.createRoot(document.getElementById('root'));
    root.render(
        <React.StrictMode>
            {/* SECURITY FIX: MsalProvider wraps the app to provide a secure 
              context for all child components (App.js). 
            */}
            <MsalProvider instance={msalInstance}>
                <App />
            </MsalProvider>
        </React.StrictMode>
    );
}).catch(error => {
    // SECURITY FIX: Graceful failure if the auth provider is blocked or misconfigured.
    console.error("MSAL Initialization Failed. Check ClientID or Network.", error);
    document.getElementById('root').innerHTML = `
        <div style="padding: 20px; color: #d13438; font-family: 'Segoe UI';">
            <h2>Authentication Error</h2>
            <p>Unable to connect to the Identity Provider. Please refresh or contact support.</p>
        </div>
    `;
});
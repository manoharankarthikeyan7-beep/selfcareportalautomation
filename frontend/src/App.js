import React from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

function App() {
    const { instance } = useMsal();

    const handleLogin = () => {
        instance.loginPopup(loginRequest).catch(e => console.error(e));
    };

    return (
        <div style={{ padding: "20px" }}>
            <h1>Azure DevOps Pipeline Generator</h1>
            
            <UnauthenticatedTemplate>
                <p>Please sign in to manage pipelines.</p>
                <button onClick={handleLogin}>Login with Azure AD</button>
            </UnauthenticatedTemplate>

            <AuthenticatedTemplate>
                <PipelineDashboard />
            </AuthenticatedTemplate>
        </div>
    );
}

// Internal component to handle logic once logged in
const PipelineDashboard = () => {
    const { accounts } = useMsal();
    const userGroups = accounts[0]?.idTokenClaims?.groups || [];

    // Mapping our Role Object IDs for UI Hiding
    const isAdmin = userGroups.includes("YOUR_ADMIN_GROUP_ID");
    const isDevOps = userGroups.includes("YOUR_DEVOPS_GROUP_ID");

    return (
        <div>
            <h2>Welcome, {accounts[0].name}</h2>
            
            {/* Show buttons only to Admins or DevOps Engineers */}
            {(isAdmin || isDevOps) && (
                <button style={{ background: "green", color: "white" }}>
                    + Create New Pipeline
                </button>
            )}

            {/* Show Delete button ONLY to Admins */}
            {isAdmin && (
                <button style={{ background: "red", color: "white", marginLeft: "10px" }}>
                    Delete Pipeline
                </button>
            )}

            <p>Role Detected: {isAdmin ? "Admin" : isDevOps ? "DevOps" : "Viewer"}</p>
        </div>
    );
};

export default App;
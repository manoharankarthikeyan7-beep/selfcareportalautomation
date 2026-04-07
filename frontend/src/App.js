import React, { useState } from "react"; // Added useState
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

function App() {
    const { instance } = useMsal();

    const handleLogin = () => {
        instance.loginPopup(loginRequest).catch(e => console.error(e));
    };

    return (
        <div style={{ padding: "20px", fontFamily: "sans-serif" }}>
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

const PipelineDashboard = () => {
    const { accounts, instance } = useMsal();
    const [pipelineName, setPipelineName] = useState("");
    const [repoId, setRepoId] = useState("");
    const [status, setStatus] = useState("");

    const userGroups = accounts[0]?.idTokenClaims?.groups || [];
    const isAdmin = userGroups.includes("YOUR_ADMIN_GROUP_ID");
    const isDevOps = userGroups.includes("YOUR_DEVOPS_GROUP_ID");

    const handleCreate = async () => {
        if (!pipelineName || !repoId) {
            alert("Please provide both a Pipeline Name and Repo ID");
            return;
        }

        setStatus("Creating...");

        try {
            // 1. Get the Access Token silently
            const authResult = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });

            // 2. Call your Backend API
            const response = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${authResult.accessToken}` // Sending the token!
                },
                body: JSON.stringify({
                    pipelineName: pipelineName,
                    repoId: repoId
                })
            });

            const result = await response.json();
            
            if (response.ok) {
                setStatus(`✅ Success! Created ID: ${result.details.id}`);
                setPipelineName(""); // Clear form
            } else {
                setStatus(`❌ Error: ${result.error}`);
            }
        } catch (error) {
            console.error(error);
            setStatus("❌ Authentication or Network error.");
        }
    };

    return (
        <div>
            <h2>Welcome, {accounts[0].name}</h2>
            <p>Role Detected: <strong>{isAdmin ? "Admin" : isDevOps ? "DevOps" : "Viewer"}</strong></p>
            <hr />

            {(isAdmin || isDevOps) ? (
                <div style={{ marginTop: "20px", background: "#f4f4f4", padding: "15px", borderRadius: "8px" }}>
                    <h3>Create New Pipeline</h3>
                    <div style={{ marginBottom: "10px" }}>
                        <input 
                            type="text" 
                            placeholder="Pipeline Name (e.g. My-New-App)" 
                            value={pipelineName}
                            onChange={(e) => setPipelineName(e.target.value)}
                            style={{ padding: "8px", width: "250px", marginRight: "10px" }}
                        />
                        <input 
                            type="text" 
                            placeholder="Repo GUID" 
                            value={repoId}
                            onChange={(e) => setRepoId(e.target.value)}
                            style={{ padding: "8px", width: "250px" }}
                        />
                    </div>
                    <button 
                        onClick={handleCreate}
                        style={{ background: "green", color: "white", padding: "10px 20px", border: "none", cursor: "pointer" }}
                    >
                        Trigger Creation
                    </button>
                    {status && <p>{status}</p>}
                </div>
            ) : (
                <p style={{ color: "red" }}>You do not have permission to create pipelines.</p>
            )}
        </div>
    );
};

export default App;
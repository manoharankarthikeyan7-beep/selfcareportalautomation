import React, { useState, useEffect } from "react"; // Added useEffect
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
    const [selectedRepoId, setSelectedRepoId] = useState(""); // Holds the selected GUID
    const [repos, setRepos] = useState([]); // List of repos from ADO
    const [status, setStatus] = useState("");

    const userGroups = accounts[0]?.idTokenClaims?.groups || [];
    const isAdmin = userGroups.includes("YOUR_ADMIN_GROUP_ID");
    const isDevOps = userGroups.includes("YOUR_DEVOPS_GROUP_ID");

    // --- Fetch Repositories on Load ---
    useEffect(() => {
        const fetchRepos = async () => {
            try {
                const authResult = await instance.acquireTokenSilent({
                    ...loginRequest,
                    account: accounts[0]
                });

                const response = await fetch("/api/repos", {
                    headers: { "Authorization": `Bearer ${authResult.accessToken}` }
                });
                const data = await response.json();
                if (Array.isArray(data)) setRepos(data);
            } catch (err) {
                console.error("Failed to load repositories", err);
            }
        };
        fetchRepos();
    }, [instance, accounts]);

    const handleCreate = async () => {
        if (!pipelineName || !selectedRepoId) {
            alert("Please provide a Name and select a Repository");
            return;
        }

        setStatus("Creating...");

        try {
            const authResult = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });

            const response = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${authResult.accessToken}`
                },
                body: JSON.stringify({
                    pipelineName: pipelineName,
                    repoId: selectedRepoId // GUID from the dropdown
                })
            });

            const result = await response.json();
            if (response.ok) {
                setStatus(`✅ Success! Created ID: ${result.details.id}`);
                setPipelineName("");
            } else {
                setStatus(`❌ Error: ${result.error}`);
            }
        } catch (error) {
            setStatus("❌ Request error.");
        }
    };

    return (
        <div>
            <h2>Welcome, {accounts[0].name}</h2>
            <p>Role Detected: <strong>{isAdmin ? "Admin" : isDevOps ? "DevOps" : "Viewer"}</strong></p>
            <hr />

            {(isAdmin || isDevOps) ? (
                <div style={{ marginTop: "20px", background: "#f4f4f4", padding: "20px", borderRadius: "8px" }}>
                    <h3>Create New Pipeline</h3>
                    
                    <div style={{ display: "flex", flexDirection: "column", gap: "15px", maxWidth: "400px" }}>
                        <label><strong>1. Select Repository:</strong></label>
                        <select 
                            value={selectedRepoId} 
                            onChange={(e) => setSelectedRepoId(e.target.value)}
                            style={{ padding: "10px" }}
                        >
                            <option value="">-- Select a Repo --</option>
                            {repos.map(repo => (
                                <option key={repo.id} value={repo.id}>{repo.name}</option>
                            ))}
                        </select>

                        <label><strong>2. Pipeline Name:</strong></label>
                        <input 
                            type="text" 
                            placeholder="e.g. Finance-App-CI" 
                            value={pipelineName}
                            onChange={(e) => setPipelineName(e.target.value)}
                            style={{ padding: "10px" }}
                        />

                        <button 
                            onClick={handleCreate}
                            style={{ background: "green", color: "white", padding: "12px", border: "none", cursor: "pointer", fontWeight: "bold" }}
                        >
                            Trigger Creation
                        </button>
                    </div>
                    {status && <p style={{ marginTop: "15px" }}>{status}</p>}
                </div>
            ) : (
                <p style={{ color: "red" }}>You do not have permission to create pipelines.</p>
            )}
        </div>
    );
};

export default App;
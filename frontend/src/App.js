import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

// Use relative paths since Backend and Frontend are unified
const BACKEND_URL = ""; 

const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [repos, setRepos] = useState([]);
    const [status, setStatus] = useState("");
    const [sourceType, setSourceType] = useState("azure");

    useEffect(() => {
        const fetchRepos = async () => {
            setStatus("Fetching repositories...");
            try {
                const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
                const endpoint = sourceType === "azure" ? "/api/repos" : "/api/github/repos";
                
                const res = await fetch(`${BACKEND_URL}${endpoint}`, {
                    headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
                });

                // Check if response is HTML (Error) or JSON (Success)
                const contentType = res.headers.get("content-type");
                if (!contentType || !contentType.includes("application/json")) {
                    throw new Error("Server returned HTML instead of JSON. Check backend routing.");
                }

                const data = await res.json();
                setRepos(data);
                setStatus("");
            } catch (err) {
                console.error("Fetch Error:", err);
                setStatus("Error: " + err.message);
            }
        };
        if (accounts.length > 0) fetchRepos();
    }, [instance, accounts, sourceType]);

    return (
        <div style={{ padding: "20px" }}>
            <h3>Source: {sourceType.toUpperCase()}</h3>
            <button onClick={() => setSourceType(sourceType === "azure" ? "github" : "azure")}>
                Switch to {sourceType === "azure" ? "GitHub" : "Azure"}
            </button>
            <p>{status}</p>
            <ul>
                {repos.map(r => <li key={r.id}>{r.name}</li>)}
            </ul>
        </div>
    );
};

function App() {
    const { instance } = useMsal();
    return (
        <div style={{ fontFamily: "Segoe UI", padding: "40px" }}>
            <h1>Pipeline Generator</h1>
            <AuthenticatedTemplate><PipelineWizard /></AuthenticatedTemplate>
            <UnauthenticatedTemplate>
                <button onClick={() => instance.loginRedirect(loginRequest)}>Login</button>
            </UnauthenticatedTemplate>
        </div>
    );
}

export default App;
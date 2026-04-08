import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [step, setStep] = useState(1);
    const [status, setStatus] = useState("");
    const [repos, setRepos] = useState([]);
    const [branches, setBranches] = useState([]);
    
    // NEW: Search state for Step 2 Dynamic Discovery
    const [searchTerm, setSearchTerm] = useState("");

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '/azure-pipelines.yml', name: '' 
    });

    useEffect(() => {
        const fetchRepos = async () => {
            try {
                const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
                const res = await fetch("/api/repos", {
                    headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
                });
                const data = await res.json();
                setRepos(data || []);
            } catch (err) { console.error("Repo fetch failed", err); }
        };
        if (accounts.length > 0) fetchRepos();
    }, [instance, accounts]);

    const handleRepoSelect = async (repo) => {
        setFormData({ ...formData, repoId: repo.id, repoName: repo.name });
        setStatus("Fetching branches...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${repo.id}/branches`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            setBranches(data || []);
            setStatus("");
            setStep(2);
        } catch (err) { setStatus("Error loading branches."); }
    };

    // (Keep your existing finalCreateCall function)
    const finalCreateCall = async () => { /* ... existing logic ... */ };

    return (
        <div style={{ background: "#f4f4f4", padding: "20px", borderRadius: "8px", marginTop: "20px" }}>
            {status && <p style={{ color: "blue" }}>{status}</p>}

            {step === 1 && (
                <div>
                    <h3>1. Select Repository</h3>
                    {/* NEW: Search Bar for Discovery */}
                    <input 
                        type="text"
                        placeholder="Type to search repositories..."
                        style={{ width: "100%", padding: "12px", marginBottom: "15px", boxSizing: "border-box", border: "1px solid #ccc", borderRadius: "4px" }}
                        onChange={(e) => setSearchTerm(e.target.value.toLowerCase())}
                    />
                    <div style={{ maxHeight: "300px", overflowY: "auto", border: "1px solid #ddd", background: "#fff" }}>
                        {repos
                            .filter(r => r.name.toLowerCase().includes(searchTerm))
                            .map(r => (
                                <button key={r.id} onClick={() => handleRepoSelect(r)} style={{ display: "block", width: "100%", padding: "10px", textAlign: "left", border: "none", borderBottom: "1px solid #eee", cursor: "pointer", background: "white" }}>
                                    {r.name}
                                </button>
                            ))
                        }
                    </div>
                </div>
            )}

            {/* Steps 2 and 3 remain the same as your working version */}
            {/* ... [Existing Step 2 & 3 UI] ... */}
        </div>
    );
};

// ... [Existing App function and export default App] ...
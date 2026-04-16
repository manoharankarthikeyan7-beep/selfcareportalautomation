import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

const BACKEND_URL = ""; 

const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [step, setStep] = useState(1);
    const [status, setStatus] = useState("");
    const [repos, setRepos] = useState([]);
    const [branches, setBranches] = useState([]);
    const [yamlFiles, setYamlFiles] = useState([]);
    const [searchTerm, setSearchTerm] = useState("");
    const [sourceType, setSourceType] = useState("azure"); 
    const [formData, setFormData] = useState({ repoId: '', repoName: '', branch: '', yamlPath: '', name: '' });

    const styles = {
        card: { background: "#fff", padding: "30px", borderRadius: "8px", border: "1px solid #ddd", marginTop: "20px" },
        input: { width: "100%", padding: "10px", marginBottom: "15px", border: "1px solid #ccc", borderRadius: "4px" },
        tabContainer: { display: "flex", marginBottom: "20px", borderBottom: "1px solid #ddd" },
        tab: (active) => ({
            padding: "10px 20px", cursor: "pointer",
            borderBottom: active ? "3px solid #0078d4" : "3px solid transparent",
            color: active ? "#0078d4" : "#666", fontWeight: active ? "600" : "400"
        }),
        repoList: { border: "1px solid #eee", borderRadius: "4px", maxHeight: "300px", overflowY: "auto" },
        repoItem: { width: "100%", padding: "12px", textAlign: "left", background: "none", border: "none", borderBottom: "1px solid #f5f5f5", cursor: "pointer", display: "flex", alignItems: "center" }
    };

    // Fetch Repositories
    useEffect(() => {
        const fetchRepos = async () => {
            setStatus("Fetching repositories...");
            try {
                const tokenRes = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
                const endpoint = sourceType === "azure" ? "/api/repos" : "/api/github/repos";
                const res = await fetch(`${BACKEND_URL}${endpoint}`, {
                    headers: { "Authorization": `Bearer ${tokenRes.accessToken}` }
                });
                const data = await res.json();
                setRepos(Array.isArray(data) ? data : []);
                setStatus("");
            } catch (err) { setStatus("Error loading repositories."); }
        };
        fetchRepos();
    }, [sourceType, instance, accounts]);

    // Handle Repo Selection & Load Branches (The Fix)
    const handleRepoSelect = async (repo) => {
        setFormData({ ...formData, repoId: repo.id, repoName: repo.name });
        setStatus(`Loading ${sourceType} branches...`);
        try {
            const tokenRes = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const encodedId = encodeURIComponent(repo.id);
            const endpoint = sourceType === "azure" ? `/api/repos/${encodedId}/branches` : `/api/github/repos/${encodedId}/branches`;
            
            const res = await fetch(`${BACKEND_URL}${endpoint}`, {
                headers: { "Authorization": `Bearer ${tokenRes.accessToken}` }
            });
            const data = await res.json();
            
            if (res.ok) {
                setBranches(data);
                setStep(2);
                setStatus("");
            } else {
                throw new Error();
            }
        } catch (err) {
            setStatus("Error loading branches.");
        }
    };

    const handleBranchChange = async (branchName) => {
        setFormData({ ...formData, branch: branchName });
        try {
            const tokenRes = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const encodedId = encodeURIComponent(formData.repoId);
            const endpoint = sourceType === "azure" ? `/api/repos/${encodedId}/yaml-files` : `/api/github/repos/${encodedId}/yaml-files`;
            const res = await fetch(`${BACKEND_URL}${endpoint}?branch=${branchName}`, {
                headers: { "Authorization": `Bearer ${tokenRes.accessToken}` }
            });
            const data = await res.json();
            setYamlFiles(data);
        } catch (err) { console.error(err); }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4", textAlign: "center" }}>{status}</p>}

            {step === 1 && (
                <div>
                    <h3>1. Select Source & Repository</h3>
                    <div style={styles.tabContainer}>
                        <div style={styles.tab(sourceType === "azure")} onClick={() => setSourceType("azure")}>Azure Repos</div>
                        <div style={styles.tab(sourceType === "github")} onClick={() => setSourceType("github")}>GitHub</div>
                    </div>
                    <input style={styles.input} placeholder={`Search ${sourceType} repositories...`} onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} />
                    <div style={styles.repoList}>
                        {repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                            <button key={r.id} style={styles.repoItem} onClick={() => handleRepoSelect(r)}>
                                <b style={{ marginRight: "10px" }}>{sourceType === "azure" ? "A" : "G"}</b> {r.name}
                            </button>
                        ))}
                    </div>
                </div>
            )}

            {step === 2 && (
                <div>
                    <h3>2. Configure Path</h3>
                    <p>Repo: <b>{formData.repoName}</b></p>
                    <label>Branch</label>
                    <select style={styles.input} onChange={(e) => handleBranchChange(e.target.value)}>
                        <option value="">-- Select Branch --</option>
                        {branches.map(b => <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>)}
                    </select>
                    <label>YAML File</label>
                    <select style={styles.input} onChange={(e) => setFormData({...formData, yamlPath: e.target.value})}>
                        <option value="">-- Select File --</option>
                        {yamlFiles.map(f => <option key={f} value={f}>{f}</option>)}
                    </select>
                    <button onClick={() => setStep(3)} style={{ padding: "10px 20px", background: "#0078d4", color: "#fff", border: "none", borderRadius: "4px" }}>Next</button>
                    <button onClick={() => setStep(1)} style={{ marginLeft: "10px", background: "none", border: "none", color: "#0078d4" }}>Back</button>
                </div>
            )}
            
            {/* Step 3 would go here */}
        </div>
    );
};

export default function App() {
    return (
        <div style={{ maxWidth: "800px", margin: "auto", padding: "40px" }}>
            <h1>Pipeline Generator</h1>
            <AuthenticatedTemplate><PipelineWizard /></AuthenticatedTemplate>
            <UnauthenticatedTemplate><p>Please Login</p></UnauthenticatedTemplate>
        </div>
    );
}
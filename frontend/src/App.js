import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [step, setStep] = useState(1);
    const [status, setStatus] = useState("");
    const [repos, setRepos] = useState([]);
    const [branches, setBranches] = useState([]);
    const [yamlFiles, setYamlFiles] = useState([]);
    const [searchTerm, setSearchTerm] = useState("");
    
    const [sourceType, setSourceType] = useState("azure"); 
    const [isManualPath, setIsManualPath] = useState(false);
    const [isUnrestricted, setIsUnrestricted] = useState(false);
    const [nameError, setNameError] = useState("");

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '', name: '' 
    });

    // --- SESSION TIMEOUT LOGIC ---
    useEffect(() => {
        const timeoutLimit = 10 * 60 * 1000;
        const timer = setTimeout(() => {
            alert("Session expired (10 minutes). Re-authenticating...");
            instance.logoutRedirect();
        }, timeoutLimit);
        return () => clearTimeout(timer);
    }, [instance]);

    const styles = {
        card: { background: "#fff", padding: "30px", borderRadius: "8px", border: "1px solid #ddd", marginTop: "20px" },
        input: { width: "100%", padding: "10px", marginBottom: "15px", border: "1px solid #ccc", borderRadius: "4px", fontSize: "14px", boxSizing: "border-box" },
        label: { display: "block", marginBottom: "5px", fontWeight: "600", fontSize: "13px", color: "#333" },
        errorText: { color: "#d13438", fontSize: "12px", marginTop: "-10px", marginBottom: "10px", fontWeight: "500" },
        toggleLink: { color: "#0078d4", cursor: "pointer", fontSize: "12px", textDecoration: "underline", marginBottom: "10px", display: "inline-block" },
        primaryBtn: { padding: "10px 20px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "600" },
        repoListWrapper: { border: "1px solid #eaeaea", borderRadius: "4px", marginTop: "10px", maxHeight: "350px", overflowY: "auto", width: "100%", display: "flex", flexDirection: "column" },
        repoItem: { display: "flex", alignItems: "center", width: "100%", padding: "14px 18px", textAlign: "left", cursor: "pointer", border: "none", background: "#fff", borderBottom: "1px solid #f3f2f1", fontSize: "14px", boxSizing: "border-box" },
        backBtn: { background: "none", border: "none", color: "#0078d4", cursor: "pointer", fontSize: "14px", padding: 0, marginLeft: '15px' },
        tabContainer: { display: "flex", marginBottom: "20px", borderBottom: "1px solid #ddd" },
        tab: (active) => ({
            padding: "10px 20px",
            cursor: "pointer",
            borderBottom: active ? "3px solid #0078d4" : "3px solid transparent",
            fontWeight: active ? "600" : "400",
            color: active ? "#0078d4" : "#666"
        })
    };

    const handleNameChange = (val) => {
        setFormData({ ...formData, name: val });
        if (isUnrestricted) { setNameError(""); return; }
        if (val.length > 48) { setNameError("Pipeline name cannot exceed 48 characters."); } 
        else { setNameError(""); }
    };

    // --- FETCH REPOS (With Array Check) ---
    useEffect(() => {
        const fetchRepos = async () => {
            setRepos([]);
            setStatus(""); 
            try {
                const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
                const endpoint = sourceType === "azure" ? "/api/repos" : "/api/github/repos";
                const res = await fetch(endpoint, { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
                const data = await res.json();

                if (res.ok && Array.isArray(data)) {
                    setRepos(data);
                } else {
                    setRepos([]);
                    setStatus(data.error || "Failed to load repositories.");
                }
            } catch (err) { 
                console.error(err);
                setRepos([]);
                setStatus("Connection error.");
            }
        };
        if (accounts.length > 0) fetchRepos();
    }, [instance, accounts, sourceType]);

    const handleRepoSelect = async (repo) => {
        setFormData({ ...formData, repoId: repo.id, repoName: repo.name });
        setStatus(`Loading ${sourceType} configuration...`);
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const endpoint = sourceType === "azure" 
                ? `/api/repos/${repo.id}/branches` 
                : `/api/github/repos/${repo.id}/branches`;
            
            const res = await fetch(endpoint, { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
            const data = await res.json();
            
            if (Array.isArray(data)) {
                setBranches(data);
                setStep(2);
                setStatus("");
            } else {
                setStatus("Invalid branch data received.");
            }
        } catch (err) { setStatus("Error loading branches."); }
    };

    const handleBranchChange = async (branchName) => {
        setFormData({ ...formData, branch: branchName, yamlPath: '' });
        if (!branchName) return;
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const baseUrl = sourceType === "azure" ? `/api/repos/${formData.repoId}` : `/api/github/repos/${formData.repoId}`;
            const res = await fetch(`${baseUrl}/yaml-files?branch=${branchName}`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            setYamlFiles(Array.isArray(data) ? data : []);
        } catch (err) { console.error(err); }
    };

    const handleCreatePipeline = async () => {
        setStatus("🚀 Creating pipeline...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${tokenResponse.accessToken}` },
                body: JSON.stringify({ ...formData, pipelineName: formData.name, sourceType, runPipeline: false })
            });
            if (res.ok) { setStatus("✅ Saved Successfully!"); } 
            else { setStatus("❌ Failed to create pipeline."); }
        } catch (err) { setStatus("❌ Connection error."); }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4", textAlign: "center" }}><b>{status}</b></p>}

            {step === 1 && (
                <div style={{ width: "100%" }}>
                    <h2 style={{ fontSize: "18px", marginBottom: "15px" }}>1. Select Source & Repository</h2>
                    <div style={styles.tabContainer}>
                        <div style={styles.tab(sourceType === "azure")} onClick={() => setSourceType("azure")}>Azure Repos</div>
                        <div style={styles.tab(sourceType === "github")} onClick={() => setSourceType("github")}>GitHub</div>
                    </div>
                    <input type="text" placeholder={`Search ${sourceType} repositories...`} style={styles.input} onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} />
                    
                    <div style={styles.repoListWrapper}>
                        {Array.isArray(repos) ? repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                            <button key={r.id} onClick={() => handleRepoSelect(r)} style={styles.repoItem}>
                                <div style={{ width: "24px", color: sourceType === "azure" ? "#0078d4" : "#24292e", fontWeight: "bold" }}>
                                    {sourceType === "azure" ? "A" : "G"}
                                </div>
                                <span style={{ flexGrow: 1 }}>{r.name}</span>
                            </button>
                        )) : <p style={{padding: '10px'}}>No repositories found.</p>}
                    </div>
                </div>
            )}

            {step === 2 && (
                <div>
                    <h2>2. Configure Path ({sourceType === "azure" ? "Azure" : "GitHub"})</h2>
                    <p>Repository: <b>{formData.repoName}</b></p>
                    <label style={styles.label}>Branch</label>
                    <select style={styles.input} value={formData.branch} onChange={(e) => handleBranchChange(e.target.value)}>
                        <option value="">-- Select Branch --</option>
                        {branches.map(b => <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>)}
                    </select>

                    <label style={styles.label}>YAML Path</label>
                    <span style={styles.toggleLink} onClick={() => setIsManualPath(!isManualPath)}>
                        {isManualPath ? "← Use file picker" : "Paste path manually →"}
                    </span>

                    {isManualPath ? (
                        <input style={styles.input} placeholder="e.g. /azure-pipelines.yml" value={formData.yamlPath} onChange={(e) => setFormData({...formData, yamlPath: e.target.value})} />
                    ) : (
                        <select style={styles.input} value={formData.yamlPath} onChange={(e) => setFormData({...formData, yamlPath: e.target.value})}>
                            <option value="">-- Select File --</option>
                            {yamlFiles.map(f => <option key={f} value={f}>{f}</option>)}
                        </select>
                    )}

                    <button onClick={() => setStep(3)} style={styles.primaryBtn} disabled={!formData.yamlPath}>Next</button>
                    <button onClick={() => setStep(1)} style={styles.backBtn}>Back</button>
                </div>
            )}

            {step === 3 && (
                <div>
                    <h2>3. Review & Name</h2>
                    <label style={styles.label}>Pipeline Name</label>
                    <div style={{ marginBottom: "10px" }}>
                        <span style={styles.toggleLink} onClick={() => { setIsUnrestricted(!isUnrestricted); setNameError(""); }}>
                            {isUnrestricted ? "Switch to Standard Mode" : "Switch to Unrestricted Mode"}
                        </span>
                    </div>
                    <input style={styles.input} value={formData.name} placeholder="Pipeline Name" onChange={(e) => handleNameChange(e.target.value)} />
                    {nameError && <p style={styles.errorText}>{nameError}</p>}
                    <button style={styles.primaryBtn} disabled={!!nameError || !formData.name} onClick={handleCreatePipeline}>Create Pipeline</button>
                    <button onClick={() => setStep(2)} style={styles.backBtn}>← Back</button>
                </div>
            )}
        </div>
    );
};

function App() {
    const { instance } = useMsal();
    return (
        <div style={{ padding: "40px", fontFamily: "Segoe UI", maxWidth: "900px", margin: "auto" }}>
            <h1>Pipeline Generator</h1>
            <UnauthenticatedTemplate>
                <div style={{ textAlign: "center", marginTop: "50px" }}>
                    <button onClick={() => instance.loginRedirect(loginRequest)} style={{ padding: "12px 24px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "600" }}>Login to Azure</button>
                </div>
            </UnauthenticatedTemplate>
            <AuthenticatedTemplate>
                <PipelineWizard />
            </AuthenticatedTemplate>
        </div>
    );
}

export default App;
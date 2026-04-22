import React, { useState, useEffect, useRef } from "react";
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

    // SECURITY FIX: Centralized session timeout logic
    useEffect(() => {
        const timeoutLimit = 10 * 60 * 1000;
        const timer = setTimeout(() => {
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

    // SECURITY FIX: Added Regex to block dangerous shell characters (; , & , |)
    const handleNameChange = (val) => {
        setFormData(prev => ({ ...prev, name: val }));
        const dangerousChars = /[;&|<>]/;
        
        if (!isUnrestricted) {
            if (val.length > 48) { 
                setNameError("Pipeline name cannot exceed 48 characters."); 
            } else if (dangerousChars.test(val)) {
                setNameError("Invalid characters detected (;, &, |, <, >).");
            } else { 
                setNameError(""); 
            }
        } else {
            setNameError("");
        }
    };

    useEffect(() => {
        // SECURITY FIX: AbortController prevents race conditions if user toggles tabs fast
        const controller = new AbortController();
        
        const fetchRepos = async () => {
            setRepos([]);
            setStatus("Loading repositories..."); 
            try {
                const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
                const endpoint = sourceType === "azure" ? "/api/repos" : "/api/github/repos";
                
                const res = await fetch(endpoint, { 
                    headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` },
                    signal: controller.signal 
                });
                
                const data = await res.json();
                if (res.ok && Array.isArray(data)) {
                    setRepos(data);
                    setStatus("");
                } else {
                    setStatus(data.error || "Failed to load repositories.");
                }
            } catch (err) { 
                if (err.name !== 'AbortError') {
                    console.error("Fetch error:", err);
                    setStatus("Error connecting to backend.");
                }
            }
        };

        if (accounts.length > 0) fetchRepos();
        return () => controller.abort(); // Cleanup
    }, [instance, accounts, sourceType]);

    const handleRepoSelect = async (repo) => {
        // SECURITY FIX: Functional state update
        setFormData(prev => ({ ...prev, repoId: repo.id, repoName: repo.name }));
        setStatus(`Loading ${sourceType} configuration...`);
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const encodedRepoId = encodeURIComponent(encodeURIComponent(repo.id));
            const endpoint = sourceType === "azure" 
                ? `/api/repos/${encodedRepoId}/branches` 
                : `/api/github/repos/${encodedRepoId}/branches`;
            
            const res = await fetch(endpoint, { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
            const data = await res.json();
            const branchList = Array.isArray(data) ? data : (data.value || []);

            if (branchList.length > 0) {
                setBranches(branchList);
                setStep(2);
                setStatus("");
            } else {
                setStatus("No branch data found.");
            }
        } catch (err) { 
            setStatus("Failed to fetch branches."); 
        }
    };

    const handleBranchChange = async (branchName) => {
        setFormData(prev => ({ ...prev, branch: branchName, yamlPath: '' }));
        if (!branchName) return;
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const encodedRepoId = encodeURIComponent(encodeURIComponent(formData.repoId));
            const baseUrl = sourceType === "azure" ? `/api/repos/${encodedRepoId}` : `/api/github/repos/${encodedRepoId}`;
            const res = await fetch(`${baseUrl}/yaml-files?branch=${branchName}`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });

            const data = await res.json();
            setYamlFiles(Array.isArray(data) ? data : []);
        } catch (err) { console.error("YAML fetch failed"); }
    };

    const handleCreatePipeline = async () => {
        // SECURITY FIX: Final check before submission
        if (nameError || !formData.name || !formData.yamlPath) {
            setStatus("❌ Please correct errors before submitting.");
            return;
        }

        setStatus("🚀 Creating pipeline...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${tokenResponse.accessToken}` },
                body: JSON.stringify({ ...formData, pipelineName: formData.name, sourceType })
            });

            const result = await res.json();
            if (res.ok) { 
                setStatus("✅ Pipeline created successfully!"); 
            } else { 
                setStatus(`❌ ${result.error || "Failed to create pipeline."}`); 
            }
        } catch (err) { 
            setStatus("❌ Connection error."); 
        }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4", textAlign: "center", fontSize: "14px" }}><b>{status}</b></p>}
            {step === 1 && (
                <div style={{ width: "100%" }}>
                    <h2 style={{ fontSize: "18px", marginBottom: "15px" }}>1. Select Source & Repository</h2>
                    <div style={styles.tabContainer}>
                        <div style={styles.tab(sourceType === "azure")} onClick={() => setSourceType("azure")}>Azure Repos</div>
                        <div style={styles.tab(sourceType === "github")} onClick={() => setSourceType("github")}>GitHub</div>
                    </div>
                    <input type="text" placeholder={`Search ${sourceType} repositories...`} style={styles.input} onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} />
                    <div style={styles.repoListWrapper}>
                        {Array.isArray(repos) && repos.length > 0 ? repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                            <button key={r.id} onClick={() => handleRepoSelect(r)} style={styles.repoItem}>
                                <div style={{ width: "24px", color: sourceType === "azure" ? "#0078d4" : "#24292e", fontWeight: "bold" }}>
                                    {sourceType === "azure" ? "A" : "G"}
                                </div>
                                <span style={{ flexGrow: 1 }}>{r.name}</span>
                            </button>
                        )) : <p style={{padding: '20px', textAlign: 'center', fontSize: '13px', color: '#666'}}>No repositories found.</p>}
                    </div>
                </div>
            )}
            {step === 2 && (
                <div>
                    <h2>2. Configure Path ({sourceType === "azure" ? "Azure" : "GitHub"})</h2>
                    <p style={{fontSize: '14px'}}>Repository: <b>{formData.repoName}</b></p>
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
                        <input style={styles.input} placeholder="e.g. /azure-pipelines.yml" value={formData.yamlPath} onChange={(e) => setFormData(prev => ({...prev, yamlPath: e.target.value}))} />
                    ) : (
                        <select style={styles.input} value={formData.yamlPath} onChange={(e) => setFormData(prev => ({...prev, yamlPath: e.target.value}))}>
                            <option value="">-- Select File --</option>
                            {yamlFiles.map(f => <option key={f} value={f}>{f}</option>)}
                        </select>
                    )}
                    <div style={{marginTop: '20px'}}>
                        <button onClick={() => setStep(3)} style={styles.primaryBtn} disabled={!formData.yamlPath || !formData.branch}>Next</button>
                        <button onClick={() => { setStep(1); setStatus(""); }} style={styles.backBtn}>Back</button>
                    </div>
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
                    <div style={{marginTop: '20px'}}>
                        <button style={styles.primaryBtn} disabled={!!nameError || !formData.name} onClick={handleCreatePipeline}>Create Pipeline</button>
                        <button onClick={() => setStep(2)} style={styles.backBtn}>← Back</button>
                    </div>
                </div>
            )}
        </div>
    );
};

function App() {
    const { instance } = useMsal();
    return (
        <div style={{ padding: "40px", fontFamily: "Segoe UI, Tahoma, Geneva, Verdana, sans-serif", maxWidth: "900px", margin: "auto" }}>
            <h1 style={{color: '#333'}}>Pipeline Generator</h1>
            <UnauthenticatedTemplate>
                <div style={{ textAlign: "center", marginTop: "50px", padding: "40px", border: "1px dashed #ccc", borderRadius: "8px" }}>
                    <p style={{marginBottom: "20px", color: "#666"}}>Please sign in with your corporate account to continue.</p>
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
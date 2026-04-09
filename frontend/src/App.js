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
    
    // Enhancement States
    const [isManualPath, setIsManualPath] = useState(false);
    const [isUnrestricted, setIsUnrestricted] = useState(false); // Toggle for restriction bypass
    const [nameError, setNameError] = useState("");

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '', name: '' 
    });

    // 10-Minute Session Timeout Enhancement
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
        backBtn: { marginTop: "20px", background: "none", border: "none", color: "#0078d4", cursor: "pointer", fontSize: "14px", padding: 0 }
    };

    // Logic: Standard (48 chars + k8s/deployment) vs Unrestricted (Anything)
    const handleNameChange = (val) => {
        setFormData({ ...formData, name: val });
        
        if (isUnrestricted) {
            setNameError(""); // No restrictions at all
            return;
        }

        const lowerVal = val.toLowerCase();
        if (val.length > 48) {
            setNameError("Standard name cannot exceed 48 characters.");
        } else if (!lowerVal.includes("k8s") && !lowerVal.includes("deployment")) {
            setNameError("Standard name must contain 'k8s' or 'deployment'.");
        } else {
            setNameError("");
        }
    };

    useEffect(() => {
        const fetchRepos = async () => {
            try {
                const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
                const res = await fetch("/api/repos", { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
                const data = await res.json();
                setRepos(data || []);
            } catch (err) { console.error(err); }
        };
        if (accounts.length > 0) fetchRepos();
    }, [instance, accounts]);

    const handleRepoSelect = async (repo) => {
        setFormData({ ...formData, repoId: repo.id, repoName: repo.name });
        setStatus("Loading configuration...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${repo.id}/branches`, { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
            const data = await res.json();
            setBranches(data || []);
            setStep(2);
            setStatus("");
        } catch (err) { setStatus("Error loading branches."); }
    };

    const handleBranchChange = async (branchName) => {
        setFormData({ ...formData, branch: branchName, yamlPath: '' });
        if (!branchName) return;
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${formData.repoId}/yaml-files?branch=${branchName}`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            setYamlFiles(data || []);
        } catch (err) { console.error(err); }
    };

    const handleCreatePipeline = async () => {
        setStatus("🚀 Creating pipeline in Azure DevOps...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { 
                    "Content-Type": "application/json", 
                    "Authorization": `Bearer ${tokenResponse.accessToken}` 
                },
                body: JSON.stringify({ 
                    ...formData, 
                    pipelineName: formData.name, 
                    runPipeline: false 
                })
            });

            if (res.ok) {
                setStatus("✅ Saved Successfully!");
            } else {
                const errData = await res.json();
                setStatus(`❌ Error: ${errData.message || "Failed to create"}`);
            }
        } catch (err) {
            setStatus("❌ Connection error.");
        }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4", textAlign: "center" }}><b>{status}</b></p>}

            {step === 1 && (
                <div style={{ width: "100%" }}>
                    <h2 style={{ fontSize: "18px", marginBottom: "15px" }}>1. Select a repository</h2>
                    <input type="text" placeholder="Filter repositories..." style={styles.input} onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} />
                    <div style={styles.repoListWrapper}>
                        {repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                            <button key={r.id} onClick={() => handleRepoSelect(r)} style={styles.repoItem}>
                                <div style={{ width: "24px", color: "#0078d4", fontWeight: "bold" }}>{r.name.charAt(0).toUpperCase()}</div>
                                <span style={{ flexGrow: 1 }}>{r.name}</span>
                            </button>
                        ))}
                    </div>
                </div>
            )}

            {step === 2 && (
                <div>
                    <h2>2. Configure Path</h2>
                    <p>Repository: <b>{formData.repoName}</b></p>
                    <label style={styles.label}>Branch</label>
                    <select style={styles.input} value={formData.branch} onChange={(e) => handleBranchChange(e.target.value)}>
                        <option value="">-- Select Branch --</option>
                        {branches.map(b => <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>)}
                    </select>

                    <label style={styles.label}>YAML Path</label>
                    <span style={styles.toggleLink} onClick={() => setIsManualPath(!isManualPath)}>
                        {isManualPath ? "← Use file picker" : "Paste path manually (Monorepo) →"}
                    </span>

                    {isManualPath ? (
                        <input 
                            style={styles.input} 
                            placeholder="e.g. /folder/azure-pipelines.yml" 
                            value={formData.yamlPath}
                            onChange={(e) => setFormData({...formData, yamlPath: e.target.value})}
                        />
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
                        <span 
                            style={styles.toggleLink} 
                            onClick={() => { 
                                setIsUnrestricted(!isUnrestricted); 
                                setNameError(""); // Clear errors when switching
                            }}
                        >
                            {isUnrestricted ? "Switch to Standard Mode (Restrictions ON)" : "Switch to Unrestricted Mode (Restrictions OFF)"}
                        </span>
                    </div>

                    <input 
                        style={styles.input} 
                        value={formData.name} 
                        placeholder={isUnrestricted ? "Enter any name" : "Must contain k8s/deployment & max 48 chars"}
                        onChange={(e) => handleNameChange(e.target.value)} 
                    />
                    {nameError && <p style={styles.errorText}>{nameError}</p>}

                    <button 
                        style={styles.primaryBtn} 
                        disabled={!!nameError || !formData.name}
                        onClick={handleCreatePipeline}
                    >
                        Create Pipeline
                    </button>
                    <button onClick={() => setStep(2)} style={styles.backBtn}>← Back to Configure</button>
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
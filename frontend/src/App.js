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
    const [yamlContent, setYamlContent] = useState("");
    
    const [variables, setVariables] = useState([]); 
    const [showSaveOptions, setShowSaveOptions] = useState(false);

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '', name: '' 
    });

    const styles = {
        card: { background: "#fff", padding: "30px", borderRadius: "8px", border: "1px solid #ddd", marginTop: "20px" },
        input: { width: "100%", padding: "10px", marginBottom: "15px", border: "1px solid #ccc", borderRadius: "4px", fontSize: "14px", boxSizing: "border-box" },
        primaryBtn: { padding: "10px 20px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "600" },
        codeArea: { background: "#1e1e1e", color: "#d4d4d4", padding: "20px", borderRadius: "4px", fontFamily: "monospace", overflow: "auto", maxHeight: "400px", textAlign: "left", fontSize: "13px", lineHeight: "1.5" },
        varBox: { marginBottom: "20px", padding: "15px", background: "#f8f9fa", border: "1px solid #eee", borderRadius: "4px" },
        splitBtnContainer: { display: "flex", position: "relative" },
        runBtn: { padding: "10px 20px", background: "#107c10", color: "white", border: "none", borderRadius: "4px 0 0 4px", cursor: "pointer", fontWeight: "600" },
        arrowBtn: { padding: "10px 12px", background: "#0b5a0b", color: "white", border: "none", borderRadius: "0 4px 4px 0", borderLeft: "1px solid #084a08", cursor: "pointer" },
        dropdownMenu: { position: "absolute", top: "42px", right: 0, background: "white", border: "1px solid #ccc", zIndex: 100, width: "140px", boxShadow: "0 4px 8px rgba(0,0,0,0.1)" },
        
        // STYLES FOR THE REPO LIST
        repoListContainer: { 
            border: "1px solid #eaeaea", 
            borderRadius: "4px", 
            marginTop: "10px", 
            maxHeight: "450px", 
            overflowY: "auto",
            backgroundColor: "#fff"
        },
        repoItem: { 
            display: "flex", 
            alignItems: "center",
            width: "100%", 
            padding: "12px 16px", 
            textAlign: "left", 
            cursor: "pointer", 
            border: "none", 
            background: "transparent", 
            borderBottom: "1px solid #f3f2f1", 
            fontSize: "14px",
            color: "#323130",
            transition: "background 0.1s"
        },
        backBtn: { marginTop: "20px", background: "none", border: "none", color: "#0078d4", cursor: "pointer", fontSize: "14px", padding: 0 }
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
        setStatus("Fetching YAML files...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${formData.repoId}/yaml-files?branch=${branchName}`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            setYamlFiles(data || []);
            setStatus("");
        } catch (err) { setStatus("Error."); }
    };

    const fetchYamlPreview = async () => {
        setStatus("Loading YAML...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${formData.repoId}/content?path=${formData.yamlPath}&branch=${formData.branch}`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            setYamlContent(data.content);
            setStep(3);
            setStatus("");
        } catch (err) { setStatus("Error."); }
    };

    const handleAction = async (shouldRun) => {
        setShowSaveOptions(false);
        setStatus(shouldRun ? "🚀 Running Pipeline..." : "💾 Saving Pipeline...");
        const formattedVars = {};
        variables.forEach(v => { if(v.name) formattedVars[v.name] = { value: v.value }; });
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${tokenResponse.accessToken}` },
                body: JSON.stringify({ ...formData, pipelineName: formData.name, variables: formattedVars, runPipeline: shouldRun })
            });
            if (res.ok) setStatus(shouldRun ? "✅ Created & Triggered!" : "✅ Saved Successfully!");
            else setStatus("❌ Action Failed.");
        } catch (err) { setStatus("❌ Error occurred."); }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4" }}><b>{status}</b></p>}

            {/* STEP 1: SELECT REPOSITORY LIST */}
            {step === 1 && (
                <div>
                    <h2 style={{ fontSize: "20px", marginBottom: "20px" }}>Select a repository</h2>
                    <input 
                        type="text" 
                        placeholder="Filter repositories" 
                        style={styles.input} 
                        onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} 
                    />
                    <div style={styles.repoListContainer}>
                        {repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                            <button 
                                key={r.id} 
                                onClick={() => handleRepoSelect(r)} 
                                style={styles.repoItem}
                                onMouseOver={(e) => e.currentTarget.style.backgroundColor = "#f3f2f1"}
                                onMouseOut={(e) => e.currentTarget.style.backgroundColor = "transparent"}
                            >
                                <span style={{ color: "#0078d4", fontWeight: "600", marginRight: "8px" }}>{r.name.charAt(0).toUpperCase()}</span>
                                {r.name}
                            </button>
                        ))}
                    </div>
                </div>
            )}

            {/* STEP 2: CONFIGURE */}
            {step === 2 && (
                <div>
                    <h2>Configure your pipeline</h2>
                    <p>Repository: <b>{formData.repoName}</b></p>
                    <label>Branch</label>
                    <select style={styles.input} value={formData.branch} onChange={(e) => handleBranchChange(e.target.value)}>
                        <option value="">-- Select Branch --</option>
                        {branches.map(b => <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>)}
                    </select>
                    <label>YAML Path</label>
                    <select style={styles.input} value={formData.yamlPath} onChange={(e) => setFormData({...formData, yamlPath: e.target.value})}>
                        <option value="">-- Select File --</option>
                        {yamlFiles.map(f => <option key={f} value={f}>{f}</option>)}
                    </select>
                    <button onClick={fetchYamlPreview} style={styles.primaryBtn} disabled={!formData.yamlPath}>Review YAML</button>
                    <button onClick={() => setStep(1)} style={{ marginLeft: "15px", background: "none", border: "none", color: "#0078d4", cursor: "pointer" }}>Back</button>
                </div>
            )}

            {/* STEP 3: REVIEW & RUN */}
            {step === 3 && (
                <div>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "15px", alignItems: "center" }}>
                        <h2 style={{ margin: 0 }}>Review your pipeline YAML</h2>
                        <div style={{ display: "flex", gap: "8px" }}>
                            <button onClick={() => setVariables([...variables, { name: '', value: '' }])} style={{ padding: "10px 15px", background: "#f0f0f0", border: "1px solid #ccc", borderRadius: "4px", cursor: "pointer", fontWeight: "600" }}>Variables</button>
                            <div style={styles.splitBtnContainer}>
                                <button onClick={() => handleAction(true)} style={styles.runBtn}>Run</button>
                                <button onClick={() => setShowSaveOptions(!showSaveOptions)} style={styles.arrowBtn}>▼</button>
                                {showSaveOptions && (
                                    <div style={styles.dropdownMenu}>
                                        <button onClick={() => handleAction(false)} style={{ display: "block", padding: "12px", border: "none", width: "100%", background: "none", textAlign: "left", cursor: "pointer" }}>Save only</button>
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>

                    {variables.length > 0 && (
                        <div style={styles.varBox}>
                            <h4>Pipeline Variables</h4>
                            {variables.map((v, i) => (
                                <div key={i} style={{ display: "flex", gap: "10px", marginBottom: "8px" }}>
                                    <input placeholder="Name" style={styles.input} value={v.name} onChange={(e) => { const n = [...variables]; n[i].name = e.target.value; setVariables(n); }} />
                                    <input placeholder="Value" style={styles.input} value={v.value} onChange={(e) => { const n = [...variables]; n[i].value = e.target.value; setVariables(n); }} />
                                    <button onClick={() => setVariables(variables.filter((_, idx) => idx !== i))} style={{ border: "none", background: "none", color: "#d13438", cursor: "pointer", fontWeight: "bold" }}>✕</button>
                                </div>
                            ))}
                        </div>
                    )}

                    <div style={styles.codeArea}><pre style={{ margin: 0 }}>{yamlContent}</pre></div>

                    <div style={{ marginTop: "20px" }}>
                        <label>Pipeline Name</label>
                        <input style={styles.input} value={formData.name} placeholder="Name your pipeline" onChange={(e) => setFormData({...formData, name: e.target.value})} />
                    </div>
                    
                    <button onClick={() => setStep(2)} style={styles.backBtn}>
                        ← Back to Configure
                    </button>
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
          <button onClick={() => instance.loginRedirect(loginRequest)} style={{ padding: "12px 24px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "600" }}>Login to Azure DevOps</button>
        </div>
      </UnauthenticatedTemplate>
      <AuthenticatedTemplate>
          <PipelineWizard />
      </AuthenticatedTemplate>
    </div>
  );
}

export default App;
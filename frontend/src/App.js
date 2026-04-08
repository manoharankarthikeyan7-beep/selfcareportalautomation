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
    
    // UI Logic States
    const [variables, setVariables] = useState([]); 
    const [showSaveOptions, setShowSaveOptions] = useState(false);

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '', name: '' 
    });

    const styles = {
        card: { background: "#fff", padding: "30px", borderRadius: "8px", border: "1px solid #ddd", marginTop: "20px" },
        input: { width: "100%", padding: "10px", marginBottom: "10px", border: "1px solid #ccc", borderRadius: "4px" },
        primaryBtn: { padding: "10px 20px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "600" },
        codeArea: { background: "#1e1e1e", color: "#d4d4d4", padding: "20px", borderRadius: "4px", fontFamily: "monospace", overflow: "auto", maxHeight: "300px", textAlign: "left" },
        varBox: { marginBottom: "20px", padding: "15px", background: "#f8f9fa", border: "1px solid #eee", borderRadius: "4px" }
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
        setStatus("Loading branches...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${repo.id}/branches`, { headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` } });
            const data = await res.json();
            setBranches(data || []);
            setStatus("");
            setStep(2);
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
        } catch (err) { setStatus("No files found."); }
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
        } catch (err) { setStatus("Error reading file."); }
    };

    const handleAction = async (shouldRun) => {
        setShowSaveOptions(false);
        setStatus(shouldRun ? "🚀 Running..." : "💾 Saving...");
        
        const formattedVars = {};
        variables.forEach(v => { if(v.name) formattedVars[v.name] = { value: v.value }; });

        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${tokenResponse.accessToken}` },
                body: JSON.stringify({ ...formData, pipelineName: formData.name, variables: formattedVars, runPipeline: shouldRun })
            });
            if (res.ok) setStatus(shouldRun ? "✅ Created & Queued!" : "✅ Saved!");
            else setStatus("❌ Failed.");
        } catch (err) { setStatus("❌ Error."); }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4" }}><b>{status}</b></p>}

            {step === 1 && (
                <div>
                    <h2>1. Select a repository</h2>
                    <input type="text" placeholder="Search..." style={styles.input} onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} />
                    <div style={{ maxHeight: "200px", overflowY: "auto" }}>
                        {repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                            <button key={r.id} onClick={() => handleRepoSelect(r)} style={{ display: "block", width: "100%", padding: "10px", textAlign: "left", cursor: "pointer", border: "none", background: "none", borderBottom: "1px solid #eee" }}>{r.name}</button>
                        ))}
                    </div>
                </div>
            )}

            {step === 2 && (
                <div>
                    <h2>2. Configure</h2>
                    <label>Branch</label>
                    <select style={styles.input} onChange={(e) => handleBranchChange(e.target.value)}>
                        <option value="">-- Select Branch --</option>
                        {branches.map(b => <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>)}
                    </select>
                    <label>YAML Path</label>
                    <select style={styles.input} onChange={(e) => setFormData({...formData, yamlPath: e.target.value})}>
                        <option value="">-- Select File --</option>
                        {yamlFiles.map(f => <option key={f} value={f}>{f}</option>)}
                    </select>
                    <button onClick={fetchYamlPreview} style={styles.primaryBtn} disabled={!formData.yamlPath}>Continue</button>
                </div>
            )}

            {step === 3 && (
                <div>
                    <div style={{ display: "flex", justifyContent: "space-between", marginBottom: "15px", alignItems: "center" }}>
                        <h2>3. Review & Run</h2>
                        <div style={{ display: "flex", gap: "8px" }}>
                            <button onClick={() => setVariables([...variables, { name: '', value: '' }])} style={{ ...styles.primaryBtn, background: "#f0f0f0", color: "#333", border: "1px solid #ccc" }}>Variables</button>
                            <div style={{ position: "relative", display: "flex" }}>
                                <button onClick={() => handleAction(true)} style={{ ...styles.primaryBtn, background: "#107c10", borderRadius: "4px 0 0 4px" }}>Run</button>
                                <button onClick={() => setShowSaveOptions(!showSaveOptions)} style={{ ...styles.primaryBtn, background: "#0b5a0b", borderRadius: "0 4px 4px 0", borderLeft: "1px solid #084a08" }}>▼</button>
                                {showSaveOptions && (
                                    <div style={{ position: "absolute", top: "45px", right: 0, background: "white", border: "1px solid #ccc", zIndex: 100, width: "120px", boxShadow: "0 4px 6px rgba(0,0,0,0.1)" }}>
                                        <button onClick={() => handleAction(false)} style={{ display: "block", padding: "10px", border: "none", width: "100%", background: "none", textAlign: "left", cursor: "pointer" }}>Save only</button>
                                    </div>
                                )}
                            </div>
                        </div>
                    </div>

                    {variables.length > 0 && (
                        <div style={styles.varBox}>
                            <h4>Variables</h4>
                            {variables.map((v, i) => (
                                <div key={i} style={{ display: "flex", gap: "10px" }}>
                                    <input placeholder="Name" style={styles.input} onChange={(e) => { const n = [...variables]; n[i].name = e.target.value; setVariables(n); }} />
                                    <input placeholder="Value" style={styles.input} onChange={(e) => { const n = [...variables]; n[i].value = e.target.value; setVariables(n); }} />
                                </div>
                            ))}
                        </div>
                    )}

                    <div style={styles.codeArea}><pre style={{ margin: 0 }}>{yamlContent}</pre></div>
                    <div style={{ marginTop: "15px" }}>
                        <label>Pipeline Name</label>
                        <input style={styles.input} placeholder="My-New-Pipeline" onChange={(e) => setFormData({...formData, name: e.target.value})} />
                    </div>
                    <button onClick={() => setStep(2)} style={{ background: "none", border: "none", color: "#0078d4", cursor: "pointer" }}>← Back to Configure</button>
                </div>
            )}
        </div>
    );
};

export default PipelineWizard;
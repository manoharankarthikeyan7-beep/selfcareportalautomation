import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [step, setStep] = useState(1);
    const [status, setStatus] = useState("");
    const [repos, setRepos] = useState([]);
    const [branches, setBranches] = useState([]);
    const [searchTerm, setSearchTerm] = useState("");
    const [yamlContent, setYamlContent] = useState("");

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '/azure-pipelines.yml', name: '' 
    });

    // Professional Styles
    const styles = {
        card: { background: "#fff", padding: "30px", borderRadius: "8px", border: "1px solid #ddd", marginTop: "20px" },
        input: { width: "100%", padding: "12px", marginBottom: "15px", boxSizing: "border-box", border: "1px solid #ccc", borderRadius: "4px" },
        primaryBtn: { padding: "10px 25px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer", fontWeight: "600" },
        codeArea: { background: "#1e1e1e", color: "#d4d4d4", padding: "20px", borderRadius: "4px", fontFamily: "monospace", overflow: "auto", maxHeight: "350px", textAlign: "left" }
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
        } catch (err) { setStatus("Error."); }
    };

    const fetchYamlPreview = async () => {
        if (!formData.branch) { alert("Select a branch!"); return; }
        setStatus("Reading YAML...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch(`/api/repos/${formData.repoId}/content?path=${formData.yamlPath}&branch=${formData.branch}`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            if (data.error) throw new Error();
            setYamlContent(data.content);
            setStep(3);
            setStatus("");
        } catch (err) { setStatus("❌ YAML file not found in this branch."); }
    };

    const finalCreateCall = async () => {
        setStatus("🚀 Provisioning Pipeline...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { "Content-Type": "application/json", "Authorization": `Bearer ${tokenResponse.accessToken}` },
                body: JSON.stringify({ pipelineName: formData.name, repoId: formData.repoId, branch: formData.branch, yamlPath: formData.yamlPath })
            });
            if (res.ok) setStatus("✅ Pipeline Created Successfully!");
            else setStatus("❌ Failed to create.");
        } catch (err) { setStatus("❌ Connection error."); }
    };

    return (
        <div style={styles.card}>
            {status && <p style={{ color: "#0078d4" }}><b>{status}</b></p>}

            {step === 1 && (
                <div>
                    <h2>1. Select a repository</h2>
                    <input type="text" placeholder="Search by name..." style={styles.input} onChange={(e) => setSearchTerm(e.target.value.toLowerCase())} />
                    <div style={{ maxHeight: "250px", overflowY: "auto", border: "1px solid #eee" }}>
                        {searchTerm.length >= 2 ? (
                            repos.filter(r => r.name.toLowerCase().includes(searchTerm)).map(r => (
                                <button key={r.id} onClick={() => handleRepoSelect(r)} style={{ display: "block", width: "100%", padding: "10px", textAlign: "left", background: "none", border: "none", borderBottom: "1px solid #eee", cursor: "pointer" }}>
                                    {r.name}
                                </button>
                            ))
                        ) : <p style={{ padding: "10px", color: "#999" }}>Start typing to search...</p>}
                    </div>
                </div>
            )}

            {step === 2 && (
                <div>
                    <h2>2. Configure</h2>
                    <p>Repository: <b>{formData.repoName}</b></p>
                    <label>Branch</label>
                    <select style={styles.input} onChange={(e) => setFormData({...formData, branch: e.target.value})}>
                        <option value="">-- Select Branch --</option>
                        {branches.map(b => <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>)}
                    </select>
                    <label>YAML Path</label>
                    <input style={styles.input} value={formData.yamlPath} onChange={(e) => setFormData({...formData, yamlPath: e.target.value})} />
                    <button onClick={fetchYamlPreview} style={styles.primaryBtn}>Review YAML</button>
                </div>
            )}

            {step === 3 && (
                <div>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                        <h2>3. Review & Run</h2>
                        <button onClick={finalCreateCall} style={{ ...styles.primaryBtn, background: "#107c10" }}>Run</button>
                    </div>
                    <div style={styles.codeArea}><pre>{yamlContent}</pre></div>
                    <div style={{ marginTop: "20px" }}>
                        <label>Pipeline Name</label>
                        <input style={styles.input} placeholder="My-New-Pipeline" onChange={(e) => setFormData({...formData, name: e.target.value})} />
                    </div>
                </div>
            )}
        </div>
    );
};

function App() {
  const { instance } = useMsal();
  return (
    <div style={{ padding: "40px", fontFamily: "Segoe UI", maxWidth: "800px", margin: "auto" }}>
      <h1>Pipeline Generator</h1>
      <UnauthenticatedTemplate>
        <button onClick={() => instance.loginRedirect(loginRequest)} style={ {padding: "10px 20px", background: "#0078d4", color: "white", border: "none", borderRadius: "4px", cursor: "pointer" }}>Login</button>
      </UnauthenticatedTemplate>
      <AuthenticatedTemplate><PipelineWizard /></AuthenticatedTemplate>
    </div>
  );
}

export default App;
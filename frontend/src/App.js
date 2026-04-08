import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

// --- STEP 2: DYNAMIC RESOURCE DISCOVERY COMPONENT ---
const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [step, setStep] = useState(1);
    const [status, setStatus] = useState("");
    const [repos, setRepos] = useState([]);
    const [branches, setBranches] = useState([]);
    const [searchTerm, setSearchTerm] = useState(""); // Search logic for Discovery

    const [formData, setFormData] = useState({ 
        repoId: '', repoName: '', branch: '', yamlPath: '/azure-pipelines.yml', name: '' 
    });

    // Fetching Repos (Discovery)
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

    const finalCreateCall = async () => {
        setStatus("Creating Pipeline...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
            const res = await fetch("/api/pipelines/create", {
                method: "POST",
                headers: { 
                    "Content-Type": "application/json",
                    "Authorization": `Bearer ${tokenResponse.accessToken}` 
                },
                body: JSON.stringify({
                    pipelineName: formData.name,
                    repoId: formData.repoId,
                    branch: formData.branch,
                    yamlPath: formData.yamlPath
                })
            });
            if (res.ok) setStatus("✅ Success!");
            else setStatus("❌ Failed.");
        } catch (err) { setStatus("❌ Auth error."); }
    };

    return (
        <div style={{ background: "#f4f4f4", padding: "20px", borderRadius: "8px", marginTop: "20px" }}>
            {status && <p style={{ color: "blue", fontWeight: "bold" }}>{status}</p>}

            {step === 1 && (
                <div>
                    <h3>1. Select Repository</h3>
                    <input 
                        type="text"
                        placeholder="Search repositories..."
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

            {step === 2 && (
                <div>
                    <h3>2. Select Branch</h3>
                    <p>Repo: <b>{formData.repoName}</b></p>
                    <select style={{ width: "100%", padding: "10px" }} onChange={(e) => setFormData({...formData, branch: e.target.value})}>
                        <option value="">-- Choose Branch --</option>
                        {branches.map(b => (
                            <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>
                        ))}
                    </select>
                    <button onClick={() => setStep(3)} style={{ marginTop: "10px", padding: "10px" }}>Next</button>
                </div>
            )}

            {step === 3 && (
                <div>
                    <h3>3. Name & Create</h3>
                    <input 
                        placeholder="New Pipeline Name" 
                        style={{ width: "100%", padding: "10px" }}
                        onChange={(e) => setFormData({...formData, name: e.target.value})} 
                    />
                    <button onClick={finalCreateCall} style={{ marginTop: "10px", padding: "10px", background: "green", color: "white", border: "none", borderRadius: "4px" }}>
                        Save & Run
                    </button>
                </div>
            )}
        </div>
    );
};

// --- STEP 1: IDENTITY HANDSHAKE COMPONENT ---
function App() {
  const { instance } = useMsal();

  const handleLogin = () => {
    instance.loginRedirect(loginRequest).catch((e) => {
      console.error("Login Error:", e);
    });
  };

  return (
    <div style={{ padding: "40px", fontFamily: "Segoe UI" }}>
      <h1>Azure DevOps Pipeline Generator</h1>
      
      <UnauthenticatedTemplate>
        <p>Please sign in to manage pipelines.</p>
        <button 
          onClick={handleLogin} 
          style={{ padding: "10px 20px", fontSize: "16px", cursor: "pointer", backgroundColor: "#0078d4", color: "white", border: "none", borderRadius: "4px" }}
        >
          Login
        </button>
      </UnauthenticatedTemplate>

      <AuthenticatedTemplate>
        <PipelineWizard />
      </AuthenticatedTemplate>
    </div>
  );
}

// THIS IS THE LINE THAT WAS MISSING OR MISPLACED:
export default App;
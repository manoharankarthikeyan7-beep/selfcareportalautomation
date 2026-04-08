import React, { useState, useEffect } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "./authConfig";

// --- SUB-COMPONENT: The Wizard ---
const PipelineWizard = () => {
    const { instance, accounts } = useMsal();
    const [step, setStep] = useState(1);
    const [status, setStatus] = useState("");
    const [repos, setRepos] = useState([]);
    const [branches, setBranches] = useState([]);
    
    const [formData, setFormData] = useState({ 
        repoId: '', 
        repoName: '', 
        branch: '', 
        yamlPath: '/azure-pipelines.yml', 
        name: '' 
    });

    // Step 1: Load Repos
    useEffect(() => {
        const fetchRepos = async () => {
            try {
                const tokenResponse = await instance.acquireTokenSilent({
                    ...loginRequest,
                    account: accounts[0]
                });
                const res = await fetch("/api/repos", {
                    headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
                });
                const data = await res.json();
                setRepos(data || []);
            } catch (err) {
                console.error("Failed to load repos:", err);
            }
        };
        if (accounts.length > 0) fetchRepos();
    }, [instance, accounts]);

    // Step 2: Select Repo & Load Branches
    const handleRepoSelect = async (repo) => {
        setFormData({ ...formData, repoId: repo.id, repoName: repo.name });
        setStatus("Loading branches...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            const res = await fetch(`/api/repos/${repo.id}/branches`, {
                headers: { "Authorization": `Bearer ${tokenResponse.accessToken}` }
            });
            const data = await res.json();
            setBranches(data || []);
            setStatus("");
            setStep(2);
        } catch (err) {
            setStatus("Error loading branches.");
        }
    };

    // Step 3: Final Save/Create Call
    const finalCreateCall = async () => {
        setStatus("Creating Pipeline...");
        try {
            const tokenResponse = await instance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
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

            if (res.ok) setStatus("✅ Pipeline successfully created!");
            else setStatus("❌ Failed to create pipeline.");
        } catch (err) {
            setStatus("❌ Auth error.");
        }
    };

    return (
        <div style={{ background: "#f4f4f4", padding: "20px", borderRadius: "8px", marginTop: "20px" }}>
            {status && <p style={{ fontWeight: "bold", color: "blue" }}>{status}</p>}

            {step === 1 && (
                <div>
                    <h3>1. Select Repository</h3>
                    {repos.map(r => (
                        <button key={r.id} onClick={() => handleRepoSelect(r)} style={{ display: "block", margin: "10px 0", padding: "10px", width: "100%", textAlign: "left" }}>
                            {r.name}
                        </button>
                    ))}
                </div>
            )}

            {step === 2 && (
                <div>
                    <h3>2. Select Branch & YAML Path</h3>
                    <p>Repo: <b>{formData.repoName}</b></p>
                    <select 
                        style={{ width: "100%", padding: "10px" }}
                        onChange={(e) => setFormData({...formData, branch: e.target.value})}
                    >
                        <option value="">-- Choose Branch --</option>
                        {branches.map(b => (
                            <option key={b.name} value={b.name}>{b.name.replace('refs/heads/', '')}</option>
                        ))}
                    </select>
                    <input 
                        style={{ width: "100%", padding: "10px", marginTop: "10px" }}
                        value={formData.yamlPath} 
                        onChange={(e) => setFormData({...formData, yamlPath: e.target.value})} 
                    />
                    <button onClick={() => setStep(3)} style={{ marginTop: "10px", padding: "10px" }}>Next</button>
                </div>
            )}

            {step === 3 && (
                <div>
                    <h3>3. Rename & Run</h3>
                    <p>Selected: {formData.repoName} ({formData.branch.replace('refs/heads/', '')})</p>
                    <input 
                        placeholder="Pipeline Name" 
                        style={{ width: "100%", padding: "10px" }}
                        onChange={(e) => setFormData({...formData, name: e.target.value})} 
                    />
                    <button onClick={finalCreateCall} style={{ marginTop: "10px", padding: "10px", background: "green", color: "white" }}>
                        Save & Run
                    </button>
                </div>
            )}
        </div>
    );
};

// --- MAIN APP COMPONENT ---
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

export default App;
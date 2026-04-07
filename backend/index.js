// A. GET ALL REPOS
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(`https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`, 
        { headers: { 'Authorization': authHeader } });
        res.json(response.data.value);
    } catch (e) { res.status(500).send(e.message); }
});

// B. GET BRANCHES FOR SELECTED REPO
app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(`https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`, 
        { headers: { 'Authorization': authHeader } });
        res.json(response.data.value);
    } catch (e) { res.status(500).send(e.message); }
});

// C. THE CREATION ENGINE (Rename + Save)
app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath } = req.body;
    try {
        const response = await axios.post(`https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`, {
            name: pipelineName, // This handles your "Rename" requirement
            configuration: {
                type: "yaml",
                path: yamlPath, // e.g., "/azure-pipelines.yml"
                repository: {
                    id: repoId,
                    type: "azureReposGit",
                    defaultBranch: branch // e.g., "refs/heads/main"
                }
            }
        }, { headers: { 'Authorization': authHeader, 'Content-Type': 'application/json' } });
        res.json(response.data);
    } catch (e) { res.status(500).json(e.response?.data || e.message); }
});
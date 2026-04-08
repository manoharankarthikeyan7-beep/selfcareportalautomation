// A. LIST ALL REPOS (Step 1)
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(`https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`, 
        { headers: { 'Authorization': authHeader } });
        res.json(response.data.value);
    } catch (e) { res.status(500).json({ error: "Check ADO_ORG_NAME variable" }); }
});

// B. LIST BRANCHES (Step 2)
app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(`https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`, 
        { headers: { 'Authorization': authHeader } });
        res.json(response.data.value);
    } catch (e) { res.status(500).send(e.message); }
});

// C. CREATE & RENAME (Step 3)
app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath } = req.body;
    try {
        const response = await axios.post(`https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`, {
            name: pipelineName, 
            configuration: {
                type: "yaml",
                path: yamlPath,
                repository: { id: repoId, type: "azureReposGit", defaultBranch: branch }
            }
        }, { headers: { 'Authorization': authHeader, 'Content-Type': 'application/json' } });
        res.json(response.data);
    } catch (e) { res.status(500).json(e.response?.data || e.message); }
});
// Add this above your /api/pipelines/create route
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`,
            { 
                headers: { 
                    'Authorization': authHeader // Uses your DEVOPS_PAT
                } 
            }
        );
        // Azure DevOps returns the list in the .value property
        res.json(response.data.value); 
    } catch (error) {
        console.error("Error fetching repos:", error.response?.data || error.message);
        res.status(500).json({ error: "Failed to fetch repositories" });
    }
});

// Update the creation route to use process.env for Org/Project
app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, yamlPath } = req.body;
    try {
        const response = await axios.post(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`,
            {
                name: pipelineName,
                configuration: {
                    type: "yaml",
                    path: yamlPath || "/azure-pipelines.yml",
                    repository: { id: repoId, type: "azureReposGit" }
                }
            },
            { headers: { 'Authorization': authHeader, 'Content-Type': 'application/json' } }
        );
        res.json({ message: 'Pipeline Created!', details: response.data });
    } catch (error) {
        res.status(500).json({ error: 'Failed to create pipeline' });
    }
});
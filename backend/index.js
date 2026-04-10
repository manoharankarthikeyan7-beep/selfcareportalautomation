const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
const path = require('path');
const axios = require('axios');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

const getAdoHeader = () => {
    const pat = process.env.DEVOPS_PAT;
    return pat ? `Basic ${Buffer.from(`:${pat}`).toString('base64')}` : null;
};

// --- AUTHENTICATION (Unchanged) ---
const client = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/discovery/v2.0/keys`
});

function getKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

const validateToken = (req, res, next) => {
  const token = req.headers.authorization?.split(' ')[1];
  if (!token) return res.status(401).send('Unauthorized');

  jwt.verify(token, getKey, {
    audience: `api://${process.env.AZURE_API_CLIENT_ID}`,
    issuer: `https://sts.windows.net/${process.env.AZURE_TENANT_ID}/`,
    algorithms: ['RS256']
  }, (err, decoded) => {
    if (err) return res.status(403).send('Invalid Token');
    if (decoded.tid !== process.env.AZURE_TENANT_ID) return res.status(403).send('Unauthorized Tenant');
    req.user = decoded;
    next();
  });
};

// --- API ROUTES ---

// 1. Get Azure Repos (Unchanged)
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "ADO Unreachable" }); }
});

// NEW: 1b. Get GitHub Repos via ADO Service Connection
app.get('/api/github/repos', validateToken, async (req, res) => {
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/sourceProviders/github/repositories?serviceConnectionId=${process.env.GITHUB_SERVICE_CONNECTION_ID}&api-version=7.1-preview.1`;
        const response = await axios.get(url, { headers: { 'Authorization': getAdoHeader() } });
        // Map GitHub response to match our frontend expectation
        const formattedRepos = response.data.repositories.map(r => ({ id: r.id, name: r.name }));
        res.json(formattedRepos.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "GitHub Repos Unreachable" }); }
});

// 2. Get Branches (Updated to handle both)
app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).send("Error fetching branches"); }
});

// NEW: 2b. Get GitHub Branches
app.get('/api/github/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        // repoId for GitHub is usually "org/repo"
        const encodedRepoId = encodeURIComponent(req.params.repoId);
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/sourceProviders/github/repositories/${encodedRepoId}/branches?serviceConnectionId=${process.env.GITHUB_SERVICE_CONNECTION_ID}&api-version=7.1-preview.1`;
        const response = await axios.get(url, { headers: { 'Authorization': getAdoHeader() } });
        const formattedBranches = response.data.branches.map(b => ({ name: b.name }));
        res.json(formattedBranches);
    } catch (e) { res.status(500).send("Error fetching GitHub branches"); }
});

// 3. Discover YAML Files (Updated to handle both)
app.get('/api/:source/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const { branch } = req.query;
    const { source, repoId } = req.params;
    const version = branch.replace('refs/heads/', '');
    
    try {
        // NOTE: For simplicity, we assume Azure Repos logic. 
        // For GitHub, ADO doesn't have a direct "items" API like Git. 
        // You might need to rely on the "Manual Path" toggle for GitHub monorepos.
        if (source === "github") return res.json([]); 

        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${repoId}/items?recursionLevel=full&versionDescriptor.version=${version}&api-version=7.1`;
        const response = await axios.get(url, { headers: { 'Authorization': getAdoHeader() } });
        const yamlFiles = response.data.value
            .filter(item => item.path.endsWith('.yml') || item.path.endsWith('.yaml'))
            .map(item => item.path);
        res.json(yamlFiles);
    } catch (e) { res.status(500).json({ error: "Could not fetch YAML files" }); }
});

// 5. Final Create + Optional Run (Updated with GitHub Support)
app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath, variables, runPipeline, sourceType } = req.body;
    
    try {
        let repoConfig;

        if (sourceType === "github") {
            repoConfig = {
                type: "github",
                id: repoId, // org/repo
                name: repoId,
                defaultBranch: branch,
                properties: {
                    connectedServiceId: process.env.GITHUB_SERVICE_CONNECTION_ID,
                    apiUrl: `https://api.github.com/repos/${repoId}`
                }
            };
        } else {
            repoConfig = { 
                id: repoId, 
                type: "azureReposGit", 
                defaultBranch: branch 
            };
        }

        // Create Pipeline Definition
        const createRes = await axios.post(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`,
            {
                name: pipelineName,
                configuration: {
                    type: "yaml", 
                    path: yamlPath,
                    repository: repoConfig
                },
                variables: variables 
            },
            { headers: { 'Authorization': getAdoHeader(), 'Content-Type': 'application/json' } }
        );

        if (runPipeline) {
            const pipelineId = createRes.data.id;
            await axios.post(
                `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines/${pipelineId}/runs?api-version=7.0`,
                { resources: { repositories: { self: { refName: branch } } } },
                { headers: { 'Authorization': getAdoHeader() } }
            );
        }
        res.json(createRes.data);
    } catch (e) { 
        console.error("Creation Error:", e.response?.data);
        res.status(500).json(e.response?.data || "Operation failed"); 
    }
});

// --- Static Files ---
app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Server online on port ${PORT}`));
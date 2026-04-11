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

// --- AUTHENTICATION ---
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

// Step 1: Get Repositories (Azure)
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "ADO Unreachable" }); }
});

// Step 1.1: Get Repositories (GitHub via Service Connection)
app.get('/api/github/repos', validateToken, async (req, res) => {
    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, 
            { dataSourceDetails: { dataSourceName: "Repos" } },
            { headers: { 'Authorization': getAdoHeader() } }
        );
        const formatted = response.data.value.map(r => ({ id: r.name, name: r.name }));
        res.json(formatted.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "GitHub Repos Unreachable" }); }
});

// Step 2: Get Branches (Azure)
app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).send("Error fetching branches"); }
});

// Step 2.1: Get Branches (GitHub)
app.get('/api/github/repos/:repoId/branches', validateToken, async (req, res) => {
    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, 
            { dataSourceDetails: { dataSourceName: "Branches", parameters: { repository: req.params.repoId } } },
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).send("Error fetching GitHub branches"); }
});

// Step 3: Discover YAML Files (Azure)
app.get('/api/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const { branch } = req.query;
    const version = branch.replace('refs/heads/', '');
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/items?recursionLevel=full&versionDescriptor.version=${version}&api-version=7.1`;
        const response = await axios.get(url, { headers: { 'Authorization': getAdoHeader() } });
        const yamlFiles = response.data.value
            .filter(item => item.path.endsWith('.yml') || item.path.endsWith('.yaml'))
            .map(item => item.path);
        res.json(yamlFiles);
    } catch (e) { res.status(500).json({ error: "Could not fetch YAML files" }); }
});

// Step 4: Discover YAML Files (GitHub) --- ADDED LOGIC ---
app.get('/api/github/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    const { branch } = req.query;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        // Using the GitHub API via ADO Proxy to list files in the repo
        const response = await axios.post(url, 
            { 
                dataSourceDetails: { 
                    dataSourceName: "SearchRepositories", 
                    parameters: { 
                        repository: req.params.repoId,
                        branch: branch 
                    } 
                } 
            },
            { headers: { 'Authorization': getAdoHeader() } }
        );
        
        // Filter for YAML files from the GitHub response
        const yamlFiles = response.data.value
            .filter(file => file.path.endsWith('.yml') || file.path.endsWith('.yaml'))
            .map(file => file.path);
            
        res.json(yamlFiles);
    } catch (e) { 
        // Fallback for GitHub discovery as it can be flaky via proxy
        res.json([]); 
    }
});

// Step 5: Final Create Pipeline (Supports both)
app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath, variables, sourceType } = req.body;
    const isGitHub = sourceType === "github";
    
    try {
        const payload = {
            name: pipelineName,
            configuration: {
                type: "yaml",
                path: yamlPath,
                repository: {
                    id: repoId,
                    type: isGitHub ? "github" : "azureReposGit",
                    defaultBranch: branch,
                    connection: isGitHub ? { id: process.env.GITHUB_SERVICE_CONNECTION_ID } : undefined
                }
            },
            variables: variables 
        };

        const createRes = await axios.post(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`,
            payload,
            { headers: { 'Authorization': getAdoHeader(), 'Content-Type': 'application/json' } }
        );
        res.json(createRes.data);
    } catch (e) { res.status(500).json(e.response?.data || "Operation failed"); }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Server online on port ${PORT}`));
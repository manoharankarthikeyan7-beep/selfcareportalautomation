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

const getGitHubHeader = () => {
    return process.env.GITHUB_TOKEN ? `token ${process.env.GITHUB_TOKEN}` : null;
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

app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "ADO Unreachable" }); }
});

app.get('/api/github/repos', validateToken, async (req, res) => {
    const ghToken = getGitHubHeader();
    if (ghToken) {
        try {
            const response = await axios.get('https://api.github.com/user/repos?per_page=100', {
                headers: { 'Authorization': ghToken, 'Accept': 'application/vnd.github.v3+json' }
            });
            return res.json(response.data.map(r => ({ id: r.full_name, name: r.full_name })).sort((a, b) => a.name.localeCompare(b.name)));
        } catch (e) { console.error("Direct GH Error"); }
    }

    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, { dataSourceDetails: { dataSourceName: "Repos" } }, { headers: { 'Authorization': getAdoHeader() } });
        const formatted = response.data.value.map(r => ({ id: r.name, name: r.name }));
        res.json(formatted.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "GitHub Repos Unreachable" }); }
});

app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).json({ error: "Error fetching branches" }); }
});

app.get('/api/github/repos/:repoId/branches', validateToken, async (req, res) => {
    const ghToken = getGitHubHeader();
    const repoPath = decodeURIComponent(req.params.repoId); 
    
    console.log(`[Backend] Fetching branches for GitHub repo: ${repoPath}`);

    if (ghToken && repoPath.includes('/')) {
        try {
            const response = await axios.get(`https://api.github.com/repos/${repoPath}/branches`, {
                headers: { 'Authorization': ghToken, 'Accept': 'application/vnd.github.v3+json' }
            });
            return res.json(response.data.map(b => ({ name: b.name })));
        } catch (e) { console.error("Direct GH Branch Error, trying proxy..."); }
    }

    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, 
            { dataSourceDetails: { dataSourceName: "Branches", parameters: { "repository": repoPath } } },
            { headers: { 'Authorization': getAdoHeader() } }
        );
        
        const rawData = response.data.value || [];
        const formattedBranches = rawData.map(b => {
            if (typeof b === 'string') return { name: b };
            return { name: b.name || b.displayValue || "Unknown Branch" };
        });
        res.json(formattedBranches);
    } catch (e) { 
        console.error("Proxy Branch Error:", e.response?.data || e.message);
        res.status(502).json({ error: "GitHub branches unreachable via proxy" }); 
    }
});

app.get('/api/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const { branch } = req.query;
    const version = branch.replace('refs/heads/', '');
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/items?recursionLevel=full&versionDescriptor.version=${version}&api-version=7.1`;
        const response = await axios.get(url, { headers: { 'Authorization': getAdoHeader() } });
        res.json(response.data.value.filter(i => i.path.endsWith('.yml') || i.path.endsWith('.yaml')).map(i => i.path));
    } catch (e) { res.status(500).json([]); }
});

app.get('/api/github/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const repoId = decodeURIComponent(req.params.repoId);
    const { branch } = req.query;
    const ghToken = getGitHubHeader();

    if (ghToken && repoId.includes('/')) {
        try {
            const url = `https://api.github.com/repos/${repoId}/git/trees/${branch}?recursive=1`;
            const response = await axios.get(url, { headers: { 'Authorization': ghToken, 'Accept': 'application/vnd.github.v3+json' } });
            const files = response.data.tree
                .filter(f => f.type === "blob" && (f.path.endsWith('.yml') || f.path.endsWith('.yaml')))
                .map(f => "/" + f.path);
            return res.json(files);
        } catch (e) { console.error("GitHub Tree Error"); }
    }

    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, {
            dataSourceDetails: { dataSourceName: "FileContents", parameters: { "repository": repoId, "sha": branch } }
        }, { headers: { 'Authorization': getAdoHeader() } });
        const files = (response.data.value || [])
            .filter(f => f.path && (f.path.endsWith('.yml') || f.path.endsWith('.yaml')))
            .map(f => f.path.startsWith('/') ? f.path : "/" + f.path);
        return res.json(files);
    } catch (e) { 
        console.error("Proxy YAML Error:", e.message); 
        res.status(502).json([]); 
    }
});

app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath, sourceType } = req.body;
    const isGitHub = sourceType === "github";
    const formattedPath = yamlPath.startsWith('/') ? yamlPath : `/${yamlPath}`;

    try {
        const payload = {
            name: pipelineName,
            folder: "\\",
            configuration: {
                type: "yaml",
                path: formattedPath,
                repository: {
                    id: repoId,
                    type: isGitHub ? "github" : "azureReposGit",
                    name: repoId,
                    defaultBranch: branch,
                    connection: isGitHub ? { id: process.env.GITHUB_SERVICE_CONNECTION_ID } : undefined
                }
            }
        };

        const createRes = await axios.post(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`,
            payload,
            { headers: { 'Authorization': getAdoHeader(), 'Content-Type': 'application/json' } }
        );
        res.json(createRes.data);
    } catch (e) { 
        const errorDetail = e.response?.data?.message || e.response?.data || e.message;
        console.error("Azure DevOps Pipeline Error:", errorDetail);
        res.status(500).json({ error: "Operation failed", details: errorDetail }); 
    }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Server online on port ${PORT}`));
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

// New Helper for Direct GitHub Access
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

// Step 1: Azure Repos
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "ADO Unreachable" }); }
});

// Step 1.1: GitHub Repos (Direct + Proxy Fallback)
app.get('/api/github/repos', validateToken, async (req, res) => {
    const ghToken = getGitHubHeader();
    if (ghToken) {
        try {
            const response = await axios.get('https://api.github.com/user/repos?per_page=100', {
                headers: { 
                    'Authorization': ghToken,
                    'Accept': 'application/vnd.github.v3+json'
                }
            });
            return res.json(response.data.map(r => ({ id: r.full_name, name: r.full_name })).sort((a, b) => a.name.localeCompare(b.name)));
        } catch (e) { console.error("Direct GH Error, falling back to Proxy..."); }
    }

    // Fallback to Service Connection Proxy
    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, { dataSourceDetails: { dataSourceName: "Repos" } }, { headers: { 'Authorization': getAdoHeader() } });
        const formatted = response.data.value.map(r => ({ id: r.name, name: r.name }));
        res.json(formatted.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "GitHub Repos Unreachable" }); }
});

// Step 2: Azure Branches
app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).send("Error fetching branches"); }
});

// Step 2.1: GitHub Branches (Direct + Proxy Fallback)
app.get('/api/github/repos/:repoId/branches', validateToken, async (req, res) => {
    const ghToken = getGitHubHeader();
    const repoPath = req.params.repoId; // e.g. "Owner/Repo"

    if (ghToken && repoPath.includes('/')) {
        try {
            const response = await axios.get(`https://api.github.com/repos/${repoPath}/branches`, {
                headers: { 
                    'Authorization': ghToken,
                    'Accept': 'application/vnd.github.v3+json' 
                }
            });
            return res.json(response.data.map(b => ({ name: b.name })));
        } catch (e) { console.error("Direct GH Branch Error, falling back..."); }
    }

    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, 
            { dataSourceDetails: { dataSourceName: "Branches", parameters: { repository: repoPath } } },
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value || []);
    } catch (e) { res.status(500).send("Error fetching GitHub branches"); }
});

// Step 3 & 4: YAML Discovery (Simplified)
app.get('/api/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const { branch } = req.
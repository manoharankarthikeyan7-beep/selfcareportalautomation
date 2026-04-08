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

// Helper for Azure DevOps Auth
const getAdoHeader = () => {
    const pat = process.env.DEVOPS_PAT;
    return pat ? `Basic ${Buffer.from(`:${pat}`).toString('base64')}` : null;
};

// --- AUTH MIDDLEWARE ---
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
    req.user = decoded;
    next();
  });
};

// --- API ROUTES FOR THE WIZARD ---

// Step 1: Get Repos
app.get('/api/repos', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories?api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// Step 2: Get Branches
app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).json({ error: e.message }); }
});

// Step 3: Create & Rename Pipeline
app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath } = req.body;
    try {
        const response = await axios.post(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/pipelines?api-version=7.0`,
            {
                name: pipelineName,
                configuration: {
                    type: "yaml",
                    path: yamlPath,
                    repository: { id: repoId, type: "azureReposGit", defaultBranch: branch }
                }
            },
            { headers: { 'Authorization': getAdoHeader(), 'Content-Type': 'application/json' } }
        );
        res.json(response.data);
    } catch (e) { res.status(500).json(e.response?.data || e.message); }
});

// --- SERVE STATIC FRONTEND ---
app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Backend online on port ${PORT}`));
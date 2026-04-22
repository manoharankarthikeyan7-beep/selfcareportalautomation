const express = require('express');
const cors = require('cors');
const helmet = require('helmet'); // ADDED: Security headers protection
const rateLimit = require('express-rate-limit'); // ADDED: DDoS/Spam protection
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
const path = require('path');
const axios = require('axios');
require('dotenv').config();

const app = express();

/**
 * SECURITY FIX: Updated Content Security Policy (CSP)
 * We must explicitly allow the browser to connect to Microsoft's login servers.
 * Without this, Helmet's default policy blocks the MSAL authentication handshake.
 */
app.use(
  helmet({
    contentSecurityPolicy: {
      directives: {
        "default-src": ["'self'"],
        "connect-src": [
          "'self'", 
          "https://login.microsoftonline.com", 
          "https://graph.microsoft.com"
        ],
        "script-src": ["'self'", "'unsafe-inline'", "https://login.microsoftonline.com"],
        "frame-src": ["'self'", "https://login.microsoftonline.com"],
        "img-src": ["'self'", "data:"],
        "style-src": ["'self'", "'unsafe-inline'"],
      },
    },
  })
);

// SECURITY FIX: Prevents a malicious actor from crashing your server by sending 
// a massive JSON file. Limits the request body to 10kb.
app.use(express.json({ limit: '10kb' })); 

// SECURITY FIX: Rate limiting prevents automated bots from spamming your 
// 'create pipeline' endpoint. Allows 100 requests per 15 mins.
const limiter = rateLimit({
    windowMs: 15 * 60 * 1000,
    max: 100,
    message: { error: "Too many requests from this IP, please try again later." }
});
app.use('/api/', limiter);

// SECURITY FIX: Hardened CORS. In production, change '*' to your specific UI URL.
app.use(cors({
    origin: process.env.ALLOWED_ORIGIN || '*', 
    methods: ['GET', 'POST']
}));

const getAdoHeader = () => {
    const pat = process.env.DEVOPS_PAT;
    return pat ? `Basic ${Buffer.from(`:${pat}`).toString('base64')}` : null;
};

const getGitHubHeader = () => {
    return process.env.GITHUB_TOKEN ? `token ${process.env.GITHUB_TOKEN}` : null;
};

// --- AUTHENTICATION ---
const client = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/discovery/v2.0/keys`,
  cache: true, // SECURITY FIX: Caching keys reduces calls to Microsoft, preventing rate limits.
  rateLimit: true
});

function getKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    if (err) return callback(err);
    callback(null, key.getPublicKey());
  });
}

const validateToken = (req, res, next) => {
  // SECURITY FIX: Strict check for 'Bearer' prefix to follow RFC standards.
  const authHeader = req.headers.authorization;
  if (!authHeader || !authHeader.startsWith('Bearer ')) return res.status(401).send('Unauthorized');
  
  const token = authHeader.split(' ')[1];

  jwt.verify(token, getKey, {
    audience: `api://${process.env.AZURE_API_CLIENT_ID}`,
    issuer: `https://sts.windows.net/${process.env.AZURE_TENANT_ID}/`,
    algorithms: ['RS256'],
    clockTolerance: 30 // SECURITY FIX: Allows 30s time drift between servers to prevent random 403 errors.
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
    } catch (e) { 
        // SECURITY FIX: Generic error message hides internal Azure URLs from the frontend.
        res.status(502).json({ error: "External Service Unavailable" }); 
    }
});

app.get('/api/github/repos', validateToken, async (req, res) => {
    const ghToken = getGitHubHeader();
    if (ghToken) {
        try {
            const response = await axios.get('https://api.github.com/user/repos?per_page=100', {
                headers: { 'Authorization': ghToken, 'Accept': 'application/vnd.github.v3+json' }
            });
            return res.json(response.data.map(r => ({ id: r.full_name, name: r.full_name })).sort((a, b) => a.name.localeCompare(b.name)));
        } catch (e) { console.error("GitHub Fetch Error Logged Privately"); }
    }

    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, { dataSourceDetails: { dataSourceName: "Repos" } }, { headers: { 'Authorization': getAdoHeader() } });
        const formatted = response.data.value.map(r => ({ id: r.name, name: r.name }));
        res.json(formatted.sort((a, b) => a.name.localeCompare(b.name)));
    } catch (e) { res.status(502).json({ error: "Source Control Unreachable" }); }
});

app.get('/api/repos/:repoId/branches', validateToken, async (req, res) => {
    try {
        const response = await axios.get(
            `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/git/repositories/${req.params.repoId}/refs?filter=heads/&api-version=7.1`,
            { headers: { 'Authorization': getAdoHeader() } }
        );
        res.json(response.data.value);
    } catch (e) { res.status(500).json({ error: "Failed to fetch branch data" }); }
});

app.get('/api/github/repos/:repoId/branches', validateToken, async (req, res) => {
    const ghToken = getGitHubHeader();
    const repoPath = decodeURIComponent(req.params.repoId); 
    
    if (ghToken && repoPath.includes('/')) {
        try {
            const response = await axios.get(`https://api.github.com/repos/${repoPath}/branches`, {
                headers: { 'Authorization': ghToken, 'Accept': 'application/vnd.github.v3+json' }
            });
            return res.json(response.data.map(b => ({ name: b.name })));
        } catch (e) { console.error("Logged: GitHub Branch Error"); }
    }

    const scId = process.env.GITHUB_SERVICE_CONNECTION_ID;
    try {
        const url = `https://dev.azure.com/${process.env.ADO_ORG_NAME}/${process.env.ADO_PROJECT_NAME}/_apis/serviceendpoint/proxy/execute?endpointId=${scId}&api-version=7.1-preview.1`;
        const response = await axios.post(url, 
            { dataSourceDetails: { dataSourceName: "Branches", parameters: { "repository": repoPath } } },
            { headers: { 'Authorization': getAdoHeader() } }
        );
        const rawData = response.data.value || [];
        res.json(rawData.map(b => ({ name: typeof b === 'string' ? b : (b.name || b.displayValue) })));
    } catch (e) { res.status(502).json({ error: "Branch Proxy Error" }); }
});

app.get('/api/repos/:repoId/yaml-files', validateToken, async (req, res) => {
    const { branch } = req.query;
    if (!branch) return res.status(400).json({ error: "Parameter 'branch' is required" });
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

    if (ghToken && repoId.includes('/') && branch) {
        try {
            const url = `https://api.github.com/repos/${repoId}/git/trees/${branch}?recursive=1`;
            const response = await axios.get(url, { headers: { 'Authorization': ghToken, 'Accept': 'application/vnd.github.v3+json' } });
            const files = response.data.tree
                .filter(f => f.type === "blob" && (f.path.endsWith('.yml') || f.path.endsWith('.yaml')))
                .map(f => "/" + f.path);
            return res.json(files);
        } catch (e) { console.error("GitHub Tree Error Logged"); }
    }
    res.json([]);
});

app.post('/api/pipelines/create', validateToken, async (req, res) => {
    const { pipelineName, repoId, branch, yamlPath, sourceType } = req.body;
    
    // SECURITY FIX: Basic Input validation to ensure required fields aren't empty
    // preventing the app from sending malformed requests to Azure.
    if (!pipelineName || !repoId || !branch || !yamlPath) {
        return res.status(400).json({ error: "Missing required pipeline fields" });
    }

    const isGitHub = sourceType === "github";
    const cleanBranch = branch.replace('refs/heads/', '');
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
                    name: repoId,
                    fullName: repoId,
                    type: isGitHub ? "github" : "azureReposGit",
                    defaultBranch: cleanBranch,
                    connection: isGitHub ? { id: process.env.GITHUB_SERVICE_CONNECTION_ID } : undefined,
                    properties: isGitHub ? {
                        apiUrl: `https://api.github.com/repos/${repoId}`,
                        branchesUrl: `https://api.github.com/repos/${repoId}/branches`,
                        cloneUrl: `https://github.com/${repoId}.git`,
                        connectedServiceId: process.env.GITHUB_SERVICE_CONNECTION_ID,
                        defaultBranch: cleanBranch,
                        fullName: repoId,
                        isPrivate: "true"
                    } : undefined
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
        const errorDetail = e.response?.data?.message || "Internal Service Error";
        console.error("Critical Failure:", errorDetail);
        res.status(500).json({ error: "Pipeline Creation Failed" }); 
    }
});

app.use(express.static(path.join(__dirname, 'public')));
app.get('*', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

const PORT = process.env.PORT || 8080;
app.listen(PORT, '0.0.0.0', () => console.log(`Secured server running on port ${PORT}`));
const express = require('express');
const cors = require('cors');
const jwt = require('jsonwebtoken');
const jwksClient = require('jwks-rsa');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// 1. SETUP AZURE AD KEY DISCOVERY
const client = jwksClient({
  jwksUri: `https://login.microsoftonline.com/${process.env.AZURE_TENANT_ID}/discovery/v2.0/keys`
});

function getKey(header, callback) {
  client.getSigningKey(header.kid, (err, key) => {
    const signingKey = key.getPublicKey();
    callback(null, signingKey);
  });
}

// 2. MIDDLEWARE: TOKEN VALIDATION (The Bouncer)
const validateToken = (req, res, next) => {
  const token = req.headers.authorization?.split(' ')[1];
  if (!token) return res.status(401).json({ error: 'No token provided' });

  jwt.verify(token, getKey, {
    audience: `api://${process.env.AZURE_API_CLIENT_ID}`,
    issuer: `https://sts.windows.net/${process.env.AZURE_TENANT_ID}/`,
    algorithms: ['RS256']
  }, (err, decoded) => {
    if (err) return res.status(403).json({ error: 'Invalid token' });
    req.user = decoded; 
    next();
  });
};

// 3. MIDDLEWARE: RBAC ENFORCEMENT
const authorizeRole = (allowedGroups) => {
  return (req, res, next) => {
    const userGroups = req.user.groups || [];
    // Check if the user's AD Group IDs match the required IDs in .env
    const hasAccess = allowedGroups.some(groupId => userGroups.includes(groupId));
    
    if (!hasAccess) return res.status(403).json({ error: 'Insufficient Permissions' });
    next();
  };
};

// 4. PROTECTED ROUTES
// Admin only: Delete
app.delete('/api/pipelines/:id', validateToken, authorizeRole([process.env.ADMIN_GROUP_ID]), (req, res) => {
    console.log(`Audit: ${req.user.upn} deleted pipeline ${req.params.id}`);
    res.json({ message: 'Pipeline deleted' });
});

// DevOps & Admin: Create
app.post('/api/pipelines/create', validateToken, authorizeRole([process.env.ADMIN_GROUP_ID, process.env.DEVOPS_GROUP_ID]), (req, res) => {
    res.json({ message: 'Pipeline creation triggered in Azure DevOps' });
});

// All Auth Users: View
app.get('/api/pipelines', validateToken, (req, res) => {
    res.json({ data: [] });
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`Backend secured on port ${PORT}`));
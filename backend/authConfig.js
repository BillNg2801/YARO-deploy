const path = require('path');
const fs = require('fs');
const msal = require('@azure/msal-node');
require('dotenv').config();

// WARNING: File-based token cache is only for local development.
// On Vercel, use /tmp (ephemeral - tokens lost on cold start).
const DATA_DIR = process.env.VERCEL
  ? '/tmp'
  : path.join(__dirname, '.data');
const CACHE_PATH = path.join(DATA_DIR, 'msal-cache.json');

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

const config = {
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    authority: 'https://login.microsoftonline.com/common',
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  },
};

const SCOPES = [
  'User.Read',
  'offline_access',
  'Mail.Send',
  'Mail.ReadWrite',
];

const pca = new msal.ConfidentialClientApplication(config);
const tokenCache = pca.getTokenCache();

// Load cache from disk at startup
if (fs.existsSync(CACHE_PATH)) {
  try {
    const cache = fs.readFileSync(CACHE_PATH, 'utf8');
    if (cache) {
      tokenCache.deserialize(cache);
    }
  } catch (err) {
    console.error('Failed to read MSAL cache from disk:', err.message);
  }
}

function persistCache() {
  try {
    const cache = tokenCache.serialize();
    fs.writeFileSync(CACHE_PATH, cache, 'utf8');
  } catch (err) {
    console.error('Failed to persist MSAL cache to disk:', err.message);
  }
}

module.exports = {
  pca,
  SCOPES,
  persistCache,
};


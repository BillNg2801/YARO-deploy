const path = require('path');
const fs = require('fs');
const mongoose = require('mongoose');
const msal = require('@azure/msal-node');
require('dotenv').config();

const CACHE_COLLECTION = 'msal_cache';
const CACHE_DOC_ID = 'tokens';

async function loadCacheFromMongo() {
  try {
    if (mongoose.connection.readyState !== 1) return null;
    const doc = await mongoose.connection.db.collection(CACHE_COLLECTION).findOne({ _id: CACHE_DOC_ID });
    return doc ? doc.data : null;
  } catch (err) {
    console.error('Failed to load MSAL cache from MongoDB:', err.message);
    return null;
  }
}

async function saveCacheToMongo(data) {
  try {
    if (mongoose.connection.readyState !== 1) {
      console.warn('MongoDB not connected; could not persist MSAL cache.');
      return;
    }
    await mongoose.connection.db.collection(CACHE_COLLECTION).updateOne(
      { _id: CACHE_DOC_ID },
      { $set: { data, updatedAt: new Date() } },
      { upsert: true }
    );
  } catch (err) {
    console.error('Failed to save MSAL cache to MongoDB:', err.message);
  }
}

// MongoDB cache plugin - tokens persist across restarts and Vercel cold starts
const mongoCachePlugin = {
  async beforeCacheAccess(cacheContext) {
    const data = await loadCacheFromMongo();
    if (data) {
      cacheContext.tokenCache.deserialize(data);
    }
  },
  async afterCacheAccess(cacheContext) {
    if (cacheContext.cacheHasChanged) {
      await saveCacheToMongo(cacheContext.tokenCache.serialize());
    }
  },
};

// File fallback for local dev startup (before first DB access)
const DATA_DIR = path.join(__dirname, '.data');
const CACHE_PATH = path.join(DATA_DIR, 'msal-cache.json');
if (!process.env.VERCEL && !fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

const config = {
  auth: {
    clientId: process.env.MICROSOFT_CLIENT_ID,
    authority: 'https://login.microsoftonline.com/common',
    clientSecret: process.env.MICROSOFT_CLIENT_SECRET,
  },
  cache: {
    cachePlugin: mongoCachePlugin,
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

// Load from file on local startup (fallback before first DB access)
if (!process.env.VERCEL && fs.existsSync(CACHE_PATH)) {
  try {
    const cache = fs.readFileSync(CACHE_PATH, 'utf8');
    if (cache) {
      tokenCache.deserialize(cache);
    }
  } catch (err) {
    console.error('Failed to read MSAL cache from disk:', err.message);
  }
}

// No-op for backward compatibility; MongoDB plugin handles persistence
function persistCache() {}

module.exports = {
  pca,
  SCOPES,
  persistCache,
};

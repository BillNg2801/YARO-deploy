const fetch = require('node-fetch');
const { pca, SCOPES, persistCache } = require('./authConfig');

async function getAccessToken() {
  const tokenCache = pca.getTokenCache();
  const accounts = await tokenCache.getAllAccounts();

  if (!accounts || accounts.length === 0) {
    const error = new Error(
      'No Microsoft account found in token cache. Visit /auth/microsoft/start to connect.'
    );
    error.code = 'NO_ACCOUNT';
    throw error;
  }

  const account = accounts[0];

  const result = await pca.acquireTokenSilent({
    scopes: SCOPES,
    account,
  });

  persistCache();

  return result.accessToken;
}

async function graphFetch(path, options = {}) {
  const accessToken = await getAccessToken();
  const url = `https://graph.microsoft.com/v1.0${path}`;

  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });

  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph API error ${response.status}: ${text}`);
  }

  const text = await response.text();
  if (!text) return null;
  return JSON.parse(text);
}

module.exports = {
  getAccessToken,
  graphFetch,
};


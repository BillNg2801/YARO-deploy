const express = require('express');
const path = require('path');
const mongoose = require('mongoose');
require('dotenv').config();

const { pca, SCOPES } = require('./authConfig');
const { graphFetch } = require('./graphClient');
const { startTelegramBot } = require('./telegramBot');
const { connectDB } = require('./db');

const app = express();
const PORT = process.env.PORT || 5000;

app.use(express.json());

// Vercel: ensure DB connects on first request
if (process.env.VERCEL) {
  app.use(async (req, res, next) => {
    if (mongoose.connection.readyState !== 1) {
      await connectDB();
    }
    next();
  });
}

// --- Auth routes ---

app.get('/auth/microsoft/start', async (req, res) => {
  try {
    const authCodeUrlParameters = {
      scopes: SCOPES,
      redirectUri: process.env.MICROSOFT_REDIRECT_URI,
    };

    const authUrl = await pca.getAuthCodeUrl(authCodeUrlParameters);
    res.redirect(authUrl);
  } catch (err) {
    console.error('Error generating auth URL:', err);
    res.status(500).send('Failed to generate Microsoft auth URL.');
  }
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;

  if (!code) {
    return res.status(400).send('Missing authorization code.');
  }

  try {
    await pca.acquireTokenByCode({
      code,
      scopes: SCOPES,
      redirectUri: process.env.MICROSOFT_REDIRECT_URI,
    });

    res.redirect('/success');
  } catch (err) {
    console.error('Error exchanging code for tokens:', err);
    res
      .status(500)
      .send('Failed to acquire tokens from Microsoft. Check server logs.');
  }
});

// --- Success page ---

app.get('/success', (req, res) => {
  res.send(`
    <html>
      <head>
        <title>Microsoft Graph Connected</title>
      </head>
      <body>
        <h1>Connected to Microsoft Graph</h1>
        <p>Your account is now linked.</p>
        <ul>
          <li><a href="/api/me" target="_blank">GET /api/me</a></li>
          <li><a href="/api/inbox?top=10" target="_blank">GET /api/inbox?top=10</a></li>
        </ul>
      </body>
    </html>
  `);
});

// --- Test API routes ---

app.get('/api/me', async (req, res) => {
  try {
    const me = await graphFetch('/me');
    res.json(me);
  } catch (err) {
    console.error('Error calling /me:', err);
    if (err.code === 'NO_ACCOUNT') {
      return res.status(400).json({
        error: 'not_connected',
        message: 'No Microsoft account connected. Visit /auth/microsoft/start.',
      });
    }
    res.status(500).json({ error: 'graph_error', message: err.message });
  }
});

app.get('/api/inbox', async (req, res) => {
  const top = parseInt(req.query.top, 10) || 10;

  const query = `/me/mailFolders/Inbox/messages?$top=${encodeURIComponent(
    top
  )}&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime,bodyPreview`;

  try {
    const messages = await graphFetch(query);
    res.json(messages);
  } catch (err) {
    console.error('Error calling inbox:', err);
    if (err.code === 'NO_ACCOUNT') {
      return res.status(400).json({
        error: 'not_connected',
        message: 'No Microsoft account connected. Visit /auth/microsoft/start.',
      });
    }
    res.status(500).json({ error: 'graph_error', message: err.message });
  }
});

app.post('/api/send-test', async (req, res) => {
  const { to, subject, content } = req.body || {};

  if (!to || !subject || !content) {
    return res.status(400).json({
      error: 'invalid_request',
      message: 'Body must include to, subject, and content fields.',
    });
  }

  const payload = {
    message: {
      subject,
      body: {
        contentType: 'Text',
        content,
      },
      toRecipients: [
        {
          emailAddress: {
            address: to,
          },
        },
      ],
    },
    saveToSentItems: 'true',
  };

  try {
    await graphFetch('/me/sendMail', {
      method: 'POST',
      body: JSON.stringify(payload),
    });

    res.json({ success: true });
  } catch (err) {
    console.error('Error sending test mail:', err);
    if (err.code === 'NO_ACCOUNT') {
      return res.status(400).json({
        error: 'not_connected',
        message: 'No Microsoft account connected. Visit /auth/microsoft/start.',
      });
    }
    res.status(500).json({ error: 'graph_error', message: err.message });
  }
});

// --- Root helper ---

app.get('/', (req, res) => {
  res.send(
    'Microsoft Graph OAuth puller is running. Start at /auth/microsoft/start.'
  );
});

// Vercel: export app for serverless; local: run listen
if (process.env.VERCEL) {
  module.exports = app;
} else {
  // Local - start server and Telegram bot
  app.listen(PORT, async () => {
    console.log(`Server listening on http://localhost:${PORT}`);
    await connectDB();
    startTelegramBot();
  });
}


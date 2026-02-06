const express = require('express');
const path = require('path');
const mongoose = require('mongoose');
require('dotenv').config();

const { pca, SCOPES } = require('./authConfig');
const { graphFetch } = require('./graphClient');
const { startTelegramBot } = require('./telegramBot');
const { connectDB } = require('./db');
const { handleMailNotification } = require('./webhook/mailHandler');
const {
  createMailSubscription,
  renewExpiringSubscriptions,
} = require('./subscription');
const {
  expandDraftWithOpenAI,
  applyEditWithOpenAI,
  sendReplyViaGraph,
  ensureFormattedDraft,
} = require('./replyHandler');

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

app.get('/success', async (req, res) => {
  try {
    await renewExpiringSubscriptions();
  } catch (e) {
    /* ignore */
  }
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
          <li><a href="/api/subscribe">Enable email notifications (one-time)</a></li>
          <li><a href="/api/setup/telegram-webhook">Set Telegram webhook (one-time)</a></li>
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

// --- Mail webhook (Microsoft Graph) ---

app.post('/api/webhook/mail', async (req, res) => {
  const validationToken = req.query.validationToken;
  if (validationToken) {
    return res.type('text/plain').send(validationToken);
  }
  try {
    await handleMailNotification(req.body);
  } catch (err) {
    console.error('Mail webhook processing error:', err);
  }
  res.status(202).send();
});

// --- Subscription creation ---

app.get('/api/subscribe', async (req, res) => {
  try {
    const result = await createMailSubscription();
    res.send(`
      <html>
        <head><title>Subscription Created</title></head>
        <body>
          <h1>Mail subscription created</h1>
          <p>New emails will trigger notifications. Subscription ID: ${result?.id || 'N/A'}</p>
          <p><a href="/success">Back to success page</a></p>
        </body>
      </html>
    `);
  } catch (err) {
    console.error('Subscribe error:', err);
    if (err.code === 'NO_ACCOUNT') {
      return res.status(400).send(
        'No Microsoft account connected. <a href="/auth/microsoft/start">Sign in first</a>.'
      );
    }
    res.status(500).send(`Error: ${err.message}`);
  }
});

// --- Subscription renewal (call periodically) ---

app.get('/api/cron/renew-subscriptions', async (req, res) => {
  try {
    await renewExpiringSubscriptions();
    res.json({ ok: true });
  } catch (err) {
    console.error('Renew error:', err);
    res.status(500).json({ error: err.message });
  }
});

// --- Telegram webhook (receives updates from Telegram) ---

async function sendTelegramReply(chatId, text) {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) return;
  const fetch = require('node-fetch');
  await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chatId, text }),
  });
}

const EMAIL_VIEWS_COLLECTION = 'email_notification_views';
const REPLY_STATE_COLLECTION = 'telegram_reply_state';

async function ensureReplyStateTTL() {
  try {
    if (mongoose.connection.readyState !== 1 || !mongoose.connection.db) return;
    const col = mongoose.connection.db.collection(REPLY_STATE_COLLECTION);
    await col.createIndex({ updatedAt: 1 }, { expireAfterSeconds: 2 * 60 * 60 });
  } catch (err) {
    if (err.code !== 85 && err.code !== 86) console.error('reply_state TTL:', err.message);
  }
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

async function sendTelegramMessage(chatId, text, options = {}) {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) return;
  const fetch = require('node-fetch');
  const body = { chat_id: chatId, text };
  if (options.parse_mode) body.parse_mode = options.parse_mode;
  if (options.reply_markup) body.reply_markup = options.reply_markup;
  const resp = await fetch(`https://api.telegram.org/bot${token}/sendMessage`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!resp.ok) console.error('Telegram sendMessage failed:', await resp.text());
}

async function telegramEditMessageText(chatId, messageId, text, options = {}) {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) return;
  const fetch = require('node-fetch');
  const body = { chat_id: chatId, message_id: messageId, text };
  if (options.parse_mode !== undefined) body.parse_mode = options.parse_mode;
  if (options.reply_markup) body.reply_markup = options.reply_markup;
  const response = await fetch(`https://api.telegram.org/bot${token}/editMessageText`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const errBody = await response.text();
    console.error('Telegram editMessageText failed:', response.status, errBody);
  }
}

async function telegramAnswerCallbackQuery(callbackQueryId) {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) return;
  const fetch = require('node-fetch');
  await fetch(`https://api.telegram.org/bot${token}/answerCallbackQuery`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ callback_query_id: callbackQueryId }),
  });
}

app.post('/api/telegram/webhook', express.json(), async (req, res) => {
  const update = req.body;
  const callbackQuery = update?.callback_query;

  try {
    if (callbackQuery) {
      const data = callbackQuery.data || '';
      const chatId = callbackQuery.message?.chat?.id;
      const messageId = callbackQuery.message?.message_id;
      const callbackQueryId = callbackQuery.id;
      if (!chatId || !messageId) {
        res.status(200).send();
        return;
      }
      if (mongoose.connection.readyState !== 1) await connectDB();
      if (!mongoose.connection.db) {
        console.error('Telegram callback: db not available');
      }
      const viewsCol = mongoose.connection.db?.collection(EMAIL_VIEWS_COLLECTION);
      const replyStateCol = mongoose.connection.db?.collection(REPLY_STATE_COLLECTION);
      ensureReplyStateTTL().catch(() => {});

      const uuid =
        data.startsWith('view_full:') ? data.slice('view_full:'.length)
        : data.startsWith('view_summary:') ? data.slice('view_summary:'.length)
        : data.startsWith('reply_start:') ? data.slice('reply_start:'.length)
        : data.startsWith('reply_back:') ? data.slice('reply_back:'.length)
        : data.startsWith('reply_send:') ? data.slice('reply_send:'.length)
        : data.startsWith('reply_edit:') ? data.slice('reply_edit:'.length)
        : data.startsWith('reply_cancel_edit:') ? data.slice('reply_cancel_edit:'.length)
        : null;

      if (uuid && viewsCol) {
        const doc = await viewsCol.findOne({ _id: uuid });
        if (!doc) {
          console.error('Telegram callback: view not found for uuid', uuid);
        } else if (data.startsWith('view_full:')) {
          const fullText = doc.fullText || '';
          const options = {
            reply_markup: {
              inline_keyboard: [[{ text: 'Go back to the summary', callback_data: `view_summary:${uuid}` }]],
            },
          };
          if (fullText.trim().startsWith('<b>')) options.parse_mode = 'HTML';
          await telegramEditMessageText(chatId, messageId, fullText, options);
        } else if (data.startsWith('view_summary:')) {
          await telegramEditMessageText(chatId, messageId, doc.summaryText, {
            parse_mode: 'HTML',
            reply_markup: {
              inline_keyboard: [
                [{ text: 'See the full email', callback_data: `view_full:${uuid}` }],
                [{ text: 'REPLY', callback_data: `reply_start:${uuid}` }],
              ],
            },
          });
        } else if (data.startsWith('reply_start:')) {
          const replyPrompt =
            (doc.summaryText || '') +
            '\n\nWhat would you like to say to ' +
            (doc.senderName || 'them') +
            '?';
          await replyStateCol.updateOne(
            { _id: chatId },
            { $set: { viewUuid: uuid, mode: 'awaiting_reply', updatedAt: new Date() } },
            { upsert: true }
          );
          await telegramEditMessageText(chatId, messageId, replyPrompt, {
            parse_mode: 'HTML',
            reply_markup: {
              inline_keyboard: [[{ text: 'Back', callback_data: `reply_back:${uuid}` }]],
            },
          });
        } else if (data.startsWith('reply_back:')) {
          await replyStateCol.deleteOne({ _id: chatId });
          await telegramEditMessageText(chatId, messageId, doc.summaryText, {
            parse_mode: 'HTML',
            reply_markup: {
              inline_keyboard: [
                [{ text: 'See the full email', callback_data: `view_full:${uuid}` }],
                [{ text: 'REPLY', callback_data: `reply_start:${uuid}` }],
              ],
            },
          });
        } else if (data.startsWith('reply_send:')) {
          const state = await replyStateCol.findOne({ _id: chatId });
          if (!state || state.viewUuid !== uuid || !state.draft) {
            await sendTelegramMessage(chatId, 'No draft found. Please start the reply flow again.');
          } else {
            try {
              await sendReplyViaGraph(doc.graphMessageId, state.draft);
              await replyStateCol.deleteOne({ _id: chatId });
              await sendTelegramMessage(chatId, 'Your email has been successfully delivered.');
            } catch (err) {
              console.error('Send reply failed:', err);
              await sendTelegramMessage(
                chatId,
                'Failed to send the email. Please try again or check your Outlook connection.'
              );
            }
          }
        } else if (data.startsWith('reply_edit:')) {
          const state = await replyStateCol.findOne({ _id: chatId });
          if (!state || state.viewUuid !== uuid || !state.draft) {
            await sendTelegramMessage(chatId, 'No draft found. Please start the reply flow again.');
          } else {
            const editPrompt = escapeHtml(state.draft) + '\n\n<b>What would you like to change?</b>';
            await replyStateCol.updateOne(
              { _id: chatId },
              { $set: { mode: 'awaiting_edit_feedback', telegramMessageId: messageId, updatedAt: new Date() } },
              { upsert: true }
            );
            await telegramEditMessageText(chatId, messageId, editPrompt, {
              parse_mode: 'HTML',
              reply_markup: {
                inline_keyboard: [[{ text: 'Back', callback_data: `reply_cancel_edit:${uuid}` }]],
              },
            });
          }
        } else if (data.startsWith('reply_cancel_edit:')) {
          const state = await replyStateCol.findOne({ _id: chatId });
          if (!state || state.viewUuid !== uuid || !state.draft) {
            await sendTelegramMessage(chatId, 'No draft found. Please start the reply flow again.');
          } else {
            await replyStateCol.updateOne(
              { _id: chatId },
              { $set: { mode: 'awaiting_send_edit', updatedAt: new Date() } },
              { upsert: true }
            );
            const draftText = ensureFormattedDraft(state.draft);
            await telegramEditMessageText(chatId, messageId, draftText, {
              reply_markup: {
                inline_keyboard: [
                  [
                    { text: 'Send', callback_data: `reply_send:${uuid}` },
                    { text: 'Edit', callback_data: `reply_edit:${uuid}` },
                  ],
                ],
              },
            });
          }
        }
      }
      await telegramAnswerCallbackQuery(callbackQueryId);
      res.status(200).send();
      return;
    }

    const text = update?.message?.text || '';
    const chatId = update?.message?.chat?.id;

    if (!chatId) {
      res.status(200).send();
      return;
    }
    if (mongoose.connection.readyState !== 1) await connectDB();

    const replyStateCol = mongoose.connection.db?.collection(REPLY_STATE_COLLECTION);
    const viewsCol = mongoose.connection.db?.collection(EMAIL_VIEWS_COLLECTION);
    const state = replyStateCol ? await replyStateCol.findOne({ _id: chatId }) : null;

    if (state && viewsCol && text && !text.startsWith('/')) {
      const viewDoc = await viewsCol.findOne({ _id: state.viewUuid });
      if (!viewDoc) {
        await replyStateCol.deleteOne({ _id: chatId });
      } else if (state.mode === 'awaiting_reply') {
        try {
          const draft = await expandDraftWithOpenAI(text, viewDoc.senderName);
          await replyStateCol.updateOne(
            { _id: chatId },
            { $set: { draft, mode: 'awaiting_send_edit', updatedAt: new Date() } },
            { upsert: true }
          );
          await sendTelegramMessage(chatId, draft, {
            reply_markup: {
              inline_keyboard: [
                [
                  { text: 'Send', callback_data: `reply_send:${state.viewUuid}` },
                  { text: 'Edit', callback_data: `reply_edit:${state.viewUuid}` },
                ],
              ],
            },
          });
        } catch (err) {
          console.error('expandDraftWithOpenAI failed:', err);
          await sendTelegramMessage(
            chatId,
            'Failed to generate the email. Please try again or check that OPENAI_API_KEY is set.'
          );
        }
        res.status(200).send();
        return;
      } else if (state.mode === 'awaiting_edit_feedback') {
        try {
          const revised = await applyEditWithOpenAI(state.draft, text);
          await replyStateCol.updateOne(
            { _id: chatId },
            { $set: { draft: revised, mode: 'awaiting_send_edit', updatedAt: new Date() } },
            { upsert: true }
          );
          const replyMarkup = {
            reply_markup: {
              inline_keyboard: [
                [
                  { text: 'Send', callback_data: `reply_send:${state.viewUuid}` },
                  { text: 'Edit', callback_data: `reply_edit:${state.viewUuid}` },
                ],
              ],
            },
          };
          if (state.telegramMessageId != null) {
            await telegramEditMessageText(chatId, state.telegramMessageId, revised, replyMarkup);
          } else {
            await sendTelegramMessage(chatId, revised, replyMarkup);
          }
        } catch (err) {
          console.error('applyEditWithOpenAI failed:', err);
          await sendTelegramMessage(
            chatId,
            'Failed to apply your changes. Please try again.'
          );
        }
        res.status(200).send();
        return;
      }
    }

    const col = mongoose.connection.db.collection('telegram_subscribers');
    const doc = await col.findOne({ _id: 'subscribers' });
    const chatIds = doc?.chatIds || [];

    // /check - reassure registered users
    if (text.startsWith('/check')) {
      if (chatIds.includes(chatId)) {
        await sendTelegramReply(
          chatId,
          'Notifications set up is completed! You will receive email summaries in this chat.'
        );
      } else {
        await sendTelegramReply(
          chatId,
          "You're not registered yet. Send /start to register for email notifications."
        );
      }
      res.status(200).send();
      return;
    }

    // /start - register for notifications
    if (!text.startsWith('/start')) {
      res.status(200).send();
      return;
    }

    if (chatIds.includes(chatId)) {
      res.status(200).send();
      return;
    }
    if (chatIds.length >= 2) {
      await sendTelegramReply(
        chatId,
        'Limit reached. Only 2 users can receive notifications.'
      );
      res.status(200).send();
      return;
    }

    chatIds.push(chatId);
    await col.updateOne(
      { _id: 'subscribers' },
      { $set: { chatIds, updatedAt: new Date() } },
      { upsert: true }
    );

    await sendTelegramReply(chatId, 'Notifications set up is completed!');
    res.status(200).send();
  } catch (err) {
    console.error('Telegram webhook error:', err);
    res.status(200).send();
  }
});

// --- Telegram setWebhook setup (one-time) ---

app.get('/api/setup/telegram-webhook', async (req, res) => {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  const baseUrl = process.env.BASE_URL || 'https://yarodeploy.vercel.app';
  const url = `${baseUrl.replace(/\/$/, '')}/api/telegram/webhook`;

  if (!token) {
    return res.status(500).send('TELEGRAM_BOT_TOKEN not set');
  }

  try {
    const fetch = require('node-fetch');
    const resp = await fetch(
      `https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(url)}`
    );
    const data = await resp.json();
    if (data.ok) {
      res.send(`<h1>Telegram webhook set</h1><p>URL: ${url}</p>`);
    } else {
      res.status(500).send(`Telegram API error: ${JSON.stringify(data)}`);
    }
  } catch (err) {
    res.status(500).send(`Error: ${err.message}`);
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


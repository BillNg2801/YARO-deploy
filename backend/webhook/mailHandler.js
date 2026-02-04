const crypto = require('crypto');
const mongoose = require('mongoose');
const fetch = require('node-fetch');
const OpenAI = require('openai').default;
const { graphFetch } = require('../graphClient');
const { connectDB } = require('../db');

const TELEGRAM_SUBSCRIBERS_COLLECTION = 'telegram_subscribers';
const TELEGRAM_SUBSCRIBERS_DOC_ID = 'subscribers';
const EMAIL_VIEWS_COLLECTION = 'email_notification_views';
const MAX_SUBSCRIBERS = 2;
const TELEGRAM_MESSAGE_MAX_LENGTH = 4096;

async function getTelegramChatIds() {
  try {
    if (mongoose.connection.readyState !== 1) return [];
    const doc = await mongoose.connection.db
      .collection(TELEGRAM_SUBSCRIBERS_COLLECTION)
      .findOne({ _id: TELEGRAM_SUBSCRIBERS_DOC_ID });
    return doc?.chatIds || [];
  } catch (err) {
    console.error('Failed to load Telegram subscribers:', err.message);
    return [];
  }
}

async function sendTelegramMessage(chatId, text, options = {}) {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) {
    console.warn('TELEGRAM_BOT_TOKEN not set; skipping Telegram send.');
    return;
  }
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const body = {
    chat_id: chatId,
    text,
    parse_mode: options.parse_mode !== undefined ? options.parse_mode : 'HTML',
  };
  if (options.reply_markup) body.reply_markup = options.reply_markup;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });
  if (!response.ok) {
    const errText = await response.text();
    console.error('Telegram send failed:', errText);
  }
}

function stripHtml(content) {
  if (!content) return '';
  return String(content).replace(/<[^>]*>/g, '');
}

const HTML_ENTITIES = {
  '&nbsp;': ' ',
  '&amp;': '&',
  '&lt;': '<',
  '&gt;': '>',
  '&quot;': '"',
  '&#39;': "'",
  '&apos;': "'",
};

function decodeHtmlEntities(content) {
  if (!content) return '';
  let s = String(content);
  for (const [entity, char] of Object.entries(HTML_ENTITIES)) {
    s = s.split(entity).join(char);
  }
  return s;
}

function normalizeWhitespace(content) {
  if (!content) return '';
  return String(content)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, ' ')
    .replace(/\n /g, '\n')
    .replace(/ \n/g, '\n')
    .trim()
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .join('\n');
}

function normalizeEmailBody(content) {
  return normalizeWhitespace(decodeHtmlEntities(content));
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function formatFullEmailBody(content) {
  if (!content) return '';
  let s = decodeHtmlEntities(content);
  s = s.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  s = s.replace(/\n{3,}/g, '\n\n').trim();
  return s
    .split('\n')
    .map((line) => line.replace(/\s+/g, ' ').trim())
    .filter(Boolean)
    .join('\n');
}

const SIGN_OFF_PATTERN = /(\n\s*)(Best|Regards|Sincerely|Thanks|Thank you|Cheers|Alright[^.]*\.|Take care)[,\s]*\s*[\s\S]*$/i;

function fallbackSummaryBlock(normalizedBody) {
  const lines = normalizedBody.split(/\n/).filter(Boolean);
  const greeting = lines[0]?.trim() || '';
  const rest = lines.slice(1).join(' ').trim().replace(SIGN_OFF_PATTERN, '');
  const snippet = rest.slice(0, 150).trim();
  if (!greeting) return snippet || '(No content)';
  return snippet ? `${greeting}\n\n${snippet}` : greeting;
}

async function rewriteEmailWithOpenAI(normalizedBody) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey || !normalizedBody.trim()) {
    return fallbackSummaryBlock(normalizedBody);
  }

  try {
    const openai = new OpenAI({ apiKey });
    const completion = await openai.chat.completions.create({
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: `Rewrite this email into exactly this format. Output ONLY:
1) One line: the greeting only (e.g. "Dear Yaroslavl," or "Hi,").
2) A blank line.
3) One or two sentences that summarize the main point of the email. Do not include any sign-off (no Best, Sincerely, Regards, Thanks, Alright, etc.). Do not include closings.

Email:
${normalizedBody}`,
        },
      ],
      max_tokens: 200,
    });
    const out = completion.choices?.[0]?.message?.content?.trim();
    return out || fallbackSummaryBlock(normalizedBody);
  } catch (err) {
    console.error('OpenAI rewrite failed:', err.message);
    return fallbackSummaryBlock(normalizedBody);
  }
}

async function isThread(conversationId) {
  if (!conversationId) return false;
  try {
    const escaped = String(conversationId).replace(/'/g, "''");
    const path = `/me/mailFolders/Inbox/messages?$filter=conversationId eq '${escaped}'&$top=2&$select=id`;
    const data = await graphFetch(path);
    const messages = data?.value || [];
    return messages.length >= 2;
  } catch (err) {
    console.error('Thread check failed:', err.message);
    return false;
  }
}

const PROCESSED_MAIL_IDS_COLLECTION = 'processed_mail_ids';

async function ensureProcessedMailIdsTTL() {
  try {
    if (mongoose.connection.readyState !== 1 || !mongoose.connection.db) return;
    const col = mongoose.connection.db.collection(PROCESSED_MAIL_IDS_COLLECTION);
    await col.createIndex({ processedAt: 1 }, { expireAfterSeconds: 30 * 24 * 60 * 60 });
  } catch (err) {
    if (err.code !== 85 && err.code !== 86) console.error('processed_mail_ids TTL index:', err.message);
  }
}

async function handleMailNotification(notification) {
  try {
    if (mongoose.connection.readyState !== 1) await connectDB();
    const value = notification?.value || [];
    if (value.length === 0) return;

    const rawIds = value.map((item) => item?.resourceData?.id).filter(Boolean);
    const messageIds = [...new Set(rawIds)];
    console.log('Mail webhook: value.length=%d, messageIds=%j', value.length, rawIds);
    ensureProcessedMailIdsTTL().catch(() => {});

    for (const messageId of messageIds) {
      try {
        await mongoose.connection.db.collection(PROCESSED_MAIL_IDS_COLLECTION).insertOne({
          _id: messageId,
          processedAt: new Date(),
        });
      } catch (insertErr) {
        if (insertErr.code === 11000) continue;
        throw insertErr;
      }

      let message;
      try {
        message = await graphFetch(
          `/me/messages/${messageId}?$select=from,body,conversationId`
        );
      } catch (err) {
        console.error('Failed to fetch message:', err.message);
        continue;
      }

      const from = message?.from?.emailAddress;
      const senderName =
        from?.name || from?.address || 'Unknown';
      const bodyObj = message?.body || {};
      const contentType = bodyObj.contentType || 'text';
      let content = bodyObj.content || '';
      if (contentType === 'html') content = stripHtml(content);
      const normalizedBody = normalizeEmailBody(content);

      const isInThread = await isThread(message.conversationId || '');
      const header = isInThread
        ? `A new email was sent from ${senderName} (thread).`
        : `A new email was sent from ${senderName}.`;

      let summaryBlock;
      if (!normalizedBody.trim()) {
        summaryBlock = '(No content)';
      } else if (normalizedBody.length <= 40 && normalizedBody.indexOf('\n') === -1) {
        summaryBlock = normalizedBody;
      } else {
        summaryBlock = await rewriteEmailWithOpenAI(normalizedBody);
      }

      const fullMessage = `<b>${header}</b>\n\n<b>📧 Email Summary:</b>\n\n${summaryBlock}`;
      const formattedBody = formatFullEmailBody(content);
      const bodyHtml = escapeHtml(formattedBody).replace(/\n/g, '<br>');
      const prefix = `<b>${escapeHtml(header)}</b><br><br><b>Full email:</b><br><br>`;
      let bodyPart = bodyHtml;
      const truncateSuffix = '... (truncated)';
      if (prefix.length + bodyPart.length + truncateSuffix.length > TELEGRAM_MESSAGE_MAX_LENGTH) {
        bodyPart = bodyPart.slice(0, TELEGRAM_MESSAGE_MAX_LENGTH - prefix.length - truncateSuffix.length) + truncateSuffix;
      }
      const fullText = prefix + bodyPart;

      const uuid = crypto.randomUUID();
      let replyMarkup = null;
      if (mongoose.connection.db) {
        const viewsCol = mongoose.connection.db.collection(EMAIL_VIEWS_COLLECTION);
        try {
          await viewsCol.createIndex(
            { createdAt: 1 },
            { expireAfterSeconds: 24 * 60 * 60 }
          );
        } catch (ttlErr) {
          if (ttlErr.code !== 85 && ttlErr.code !== 86) console.error('email_notification_views TTL:', ttlErr.message);
        }
        await viewsCol.insertOne({
          _id: uuid,
          summaryText: fullMessage,
          fullText,
          createdAt: new Date(),
        });
        replyMarkup = {
          inline_keyboard: [[{ text: 'See the full email', callback_data: `view_full:${uuid}` }]],
        };
      }

      const chatIds = await getTelegramChatIds();
      for (const chatId of chatIds) {
        await sendTelegramMessage(chatId, fullMessage, replyMarkup ? { reply_markup: replyMarkup } : {});
      }
    }
  } catch (err) {
    console.error('handleMailNotification error:', err);
  }
}

module.exports = { handleMailNotification };

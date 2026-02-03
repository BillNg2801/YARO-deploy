const mongoose = require('mongoose');
const fetch = require('node-fetch');
const OpenAI = require('openai').default;
const { graphFetch } = require('../graphClient');

const TELEGRAM_SUBSCRIBERS_COLLECTION = 'telegram_subscribers';
const TELEGRAM_SUBSCRIBERS_DOC_ID = 'subscribers';
const MAX_SUBSCRIBERS = 2;

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

async function sendTelegramMessage(chatId, text) {
  const token = process.env.TELEGRAM_BOT_TOKEN;
  if (!token) {
    console.warn('TELEGRAM_BOT_TOKEN not set; skipping Telegram send.');
    return;
  }
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ chat_id: chatId, text, parse_mode: 'HTML' }),
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

async function summarizeWithOpenAI(text) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey || !text.trim()) return text.slice(0, 200);

  try {
    const openai = new OpenAI({ apiKey });
    const completion = await openai.chat.completions.create({
      model: 'gpt-3.5-turbo',
      messages: [
        {
          role: 'user',
          content: `Summarize this email in 1-2 sentences. Do not include closings like Sincerely or Best regards. Output only the summary.\n\n${text}`,
        },
      ],
      max_tokens: 150,
    });
    const summary = completion.choices?.[0]?.message?.content?.trim();
    return summary || text.slice(0, 200);
  } catch (err) {
    console.error('OpenAI summarization failed:', err.message);
    return text.slice(0, 200);
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

async function handleMailNotification(notification) {
  try {
    const value = notification?.value || [];
    if (value.length === 0) return;

    for (const item of value) {
      const resourceData = item?.resourceData;
      if (!resourceData?.id) continue;

      const messageId = resourceData.id;

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

      const isInThread = await isThread(message.conversationId || '');
      const header = isInThread
        ? `A new email was sent from ${senderName} (thread).`
        : `A new email was sent from ${senderName}.`;

      const lines = content.split(/\r?\n/).filter(Boolean);
      const greeting = lines[0]?.trim() || '';
      const rest = lines.slice(1).join('\n').trim();

      let summary = rest;
      if (rest.length > 50) {
        summary = await summarizeWithOpenAI(rest);
      }
      const bodyText = greeting ? `${greeting}\n${summary}` : summary;

      const fullMessage = `${header}\n\n${bodyText}`;

      const chatIds = await getTelegramChatIds();
      for (const chatId of chatIds) {
        await sendTelegramMessage(chatId, fullMessage);
      }
    }
  } catch (err) {
    console.error('handleMailNotification error:', err);
  }
}

module.exports = { handleMailNotification };

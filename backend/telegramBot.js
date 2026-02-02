const TelegramBot = require('node-telegram-bot-api');
const { graphFetch } = require('./graphClient');

const token = process.env.TELEGRAM_BOT_TOKEN;
let bot = null;

function startTelegramBot() {
  if (!token) {
    console.warn('TELEGRAM_BOT_TOKEN not set; Telegram bot disabled.');
    return null;
  }

  bot = new TelegramBot(token, { polling: true });

  bot.onText(/\/start/, (msg) => {
    const chatId = msg.chat.id;
    bot.sendMessage(
      chatId,
      '👋 YARO OUTLOOK bot connected.\n\nCommands:\n/start - Show this message\n/inbox - Recent inbox count (last 30 days)\n/me - Your linked Outlook account'
    );
  });

  bot.onText(/\/me/, async (msg) => {
    const chatId = msg.chat.id;
    try {
      const me = await graphFetch('/me');
      const name = me.displayName || me.userPrincipalName || 'Unknown';
      const email = me.mail || me.userPrincipalName || '—';
      bot.sendMessage(chatId, `📧 Linked account:\n\nName: ${name}\nEmail: ${email}`);
    } catch (err) {
      const message =
        err.code === 'NO_ACCOUNT'
          ? 'No Microsoft account connected. Visit the app and sign in at /auth/microsoft/start'
          : `Error: ${err.message}`;
      bot.sendMessage(chatId, message);
    }
  });

  bot.onText(/\/inbox/, async (msg) => {
    const chatId = msg.chat.id;
    try {
      const top = 999;
      const query = `/me/mailFolders/Inbox/messages?$top=${top}&$orderby=receivedDateTime desc&$select=receivedDateTime`;
      const data = await graphFetch(query);
      const messages = data.value || [];
      const cutoff = new Date();
      cutoff.setDate(cutoff.getDate() - 30);
      const lastMonth = messages.filter((m) => new Date(m.receivedDateTime) >= cutoff);
      const count = lastMonth.length;
      bot.sendMessage(
        chatId,
        `📬 Inbox (last 30 days): ${count} emails\n(Shown from most recent ${top} messages.)`
      );
    } catch (err) {
      const message =
        err.code === 'NO_ACCOUNT'
          ? 'No Microsoft account connected. Visit the app and sign in.'
          : `Error: ${err.message}`;
      bot.sendMessage(chatId, message);
    }
  });

  bot.on('polling_error', (err) => {
    console.error('Telegram polling error:', err.message);
  });

  console.log('Telegram bot started (polling).');
  return bot;
}

module.exports = { startTelegramBot };

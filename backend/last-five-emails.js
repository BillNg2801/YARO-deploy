require('dotenv').config();
const { graphFetch } = require('./graphClient');

async function main() {
  const data = await graphFetch(
    '/me/mailFolders/Inbox/messages?$top=5&$orderby=receivedDateTime desc&$select=subject,from,receivedDateTime'
  );
  const messages = data.value || [];
  if (messages.length === 0) {
    console.log('No emails in inbox.');
    return;
  }
  console.log('Last 5 emails received:\n');
  messages.forEach((m, i) => {
    const from = m.from?.emailAddress?.name || m.from?.emailAddress?.address || '?';
    const subj = (m.subject || '(No subject)').slice(0, 60);
    const dt = m.receivedDateTime ? new Date(m.receivedDateTime).toLocaleString() : '—';
    console.log(`${i + 1}. ${subj}`);
    console.log(`   From: ${from}`);
    console.log(`   Date & time: ${dt}\n`);
  });
}

main().catch((err) => {
  console.error(err.code === 'NO_ACCOUNT' ? 'Not logged in. Run the app and visit /auth/microsoft/start first.' : err.message);
  process.exit(1);
});

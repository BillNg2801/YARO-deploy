require('dotenv').config();
const { connectDB } = require('./db');
const { graphFetch } = require('./graphClient');

async function main() {
  await connectDB();
  const data = await graphFetch(
    '/me/mailFolders/SentItems/messages?$top=1&$orderby=sentDateTime desc&$select=subject,body,toRecipients,sentDateTime'
  );
  const msg = data?.value?.[0];
  if (!msg) {
    console.log('No sent messages found.');
    return;
  }
  const to = (msg.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(', ');
  console.log('To:', to);
  console.log('Subject:', msg.subject);
  console.log('Sent:', msg.sentDateTime);
  console.log('\n--- Body ---\n');
  console.log((msg.body?.contentType === 'html' ? (msg.body?.content || '').replace(/<[^>]*>/g, ' ') : msg.body?.content) || '(empty)');
}

main().catch(e => { console.error(e); process.exit(1); });

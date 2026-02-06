const OpenAI = require('openai').default;
const { graphFetch } = require('./graphClient');

async function expandDraftWithOpenAI(userMessage, senderName) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY not set');
  }

  const openai = new OpenAI({ apiKey });
  const completion = await openai.chat.completions.create({
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'user',
        content: `You are a professional email assistant for Naked Car Studio. The user wants to reply to an email.
Convert their short message into a polite, respectful, professional email.

Rules:
- Start with "Dear [recipient name]," (e.g. Dear Nguyen Phan Anh or Dear Phan Anh). Put a comma only at the end of the line after the full name; do not put commas between parts of the name. Do not use "Hi".
- Use proper paragraph breaks (blank line between paragraphs)
- Always end with a two-line sign-off: first line a closing phrase (e.g. "Best regards,"), second line exactly "Naked Car Studio". The sender is always Naked Car Studio. Do not end the email with any punctuation (no period or other mark after "Naked Car Studio" or after the closing phrase).
- Output plain text only, well-formatted with double newlines between paragraphs
- Keep tone professional and friendly

User's message: ${userMessage}
Recipient name (for greeting): ${senderName || 'there'}`,
      },
    ],
    max_tokens: 500,
  });

  const draft = completion.choices?.[0]?.message?.content?.trim() || '';
  return ensureFormattedDraft(draft);
}

async function applyEditWithOpenAI(draft, feedback) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error('OPENAI_API_KEY not set');
  }

  const openai = new OpenAI({ apiKey });
  const completion = await openai.chat.completions.create({
    model: 'gpt-3.5-turbo',
    messages: [
      {
        role: 'user',
        content: `You are a professional email assistant. The user wants to modify this draft email.

Current draft:
${draft}

User's edit request: ${feedback}

Apply the changes. Start with "Dear [recipient name]," (not "Hi"); comma only at the end of the line after the full name, not between parts of the name. End with a two-line sign-off: first line a closing phrase (e.g. "Best regards,"), second line exactly "Naked Car Studio". Do not end the email with any punctuation (no period after "Naked Car Studio" or the closing phrase).
Output the revised email only, well-formatted with double newlines between paragraphs.`,
      },
    ],
    max_tokens: 800,
  });

  const revised = completion.choices?.[0]?.message?.content?.trim() || draft;
  return ensureFormattedDraft(revised);
}

function ensureFormattedDraft(text) {
  if (!text || typeof text !== 'string') return text;
  return text
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

function plainTextToHtml(plainText) {
  if (!plainText || typeof plainText !== 'string') return '';
  const trimmed = plainText.trim();
  if (!trimmed) return '<p></p>';
  const paragraphs = trimmed.split(/\n\n+/);
  return paragraphs
    .map((p) => p.trim())
    .filter(Boolean)
    .map((p) => '<p>' + escapeHtml(p).replace(/\n/g, '<br>') + '</p>')
    .join('\n');
}

async function sendReplyViaGraph(graphMessageId, body) {
  const formatted = ensureFormattedDraft(body);
  const htmlContent = plainTextToHtml(formatted);
  await graphFetch(`/me/messages/${graphMessageId}/reply`, {
    method: 'POST',
    body: JSON.stringify({
      message: {
        body: {
          contentType: 'html',
          content: htmlContent,
        },
      },
    }),
  });
}

module.exports = {
  expandDraftWithOpenAI,
  applyEditWithOpenAI,
  sendReplyViaGraph,
  ensureFormattedDraft,
};

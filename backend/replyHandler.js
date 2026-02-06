const OpenAI = require('openai').default;
const { graphFetch } = require('./graphClient');

const SIGN_OFF = 'Best regards,\nNaked Car Studio';

// Remove any trailing sign-off (full block or orphan "Best regards,") so we can append the correct one once
function stripTrailingSignOff(text) {
  if (!text || typeof text !== 'string') return text;
  let s = text.replace(/\n*\s*Best regards,?\s*\n+\s*Naked Car Studio\s*$/i, '');
  s = s.replace(/\n*\s*Best regards,?\s*$/i, '');
  return s.trimEnd();
}

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
- Start with a greeting (e.g. "Dear [Name]," or "Hi,")
- Use proper paragraph breaks (blank line between paragraphs)
- Do not include any sign-off or closing (no "Best regards," or "Naked Car Studio"). End your reply with the last sentence of the body. The system will add the correct sign-off automatically.
- Output plain text only, well-formatted with double newlines between paragraphs
- Keep tone professional and friendly

User's message: ${userMessage}
Recipient name (for greeting): ${senderName || 'there'}`,
      },
    ],
    max_tokens: 500,
  });

  let draft = completion.choices?.[0]?.message?.content?.trim() || '';
  draft = stripTrailingSignOff(draft);
  draft = (draft ? draft.replace(/\s*$/, '') + '\n\n' : '') + SIGN_OFF;
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

Apply the changes. Do not include any sign-off. End with the last sentence of the body. The system will add the correct sign-off automatically.
Output the revised email only, well-formatted with double newlines between paragraphs.`,
      },
    ],
    max_tokens: 800,
  });

  let revised = completion.choices?.[0]?.message?.content?.trim() || draft;
  revised = stripTrailingSignOff(revised);
  revised = (revised ? revised.replace(/\s*$/, '') + '\n\n' : '') + SIGN_OFF;
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

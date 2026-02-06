const OpenAI = require('openai').default;
const { graphFetch } = require('./graphClient');

const SIGN_OFF = 'Best regards,\nNaked Car Studio';

// Remove the last two non-empty lines (AI sign-off); we then append our own sign-off
function removeLastTwoNonEmptyLines(text) {
  if (!text || typeof text !== 'string') return text;
  const lines = text.split('\n');
  const nonEmptyIndices = [];
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].trim() !== '') nonEmptyIndices.push(i);
  }
  const toRemove = nonEmptyIndices.slice(-2);
  toRemove.forEach((i) => (lines[i] = ''));
  return lines.join('\n').replace(/\n{3,}/g, '\n\n').trimEnd();
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
- End with exactly a two-line sign-off (e.g. a closing phrase on one line like "Warm regards," and your name or company on the next line). The system will replace it with the correct sign-off.
- Output plain text only, well-formatted with double newlines between paragraphs
- Keep tone professional and friendly

User's message: ${userMessage}
Recipient name (for greeting): ${senderName || 'there'}`,
      },
    ],
    max_tokens: 500,
  });

  let draft = completion.choices?.[0]?.message?.content?.trim() || '';
  draft = removeLastTwoNonEmptyLines(draft);
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

Apply the changes. End with exactly a two-line sign-off (closing phrase on one line, name or company on the next). The system will replace it with the correct sign-off.
Output the revised email only, well-formatted with double newlines between paragraphs.`,
      },
    ],
    max_tokens: 800,
  });

  let revised = completion.choices?.[0]?.message?.content?.trim() || draft;
  revised = removeLastTwoNonEmptyLines(revised);
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

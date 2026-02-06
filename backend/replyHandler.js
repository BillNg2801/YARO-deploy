const OpenAI = require('openai').default;
const { graphFetch } = require('./graphClient');

const SIGN_OFF = 'Best regards,\nNaked Car Studio';

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
- Always end with exactly: "Best regards," then newline then "Naked Car Studio"
- Output plain text only, well-formatted with double newlines between paragraphs
- Keep tone professional and friendly

User's message: ${userMessage}
Recipient name (for greeting): ${senderName || 'there'}`,
      },
    ],
    max_tokens: 500,
  });

  let draft = completion.choices?.[0]?.message?.content?.trim() || '';
  if (!draft.endsWith(SIGN_OFF)) {
    draft = draft.replace(/\s*$/, '') + (draft ? '\n\n' : '') + SIGN_OFF;
  }
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

Apply the changes. Keep the same format: greeting, body paragraphs, end with "Best regards," then newline then "Naked Car Studio".
Output the revised email only, well-formatted with double newlines between paragraphs.`,
      },
    ],
    max_tokens: 800,
  });

  let revised = completion.choices?.[0]?.message?.content?.trim() || draft;
  if (!revised.endsWith(SIGN_OFF)) {
    revised = revised.replace(/\s*$/, '') + (revised ? '\n\n' : '') + SIGN_OFF;
  }
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

async function sendReplyViaGraph(graphMessageId, body) {
  const formatted = ensureFormattedDraft(body);
  await graphFetch(`/me/messages/${graphMessageId}/reply`, {
    method: 'POST',
    body: JSON.stringify({ comment: formatted }),
  });
}

module.exports = {
  expandDraftWithOpenAI,
  applyEditWithOpenAI,
  sendReplyViaGraph,
  ensureFormattedDraft,
};

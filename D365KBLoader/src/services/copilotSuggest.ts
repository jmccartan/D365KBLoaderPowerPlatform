import sanitizeHtml from 'sanitize-html';
import type { ProcessedArticle, ArticleSuggestion } from '../types';

/**
 * Mock "Copilot" suggestion engine.
 *
 * Applies a few deterministic, realistic improvements that a real LLM-driven
 * editor might propose. Used in mock mode and as a fallback if the real
 * Copilot connector isn't wired up.
 *
 * The real-service equivalent should call Azure OpenAI (or a Dataverse AI
 * Prompt action) with the article HTML and parse the model response into the
 * same ArticleSuggestion shape.
 */
export function buildMockSuggestion(article: ProcessedArticle): ArticleSuggestion {
  const original = article.html ?? '';
  const changes: string[] = [];

  let next = original;

  // 1) Normalize whitespace between tags
  const collapsed = next.replace(/\s+\n/g, '\n').replace(/[ \t]{2,}/g, ' ');
  if (collapsed !== next) {
    changes.push('Normalized whitespace and collapsed double spaces.');
    next = collapsed;
  }

  // 2) Ensure the article starts with an <h1> derived from the title
  const hasH1 = /<h1[\s>]/i.test(next);
  if (!hasH1 && article.title) {
    next = `<h1>${escapeHtml(article.title)}</h1>\n${next}`;
    changes.push('Added a top-level heading using the article title.');
  }

  // 3) Add a short "Summary" callout if there isn't one already
  const hasSummary = /<h2[^>]*>\s*Summary\b/i.test(next) || /class=["'][^"']*summary/i.test(next);
  if (!hasSummary) {
    const firstSentence = extractFirstSentence(stripTags(next));
    if (firstSentence) {
      const callout = `<h2>Summary</h2>\n<p><em>${escapeHtml(firstSentence)}</em></p>\n`;
      // Insert after the first <h1> if present, else at the top
      const h1Match = next.match(/<\/h1>/i);
      if (h1Match && h1Match.index !== undefined) {
        const idx = h1Match.index + h1Match[0].length;
        next = next.slice(0, idx) + '\n' + callout + next.slice(idx);
      } else {
        next = callout + next;
      }
      changes.push('Added a "Summary" section above the body for quicker scanning.');
    }
  }

  // 4) Convert plain-text "Step 1 / Step 2 / Step 3 …" runs into an ordered list
  const stepRegex = /(?:<p>\s*)?Step\s+(\d+)[:.\-)]\s*([\s\S]*?)(?=(?:<\/p>|<p>\s*Step\s+\d|$))/gi;
  const stepMatches = [...next.matchAll(stepRegex)];
  if (stepMatches.length >= 2) {
    const items = stepMatches
      .map(m => `<li>${stripTags(m[2]).trim()}</li>`)
      .join('\n');
    const list = `<ol>\n${items}\n</ol>`;
    // Replace the first matched span through the last with the list
    const first = stepMatches[0];
    const last = stepMatches[stepMatches.length - 1];
    if (first.index !== undefined && last.index !== undefined) {
      const start = first.index;
      const end = last.index + last[0].length;
      next = next.slice(0, start) + list + next.slice(end);
      changes.push(`Converted ${stepMatches.length} "Step N" lines into a numbered list.`);
    }
  }

  // 5) Wrap orphan text nodes (non-tagged lines) in <p>
  const wrapped = wrapOrphanText(next);
  if (wrapped !== next) {
    changes.push('Wrapped loose text in <p> tags.');
    next = wrapped;
  }

  // 6) Add a small "Related" section if absent
  const hasRelated = /<h2[^>]*>\s*Related\b/i.test(next);
  if (!hasRelated) {
    next += `\n<h2>Related articles</h2>\n<ul>\n  <li><em>Add links to related knowledgebase entries here.</em></li>\n</ul>`;
    changes.push('Added a placeholder "Related articles" section for cross-linking.');
  }

  // Final sanitize — same pipeline as the importer to keep output safe
  const cleaned = sanitizeHtml(next, {
    allowedTags: sanitizeHtml.defaults.allowedTags.concat(['h1', 'h2', 'img']),
    allowedAttributes: { ...sanitizeHtml.defaults.allowedAttributes, '*': ['class'] },
  });

  // If we actually made no changes, return a tidy "looks good" suggestion
  if (changes.length === 0) {
    return {
      html: cleaned,
      summary: 'No improvements suggested — the article already looks well-structured.',
      changes: [],
    };
  }

  const suggestedTitle = improveTitle(article.title);
  if (suggestedTitle && suggestedTitle !== article.title) {
    changes.unshift(`Tightened the title: "${article.title}" → "${suggestedTitle}".`);
  }

  return {
    html: cleaned,
    title: suggestedTitle !== article.title ? suggestedTitle : undefined,
    summary:
      `Suggested ${changes.length} improvement${changes.length === 1 ? '' : 's'} to make this article ` +
      `easier to scan and more consistent with KB best practices.`,
    changes,
  };
}

function escapeHtml(s: string): string {
  return s.replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]!));
}

function stripTags(s: string): string {
  return s.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim();
}

function extractFirstSentence(s: string): string {
  if (!s) return '';
  const m = s.match(/[^.!?]+[.!?]/);
  const first = m ? m[0].trim() : s.slice(0, 160).trim();
  return first.length > 220 ? first.slice(0, 217) + '…' : first;
}

function wrapOrphanText(html: string): string {
  // Split on closing block tags and wrap any lines that are pure text in <p>
  return html.replace(/(^|>)\s*([^<>\n][^<\n]{2,})\s*(?=<|$)/g, (full, lead, text) => {
    // Skip if it's already inside a <p>/<li>/<h*> immediately before
    if (/<\/(p|li|h[1-6]|td|th|blockquote)>\s*$/i.test(lead)) return full;
    if (/^(p|li|h[1-6]|td|th|blockquote|ul|ol|table|thead|tbody|tr)$/i.test(lead.slice(1))) return full;
    return `${lead}<p>${text.trim()}</p>`;
  });
}

function improveTitle(title: string): string {
  if (!title) return title;
  let t = title.replace(/[-_]+/g, ' ').trim();
  // Title-case-ish, but preserve all-caps acronyms like VPN, KB, HR
  t = t.split(/\s+/).map(word => {
    if (/^[A-Z]{2,5}$/.test(word)) return word;
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  }).join(' ');
  return t;
}

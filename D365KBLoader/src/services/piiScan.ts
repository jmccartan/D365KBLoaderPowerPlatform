import type { PIIFinding } from '../types';

const MAX_SNIPPETS = 3;
const SNIPPET_RADIUS = 28;

type FindingKind = PIIFinding['kind'];

interface MatchInfo {
  kind: FindingKind;
  value: string;
  index: number;
}

const EMAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;
const SSN_RE = /\b\d{3}-\d{2}-\d{4}\b/g;
const CREDIT_CARD_RE = /(?<!\d)(?:\d[ -]?){13,19}(?!\d)/g;

export function findPii(input: string): PIIFinding[] {
  const text = normalizeInput(input);
  const matches: MatchInfo[] = [
    ...collectMatches(text, EMAIL_RE, 'email'),
    ...collectMatches(text, SSN_RE, 'ssn'),
    ...collectCreditCardMatches(text),
  ].sort((left, right) => left.index - right.index);

  if (matches.length === 0) {
    return [];
  }

  const grouped = new Map<FindingKind, PIIFinding>();
  for (const match of matches) {
    const existing = grouped.get(match.kind) ?? {
      kind: match.kind,
      count: 0,
      snippets: [],
    };
    existing.count += 1;
    const snippet = extractSnippet(text, match.index, match.index + match.value.length);
    if (snippet && existing.snippets.length < MAX_SNIPPETS && !existing.snippets.includes(snippet)) {
      existing.snippets.push(snippet);
    }
    grouped.set(match.kind, existing);
  }

  return Array.from(grouped.values());
}

function collectMatches(text: string, pattern: RegExp, kind: FindingKind): MatchInfo[] {
  pattern.lastIndex = 0;
  const matches: MatchInfo[] = [];
  let found: RegExpExecArray | null;
  while ((found = pattern.exec(text)) !== null) {
    matches.push({ kind, value: found[0], index: found.index });
  }
  return matches;
}

function collectCreditCardMatches(text: string): MatchInfo[] {
  CREDIT_CARD_RE.lastIndex = 0;
  const matches: MatchInfo[] = [];
  let found: RegExpExecArray | null;
  while ((found = CREDIT_CARD_RE.exec(text)) !== null) {
    const digitsOnly = found[0].replace(/\D/g, '');
    if (digitsOnly.length < 13 || digitsOnly.length > 19) continue;
    // Cut down false positives (phone numbers, order ids, etc.) by requiring
    // a Luhn checksum match — every issued PAN passes Luhn by spec.
    if (!isLuhnValid(digitsOnly)) continue;
    matches.push({ kind: 'credit-card', value: found[0], index: found.index });
  }
  return matches;
}

/** Standard Luhn check used to validate credit-card numbers. */
function isLuhnValid(digits: string): boolean {
  let sum = 0;
  let alt = false;
  for (let i = digits.length - 1; i >= 0; i--) {
    let n = digits.charCodeAt(i) - 48;
    if (n < 0 || n > 9) return false;
    if (alt) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
    alt = !alt;
  }
  return sum % 10 === 0;
}

function normalizeInput(input: string): string {
  return input
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function extractSnippet(text: string, start: number, end: number): string {
  const from = Math.max(0, start - SNIPPET_RADIUS);
  const to = Math.min(text.length, end + SNIPPET_RADIUS);
  const prefix = from > 0 ? '…' : '';
  const suffix = to < text.length ? '…' : '';
  return `${prefix}${text.slice(from, to).trim()}${suffix}`;
}

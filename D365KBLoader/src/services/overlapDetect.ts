import type { ProcessedArticle, ExistingKbArticle, OverlapMatch } from '../types';

/**
 * Lightweight client-side overlap detection between candidate import articles
 * and a set of existing D365 knowledgebase articles.
 *
 * Uses a Jaccard-style token overlap on title + first body chunk, with stopword
 * filtering. Returns the top matches per candidate (score >= threshold).
 *
 * The real service should call Dataverse to fetch candidate existing rows
 * (`knowledgearticle`) and then call `scoreOverlaps` to produce matches.
 */
export function scoreOverlaps(
  candidates: ProcessedArticle[],
  existing: ExistingKbArticle[],
  opts: { threshold?: number; maxPerCandidate?: number } = {},
): Record<string, OverlapMatch[]> {
  const threshold = opts.threshold ?? 0.22;
  const maxPerCandidate = opts.maxPerCandidate ?? 3;

  const existingTokens = existing.map(e => ({
    article: e,
    titleTokens: tokenize(e.title),
    bodyTokens: tokenize(e.excerpt ?? ''),
  }));

  const result: Record<string, OverlapMatch[]> = {};

  for (const c of candidates) {
    const candTitleTokens = tokenize(c.title);
    const candBodyTokens = tokenize(stripTags(c.html).slice(0, 800));
    const matches: OverlapMatch[] = [];

    for (const e of existingTokens) {
      const titleSim = jaccard(candTitleTokens, e.titleTokens);
      const bodySim = jaccard(candBodyTokens, e.bodyTokens);
      // Title weighted higher than body — KB matches usually share a clear topic
      const score = +(titleSim * 0.65 + bodySim * 0.35).toFixed(3);
      if (score < threshold) continue;

      const sharedTitle = intersect(candTitleTokens, e.titleTokens).slice(0, 5);
      const sharedBody = intersect(candBodyTokens, e.bodyTokens)
        .filter(t => !sharedTitle.includes(t))
        .slice(0, 6);

      const reasons: string[] = [];
      if (titleSim > 0) reasons.push(`Title similarity ${pct(titleSim)}`);
      if (bodySim > 0) reasons.push(`Body similarity ${pct(bodySim)}`);
      if (sharedTitle.length) reasons.push(`Shared title terms: ${sharedTitle.join(', ')}`);
      if (sharedBody.length) reasons.push(`Shared keywords: ${sharedBody.join(', ')}`);

      matches.push({ article: e.article, score, reasons });
    }

    matches.sort((a, b) => b.score - a.score);
    result[c.id] = matches.slice(0, maxPerCandidate);
  }

  return result;
}

/* --------------------------- helpers --------------------------- */

const STOPWORDS = new Set([
  'a', 'an', 'and', 'are', 'as', 'at', 'be', 'by', 'for', 'from', 'has', 'have',
  'how', 'i', 'in', 'is', 'it', 'its', 'of', 'on', 'or', 'that', 'the', 'this',
  'to', 'was', 'were', 'with', 'you', 'your', 'we', 'our', 'will', 'can', 'do',
  'does', 'if', 'into', 'over', 'so', 'than', 'then', 'there', 'these', 'they',
  'use', 'using', 'about', 'when', 'what', 'which', 'who', 'why', 'where',
  'guide', 'article', 'document', 'doc', 'page',
]);

function tokenize(s: string): Set<string> {
  if (!s) return new Set();
  return new Set(
    s.toLowerCase()
      .replace(/[^\p{L}\p{N}\s]/gu, ' ')
      .split(/\s+/)
      .filter(t => t.length >= 3 && !STOPWORDS.has(t)),
  );
}

function jaccard(a: Set<string>, b: Set<string>): number {
  if (a.size === 0 && b.size === 0) return 0;
  let inter = 0;
  for (const t of a) if (b.has(t)) inter++;
  const union = a.size + b.size - inter;
  return union === 0 ? 0 : inter / union;
}

function intersect(a: Set<string>, b: Set<string>): string[] {
  const out: string[] = [];
  for (const t of a) if (b.has(t)) out.push(t);
  return out;
}

function stripTags(html: string): string {
  return html.replace(/<[^>]+>/g, ' ').replace(/&nbsp;/g, ' ').replace(/\s+/g, ' ').trim();
}

function pct(n: number): string {
  return `${Math.round(n * 100)}%`;
}

import { describe, expect, it } from 'vitest';
import { scoreOverlaps } from '../overlapDetect';
import type { ExistingKbArticle, ProcessedArticle } from '../../types';

function makeCandidate(id: string, title: string, html = ''): ProcessedArticle {
  return {
    id,
    source: { id, name: `${id}.html`, path: `/${id}.html`, size: 0, modified: '', kind: 'html' },
    title,
    html,
    rawHtml: html,
    warnings: [],
    findings: [],
    selected: true,
    loadStatus: 'pending',
  };
}

const EXISTING: ExistingKbArticle[] = [
  { id: 'e1', title: 'VPN setup for remote workers', excerpt: 'Configure the corporate VPN client and connect.' },
  { id: 'e2', title: 'How to reset your Windows password', excerpt: 'Step by step password reset via self service portal.' },
  { id: 'e3', title: 'Onboarding checklist for new employees', excerpt: 'Day one checklist for new hires.' },
];

describe('scoreOverlaps', () => {
  it('returns no matches for unrelated content', () => {
    const out = scoreOverlaps([makeCandidate('c1', 'How to brew espresso')], EXISTING);
    expect(out.c1).toEqual([]);
  });

  it('ranks the most similar existing article first', () => {
    const out = scoreOverlaps(
      [makeCandidate('c1', 'VPN Setup', 'Set up the corporate vpn client')],
      EXISTING,
    );
    expect(out.c1.length).toBeGreaterThan(0);
    expect(out.c1[0].article.id).toBe('e1');
  });

  it('produces human readable reasons including shared terms', () => {
    const out = scoreOverlaps(
      [makeCandidate('c1', 'Reset password walkthrough', 'Reset your windows password using the portal.')],
      EXISTING,
    );
    const top = out.c1[0];
    expect(top.reasons.join(' ')).toMatch(/Shared|similarity/i);
  });
});

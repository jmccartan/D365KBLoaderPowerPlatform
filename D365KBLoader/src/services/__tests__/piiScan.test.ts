import { describe, expect, it } from 'vitest';
import { findPii } from '../piiScan';

describe('findPii', () => {
  it('returns nothing for clean text', () => {
    expect(findPii('Hello world, nothing sensitive here.')).toEqual([]);
  });

  it('detects an email address', () => {
    const out = findPii('Contact us at support@contoso.com please.');
    const email = out.find(f => f.kind === 'email');
    expect(email).toBeDefined();
    expect(email!.count).toBe(1);
  });

  it('detects a US SSN', () => {
    const out = findPii('SSN on file: 123-45-6789.');
    expect(out.find(f => f.kind === 'ssn')?.count).toBe(1);
  });

  it('Luhn-valid 16-digit card is flagged', () => {
    // 4111 1111 1111 1111 is the canonical Visa test number — passes Luhn.
    const out = findPii('Card: 4111-1111-1111-1111 ends in 1111.');
    expect(out.find(f => f.kind === 'credit-card')?.count).toBe(1);
  });

  it('Luhn-invalid 16-digit run is NOT flagged', () => {
    // 1234-5678-9012-3456 fails Luhn — should be ignored (false-positive guard)
    const out = findPii('Order id 1234-5678-9012-3456 was shipped.');
    expect(out.find(f => f.kind === 'credit-card')).toBeUndefined();
  });

  it('a US phone number is NOT flagged as a credit card', () => {
    const out = findPii('Call us at 415-555-0123 between 9 and 5.');
    expect(out.find(f => f.kind === 'credit-card')).toBeUndefined();
  });

  it('groups multiple emails into one finding with count', () => {
    const out = findPii('a@b.com and c@d.com and e@f.com');
    const email = out.find(f => f.kind === 'email')!;
    expect(email.count).toBe(3);
    expect(email.snippets.length).toBeGreaterThan(0);
  });
});

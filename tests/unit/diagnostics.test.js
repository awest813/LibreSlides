import { describe, it, expect } from 'vitest';
import { noteError, diagnostics } from '../../src/diagnostics.js';

globalThis.matchMedia ||= () => ({ matches: false, addEventListener() {}, removeEventListener() {} });

describe('bug report details', () => {
  it('lists recent errors readably, whatever was thrown', async () => {
    const quota = new Error('The quota has been exceeded.');
    quota.name = 'QuotaExceededError';
    delete quota.stack;
    Object.defineProperty(quota, 'stack', { value: undefined });
    noteError(quota, 'save');
    noteError({ code: 42, reason: 'nope' }, 'route');
    noteError('plain words');
    const report = await diagnostics();
    expect(report).toContain('[save] QuotaExceededError: The quota has been exceeded.');
    expect(report).toContain('[route] {"code":42,"reason":"nope"}');
    expect(report).toContain('plain words');
    expect(report).not.toContain('[object Object]');
  });
});

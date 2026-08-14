import { describe, it, expect } from 'vitest';
import { describeRequestForLog } from '../src/graph-client.js';

/**
 * Regression guard for SEC-2026-001.
 *
 * The original code logged `JSON.stringify(options)`, which wrote every
 * outbound Graph request body to the log — email contents from send-mail,
 * base64 file content from attachment uploads — plus the caller's access token
 * in multi-account mode. It also defeated the redaction executeGraphTool does
 * one layer above it.
 *
 * These tests assert on absence, which is the only thing that matters here: the
 * log line must not be able to carry payload or credentials, whatever is passed.
 */
describe('graph request log line (SEC-2026-001)', () => {
  it('never emits the request body', () => {
    const secretEmail = 'Dear Bob, the acquisition price is $4.2M, regards Alice';
    const line = describeRequestForLog({
      method: 'POST',
      body: JSON.stringify({ message: { body: { content: secretEmail } } }),
    });

    expect(line).not.toContain(secretEmail);
    expect(line).not.toContain('acquisition');
    expect(line).not.toContain('4.2M');
    // Size is still reported, which is the part with diagnostic value.
    expect(line).toMatch(/bodyBytes=\d+/);
  });

  it('never emits an access token, and says so explicitly', () => {
    const token = 'eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.PAYLOAD.SIGNATURE';
    const line = describeRequestForLog({ method: 'GET', accessToken: token });

    expect(line).not.toContain(token);
    expect(line).not.toContain('eyJ');
    expect(line).toContain('accessToken=[REDACTED]');
  });

  it('never emits header values, only their names', () => {
    const line = describeRequestForLog({
      method: 'GET',
      headers: { Authorization: 'Bearer super-secret-value', ConsistencyLevel: 'eventual' },
    });

    expect(line).not.toContain('super-secret-value');
    expect(line).not.toContain('Bearer');
    // Names are retained — knowing ConsistencyLevel was applied is genuinely useful.
    expect(line).toContain('Authorization');
    expect(line).toContain('ConsistencyLevel');
  });

  it('handles base64 attachment payloads without echoing them', () => {
    const base64 = Buffer.from('a'.repeat(5000)).toString('base64');
    const line = describeRequestForLog({ method: 'POST', body: JSON.stringify({ contentBytes: base64 }) });

    expect(line).not.toContain(base64.slice(0, 40));
    expect(line.length).toBeLessThan(120);
  });

  it('still reports what a diagnostician actually needs', () => {
    const line = describeRequestForLog({ method: 'PATCH', body: '{"isRead":true}' });
    expect(line).toContain('method=PATCH');
    expect(line).toMatch(/bodyBytes=\d+/);
  });

  it('defaults to GET and stays terse when there is nothing to report', () => {
    expect(describeRequestForLog()).toBe('method=GET');
    expect(describeRequestForLog({})).toBe('method=GET');
  });

  it('cannot be made to leak by an unusual body type', () => {
    // Non-string bodies previously round-tripped through JSON.stringify.
    const line = describeRequestForLog({
      method: 'POST',
      body: { nested: { secret: 'do-not-log-me' } } as unknown as string,
    });
    expect(line).not.toContain('do-not-log-me');
  });
});

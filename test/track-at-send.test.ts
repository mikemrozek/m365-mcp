import { describe, it, expect } from 'vitest';
import { applyTrackAtSend } from '../src/graph-tools.js';
import { STATE_PROPERTY_ID, META_PROPERTY_ID } from '../src/reply-tracking-tools.js';

type Props = { id: string; value: string }[];

describe('track at send (tsq.28)', () => {
  it('marks the message inside a sendMail body', () => {
    const out = applyTrackAtSend(
      'send-mail',
      { message: { subject: 'Q4', toRecipients: [] }, saveToSentItems: true },
      { dueBy: '2026-10-20T00:00:00Z', waitingOn: ['Bob@TSQ.com'] }
    ) as {
      message: { subject: string; singleValueExtendedProperties: Props };
      saveToSentItems: boolean;
    };
    expect(out.saveToSentItems).toBe(true);
    expect(out.message.subject).toBe('Q4');
    const props = out.message.singleValueExtendedProperties;
    expect(props.find((p) => p.id === STATE_PROPERTY_ID)?.value).toBe('awaiting');
    const meta = JSON.parse(props.find((p) => p.id === META_PROPERTY_ID)!.value);
    expect(meta).toMatchObject({ v: 1, dueBy: '2026-10-20T00:00:00Z', waitingOn: ['bob@tsq.com'] });
  });

  it('marks a draft body directly, since the body is the message', () => {
    const out = applyTrackAtSend('create-draft-email', { subject: 'Hi' }, {}) as {
      subject: string;
      singleValueExtendedProperties: Props;
    };
    expect(out.subject).toBe('Hi');
    expect(out.singleValueExtendedProperties).toHaveLength(2);
  });

  it('creates the message object for a reply sent with only a comment', () => {
    const out = applyTrackAtSend('reply-mail-message', { comment: 'Thanks' }, {}) as {
      comment: string;
      message: { singleValueExtendedProperties: Props };
    };
    expect(out.comment).toBe('Thanks');
    expect(out.message.singleValueExtendedProperties).toHaveLength(2);
  });

  it('keeps unrelated extended properties and replaces stale tracking ones', () => {
    const out = applyTrackAtSend(
      'create-draft-email',
      {
        singleValueExtendedProperties: [
          { id: 'String {x} Name Other', value: 'keep' },
          { id: STATE_PROPERTY_ID, value: 'done' },
        ],
      },
      {}
    ) as { singleValueExtendedProperties: Props };
    const props = out.singleValueExtendedProperties;
    expect(props.map((p) => p.value)).toContain('keep');
    expect(props.filter((p) => p.id === STATE_PROPERTY_ID)).toEqual([
      { id: STATE_PROPERTY_ID, value: 'awaiting' },
    ]);
  });

  it('leaves tools that do not create a message untouched', () => {
    const body = { subject: 'x' };
    expect(applyTrackAtSend('create-calendar-event', body, {})).toBe(body);
  });
});

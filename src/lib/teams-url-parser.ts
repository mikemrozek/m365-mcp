/**
 * Converts any Teams meeting URL format into a standard joinWebUrl
 * usable with the list-online-meetings $filter=joinWebUrl eq '...' query.
 *
 * Supported formats:
 * - Short URL: https://teams.microsoft.com/meet/29752586464443?p=...
 * - Full joinWebUrl: https://teams.microsoft.com/l/meetup-join/19%3ameeting_.../0?context=...
 * - Recap URL: https://teams.microsoft.com/v2/#/meetingrecap?threadId=...&tenantId=...&organizerId=...
 */
export function parseTeamsUrl(url: string): string {
  // Format 1 & 2: Already a joinWebUrl or short /meet/ URL — pass through
  if (url.includes('/meet/') || url.includes('/meetup-join/')) {
    return url;
  }

  // Format 3: Recap URL — extract params and reconstruct joinWebUrl
  if (url.toLowerCase().includes('meetingrecap')) {
    const params = Object.fromEntries(
      [...url.matchAll(/([a-zA-Z]+)=([^&#]+)/g)].map((m) => [m[1], m[2]])
    );
    const threadId = decodeURIComponent(params.threadId || '');
    const tenantId = params.tenantId || '';
    const organizerId = params.organizerId || '';

    if (!threadId || !tenantId || !organizerId) {
      throw new Error('Invalid recap URL: missing threadId, tenantId, or organizerId parameter');
    }

    const threadEnc = encodeURIComponent(threadId).replace(/%3A/gi, '%3a').replace(/%40/gi, '%40');
    const ctx = JSON.stringify({ Tid: tenantId, Oid: organizerId });
    const ctxEnc = encodeURIComponent(ctx);

    return `https://teams.microsoft.com/l/meetup-join/${threadEnc}/0?context=${ctxEnc}`;
  }

  // Unknown format — return as-is
  return url;
}

/**
 * Pull the meeting chat thread id (`19:meeting_…@thread.v2`) out of a Teams URL.
 *
 * The thread id is the one identifier every participant already holds and the
 * one Note Taker keys on (see docs/2026-0825-0907 meeting-id inventory). It is
 * present verbatim in recap URLs (`?threadId=`) and in the path of full
 * `/meetup-join/` URLs; short `/meet/<code>` URLs do not carry it and need a
 * Graph lookup instead, so this returns undefined for those.
 */
export function threadIdFromTeamsUrl(url: string): string | undefined {
  const recap = url.match(/[?&]threadId=([^&#]+)/i);
  if (recap) {
    const id = decodeURIComponent(recap[1]);
    return isMeetingThreadId(id) ? id : undefined;
  }
  const join = url.match(/\/meetup-join\/([^/?#]+)/i);
  if (join) {
    const id = decodeURIComponent(join[1]);
    return isMeetingThreadId(id) ? id : undefined;
  }
  return undefined;
}

/** A Teams meeting chat thread id, as returned in onlineMeeting.chatInfo.threadId. */
export function isMeetingThreadId(value: string): boolean {
  return /^19:meeting_[A-Za-z0-9_-]+@thread\.v2$/.test(value);
}

import { readFileSync } from 'fs';

/**
 * An explicit, named tool allowlist. Unlike the ENABLED_TOOLS regex, this is an
 * exact-match list of tool names — one reviewable source of truth per connector
 * tier. See docs/2026-0707-1502 m365-mcp-tiered-rollout-plan.md (Phase 1).
 */
export interface LoadedToolAllowlist {
  /** Exact tool names, de-duplicated and sorted. */
  names: string[];
  /** Human-readable description of where the list came from (for startup logging). */
  source: string;
}

/**
 * Parse a TOOL_ALLOWLIST value into an exact-match list of tool names.
 *
 * The raw value is one of:
 *   - inline JSON array:   ["get-current-user","list-mail-messages"]
 *   - inline JSON object:  {"tools":["get-current-user", ...]}   (extra keys ignored)
 *   - a filesystem path to a file containing either of the above (a mounted config)
 *
 * Matching is always exact — no regex, no patterns. Throws on any malformed input
 * so the caller can fail closed at startup rather than silently exposing every tool.
 */
export function loadToolAllowlist(raw: string): LoadedToolAllowlist {
  const trimmed = raw.trim();
  if (!trimmed) {
    throw new Error('TOOL_ALLOWLIST is empty.');
  }

  let jsonText: string;
  let source: string;
  if (trimmed.startsWith('[') || trimmed.startsWith('{')) {
    jsonText = trimmed;
    source = 'inline JSON';
  } else {
    // Treat anything that isn't inline JSON as a path to a config file.
    try {
      jsonText = readFileSync(trimmed, 'utf8');
    } catch (error) {
      throw new Error(
        `TOOL_ALLOWLIST looks like a file path but could not be read: "${trimmed}" (${(error as Error).message})`
      );
    }
    source = `file ${trimmed}`;
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(jsonText);
  } catch (error) {
    throw new Error(`TOOL_ALLOWLIST is not valid JSON (${source}): ${(error as Error).message}`);
  }

  let rawList: unknown;
  if (Array.isArray(parsed)) {
    rawList = parsed;
  } else if (parsed && typeof parsed === 'object' && Array.isArray((parsed as { tools?: unknown }).tools)) {
    rawList = (parsed as { tools: unknown }).tools;
  } else {
    throw new Error(
      `TOOL_ALLOWLIST (${source}) must be a JSON array of tool names or an object with a "tools" array.`
    );
  }

  const seen = new Set<string>();
  for (const entry of rawList as unknown[]) {
    if (typeof entry !== 'string' || !entry.trim()) {
      throw new Error(
        `TOOL_ALLOWLIST (${source}) contains a non-string or empty tool name: ${JSON.stringify(entry)}`
      );
    }
    seen.add(entry.trim());
  }

  const names = Array.from(seen).sort((a, b) => a.localeCompare(b));
  return { names, source: `${source} (${names.length} tools)` };
}

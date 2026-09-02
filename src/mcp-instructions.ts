/** Shared context for MCP `initialize.instructions` (hosts that forward it to the model). */
export type McpInstructionsContext = {
  orgMode: boolean;
  readOnly: boolean;
  multiAccount: boolean;
};

function buildGeneralMcpInstructions(opts: McpInstructionsContext): string {
  const parts = [
    'Microsoft 365 MCP exposes Microsoft Graph through MCP tools. Use each tool name, description, and parameter schema as the source of truth.',
    'Microsoft Graph OData: do not combine $filter with $search on the same request. For lists, prefer modest $top (or top) and $select; avoid very large pages unless the user needs them.',
    'Mail and message $search uses KQL; the $search query parameter value must be double-quoted per Graph (see search-query-parameter in Microsoft Graph docs).',
    'When you need an organizational user or recipient address, resolve it with list-users (or another directory tool); do not invent SMTP addresses.',
    'Directory $search on collections such as /users or /groups requires ConsistencyLevel: eventual when the tool exposes that header.',
    'Teams chat and channel messages: prefer HTML contentType in the body; plain text is often mangled by Graph.',
    // Stated once here rather than repeated on all 136 tools. Measured 2026-09-02:
    // the shared query parameters carried ~24,000 tokens of duplicated prose across
    // the catalogue, which is also noise for tool search, since discovery matches on
    // parameter names and descriptions as well as the description itself.
    'Shared query parameters, the same on every list tool: top is page size (start at 5-15); select limits returned fields and is worth passing on every list; skip and orderby page and sort; fetchAllPages merges up to 100 pages and should be reserved for a genuine full export, since it can return an enormous payload.',
    'Response-shape parameters, available on every tool: includeHeaders also returns response headers such as the ETag needed for a conditional update; excludeResponse returns only success or failure, which is worth passing when a write returns a large entity nobody needs.',
    'Advanced query mode: count=true sends ConsistencyLevel: eventual, and is required for contains() filters and for filtering on flag or flagStatus.',
    'Write tools describe their request body as a worked example on the body parameter itself, rather than as a full schema. Send the fields named there; Graph validates the rest, and its error text names the offending field.',
  ];
  if (opts.readOnly) parts.push('This server is read-only; write operations are disabled.');
  if (opts.multiAccount)
    parts.push('Multiple accounts: pass the account parameter when required (see list-accounts).');
  if (!opts.orgMode)
    parts.push('Work/school-only tools require starting the server with --org-mode.');
  return parts.join(' ');
}

const DISCOVERY_MODE_INSTRUCTIONS_ADDON =
  'DISCOVERY MODE ADD-ON: Graph is reached via search-tools → get-tool-schema → execute-tool (plus auth helpers). ' +
  'Workflow: (1) call search-tools with short natural-language keywords (BM25-ranked); ' +
  '(2) call get-tool-schema(tool_name) to see the parameters, required fields, and enum values; ' +
  '(3) call execute-tool with tool_name exactly as returned and parameters shaped per the schema. ' +
  'Skipping get-tool-schema is the leading cause of Graph 400 errors here. ' +
  'If search-tools returns no matches, retry with shorter or different keywords.';

/**
 * Full MCP `initialize.instructions` string: general guidance for every mode, plus a discovery-only suffix when applicable.
 */
export function buildMcpServerInstructions(
  opts: McpInstructionsContext & { discovery: boolean }
): string {
  const general = buildGeneralMcpInstructions(opts);
  if (!opts.discovery) return general;
  return `${general} ${DISCOVERY_MODE_INSTRUCTIONS_ADDON}`;
}

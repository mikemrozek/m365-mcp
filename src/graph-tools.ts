import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import logger from './logger.js';
import GraphClient from './graph-client.js';
import AuthManager from './auth.js';
import { api } from './generated/client.js';
import { z } from 'zod';
import { readFileSync } from 'fs';
import { createHash } from 'crypto';
import path from 'path';
import { fileURLToPath } from 'url';
import { TOOL_CATEGORIES } from './tool-categories.js';
import { getRequestTokens } from './request-context.js';
import { parseTeamsUrl } from './lib/teams-url-parser.js';
import { buildBM25Index, scoreQuery, tokenize, type BM25Index } from './lib/bm25.js';
export interface DiscoverySearchIndex {
  bm25: BM25Index;
  nameTokens: Map<string, Set<string>>;
}
import { describeToolSchema } from './lib/tool-schema.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

interface EndpointConfig {
  pathPattern: string;
  method: string;
  toolName: string;
  scopes?: string[];
  workScopes?: string[];
  returnDownloadUrl?: boolean;
  supportsTimezone?: boolean;
  supportsExpandExtendedProperties?: boolean;
  llmTip?: string;
  skipEncoding?: string[]; // Parameter names that should NOT be URL-encoded (for function-style API calls)
  contentType?: string;
  acceptType?: string; // Custom Accept header for endpoints returning non-JSON content (e.g., text/vtt)
  readOnly?: boolean; // When true, allow this endpoint in read-only mode even if method is not GET
}

const endpointsData = JSON.parse(
  readFileSync(path.join(__dirname, 'endpoints.json'), 'utf8')
) as EndpointConfig[];

/** When set to a positive integer, caps Graph `$top` on list requests (see README). */
function maxTopFromEnv(): number | undefined {
  const raw = process.env.MS365_MCP_MAX_TOP;
  if (raw === undefined || raw === '') return undefined;
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 1) {
    logger.warn(
      `Ignoring invalid MS365_MCP_MAX_TOP=${JSON.stringify(raw)} (use a positive integer)`
    );
    return undefined;
  }
  return n;
}

function clampTopQueryParam(queryParams: Record<string, string>): void {
  const cap = maxTopFromEnv();
  if (cap === undefined || queryParams['$top'] === undefined) return;
  const requested = Number.parseInt(queryParams['$top'], 10);
  if (!Number.isFinite(requested) || requested <= cap) return;
  logger.info(`Clamping $top from ${requested} to ${cap} (MS365_MCP_MAX_TOP)`);
  queryParams['$top'] = String(cap);
}

type TextContent = {
  type: 'text';
  text: string;
  [key: string]: unknown;
};

type ImageContent = {
  type: 'image';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type AudioContent = {
  type: 'audio';
  data: string;
  mimeType: string;
  [key: string]: unknown;
};

type ResourceTextContent = {
  type: 'resource';
  resource: {
    text: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceBlobContent = {
  type: 'resource';
  resource: {
    blob: string;
    uri: string;
    mimeType?: string;
    [key: string]: unknown;
  };
  [key: string]: unknown;
};

type ResourceContent = ResourceTextContent | ResourceBlobContent;

type ContentItem = TextContent | ImageContent | AudioContent | ResourceContent;

interface CallToolResult {
  content: ContentItem[];
  _meta?: Record<string, unknown>;
  isError?: boolean;

  [key: string]: unknown;
}

async function executeGraphTool(
  tool: (typeof api.endpoints)[0],
  config: EndpointConfig | undefined,
  graphClient: GraphClient,
  params: Record<string, unknown>,
  authManager?: AuthManager
): Promise<CallToolResult> {
  logger.info(`Tool ${tool.alias} called with params: ${JSON.stringify(params)}`);
  try {
    // Resolve account-specific token if `account` parameter is provided (or auto-resolve for single account).
    // Skip in OAuth/HTTP mode — let the request context drive token selection via GraphClient.
    // Also skip when a request-context token exists (HTTP/OAuth flow where token comes from middleware).
    let accountAccessToken: string | undefined;
    if (authManager && !authManager.isOAuthModeEnabled() && !getRequestTokens()) {
      const accountParam = params.account as string | undefined;
      try {
        accountAccessToken = await authManager.getTokenForAccount(accountParam);
      } catch (err) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: (err as Error).message }),
            },
          ],
          isError: true,
        };
      }
    }

    const parameterDefinitions = tool.parameters || [];

    let path = tool.path;
    const queryParams: Record<string, string> = {};
    const headers: Record<string, string> = {};
    let body: unknown = null;

    for (const [paramName, paramValue] of Object.entries(params)) {
      // Skip control parameters - not part of the Microsoft Graph API
      if (
        [
          'account',
          'fetchAllPages',
          'includeHeaders',
          'excludeResponse',
          'timezone',
          'expandExtendedProperties',
        ].includes(paramName)
      ) {
        continue;
      }

      // Ok, so, MCP clients (such as claude code) doesn't support $ in parameter names,
      // and others might not support __, so we strip them in hack.ts and restore them here
      const odataParams = [
        'filter',
        'select',
        'expand',
        'orderby',
        'skip',
        'top',
        'count',
        'search',
        'format',
      ];
      // Handle both "top" and "$top" formats - strip $ if present, then re-add it
      const normalizedParamName = paramName.startsWith('$') ? paramName.slice(1) : paramName;
      const isOdataParam = odataParams.includes(normalizedParamName.toLowerCase());
      const fixedParamName = isOdataParam ? `$${normalizedParamName.toLowerCase()}` : paramName;
      // Convert kebab-case param names to camelCase for path param matching.
      // endpoints.json uses {message-id} but hack.ts extracts :messageId (camelCase) from the path.
      // LLMs may pass "message-id" (kebab) — we normalize so both forms work.
      const camelCaseParamName = paramName.replace(/-([a-z])/g, (_, c: string) => c.toUpperCase());

      // Look up param definition using normalized name (without $) for OData params,
      // or camelCase equivalent for kebab-case path params
      const paramDef = parameterDefinitions.find(
        (p) =>
          p.name === paramName ||
          p.name === camelCaseParamName ||
          (isOdataParam && p.name === normalizedParamName)
      );

      if (paramDef) {
        switch (paramDef.type) {
          case 'Path': {
            // Check if this parameter should skip URL encoding (for function-style API calls)
            const shouldSkipEncoding = config?.skipEncoding?.includes(paramName) ?? false;
            // Use encodeURIComponent but preserve '=' which is valid in path segments (RFC 3986)
            // and commonly appears in Microsoft Graph base64-encoded resource IDs.
            // Without this, IDs like "AAMk...AAA=" become "AAMk...AAA%3D" causing 404 errors.
            // First we encode, then unencode. Crazy, check out https://github.com/Softeria/ms-365-mcp-server/issues/245
            const encodedValue = shouldSkipEncoding
              ? (paramValue as string)
              : encodeURIComponent(paramValue as string).replace(/%3D/g, '=');

            // Replace both the original param name and the camelCase variant
            // to handle {message-id} (endpoints.json) and :messageId (generated client) formats
            path = path
              .replace(`{${paramName}}`, encodedValue)
              .replace(`:${paramName}`, encodedValue)
              .replace(`{${camelCaseParamName}}`, encodedValue)
              .replace(`:${camelCaseParamName}`, encodedValue);
            break;
          }

          case 'Query':
            if (paramValue !== '' && paramValue != null) {
              queryParams[fixedParamName] = `${paramValue}`;
            }
            break;

          case 'Body':
            if (paramDef.schema) {
              const parseResult = paramDef.schema.safeParse(paramValue);
              if (!parseResult.success) {
                const wrapped = { [paramName]: paramValue };
                const wrappedResult = paramDef.schema.safeParse(wrapped);
                if (wrappedResult.success) {
                  logger.info(
                    `Auto-corrected parameter '${paramName}': AI passed nested field directly, wrapped it as {${paramName}: ...}`
                  );
                  body = wrapped;
                } else {
                  body = paramValue;
                }
              } else {
                body = paramValue;
              }
            } else {
              body = paramValue;
            }
            break;

          case 'Header':
            headers[fixedParamName] = `${paramValue}`;
            break;
        }
      } else if (paramName === 'body') {
        body = paramValue;
        logger.info(`Set body param: ${JSON.stringify(body)}`);
      } else if (
        path.includes(`:${paramName}`) ||
        path.includes(`{${paramName}}`) ||
        path.includes(`:${camelCaseParamName}`) ||
        path.includes(`{${camelCaseParamName}}`)
      ) {
        // Fallback: path param not declared in tool.parameters (generated client omits them).
        // Replace placeholder directly so the URL is valid.
        const encodedValue = encodeURIComponent(paramValue as string).replace(/%3D/g, '=');
        path = path
          .replace(`{${paramName}}`, encodedValue)
          .replace(`:${paramName}`, encodedValue)
          .replace(`{${camelCaseParamName}}`, encodedValue)
          .replace(`:${camelCaseParamName}`, encodedValue);
        logger.info(`Path param fallback: replaced :${camelCaseParamName} with encoded value`);
      }
    }

    // Defense-in-depth for Bug 2: Graph rejects $top on /delta() endpoints with
    // HTTP 400. The user-facing schema for -delta tools strips top/$top, so
    // freshly-connected clients can't send it. But cached/stale clients (and
    // ad-hoc callers) might still try. Drop it server-side before clamping or
    // sending, regardless of where it came from.
    if (tool.alias.endsWith('-delta')) {
      delete queryParams['$top'];
    }

    clampTopQueryParam(queryParams);

    const preferValues: string[] = [];

    // Handle timezone parameter for calendar endpoints
    if (config?.supportsTimezone && params.timezone) {
      preferValues.push(`outlook.timezone="${params.timezone}"`);
      logger.info(`Setting timezone preference: outlook.timezone="${params.timezone}"`);
    }

    const bodyFormat = process.env.MS365_MCP_BODY_FORMAT || 'text';
    if (bodyFormat !== 'html' && tool.method.toUpperCase() === 'GET') {
      preferValues.push(`outlook.body-content-type="${bodyFormat}"`);
    }

    if (preferValues.length > 0) {
      headers['Prefer'] = preferValues.join(', ');
    }

    // Handle expandExtendedProperties parameter for calendar endpoints
    if (config?.supportsExpandExtendedProperties && params.expandExtendedProperties === true) {
      const expandValue = 'singleValueExtendedProperties';
      if (queryParams['$expand']) {
        queryParams['$expand'] += `,${expandValue}`;
      } else {
        queryParams['$expand'] = expandValue;
      }
      logger.info(`Adding $expand=${expandValue} for extended properties`);
    }

    if (config?.contentType) {
      headers['Content-Type'] = config.contentType;
      logger.info(`Setting custom Content-Type: ${config.contentType}`);
    }

    if (config?.acceptType) {
      headers['Accept'] = config.acceptType;
      logger.info(`Setting custom Accept: ${config.acceptType}`);
    }

    if (Object.keys(queryParams).length > 0) {
      const queryString = Object.entries(queryParams)
        .map(([key, value]) => `${key}=${encodeURIComponent(value).replace(/%2C/gi, ',')}`)
        .join('&');
      path = `${path}${path.includes('?') ? '&' : '?'}${queryString}`;
    }

    const options: {
      method: string;
      headers: Record<string, string>;
      body?: string;
      rawResponse?: boolean;
      includeHeaders?: boolean;
      excludeResponse?: boolean;
      queryParams?: Record<string, string>;
      accessToken?: string;
    } = {
      method: tool.method.toUpperCase(),
      headers,
    };

    if (options.method !== 'GET' && body) {
      if (config?.contentType === 'text/html') {
        if (typeof body === 'string') {
          options.body = body;
        } else if (typeof body === 'object' && 'content' in body) {
          options.body = (body as { content: string }).content;
        } else {
          options.body = String(body);
        }
      } else {
        options.body = typeof body === 'string' ? body : JSON.stringify(body);
      }
    }

    const isProbablyMediaContent =
      tool.errors?.some((error) => error.description === 'Retrieved media content') ||
      path.endsWith('/content');

    if (config?.returnDownloadUrl && path.endsWith('/content')) {
      path = path.replace(/\/content$/, '');
      logger.info(
        `Auto-returning download URL for ${tool.alias} (returnDownloadUrl=true in endpoints.json)`
      );
    } else if (isProbablyMediaContent) {
      options.rawResponse = true;
    }

    // Set includeHeaders if requested
    if (params.includeHeaders === true) {
      options.includeHeaders = true;
    }

    // Set excludeResponse if requested
    if (params.excludeResponse === true) {
      options.excludeResponse = true;
    }

    // Pass account-resolved token if available
    if (accountAccessToken) {
      options.accessToken = accountAccessToken;
    }

    // Redact accessToken from log output to prevent credential leakage
    const { accessToken: _redacted, ...safeOptions } = options;
    logger.info(
      `Making graph request to ${path} with options: ${JSON.stringify(safeOptions)}${_redacted ? ' [accessToken=REDACTED]' : ''}`
    );

    let response = await graphClient.graphRequest(path, options);

    const fetchAllPages = params.fetchAllPages === true;
    if (fetchAllPages && response?.content?.[0]?.text) {
      try {
        let combinedResponse = JSON.parse(response.content[0].text);
        let allItems = combinedResponse.value || [];
        let nextLink = combinedResponse['@odata.nextLink'];
        let pageCount = 1;
        const maxPages = 100;
        const maxItems = 10_000;

        let deltaLink: string | undefined = combinedResponse['@odata.deltaLink'];
        let bailedOut: 'maxPages' | 'maxItems' | null = null;
        while (nextLink) {
          // Check caps BEFORE fetching the next page so the loop exits cleanly
          // and the last-known nextLink is preserved for caller resume. Previously
          // the cap check was in the while condition, which combined with the
          // unconditional `delete @odata.nextLink` below stranded callers mid-sync
          // (value array returned, but no resume token).
          if (pageCount >= maxPages) {
            bailedOut = 'maxPages';
            break;
          }
          if (allItems.length >= maxItems) {
            bailedOut = 'maxItems';
            break;
          }

          logger.info(`Fetching page ${pageCount + 1} from: ${nextLink}`);

          // Extract path + query string from the nextLink URL.
          // Pass the full path (with query string) as the endpoint so that
          // $skiptoken and other pagination params are preserved.
          // Previously, query params were extracted into nextOptions.queryParams
          // but graphRequest/performRequest never read that field — they were lost.
          const url = new URL(nextLink);
          const nextPath = url.pathname.replace('/v1.0', '') + url.search;
          const nextOptions = { ...options };

          const nextResponse = await graphClient.graphRequest(nextPath, nextOptions);
          if (nextResponse?.content?.[0]?.text) {
            const nextJsonResponse = JSON.parse(nextResponse.content[0].text);
            if (nextJsonResponse.value && Array.isArray(nextJsonResponse.value)) {
              allItems = allItems.concat(nextJsonResponse.value);
            }
            nextLink = nextJsonResponse['@odata.nextLink'];
            if (nextJsonResponse['@odata.deltaLink']) {
              deltaLink = nextJsonResponse['@odata.deltaLink'];
            }
            pageCount++;
          } else {
            break;
          }
        }

        if (bailedOut) {
          logger.warn(
            `fetchAllPages hit cap '${bailedOut}' (pages=${pageCount}/${maxPages}, ` +
              `items=${allItems.length}/${maxItems}); preserving @odata.nextLink for caller resume.`
          );
        }

        combinedResponse.value = allItems;
        if (combinedResponse['@odata.count']) {
          combinedResponse['@odata.count'] = allItems.length;
        }

        // Only strip @odata.nextLink if the loop exited because Graph stopped
        // emitting it. If we bailed out due to caps, preserve the last-known
        // nextLink so the caller can resume.
        if (bailedOut && nextLink) {
          combinedResponse['@odata.nextLink'] = nextLink;
        } else {
          delete combinedResponse['@odata.nextLink'];
        }

        // Carry the @odata.deltaLink from the final page so callers can resume
        // a delta sync. Without this, fetchAllPages on a /delta endpoint silently
        // drops the resume token and forces callers to re-list from scratch.
        if (deltaLink) {
          combinedResponse['@odata.deltaLink'] = deltaLink;
        }

        // Diagnostic marker so callers can detect partial results and the reason.
        if (bailedOut) {
          combinedResponse['_tsq_pagingBailedOut'] = {
            reason: bailedOut,
            pageCount,
            itemCount: allItems.length,
            maxPages,
            maxItems,
            note: 'Returned partial results due to pagination cap. Use @odata.nextLink to resume.',
          };
        }

        response.content[0].text = JSON.stringify(combinedResponse);

        logger.info(
          `Pagination complete: collected ${allItems.length} items across ${pageCount} pages` +
            (bailedOut ? ` (bailed out on ${bailedOut})` : '')
        );
      } catch (e) {
        logger.error(`Error during pagination: ${e}`);
      }
    }

    if (response?.content?.[0]?.text) {
      const responseText = response.content[0].text;
      logger.info(`Response size: ${responseText.length} characters`);

      try {
        const jsonResponse = JSON.parse(responseText);
        if (jsonResponse.value && Array.isArray(jsonResponse.value)) {
          logger.info(`Response contains ${jsonResponse.value.length} items`);
        }
        if (jsonResponse['@odata.nextLink']) {
          logger.info(`Response has pagination nextLink: ${jsonResponse['@odata.nextLink']}`);
        }
      } catch {
        // Non-JSON response
      }
    }

    // Convert McpResponse to CallToolResult with the correct structure
    const content: ContentItem[] = response.content.map((item) => ({
      type: 'text' as const,
      text: item.text,
    }));

    return {
      content,
      _meta: response._meta,
      isError: response.isError,
    };
  } catch (error) {
    logger.error(`Error in tool ${tool.alias}: ${(error as Error).message}`);
    return {
      content: [
        {
          type: 'text',
          text: JSON.stringify({
            error: `Error in tool ${tool.alias}: ${(error as Error).message}`,
          }),
        },
      ],
      isError: true,
    };
  }
}

export function registerGraphTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = []
): number {
  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
      logger.info(`Tool filtering enabled with pattern: ${enabledToolsPattern}`);
    } catch {
      logger.error(`Invalid tool filter regex pattern: ${enabledToolsPattern}. Ignoring filter.`);
    }
  }

  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);
    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      logger.info(`Skipping work account tool ${tool.alias} - not in org mode`);
      skippedCount++;
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      // Allow POST endpoints that are explicitly marked as readOnly in endpoints.json
      // (e.g. get-schedule, find-meeting-times which are read-only queries via POST).
      // PATCH/DELETE are always blocked in read-only mode.
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        logger.info(`Skipping write operation ${tool.alias} in read-only mode`);
        skippedCount++;
        continue;
      }
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(tool.alias)) {
      logger.info(`Skipping tool ${tool.alias} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }

    const paramSchema: Record<string, z.ZodTypeAny> = {};
    if (tool.parameters && tool.parameters.length > 0) {
      for (const param of tool.parameters) {
        paramSchema[param.name] = param.schema || z.any();
      }
    }

    // Graph rejects $top on /delta() endpoints (HTTP 400). Page size is controlled
    // internally by Graph or via the Prefer: odata.maxpagesize header (not exposed
    // by this server). Strip $top/top from delta tool schemas so the model never
    // sends it.
    const isDeltaTool = tool.alias.endsWith('-delta');
    if (isDeltaTool) {
      delete paramSchema['top'];
      delete paramSchema['$top'];
    }

    // Extract path parameters from the path pattern (e.g., :todoTaskListId from /me/todo/lists/:todoTaskListId/tasks)
    // The generated client omits these from tool.parameters, so we add them manually.
    const pathParamMatches = tool.path.matchAll(/:([a-zA-Z]+)/g);
    for (const match of pathParamMatches) {
      const pathParamName = match[1];
      if (!(pathParamName in paramSchema)) {
        paramSchema[pathParamName] = z.string().describe(`Path parameter: ${pathParamName}`);
      }
    }

    if (tool.method.toUpperCase() === 'GET' && tool.path.includes('/')) {
      paramSchema['fetchAllPages'] = z
        .boolean()
        .describe(
          'Follow @odata.nextLink and merge up to 100 pages into one response. ' +
            'Can return enormous payloads—only when the user explicitly needs a full export. ' +
            'Prefer a small $top first, then paginate or narrow with $filter/$search.'
        )
        .optional();
    }

    // Override OData parameter descriptions with spec-gap guidance
    if (paramSchema['filter'] !== undefined || paramSchema['$filter'] !== undefined) {
      const key = paramSchema['$filter'] !== undefined ? '$filter' : 'filter';
      paramSchema[key] = z
        .string()
        .describe(
          'OData filter expression. Add $count=true for advanced filters (flag/flagStatus, contains()). Cannot combine with $search.'
        )
        .optional();
    }
    if (paramSchema['search'] !== undefined || paramSchema['$search'] !== undefined) {
      const key = paramSchema['$search'] !== undefined ? '$search' : 'search';
      paramSchema[key] = z
        .string()
        .describe('KQL search query — wrap value in double quotes. Cannot combine with $filter.')
        .optional();
    }
    if (paramSchema['select'] !== undefined || paramSchema['$select'] !== undefined) {
      const key = paramSchema['$select'] !== undefined ? '$select' : 'select';
      paramSchema[key] = z
        .string()
        .describe('Comma-separated fields to return, e.g. id,subject,from,receivedDateTime')
        .optional();
    }
    if (paramSchema['orderby'] !== undefined || paramSchema['$orderby'] !== undefined) {
      const key = paramSchema['$orderby'] !== undefined ? '$orderby' : 'orderby';
      paramSchema[key] = z
        .string()
        .describe('Sort expression, e.g. receivedDateTime desc')
        .optional();
    }
    if (paramSchema['top'] !== undefined || paramSchema['$top'] !== undefined) {
      const key = paramSchema['$top'] !== undefined ? '$top' : 'top';
      paramSchema[key] = z
        .number()
        .describe(
          'Page size (Graph $top). Start small (e.g. 5–15) so responses fit the model context; ' +
            'raise only if needed. Use $select to return fewer fields per item. ' +
            'For more rows, use @odata.nextLink from the response instead of a very large $top.'
        )
        .optional();
    }
    if (paramSchema['skip'] !== undefined || paramSchema['$skip'] !== undefined) {
      const key = paramSchema['$skip'] !== undefined ? '$skip' : 'skip';
      paramSchema[key] = z
        .number()
        .describe('Items to skip for pagination. Not supported with $search.')
        .optional();
    }
    if (paramSchema['count'] !== undefined || paramSchema['$count'] !== undefined) {
      const countKey = paramSchema['$count'] !== undefined ? '$count' : 'count';
      paramSchema[countKey] = z
        .boolean()
        .describe(
          'Set true to enable advanced query mode (ConsistencyLevel: eventual). Required for complex $filter on flag/flagStatus or contains().'
        )
        .optional();
    }

    // Add account parameter for multi-account mode.
    // Layer 2: Account names are surfaced in the description (not as a strict enum) so the LLM
    // sees available accounts upfront without a round-trip, but accounts added mid-session via
    // --login are still accepted — getTokenForAccount() handles validation at runtime.
    if (multiAccount) {
      const accountHint =
        accountNames.length > 0 ? `Known accounts: ${accountNames.join(', ')}. ` : '';
      paramSchema['account'] = z
        .string()
        .describe(
          `${accountHint}Microsoft account email to use for this request. ` +
            `Required when multiple accounts are configured. ` +
            `Use the list-accounts tool to discover all currently available accounts.`
        )
        .optional();
    }

    // Add includeHeaders parameter for all tools to capture ETags and other headers
    paramSchema['includeHeaders'] = z
      .boolean()
      .describe('Include response headers (including ETag) in the response metadata')
      .optional();

    // Add excludeResponse parameter to only return success/failure indication
    paramSchema['excludeResponse'] = z
      .boolean()
      .describe('Exclude the full response body and only return success or failure indication')
      .optional();

    // Add timezone parameter for calendar endpoints that support it
    if (endpointConfig?.supportsTimezone) {
      paramSchema['timezone'] = z
        .string()
        .describe(
          'IANA timezone name (e.g., "America/New_York", "Europe/London", "Asia/Tokyo") for calendar event times. If not specified, times are returned in UTC.'
        )
        .optional();
    }

    // Add expandExtendedProperties parameter for calendar endpoints that support it
    if (endpointConfig?.supportsExpandExtendedProperties) {
      paramSchema['expandExtendedProperties'] = z
        .boolean()
        .describe(
          'When true, expands singleValueExtendedProperties on each event. Use this to retrieve custom extended properties (e.g., sync metadata) stored on calendar events.'
        )
        .optional();
    }

    // Build the tool description, optionally appending LLM tips
    let toolDescription =
      tool.description || `Execute ${tool.method.toUpperCase()} request to ${tool.path}`;
    if (endpointConfig?.llmTip) {
      toolDescription += `\n\n💡 TIP: ${endpointConfig.llmTip}`;
    }

    try {
      server.tool(
        tool.alias,
        toolDescription,
        paramSchema,
        {
          title: tool.alias,
          readOnlyHint: tool.method.toUpperCase() === 'GET',
          destructiveHint: ['POST', 'PATCH', 'DELETE'].includes(tool.method.toUpperCase()),
          openWorldHint: true, // All tools call Microsoft Graph API
        },
        async (params) => executeGraphTool(tool, endpointConfig, graphClient, params, authManager)
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool ${tool.alias}: ${(error as Error).message}`);
      failedCount++;
    }
  }

  if (multiAccount) {
    logger.info('Multi-account mode: "account" parameter injected into all tool schemas');
  }

  // Register parse-teams-url utility tool (no Graph API call)
  if (!enabledToolsRegex || enabledToolsRegex.test('parse-teams-url')) {
    try {
      server.tool(
        'parse-teams-url',
        'Converts any Teams meeting URL format (short /meet/, full /meetup-join/, or recap ?threadId=) into a standard joinWebUrl. Use this before list-online-meetings when the user provides a recap or short URL.',
        {
          url: z.string().describe('Teams meeting URL in any format'),
        },
        {
          title: 'parse-teams-url',
          readOnlyHint: true,
          openWorldHint: false,
        },
        async ({ url }) => {
          try {
            const joinWebUrl = parseTeamsUrl(url);
            return { content: [{ type: 'text', text: joinWebUrl }] };
          } catch (error) {
            return {
              content: [
                { type: 'text', text: JSON.stringify({ error: (error as Error).message }) },
              ],
              isError: true,
            };
          }
        }
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool parse-teams-url: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // list-conversation-messages and list-drafts are registered as custom tools because
  // their Graph URLs collide with existing endpoints.json paths
  // (/me/messages and /me/mailFolders/{id}/messages). The simplified-openapi
  // generator collapses path+method duplicates into a single operationId, so a
  // second endpoints.json entry on the same path would never reach the client.
  // Both delegate to executeGraphTool with a synthetic Endpoint so they inherit
  // pagination, auth, OData translation, and $top clamping.

  const mailListReadOnlyHints = {
    readOnlyHint: true,
    destructiveHint: false,
    openWorldHint: true,
  };

  if (!enabledToolsRegex || enabledToolsRegex.test('list-conversation-messages')) {
    try {
      const conversationMessagesTool = {
        alias: 'list-conversation-messages',
        method: 'get' as const,
        path: '/me/messages',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'filter', type: 'Query' as const, schema: z.string() },
          { name: 'select', type: 'Query' as const, schema: z.string() },
          { name: 'orderby', type: 'Query' as const, schema: z.string() },
          { name: 'top', type: 'Query' as const, schema: z.number() },
          { name: 'count', type: 'Query' as const, schema: z.boolean() },
          { name: 'ConsistencyLevel', type: 'Header' as const, schema: z.string() },
        ],
        response: z.any(),
      };
      const conversationMessagesConfig: EndpointConfig = {
        pathPattern: '/me/messages',
        method: 'get',
        toolName: 'list-conversation-messages',
        scopes: ['Mail.Read'],
        workScopes: ['Mail.Read'],
      };

      server.tool(
        'list-conversation-messages',
        'Lists every message in a single email conversation thread by conversationId. ' +
          "The conversationId is stable across subject line changes — get it from any " +
          "message's conversationId field (get-mail-message, list-mail-messages, etc.). " +
          "Backed by /me/messages with $filter=conversationId eq '{id}' (the server adds " +
          'this filter for you, so callers must NOT pass their own filter parameter). ' +
          'Always pass `select` to limit returned fields ' +
          '(recommended: id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId). ' +
          'To sort results, pass `orderby` (e.g. "receivedDateTime asc" or ' +
          '"receivedDateTime desc"). Microsoft Graph rejects $filter=conversationId + ' +
          '$orderby even in advanced query mode (InefficientFilter), so the server sorts ' +
          'the returned page(s) client-side. For correct sort order across long threads, ' +
          'also pass `fetchAllPages: true` so the full set is sorted, not just the first page. ' +
          "The `count: true` parameter enables Graph's advanced query mode (ConsistencyLevel: " +
          'eventual + $count=true). It is NOT required for orderby (the server handles that ' +
          'client-side) but may be useful when combined with other advanced query features.',
        {
          conversationId: z
            .string()
            .describe(
              "The conversationId to filter by. Read it from any message's conversationId field " +
                '(get-mail-message, list-mail-messages, etc.).'
            ),
          select: z
            .string()
            .optional()
            .describe(
              'Comma-separated fields to return, e.g. id,subject,from,toRecipients,receivedDateTime,bodyPreview,conversationId. ' +
                'Always set this to keep responses small.'
            ),
          orderby: z
            .string()
            .optional()
            .describe(
              "Sort expression, e.g. 'receivedDateTime asc' or 'receivedDateTime desc'. " +
                'The server sorts client-side because Graph rejects $filter + $orderby on /me/messages. ' +
                'For complete sort across long threads, also pass fetchAllPages: true.'
            ),
          top: z
            .number()
            .optional()
            .describe(
              'Page size (Graph $top). Start small (5–15) so responses fit context; raise only if needed.'
            ),
          count: z
            .boolean()
            .optional()
            .describe(
              'Set true to enable Graph advanced query mode (ConsistencyLevel: eventual + $count=true). ' +
                'NOT required for orderby (the server sorts client-side). Useful when combining with other advanced query features.'
            ),
          fetchAllPages: z
            .boolean()
            .optional()
            .describe(
              'Follow @odata.nextLink and merge up to 100 pages into one response. ' +
                'Can return enormous payloads — only when the user explicitly needs the whole thread.'
            ),
        },
        {
          title: 'list-conversation-messages',
          ...mailListReadOnlyHints,
        },
        async (params) => {
          const conversationId = String(params.conversationId ?? '').trim();
          if (!conversationId) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    error: 'conversationId is required and must be non-empty.',
                  }),
                },
              ],
              isError: true,
            };
          }
          // Capture the requested orderby BEFORE we strip it from the outbound call.
          // Graph rejects $filter=conversationId + $orderby with InefficientFilter even
          // in advanced query mode, so we sort client-side after fetch.
          const requestedOrderby = params.orderby;

          // Escape single quotes inside the OData string literal by doubling them.
          const escapedId = conversationId.replace(/'/g, "''");
          const callParams: Record<string, unknown> = {
            filter: `conversationId eq '${escapedId}'`,
          };
          if (params.select !== undefined) callParams.select = params.select;
          // NOTE: deliberately NOT setting orderby on callParams — see comment above.
          if (params.top !== undefined) callParams.top = params.top;
          if (params.count === true) {
            callParams.count = true;
            callParams.ConsistencyLevel = 'eventual';
          }
          if (params.fetchAllPages !== undefined) callParams.fetchAllPages = params.fetchAllPages;

          const result = await executeGraphTool(
            conversationMessagesTool,
            conversationMessagesConfig,
            graphClient,
            callParams,
            authManager
          );

          if (!requestedOrderby) {
            return result;
          }

          try {
            const textBlock = result.content?.[0];
            if (!textBlock || textBlock.type !== 'text' || typeof textBlock.text !== 'string') {
              return result;
            }
            const parsed = JSON.parse(textBlock.text);
            if (!parsed || !Array.isArray(parsed.value)) {
              return result;
            }

            const orderbyParts = requestedOrderby.trim().split(/\s+/);
            const field = orderbyParts[0];
            const direction = (orderbyParts[1] || 'asc').toLowerCase();
            const ascending = direction !== 'desc';

            parsed.value.sort((a: Record<string, unknown>, b: Record<string, unknown>) => {
              const aVal = a?.[field];
              const bVal = b?.[field];
              if (aVal === undefined || aVal === null) return ascending ? 1 : -1;
              if (bVal === undefined || bVal === null) return ascending ? -1 : 1;
              if (aVal < bVal) return ascending ? -1 : 1;
              if (aVal > bVal) return ascending ? 1 : -1;
              return 0;
            });

            parsed._tsq_clientSorted = {
              field,
              direction: ascending ? 'asc' : 'desc',
              note: 'Sorted client-side by the MCP server because Graph rejects this $filter + $orderby combination with InefficientFilter, even in advanced query mode.',
            };

            return {
              ...result,
              content: [
                {
                  ...textBlock,
                  text: JSON.stringify(parsed),
                },
              ],
            };
          } catch (err) {
            logger.warn(
              `Failed to client-side sort list-conversation-messages result: ${(err as Error).message}`
            );
            return result;
          }
        }
      );
      registeredCount++;
    } catch (error) {
      logger.error(
        `Failed to register tool list-conversation-messages: ${(error as Error).message}`
      );
      failedCount++;
    }
  }

  if (!enabledToolsRegex || enabledToolsRegex.test('list-drafts')) {
    try {
      const draftsTool = {
        alias: 'list-drafts',
        method: 'get' as const,
        path: '/me/mailFolders/drafts/messages',
        requestFormat: 'json' as const,
        parameters: [
          { name: 'select', type: 'Query' as const, schema: z.string() },
          { name: 'orderby', type: 'Query' as const, schema: z.string() },
          { name: 'top', type: 'Query' as const, schema: z.number() },
        ],
        response: z.any(),
      };
      const draftsConfig: EndpointConfig = {
        pathPattern: '/me/mailFolders/drafts/messages',
        method: 'get',
        toolName: 'list-drafts',
        scopes: ['Mail.Read'],
        workScopes: ['Mail.Read'],
      };

      server.tool(
        'list-drafts',
        "Lists unsent draft emails from the user's Drafts folder. Backed by " +
          '/me/mailFolders/drafts/messages (well-known folder name). Useful for auditing ' +
          'unsent drafts or resuming a draft started earlier. Always pass `select` to keep ' +
          'responses small (recommended: id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview). ' +
          'Use orderby=lastModifiedDateTime desc to see most recently edited drafts first.',
        {
          select: z
            .string()
            .optional()
            .describe(
              'Comma-separated fields to return, e.g. id,subject,toRecipients,createdDateTime,lastModifiedDateTime,bodyPreview.'
            ),
          orderby: z
            .string()
            .optional()
            .describe(
              "Sort expression, e.g. 'lastModifiedDateTime desc' or 'createdDateTime desc'."
            ),
          top: z
            .number()
            .optional()
            .describe(
              'Page size (Graph $top). Start small (5–15) so responses fit context; raise only if needed.'
            ),
          fetchAllPages: z
            .boolean()
            .optional()
            .describe(
              'Follow @odata.nextLink and merge up to 100 pages into one response. Only use if the user needs every draft.'
            ),
        },
        {
          title: 'list-drafts',
          ...mailListReadOnlyHints,
        },
        async (params) => {
          const callParams: Record<string, unknown> = {};
          if (params.select !== undefined) callParams.select = params.select;
          if (params.orderby !== undefined) callParams.orderby = params.orderby;
          if (params.top !== undefined) callParams.top = params.top;
          if (params.fetchAllPages !== undefined) callParams.fetchAllPages = params.fetchAllPages;
          return executeGraphTool(draftsTool, draftsConfig, graphClient, callParams, authManager);
        }
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool list-drafts: ${(error as Error).message}`);
      failedCount++;
    }
  }

  if (!enabledToolsRegex || enabledToolsRegex.test('get-messages-batch')) {
    try {
      server.tool(
        'get-messages-batch',
        'Fetches multiple email messages in a single Graph /$batch request. ' +
          'Much faster than calling get-mail-message N times when you need ' +
          'full bodies for several messages (e.g., daily digests, thread ' +
          'summaries, audit reports). Up to 20 messages per call (Graph $batch limit). ' +
          'Pass `select` to limit returned fields and keep payloads small ' +
          '(recommended: id,subject,from,toRecipients,receivedDateTime,bodyPreview,body). ' +
          'Returns { results, failures, summary }: results[] holds the successful ' +
          'sub-responses each tagged with their messageId; failures[] holds any ' +
          'sub-requests Graph returned non-2xx for (e.g. unknown message ID). The ' +
          'outer call still returns HTTP 200 — Graph processes the batch and ' +
          'returns partial success.',
        {
          messageIds: z
            .array(z.string())
            .min(1)
            .max(20)
            .describe(
              'List of message IDs to fetch. Maximum 20 per call (Graph $batch limit). ' +
                'Get IDs from list-mail-messages, list-conversation-messages, search-query, etc.'
            ),
          select: z
            .string()
            .optional()
            .describe(
              'Comma-separated fields to return for each message. Strongly recommended ' +
                'to keep responses small. Example: id,subject,from,receivedDateTime,body.'
            ),
        },
        {
          title: 'get-messages-batch',
          ...mailListReadOnlyHints,
        },
        async (params) => {
          const messageIds = params.messageIds.filter((id) => typeof id === 'string' && id.trim());
          if (messageIds.length === 0) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({ error: 'messageIds must contain at least one non-empty ID.' }),
                },
              ],
              isError: true,
            };
          }
          if (messageIds.length > 20) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({ error: 'messageIds limited to 20 per call (Graph $batch limit).' }),
                },
              ],
              isError: true,
            };
          }

          // Build the per-message URL. Graph $batch URLs are relative to /v1.0 and
          // may include a query string. encodeURIComponent escapes '=' in message
          // IDs to %3D — Graph accepts both forms but unescaping keeps logs readable.
          const selectClause = params.select
            ? `?$select=${encodeURIComponent(params.select).replace(/%2C/gi, ',')}`
            : '';
          const subRequests = messageIds.map((id, idx) => ({
            id: String(idx + 1),
            method: 'GET',
            url: `/me/messages/${encodeURIComponent(id).replace(/%3D/g, '=')}${selectClause}`,
          }));

          try {
            const batchResponse = await graphClient.batchRequest(subRequests);

            const results: Array<{ messageId: string; data: unknown }> = [];
            const failures: Array<{ messageId: string; status: number; error: unknown }> = [];

            for (const subResp of batchResponse.responses ?? []) {
              const idx = parseInt(subResp.id, 10) - 1;
              const messageId = messageIds[idx];
              if (subResp.status >= 200 && subResp.status < 300) {
                results.push({ messageId, data: subResp.body });
              } else {
                failures.push({ messageId, status: subResp.status, error: subResp.body });
              }
            }

            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    results,
                    failures,
                    summary: {
                      total: messageIds.length,
                      succeeded: results.length,
                      failed: failures.length,
                    },
                  }),
                },
              ],
            };
          } catch (err) {
            logger.error(`get-messages-batch failed: ${(err as Error).message}`);
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({ error: (err as Error).message }),
                },
              ],
              isError: true,
            };
          }
        }
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool get-messages-batch: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // download-mail-attachment writes to OneDrive (staging upload), so it is a
  // write tool — skip it entirely in read-only mode, mirroring the endpoint
  // loop's `readOnly && method !== 'GET'` gate above.
  if (readOnly) {
    logger.info('Skipping write tool download-mail-attachment - read-only mode');
    skippedCount++;
  } else if (!enabledToolsRegex || enabledToolsRegex.test('download-mail-attachment')) {
    try {
      // Simple OneDrive upload (PUT .../content) handles a single file up to
      // 250 MB, which covers every mail attachment (Exchange caps well below
      // this), so we avoid the complexity of chunked upload sessions.
      const MAX_ATTACHMENT_BYTES = 250 * 1024 * 1024;
      server.tool(
        'download-mail-attachment',
        'Downloads a mail attachment via OneDrive staging — returns a pre-authed ' +
          'download URL plus metadata, NOT the file bytes themselves. Use this for ' +
          'any attachment large enough that returning base64 through get-mail-attachment ' +
          'would overflow context (realistically anything over ~50KB).\n\n' +
          'Flow: (1) list-mail-attachments to find the attachmentId; ' +
          '(2) download-mail-attachment to stage the file to OneDrive and get a URL; ' +
          '(3) fetch the URL directly (curl/HTTP) to wherever you need the file.\n\n' +
          'Returns { downloadUrl, name, size, contentType, sha256, expiresInSeconds, stagedDriveItemId }. ' +
          'The downloadUrl is short-lived (~1 hour) and pre-authenticated — no Authorization ' +
          'header needed when fetching it. The bytes are NEVER returned through this tool, so ' +
          'model context is unaffected by attachment size. Only supports fileAttachment; for ' +
          'itemAttachment or referenceAttachment use get-mail-attachment instead. Verify the ' +
          'download with the returned sha256.',
        {
          messageId: z
            .string()
            .min(1)
            .describe('The ID of the mail message that owns the attachment.'),
          attachmentId: z
            .string()
            .min(1)
            .describe(
              'The ID of the attachment to download. Get it from list-mail-attachments.'
            ),
          userId: z
            .string()
            .optional()
            .describe(
              "For shared/other mailboxes: the user id or UPN whose message this is. " +
                "Omit (or pass 'me') for the signed-in user's own mailbox. The staged " +
                "copy is always written to the signed-in user's OneDrive."
            ),
        },
        {
          title: 'download-mail-attachment',
          readOnlyHint: false,
          destructiveHint: false,
          openWorldHint: true,
        },
        async (params) => {
          const messageId = (params.messageId ?? '').trim();
          const attachmentId = (params.attachmentId ?? '').trim();
          if (!messageId || !attachmentId) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    error: 'messageId and attachmentId are both required and must be non-empty.',
                  }),
                },
              ],
              isError: true,
            };
          }

          const userId = (params.userId ?? '').trim();
          const mailBase =
            userId && userId.toLowerCase() !== 'me'
              ? `/users/${encodeURIComponent(userId)}`
              : '/me';
          const attachmentPath = `${mailBase}/messages/${encodeURIComponent(
            messageId
          )}/attachments/${encodeURIComponent(attachmentId)}`;

          try {
            // 1. Metadata first (no contentBytes — $select keeps it small and
            //    lets us reject non-file attachments before pulling any bytes).
            const metadata = (await graphClient.makeRequest(
              `${attachmentPath}?$select=id,name,contentType,size,isInline`
            )) as {
              '@odata.type'?: string;
              name?: string;
              contentType?: string;
              size?: number;
            };

            const odataType = metadata['@odata.type'];
            if (odataType && odataType !== '#microsoft.graph.fileAttachment') {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: `download-mail-attachment only supports fileAttachment, but this is ${odataType}. Use get-mail-attachment for itemAttachment, or follow the @odata reference for referenceAttachment.`,
                    }),
                  },
                ],
                isError: true,
              };
            }

            if (typeof metadata.size === 'number' && metadata.size > MAX_ATTACHMENT_BYTES) {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: `Attachment is ${metadata.size} bytes, which exceeds the ${MAX_ATTACHMENT_BYTES}-byte (250 MB) staging limit.`,
                    }),
                  },
                ],
                isError: true,
              };
            }

            // 2. Stream the raw bytes from /$value (server-side only — never
            //    serialized into the tool response).
            const { buffer, contentType: rawContentType } = await graphClient.fetchBinary(
              `${attachmentPath}/$value`
            );
            if (buffer.byteLength > MAX_ATTACHMENT_BYTES) {
              return {
                content: [
                  {
                    type: 'text',
                    text: JSON.stringify({
                      error: `Attachment is ${buffer.byteLength} bytes, which exceeds the ${MAX_ATTACHMENT_BYTES}-byte (250 MB) staging limit.`,
                    }),
                  },
                ],
                isError: true,
              };
            }
            const contentType =
              metadata.contentType || rawContentType || 'application/octet-stream';

            // 3. Build a collision-safe staging filename and upload to a marked
            //    folder under the user's OneDrive root. Using /drive/root (not
            //    special/approot) because the granted scope is Files.ReadWrite,
            //    not Files.ReadWrite.AppFolder — approot is the AppFolder construct
            //    and can 404 under the broad scope. The path auto-creates the folder.
            const rawName = (metadata.name && String(metadata.name)) || `attachment-${attachmentId}`;
            const safeName =
              rawName.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim() || 'attachment';
            const uniquePrefix = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
            const stagedName = `${uniquePrefix}-${safeName}`;
            const uploadPath = `/me/drive/root:/Apps/TSQ-M365-MCP-staging/${encodeURIComponent(
              stagedName
            )}:/content`;

            const driveItem = (await graphClient.putBinary(
              uploadPath,
              buffer,
              contentType
            )) as { id?: string; '@microsoft.graph.downloadUrl'?: string };

            const stagedDriveItemId = driveItem.id;
            let downloadUrl = driveItem['@microsoft.graph.downloadUrl'];

            // The PUT response often omits the downloadUrl; fetch it explicitly
            // from the item if so.
            if (!downloadUrl && stagedDriveItemId) {
              const fetched = (await graphClient.makeRequest(
                `/me/drive/items/${encodeURIComponent(stagedDriveItemId)}`
              )) as { '@microsoft.graph.downloadUrl'?: string };
              downloadUrl = fetched['@microsoft.graph.downloadUrl'];
            }

            const sha256 = createHash('sha256').update(buffer).digest('hex');

            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    downloadUrl,
                    name: rawName,
                    size: buffer.byteLength,
                    contentType,
                    sha256,
                    expiresInSeconds: 3600,
                    stagedDriveItemId,
                    note: 'The downloadUrl is pre-authed and short-lived (~1 hour). Fetch it directly via curl/HTTP to download the bytes to your own disk — no Authorization header needed. The bytes are NOT returned through this tool. Verify the download against sha256.',
                  }),
                },
              ],
            };
          } catch (err) {
            logger.error(`download-mail-attachment failed: ${(err as Error).message}`);
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({ error: (err as Error).message }),
                },
              ],
              isError: true,
            };
          }
        }
      );
      registeredCount++;
    } catch (error) {
      logger.error(`Failed to register tool download-mail-attachment: ${(error as Error).message}`);
      failedCount++;
    }
  }

  // Layer 3 (list-accounts tool) is registered by registerAuthTools in auth-tools.ts.
  // It is the canonical owner of account discovery — no duplicate registration here.

  logger.info(
    `Tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}

export function buildToolsRegistry(
  readOnly: boolean,
  orgMode: boolean
): Map<string, { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }> {
  const toolsMap = new Map<
    string,
    { tool: (typeof api.endpoints)[0]; config: EndpointConfig | undefined }
  >();

  for (const tool of api.endpoints) {
    const endpointConfig = endpointsData.find((e) => e.toolName === tool.alias);

    if (!orgMode && endpointConfig && !endpointConfig.scopes && endpointConfig.workScopes) {
      continue;
    }

    const method = tool.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      if (!(method === 'POST' && endpointConfig?.readOnly)) {
        continue;
      }
    }

    toolsMap.set(tool.alias, { tool, config: endpointConfig });
  }

  return toolsMap;
}

/**
 * Builds a BM25 index over the tool registry. Name tokens are weighted 3x and llmTip
 * tokens 2x via repetition, so a tool whose name matches the query outranks one that
 * merely mentions the query term in its Microsoft-supplied description.
 */
export function buildDiscoverySearchIndex(
  toolsRegistry: ReturnType<typeof buildToolsRegistry>
): DiscoverySearchIndex {
  // Cap contribution from the `description` and `llmTip` fields so a verbose llmTip
  // (e.g. the KQL search-syntax guide on list-mail-messages, ~300 tokens) doesn't
  // inflate a tool's doc length and crush BM25's length normalization. Names and
  // paths are short and reliable, so they stay uncapped and are repeated to carry
  // the bulk of the ranking signal. Tip excerpt (12 tokens) is enough to capture
  // the first "what this tool does" phrase without swamping the doc.
  const TIP_EXCERPT_TOKENS = 12;
  const DESC_CAP_TOKENS = 40;
  const docs: Array<{ id: string; tokens: string[] }> = [];
  const nameTokens = new Map<string, Set<string>>();
  for (const [name, { tool, config }] of toolsRegistry) {
    const nt = tokenize(name);
    nameTokens.set(name, new Set(nt));
    const pathTokens = tokenize(tool.path);
    const descTokens = tokenize(tool.description).slice(0, DESC_CAP_TOKENS);
    const tipTokens = tokenize(config?.llmTip).slice(0, TIP_EXCERPT_TOKENS);
    const tokens = [
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...nt,
      ...pathTokens,
      ...pathTokens,
      ...tipTokens,
      ...descTokens,
    ];
    docs.push({ id: name, tokens });
  }
  return { bm25: buildBM25Index(docs), nameTokens };
}

/**
 * BM25 + a "name precision" bonus: reward tools whose names contain a high fraction
 * of the query tokens (and consist mostly of query-matching tokens). This counteracts
 * cases where a tool with a longer or more off-topic description outranks a tool
 * whose name directly matches — a common problem because many endpoint descriptions
 * are the wrong Graph prose pasted in.
 */
export function scoreDiscoveryQuery(
  query: string,
  index: DiscoverySearchIndex
): Array<{ id: string; score: number }> {
  const queryTokenSet = new Set(tokenize(query));
  if (queryTokenSet.size === 0) return [];
  const ranked = scoreQuery(query, index.bm25);
  const NAME_BONUS_WEIGHT = 2;
  for (const r of ranked) {
    const nt = index.nameTokens.get(r.id);
    if (!nt || nt.size === 0) continue;
    let matchedIdf = 0;
    let matchedCount = 0;
    for (const qt of queryTokenSet) {
      if (nt.has(qt)) {
        matchedCount++;
        matchedIdf += index.bm25.idf.get(qt) ?? 0;
      }
    }
    if (matchedCount === 0) continue;
    const precision = matchedCount / nt.size;
    r.score += precision * matchedIdf * NAME_BONUS_WEIGHT;
  }
  ranked.sort((a, b) => b.score - a.score);
  return ranked;
}

export function registerDiscoveryTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  orgMode: boolean = false,
  authManager?: AuthManager,
  _multiAccount: boolean = false
): void {
  const toolsRegistry = buildToolsRegistry(readOnly, orgMode);
  const searchIndex = buildDiscoverySearchIndex(toolsRegistry);
  logger.info(`Discovery mode: ${toolsRegistry.size} tools available in registry`);

  const categoryNames = Object.keys(TOOL_CATEGORIES).join(', ');

  const toResultEntry = (name: string) => {
    const entry = toolsRegistry.get(name);
    if (!entry) return null;
    const { tool, config } = entry;
    return {
      name,
      method: tool.method.toUpperCase(),
      path: tool.path,
      description: tool.description || `${tool.method.toUpperCase()} ${tool.path}`,
      ...(config?.llmTip ? { llmTip: config.llmTip } : {}),
    };
  };

  server.tool(
    'search-tools',
    `Search through ${toolsRegistry.size} Microsoft Graph API tools. Ranks results by BM25 over tool name, llmTip, description, and path (tokenized on hyphens, camelCase, and whitespace). After picking a tool, call get-tool-schema to see its parameters, then execute-tool to invoke it.`,
    {
      query: z
        .string()
        .describe(
          'Natural-language query. Tokenized and BM25-ranked. E.g. "send email", "create calendar event", "list unread messages".'
        )
        .optional(),
      category: z.string().describe(`Optional pre-filter by category: ${categoryNames}`).optional(),
      limit: z.number().describe('Maximum results (default: 10, max: 50)').optional(),
    },
    {
      title: 'search-tools',
      readOnlyHint: true,
      openWorldHint: true,
    },
    async ({ query, category, limit = 10 }) => {
      const maxLimit = Math.min(Math.max(limit, 1), 50);
      const categoryDef = category ? TOOL_CATEGORIES[category] : undefined;
      const categoryFilter = (name: string) => !categoryDef || categoryDef.pattern.test(name);

      let orderedNames: string[];
      if (query && query.trim().length > 0) {
        const ranked = scoreDiscoveryQuery(query, searchIndex);
        orderedNames = ranked.map((r) => r.id).filter(categoryFilter);
      } else {
        orderedNames = [...toolsRegistry.keys()].filter(categoryFilter);
      }

      const tools = orderedNames.slice(0, maxLimit).map(toResultEntry).filter(Boolean);

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(
              {
                found: tools.length,
                total: toolsRegistry.size,
                tools,
                tip: 'Call get-tool-schema(tool_name) to see parameters before invoking execute-tool.',
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );

  server.tool(
    'get-tool-schema',
    'Returns the full parameter schema (name, placement, required, JSON Schema) for a tool discovered via search-tools. Call this before execute-tool so you know what parameters to pass and what enum values are valid.',
    {
      tool_name: z.string().describe('Exact tool name from search-tools (e.g. "send-mail")'),
    },
    {
      title: 'get-tool-schema',
      readOnlyHint: true,
      openWorldHint: false,
    },
    async ({ tool_name }) => {
      const entry = toolsRegistry.get(tool_name);
      if (!entry) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: 'Use search-tools to find available tools.',
              }),
            },
          ],
          isError: true,
        };
      }
      const schema = describeToolSchema(entry.tool, entry.config?.llmTip);
      return {
        content: [{ type: 'text', text: JSON.stringify(schema, null, 2) }],
      };
    }
  );

  server.tool(
    'execute-tool',
    'Execute a Microsoft Graph API tool by name. Workflow: search-tools → get-tool-schema → execute-tool. Call get-tool-schema first for any tool you have not seen before — passing the wrong shape to parameters will fail validation or return a Graph 400. For list endpoints, prefer modest $top plus $select.',
    {
      tool_name: z.string().describe('Name of the tool to execute (e.g., "list-mail-messages")'),
      parameters: z
        .record(z.any())
        .describe(
          'Parameters shaped per get-tool-schema. Path/query/header params go at the top level; request bodies go under "body".'
        )
        .optional(),
    },
    {
      title: 'execute-tool',
      readOnlyHint: false,
      destructiveHint: true,
      openWorldHint: true,
    },
    async ({ tool_name, parameters = {} }) => {
      const toolData = toolsRegistry.get(tool_name);
      if (!toolData) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Tool not found: ${tool_name}`,
                tip: 'Use search-tools to find available tools.',
              }),
            },
          ],
          isError: true,
        };
      }

      return executeGraphTool(toolData.tool, toolData.config, graphClient, parameters, authManager);
    }
  );

  // Layer 3 (list-accounts) is registered by registerAuthTools — no duplicate here.
}

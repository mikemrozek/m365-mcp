import logger from '../logger.js';

/**
 * Client for chat completions against a model deployed in Microsoft Foundry
 * (Objective 10 — delegated inference, tsq.25).
 *
 * The boundary, stated once: the connector calls Foundry as ITSELF — a managed
 * identity token for the `https://ai.azure.com` audience — never as the user.
 * The model receives exactly the content the tool assembled: no tokens, no
 * tools, no Graph reach. The human caller travels only inside
 * `user_security_context` so tenant Purview/DLP can attribute the interaction;
 * whether Purview actually evaluates that field on a non-OpenAI deployment is
 * UNVERIFIED (2026-0921 Part A findings) and nothing here may assume it does.
 *
 * Request shape per the Foundry Models v1 route, proven live 2026-09-21 from
 * inside the container runtime: POST {endpoint}/openai/v1/chat/completions,
 * deployment name in the body's `model` field, no api-version parameter.
 */

export interface InferenceConfig {
  /** Foundry resource endpoint, no trailing slash (…services.ai.azure.com). */
  endpoint: string;
  /** Deployment name, sent as the body's `model` field. */
  deployment: string;
  /** Entra object ids allowed to use the tool. Empty means nobody — fail closed. */
  pilotOids: string[];
  maxInputChars: number;
  dailyTokenBudget: number;
  /**
   * Send `user_security_context` on the request. Default FALSE, from live evidence
   * (2026-09-22, first pilot night): the DeepSeek deployment on the /openai/v1 route
   * rejects it — HTTP 400 "Unrecognized request argument supplied:
   * user_security_context" — which failed every call. Isolated by A/B from inside
   * the container the same night. Consequence, stated plainly: Purview cannot
   * attribute these calls to the human until Microsoft supports the field here, so
   * the server-side pre-send check is the ONLY policy layer. Flip
   * INFERENCE_SEND_USER_CONTEXT=true only after re-testing against the live
   * deployment.
   */
  sendUserContext: boolean;
}

const DEFAULT_MAX_INPUT_CHARS = 200_000;
const DEFAULT_DAILY_TOKEN_BUDGET = 2_000_000;

export function inferenceConfigFromEnv(
  env: Record<string, string | undefined> = process.env
): InferenceConfig | undefined {
  const endpoint = env.INFERENCE_ENDPOINT?.trim();
  if (!endpoint) return undefined;

  const deployment = env.INFERENCE_DEPLOYMENT?.trim();
  if (!deployment) {
    logger.warn('INFERENCE_ENDPOINT is set but INFERENCE_DEPLOYMENT is not; inference stays off');
    return undefined;
  }

  // A malformed pilot list means an EMPTY pilot list, never "everyone".
  let pilotOids: string[] = [];
  const rawPilot = env.INFERENCE_PILOT_OIDS?.trim();
  if (rawPilot) {
    try {
      const parsed: unknown = JSON.parse(rawPilot);
      if (Array.isArray(parsed) && parsed.every((x) => typeof x === 'string')) {
        pilotOids = parsed.map((x) => x.trim()).filter(Boolean);
      } else {
        logger.warn('INFERENCE_PILOT_OIDS is not a JSON array of strings; pilot list is EMPTY');
      }
    } catch {
      logger.warn('INFERENCE_PILOT_OIDS is not valid JSON; pilot list is EMPTY (fail closed)');
    }
  }

  const positiveInt = (raw: string | undefined, fallback: number): number => {
    const n = Number.parseInt(raw ?? '', 10);
    return Number.isFinite(n) && n > 0 ? n : fallback;
  };

  return {
    endpoint: endpoint.replace(/\/+$/, ''),
    deployment,
    pilotOids,
    maxInputChars: positiveInt(env.INFERENCE_MAX_INPUT_CHARS, DEFAULT_MAX_INPUT_CHARS),
    dailyTokenBudget: positiveInt(env.INFERENCE_DAILY_TOKEN_BUDGET, DEFAULT_DAILY_TOKEN_BUDGET),
    sendUserContext: (env.INFERENCE_SEND_USER_CONTEXT ?? '').trim().toLowerCase() === 'true',
  };
}

export type FoundryRefusal =
  /** Ours: the managed identity could not mint a token for the audience. */
  | 'token_unavailable'
  /** Network failure or the 90 s timeout. */
  | 'unreachable'
  /** 429 from the provider. */
  | 'rate_limited'
  /** 401/403 — the provider rejected OUR bearer. A deployment fault, never the caller's. */
  | 'unauthorized'
  /** Any other non-2xx. */
  | 'provider_error'
  /** 2xx with a shape we do not recognise. */
  | 'unexpected';

export class FoundryError extends Error {
  constructor(
    message: string,
    readonly refusal: FoundryRefusal,
    readonly status?: number,
    readonly code?: string,
    readonly retryAfterSeconds?: number
  ) {
    super(message);
    this.name = 'FoundryError';
  }
}

/**
 * The caller, asserted to the provider so tenant DLP can attribute the call to
 * the human. Wire format is snake_case; the service silently drops misspelled
 * fields, which is why this is a typed object and not ad-hoc keys.
 */
export interface UserSecurityContext {
  application_name: string;
  /** Entra user OBJECT id — never a UPN, never anything personal. */
  end_user_id: string;
  end_user_tenant_id?: string;
}

export interface ChatUsage {
  promptTokens: number;
  completionTokens: number;
  totalTokens: number;
  /** Nonzero would be evidence prompt caching is active despite being undocumented for this model. */
  cachedTokens?: number;
}

export interface ChatResult {
  content: string;
  usage: ChatUsage;
}

export type TokenSource = (scope: string) => Promise<string>;

const TOKEN_SCOPE = 'https://ai.azure.com/.default';

export class FoundryClient {
  private readonly tokenSource: TokenSource;
  private readonly fetchImpl: typeof fetch;
  private readonly timeoutMs: number;

  constructor(
    private readonly config: Pick<InferenceConfig, 'endpoint'>,
    options: { tokenSource?: TokenSource; fetchImpl?: typeof fetch; timeoutMs?: number } = {}
  ) {
    this.tokenSource = options.tokenSource ?? defaultCredentialTokenSource();
    this.fetchImpl = options.fetchImpl ?? fetch;
    this.timeoutMs = options.timeoutMs ?? 90_000;
  }

  async chatComplete(args: {
    model: string;
    system?: string;
    user: string;
    userContext?: UserSecurityContext;
  }): Promise<ChatResult> {
    let token: string;
    try {
      token = await this.tokenSource(TOKEN_SCOPE);
    } catch (error) {
      throw new FoundryError(
        `Could not obtain a credential for the inference provider: ${(error as Error).message}`,
        'token_unavailable'
      );
    }

    const body: Record<string, unknown> = {
      model: args.model,
      messages: [
        ...(args.system ? [{ role: 'system', content: args.system }] : []),
        { role: 'user', content: args.user },
      ],
      temperature: 0.2,
    };
    if (args.userContext) body.user_security_context = args.userContext;

    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), this.timeoutMs);
    let response: Response;
    try {
      response = await this.fetchImpl(`${this.config.endpoint}/openai/v1/chat/completions`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        signal: controller.signal,
      });
    } catch (error) {
      throw new FoundryError(
        controller.signal.aborted
          ? `The inference provider did not answer within ${Math.round(this.timeoutMs / 1000)}s`
          : `The inference provider is unreachable: ${(error as Error).message}`,
        'unreachable'
      );
    } finally {
      clearTimeout(timer);
    }

    if (!response.ok) {
      // Take only the status and a symbolic code. The provider's message text can
      // echo the content we sent, and nothing that reaches a log or a tool error
      // may carry content — same bar as usage-log (92e6104).
      const parsed = await readJsonQuietly(response);
      const code = codeOf(parsed);
      logger.warn(`Inference provider refused: HTTP ${response.status}${code ? ` ${code}` : ''}`);
      if (response.status === 429) {
        const header = Number(response.headers?.get?.('Retry-After'));
        const retryAfter = Number.isFinite(header) && header > 0 ? header : undefined;
        throw new FoundryError(
          `The inference provider is rate limited (HTTP 429)${retryAfter ? `; retry after ${retryAfter}s` : ''}.`,
          'rate_limited',
          429,
          code,
          retryAfter
        );
      }
      if (response.status === 401 || response.status === 403) {
        throw new FoundryError(
          "The inference provider rejected the connector's own credential — a deployment fault (role, audience or endpoint), not something the caller can fix.",
          'unauthorized',
          response.status,
          code
        );
      }
      throw new FoundryError(
        `The inference provider returned HTTP ${response.status}${code ? ` (${code})` : ''}.`,
        'provider_error',
        response.status,
        code
      );
    }

    const parsed = await readJsonQuietly(response);
    const choices = parsed?.choices;
    const first = Array.isArray(choices) ? (choices[0] as Record<string, unknown>) : undefined;
    const message = first?.message as Record<string, unknown> | undefined;
    const content = message?.content;
    if (typeof content !== 'string') {
      throw new FoundryError(
        'The inference provider returned a response without message content.',
        'unexpected',
        response.status
      );
    }

    const usage = (parsed?.usage ?? {}) as Record<string, unknown>;
    const num = (v: unknown): number => (typeof v === 'number' && Number.isFinite(v) ? v : 0);
    const details = (usage.prompt_tokens_details ?? {}) as Record<string, unknown>;
    const cached = num(details.cached_tokens) || num(usage.cached_tokens);
    return {
      content,
      usage: {
        promptTokens: num(usage.prompt_tokens),
        completionTokens: num(usage.completion_tokens),
        totalTokens:
          num(usage.total_tokens) || num(usage.prompt_tokens) + num(usage.completion_tokens),
        ...(cached ? { cachedTokens: cached } : {}),
      },
    };
  }
}

async function readJsonQuietly(response: Response): Promise<Record<string, unknown> | undefined> {
  try {
    const parsed: unknown = await response.json();
    return parsed && typeof parsed === 'object' ? (parsed as Record<string, unknown>) : undefined;
  } catch {
    return undefined;
  }
}

/** The provider's error code as a bare token, never its message text. */
function codeOf(parsed: Record<string, unknown> | undefined): string | undefined {
  const error = parsed?.error;
  const raw =
    typeof error === 'string' ? error : (error as Record<string, unknown> | undefined)?.code;
  if (typeof raw !== 'string') return undefined;
  return raw.replace(/[^A-Za-z0-9_.:-]/g, '').slice(0, 64) || undefined;
}

/**
 * Mints tokens with Azure's credential chain — the same pattern as the Note
 * Taker client. In the container app it resolves to the user-assigned managed
 * identity (the chain honours the AZURE_CLIENT_ID env var already set there);
 * on a developer machine it falls through to `az login`.
 */
function defaultCredentialTokenSource(): TokenSource {
  let credential: { getToken(scopes: string): Promise<{ token: string } | null> } | undefined;
  return async (scope) => {
    if (!credential) {
      const { DefaultAzureCredential } = await import('@azure/identity');
      credential = new DefaultAzureCredential({});
    }
    const result = await credential.getToken(scope);
    if (!result?.token) throw new Error('credential chain returned no token');
    return result.token;
  };
}

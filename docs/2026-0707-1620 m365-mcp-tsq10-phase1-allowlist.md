# TSQ v0.88.2-tsq.10 — Phase 1: explicit tool allowlist

**Date:** 2026-07-07
**Branch:** deploy-tsq-v0.88.2
**Plan:** `docs/2026-0707-1502 m365-mcp-tiered-rollout-plan.md` (Phase 1)
**Status:** Code complete + validated locally. Config split is a DRAFT pending Scott sign-off. Not yet deployed.

## What shipped (code)

An explicit, exact-match `TOOL_ALLOWLIST` mechanism that replaces the `ENABLED_TOOLS`
keyword regex per connector tier. Backward compatible — `ENABLED_TOOLS` still works when
`TOOL_ALLOWLIST` is unset, so the current production connector is untouched.

- **`src/tool-allowlist.ts`** (new) — parses `TOOL_ALLOWLIST` into an exact-match name list.
  Accepts an inline JSON array, an inline `{"tools":[...]}` object, or a path to a file
  containing either. De-dupes + sorts. Throws on any malformed input (fail closed).
- **`src/cli.ts`** — `--tool-allowlist <value>` / `TOOL_ALLOWLIST` env. When set it wins over
  `ENABLED_TOOLS` (logs a warning if both are set) and the regex is cleared. Parse failure
  → `process.exit(1)` at startup, never a silent fall-through to exposing all tools.
- **`src/graph-tools.ts`** — single `isToolEnabled()` predicate gates the endpoint loop **and**
  all 5 hand-written custom tools (parse-teams-url, list-conversation-messages, list-drafts,
  get-messages-batch, download-mail-attachment) through the same check. Governance logging at
  startup: filter source + the exact sorted list of registered tool names.
- **`src/auth.ts`** — `buildScopesFromEndpoints()` honors the allowlist so delegated scopes are
  subset to exactly the tier's tools.
- **`src/server.ts`, `src/index.ts`** — thread `toolAllowlist` through; `--list-permissions`
  now reports the resolved allowlist.
- **`test/tool-allowlist.test.ts`** (new) — 14 tests: loader parsing/validation + registration
  filtering (exact match, no substring leak, regex ignored when allowlist present, empty list
  registers nothing, read-only still honored). All pass. No new failures in the existing suite
  (the 8 failing tests in tool-filtering/other files are pre-existing and unrelated — they
  predate get-messages-batch/download-mail-attachment).

## Config files (DRAFT — `config/allowlists/`)

All 257 org-mode tools assigned to exactly one tier (verified: no overlap, no gaps).
Each config registers **exactly** its listed tools (validated against the real
`registerGraphTools`).

| Tier | File | Tools | Scopes (delegated) |
|---|---|---:|---:|
| Core | `core.json` | 68 | 22 |
| Docs & Excel Authoring | `docs-excel.json` | 56 | — |
| Collaboration Admin | `collab-admin.json` | 72 | — |
| Mailbox & Calendar Mgmt | `mailbox-calendar.json` | 61 | — |

See `config/allowlists/README.md` for the tier rationale and the list of judgment calls to
review (reactions/pins, shared mailbox, specific-calendar tools, Planner vs To Do, subscriptions).

**Scope-minimization pass done:** `list-team-members` moved Core → Collab Admin (it was the sole
holder of the broad `TeamMember.Read.All` scope — May scope-review item), dropping Core to 22
scopes. `search-query` is the sole reason Core needs `Sites.Read.All` + `Files.Read.All` (accepted
cost of org-wide search). **Entra-app note:** Core's app registration must additionally grant
`Files.ReadWrite` for `download-mail-attachment` (custom-tool scopes aren't auto-derived).

## Measurement (headline for Scott)

Two local approximations (no tokenizer yet; ~tokens = serialized chars / 4). The
**full-schema** measure serializes the actual `tools/list` manifest entry per tool
(name + description + `inputSchema` JSON, via `zodToJsonSchema`) — i.e. what the client
really receives. The **description-only** measure (name + description + parameter
descriptions) is a conservative lower bound. A precise tokenizer count comes at the
test-deploy gate against the live manifest.

Full-schema manifest (primary):

| Set | Tools | ~Schema tokens |
|---|---:|---:|
| Full org surface (no filter) | 257 | ~303,000 |
| **Current production** (tsq.9 `ENABLED_TOOLS` regex) | **166** | **~202,000** |
| **Core allowlist** | **68** | **~69,000** |

**Core vs current production: −59% tools, −66% schema tokens (~133K tokens saved per
conversation).** Description-only lower bound agrees on direction: −55% tokens
(~42K → ~19K). The large absolute baseline (~200K tokens of tool schema on every
conversation today) is itself the case for the split.

## Test deploy (Phase 1 validation gate) — live 2026-07-10

- Image `ms-365-mcp-server:v0.88.2-tsq.10` built in ACR from commit `d232cab`.
- Test Container App **`tsqm365mcp-core-test`** in `rg-tsq-m365-mcp` / env `tsqm365mcp-cae`,
  separate from prod `tsqm365mcp-app`. Single replica, ephemeral `/tmp` token cache,
  reuses UAMI `tsqm365mcp-uami` (Key Vault + ACR) and Entra app `M365 MCP`
  (`eae48fd3-…`). Startup: `--http 3000 --org-mode -v`, `TOOL_ALLOWLIST=<inline core>`.
- URL: `https://tsqm365mcp-core-test.jollywave-3fc86824.eastus2.azurecontainerapps.io/mcp`
- The test app's `/oauth/callback` was added to the shared Entra app's redirect URIs
  (prod + claude.ai URIs preserved).
- **Live evidence the allowlist is active:** `/.well-known/oauth-protected-resource`
  advertises exactly Core's 22 scopes — **no `TeamMember.Read.All`** (confirms the
  `list-team-members` move in the real deployment). `GET /` → 200.
- Registered-tool count (68) is logged per session (`-v`); confirmed on claude.ai connect.
- **Teardown when done:** `az containerapp delete -n tsqm365mcp-core-test -g rg-tsq-m365-mcp`
  and drop the test redirect URI from the Entra app.

## Not done yet / next

- Scott review of the draft tier split (esp. the judgment calls in the README).
- Precise token measurement on a real test container app + a live claude.ai connection showing
  only Core tools (Phase 1 validation gate).
- Version bump to `v0.88.2-tsq.10`, ACR build, deploy to a **test** container app (not prod).
- Phase 2: stand up the Core connector, `request-m365-capability` v1, Entra app decision.

## How to run locally

```
# by file
TOOL_ALLOWLIST=config/allowlists/core.json  node dist/index.js --http --org-mode
# inline
TOOL_ALLOWLIST='["get-current-user","send-mail"]'  node dist/index.js --list-permissions --org-mode
```

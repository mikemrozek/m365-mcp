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

Approximation = char count (tool name + description + parameter descriptions) / 4.
Consistent across sets; a precise manifest-token count comes at the test-deploy gate.

| Set | Tools | ~Schema tokens |
|---|---:|---:|
| Full org surface (no filter) | 257 | ~64,200 |
| **Current production** (tsq.9 `ENABLED_TOOLS` regex) | **166** | **~42,300** |
| **Core allowlist** | **68** | **~18,900** |

**Core vs current production: −59% tools, −55% schema tokens (~23K tokens saved per conversation).**

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

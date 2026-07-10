# M365 MCP Tiered Connector Rollout — Phased Implementation Plan

**Date:** June 2026
**Owner:** Mike Mrozek
**Sponsor:** Scott Schatz
**Baseline:** Single connector, ~170 tool schemas loaded into every claude.ai conversation
**Target:** Core connector (~60 tools) for everyone; three Advanced connectors gated by Entra group; measured context reduction

---

## Confirmed decisions (from Scott, June 2026)

1. **Multiple deployments, one codebase.** All four connectors are the SAME server image deployed four times, each with a different allowlist config and Entra group assignment. One image, four configs — not four projects.
2. **Explicit named allowlist replaces the keyword regex.** Per-tool control, read/write separation, one reviewable source of truth per tier.
3. **All four TSQ custom tools stay in Core:** `download-mail-attachment`, `get-messages-batch`, `list-conversation-messages`, `list-drafts`.
4. **Access on-ramp v1 is simple:** `request-m365-capability` files a ticket / Teams / email notification to IT. Automated Entra access packages come in v2.
5. **Current deployment stays running** until the new ones are proven.
6. **Phased rollout with measurement:** Core first, measure actual context drop against the ~170 baseline, then Advanced tiers one at a time.

---

## Architecture: one image, four configs

```
                    ┌─────────────────────────────┐
                    │   ONE codebase, ONE image    │
                    │  ms-365-mcp-server:vX.Y.Z    │
                    │  (with allowlist support)    │
                    └──────────────┬──────────────┘
                                   │ deployed 4x
        ┌──────────────┬───────────┴──────┬──────────────────┐
        ▼              ▼                  ▼                  ▼
  ┌───────────┐  ┌────────────┐  ┌──────────────┐  ┌──────────────┐
  │   CORE    │  │ DOCS/EXCEL │  │ COLLAB ADMIN │  │ MAILBOX/CAL  │
  │           │  │ AUTHORING  │  │              │  │  MANAGEMENT  │
  ├───────────┤  ├────────────┤  ├──────────────┤  ├──────────────┤
  │ TOOL_     │  │ TOOL_      │  │ TOOL_        │  │ TOOL_        │
  │ ALLOWLIST │  │ ALLOWLIST  │  │ ALLOWLIST    │  │ ALLOWLIST    │
  │ =core.json│  │ =docs.json │  │ =collab.json │  │ =mailbox.json│
  ├───────────┤  ├────────────┤  ├──────────────┤  ├──────────────┤
  │ Entra:    │  │ Entra:     │  │ Entra:       │  │ Entra:       │
  │ org-wide  │  │ [group TBD]│  │ [group TBD]  │  │ [group TBD]  │
  └───────────┘  └────────────┘  └──────────────┘  └──────────────┘
```

Each connector is:
- The same container image from the same ACR
- A separate Container App (its own URL) in the same Container Apps Environment
- Configured via env var pointing at its allowlist (env var containing the JSON list, or a mounted config — implementation detail for Phase 1)
- Its own Entra app registration (or shared app with per-connector assignment — decision in Phase 2)
- Assigned in claude.ai to the corresponding Entra group

A user in no advanced groups sees only Core (~60 schemas). An analyst sees Core + Docs Authoring. Context scales with role.

---

## Phase 0 — Baseline measurement (before any changes)

**Goal:** Capture the "before" numbers so the context reduction is provable.

1. Count the exact number of tools currently exposed by the production connector. (Server logs report "N registered, M skipped" at startup — capture N.)
2. Measure the actual context cost: start a fresh claude.ai conversation with only the M365 connector enabled, and record the token count consumed by tool schemas. (Approximation: export the tool manifest, count tokens. Claude Code can script this.)
3. Record per-tool schema sizes to identify the heaviest schemas — useful for prioritizing what stays out of Core.

**Deliverable:** A short baseline doc: total tool count, total schema token cost, top-10 heaviest schemas.

**Effort:** Half a day.

---

## Phase 1 — Allowlist mechanism (code change, the real work)

**Goal:** Replace the `ENABLED_TOOLS` keyword regex with an explicit named allowlist, in our fork.

**Design:**

- New env var `TOOL_ALLOWLIST` containing either a JSON array of exact tool names, or a path/URL to a JSON config. Exact-match against tool names — no patterns.
- Precedence: if `TOOL_ALLOWLIST` is set, it wins and `ENABLED_TOOLS` is ignored (log a warning if both are set). If only `ENABLED_TOOLS` is set, behave exactly as today (backward compatible — the current production deployment keeps working unchanged).
- The allowlist applies at tool-registration time (same place the regex filter runs today), so unlisted tools never register and never advertise schemas.
- Custom tools (`download-mail-attachment`, `get-messages-batch`, `list-conversation-messages`, `list-drafts`, and the new `request-m365-capability`) go through the same allowlist check as endpoint tools — no special-casing, one mechanism.
- Startup logging: log the allowlist source, the count of tools registered, and the full sorted list of registered tool names — this is the governance audit trail Scott wants.

**Also in this phase:** build the four allowlist config files from Scott's brief (Core, Docs/Excel Authoring, Collab Admin, Mailbox/Calendar Management), with the two corrections Scott confirmed:
- `list-conversation-messages` → Core (was mistakenly in the drop list)
- `list-drafts` → Core

**Validation before proceeding:**
- Deploy to a TEST container app (not production) with the Core allowlist
- Verify the registered tool count matches the Core list exactly
- Verify a claude.ai connection to the test connector shows only Core tools
- Re-measure schema token cost against Phase 0 baseline → this is the headline number

**Deliverable:** Fork release (v0.88.2-tsq.10 or a new versioning scheme for this project), four allowlist configs, measured before/after tool count and token cost.

**Effort:** 2-4 days including testing.

---

## Phase 2 — Core connector rollout

**Goal:** Stand up the Core connector as a new production deployment and migrate users to it.

1. **New Container App** (`tsqm365mcp-core`) in the existing environment, same image, `TOOL_ALLOWLIST=core`.
2. **Entra app decision:** new app registration for Core vs. reusing the existing one. Recommendation: new registration named for the connector (TSQ-M365-MCP-Core) so assignment, consent, and audit are per-connector. The scopes for Core are a subset of what we have today (read + communicate; no file-write beyond what download-mail-attachment staging needs).
3. **Assignment:** org-wide (or a broad pilot group first — recommend piloting with 5-10 users before org-wide).
4. **The meta-tool:** implement `request-m365-capability` v1 — takes a capability name or free-text need, returns the catalog of advanced modules, and files the request via ticket/Teams/email to IT (mechanism TBD by Mike — simplest reliable option).
5. **Parallel running:** the existing full connector stays live for Mike and Scott during burn-in. Nothing is cut over until Core is proven.

**Validation:**
- Pilot users confirm daily workflows function on Core alone
- Measure real conversation context usage for pilot users
- request-m365-capability files a real request end-to-end

**Deliverable:** Core connector live, pilot feedback, measured context numbers in production use.

**Effort:** 2-3 days build/config + 1-2 weeks pilot.

---

## Phase 3 — Advanced tiers, one at a time

**Goal:** Stand up the three Advanced connectors in priority order, each gated by its Entra group.

Suggested order (by likely demand):
1. **M365 Docs & Excel Authoring** — analysts/finance are the most likely immediate users
2. **M365 Mailbox & Calendar Management** — assistants/schedulers
3. **M365 Collaboration Admin** — team owners (smallest audience, most sensitive operations)

For each tier:
1. Create the Entra group (names from Mike) and define who approves membership
2. Deploy the Container App with the tier's allowlist
3. Entra app registration + admin consent scoped to what that tier needs (e.g., Collab Admin is where TeamMember.ReadWrite.All finally gets properly gated — resolving the "hide" item from the May scope review)
4. Assign the connector to the Entra group in claude.ai
5. Validate: a group member sees Core + the tier; a non-member sees Core only

**Scope cleanup opportunity:** This is the moment to rationalize the Entra scopes per connector. The current single app carries 35 scopes for everything; per-tier apps can carry only what their tools need. Document the scope set per tier.

**Deliverable per tier:** Live connector, group assignment working, scope documentation.

**Effort:** 1-2 days per tier once the pattern is established by the first.

---

## Phase 4 — Decommission & governance

1. **Retire the legacy full connector** (or park it as an admin-only connector for Mike) once all tiers are proven.
2. **Update all documentation:** deployment & risk doc, feature overview, user setup instructions — all currently describe the single-connector world.
3. **Quarterly review process:** Scott asked for the split to be revisited quarterly against usage. Define the lightweight telemetry: per-tool invocation counts from server logs, reviewed quarterly, tools moved between tiers as reality dictates.
4. **v2 backlog:** automated Entra access packages for request-m365-capability; consider consolidating deployments if claude.ai's connector model ever supports per-user schema variation on one URL.

---

## Measurement commitments (Scott's requirement)

The plan reports these numbers at each gate:

| Measurement | When | Target |
|---|---|---|
| Baseline tool count + schema token cost | Phase 0 | ~170 tools documented |
| Core tool count + schema token cost | Phase 1 test deploy | ~60 tools + custom readers; token reduction quantified |
| Real conversation context usage (pilot) | Phase 2 | Confirm reduction holds in practice |
| Per-tier tool counts | Phase 3 | Each tier documented |
| Per-tool usage telemetry | Phase 4 quarterly | Informs tier adjustments |

---

## Risks & mitigations

| Risk | Mitigation |
|---|---|
| Allowlist code change introduces regressions in tool registration | Backward-compatible design (ENABLED_TOOLS still works); test deploy before touching production; existing connector untouched until Core proven |
| Four connectors = 4x OAuth reconnect pain when deploying updates | Batch updates; deploy all four from one pipeline run; communicate maintenance windows |
| Users confused by which connector does what | Naming + the request-m365-capability catalog; feature overview rewrite in Phase 4 |
| Scope creep per Entra app multiplies admin-consent surface | Per-tier scope documentation; consent reviewed at each tier's deployment |
| Pilot reveals Core is missing a daily-use tool | That's what the pilot is for; moving a tool into Core is a config change + redeploy, minutes not days |

---

## Open items (Mike to resolve during Phase 1-2)

1. Real Entra group names for the three tiers + membership approvers
2. request-m365-capability v1 delivery mechanism: ticket system, Teams message, or email to IT
3. Allowlist config delivery: env var JSON vs. mounted file vs. URL (implementation detail, decide during Phase 1)
4. Per-tier Entra app registrations: confirm new-app-per-tier approach vs. shared app
5. Pilot group composition for Phase 2

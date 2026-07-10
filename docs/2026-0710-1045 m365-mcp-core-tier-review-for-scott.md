# M365 Core connector — tier split review

**For:** Scott Schatz **From:** Mike Mrozek **Date:** 2026-07-10
**Decision needed:** sign off (or adjust) which tools land in the everyone-gets-it **Core** connector vs. the three gated Advanced tiers.

---

## Where we are

Phase 1 is built and **proven on a live test connector**. One server image now supports an exact-match tool allowlist per tier. All 257 tools are assigned to exactly one tier:

| Connector | Tools | Who gets it |
|---|---:|---|
| **Core** | 68 | Everyone (org-wide) |
| Docs & Excel Authoring | 56 | Gated by Entra group |
| Collaboration Admin | 72 | Gated by Entra group |
| Mailbox & Calendar Mgmt | 61 | Gated by Entra group |

**Measured result:** Core cuts the per-conversation tool-schema load **~66%** (≈202K → ≈69K tokens) vs. today's single 166-tool connector. The split is proven; **only the Core boundary needs your call.** Moving a tool later is a config edit + redeploy (minutes), and the Phase 2 pilot exists to catch anything we mis-placed.

---

## The 7 judgment calls (everything else is unambiguous)

For each: *does a typical everyone-user need this daily?* My recommendation is in **bold**.

| # | Tools in question | Currently | Recommend | Rationale |
|---|---|---|---|---|
| 1 | Message **reactions** ×4 (`set/unset-*-reaction`) | Collab Admin | **→ Core** | Daily Teams action; adds ~no new scope (Core already has chat-write). |
| 2 | **Shared-mailbox** read/send ×4 | Mailbox Mgmt | **→ Core** *(if pilot includes assistants/execs)* | Common for assistants; no new scope. Audience-dependent. |
| 3 | **Non-default calendar** tools ×4 (`*-specific-calendar-*`) | Mailbox Mgmt | **→ Core** | Core today only handles the *default* calendar; multi-calendar users need these daily. |
| 4 | Message **pins** ×3 (`pin/unpin/list-pinned`) | Collab Admin | **Keep gated** | Less common than reactions; fine to gate. |
| 5 | `list-team-members` (team roster) | Collab Admin | **Keep gated** | Sole holder of the broad `TeamMember.Read.All` scope (your May scope-review item) — keeping it out keeps that scope off the org-wide connector. |
| 6 | **SharePoint** site/list reads (~20 tools) | Docs & Excel | **Keep gated** | Large surface; would bloat Core. Core already has OneDrive search + file read. |
| 7 | **To Do** personal tasks ×9 | Mailbox Mgmt | **Keep gated** | Keeps all task tools together; not universal daily use. |

**Net if you accept the recommendations:** Core goes 68 → **~80 tools** (still ≈ −60% tokens vs. today). Or keep Core lean at 68 and let the pilot promote on real demand — your preference:

- ☐ **Option A — Ship lean (68).** Promote items 1–3 only if the pilot shows friction.
- ☐ **Option B — Pre-load the daily ones (~80).** Apply recommended promotions 1–3 now.

---

## Three structural questions (optional)

1. **Collaboration Admin is broad (72 tools)** — Teams admin + Groups + online meetings/recordings + Planner + subscriptions. Fine as one "power" tier, or split? *(Recommend: keep for now; revisit at the quarterly review you asked for.)*
2. **Change-notification webhooks** (`*-subscription` ×6) — an integration/admin concern, not user collaboration. Keep in Collab Admin, or drop from all user tiers? *(Recommend: keep gated; candidate to remove entirely if unused.)*
3. **Docs & Excel Authoring** bundles Excel + OneDrive authoring + SharePoint lists + OneNote. Hold together, or separate files/SharePoint from Excel/OneNote? *(Recommend: keep as one authoring connector.)*

---

## How to respond

Just mark **Option A or B** above, flag any of items 1–7 you'd move, and note the pilot audience (drives item 2). I'll re-generate the configs and re-validate in minutes. Full tool-by-tier lists are in `config/allowlists/` if you want to inspect any tier in detail.

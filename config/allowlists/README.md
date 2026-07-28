# Tool allowlists — tiered connector rollout

Each file is an explicit, exact-match list of MCP tool names for one connector tier.
Supply it to the server via the `TOOL_ALLOWLIST` env var (or `--tool-allowlist`),
pointing at either the file path or the inline JSON. When set, it takes precedence
over the legacy `ENABLED_TOOLS` regex. See
`docs/2026-0707-1502 m365-mcp-tiered-rollout-plan.md`.

```
TOOL_ALLOWLIST=config/allowlists/core.json  ms-365-mcp-server --http --org-mode
```

## Tiers

| File | Connector | Tools | Purpose |
|---|---|---:|---|
| `core.json` | M365 Core | 115 | Read + communicate daily driver — everyone gets this. Mail/calendar/chat/Teams read, send & everyday management (folders, rules, message move/flag), shared mailbox & calendar, non-default calendars, Teams reactions/pins & online-meeting scheduling, people lookup, basic OneDrive read, search, and the 4 TSQ custom tools. |
| `docs-excel.json` | M365 Docs & Excel Authoring | 56 | Excel read/write, OneDrive/SharePoint file authoring, SharePoint lists, OneNote. |
| `collab-admin.json` | M365 Collaboration Admin | 59 | Teams/Group administration, meeting recordings/transcripts/attendance, Planner, subscriptions/webhooks, hosted-content. Smallest audience, most sensitive ops. |
| `mailbox-calendar.json` | M365 Mailbox & Calendar Management | 27 | Contacts, Outlook categories, To Do tasks, rooms/places, mailbox-settings write, and delta-sync endpoints. |

Every one of the 257 org-mode tools is assigned to exactly one tier (no overlap, no gaps).
A user in no advanced group sees only Core (115 schemas, ~28% fewer schema tokens than the
current 166-tool production connector). Advanced group members see Core **plus** their tier.

## Conservative-Core decision (management, July 2026)

Management reviewed the tiered-rollout update and directed: **err conservative on the number
of tools reduced — cut fewer now, evaluate later from real usage data.** Accordingly Core was
widened from the initial 68-tool draft to 115 by promoting everyday mail/calendar management,
shared mailbox & calendar, non-default ("specific") calendars, Teams reactions/pins, message
replies, and online-meeting scheduling out of the advanced tiers. Gated (advanced) surface is
now limited to genuinely advanced/sensitive/rare ops: Excel/OneNote/SharePoint authoring,
Teams/Group administration, meeting recordings/transcripts/attendance, subscriptions/webhooks,
contacts, To Do, and delta-sync.

Further cuts are **data-driven**: the server now writes a structured per-invocation usage log
(`usage.log`, one JSON object per call: `tool`, `outcome`, and caller `oid`/`upn`/`tid`) so we
can report tool usage by user periodically and retire genuinely unused tools with evidence.
See `src/usage-log.ts`. Disable with `USAGE_LOG=off`; relocate with `MS365_MCP_USAGE_LOG_DIR`.

## Scope minimization (Core = 22 delegated scopes)

A scope-minimization pass was run on Core (map each consented scope → the tools requiring it):

- **`list-team-members` was moved from Core → Collab Admin.** It was the *sole* holder of the
  broad `TeamMember.Read.All` scope (a May scope-review item); moving it keeps the whole
  `TeamMember.*` family gated in Collab Admin and drops that scope from Core.
- **`search-query` is a scope magnet** — it single-handedly justifies Core's two broadest read
  scopes, `Sites.Read.All` and `Files.Read.All` (no other Core tool needs `.All`). This is the
  accepted cost of org-wide Microsoft Search; not a trim candidate.
- **Entra-app note (not an allowlist concern):** the 5 custom tools do not feed scope derivation,
  so Core's app registration must additionally grant **`Files.ReadWrite`** for
  `download-mail-attachment` (OneDrive staging). That scope collapses `Files.Read`.

## Confirmed placements (Scott, June 2026)

- `list-conversation-messages` → **Core** (was mistakenly on the drop list)
- `list-drafts` → **Core**
- All four TSQ custom tools stay in Core: `download-mail-attachment`, `get-messages-batch`,
  `list-conversation-messages`, `list-drafts`.

## Status: DRAFT for review (conservative revision)

The Core / Advanced split reflects the July 2026 conservative-Core decision above. Moving a
tool between tiers is a config edit + redeploy (minutes). Judgment calls, as now resolved:

- **Reactions / pins** (`set-*-reaction`, `pin-chat-message`, …) → **Core** (common Teams UX).
- **Shared mailbox** read/send → **Core** (assistant/delegate workflow is common).
- **Non-default ("specific") calendar** tools → **Core** (multi-calendar users).
- **Online-meeting scheduling** (`create/get/list-online-meeting`) → **Core**; recordings,
  transcripts, and attendance reports stay in **Collab Admin** (sensitive).
- **Planner** → Collab Admin (team boards); **To Do** → Mailbox Management (personal tasks, kept
  gated pending usage data).
- **Subscriptions / change-notification webhooks** → Collab Admin (sensitive integration surface).
- **Contacts / Outlook categories / rooms** → Mailbox Management (kept gated pending usage data).
- OneDrive/SharePoint: only basic OneDrive *read* is in Core; all authoring, SharePoint lists,
  and OneNote are in Docs & Excel Authoring.

Open for Scott's review: whether **contacts** (read) and **online-meeting `update`** should
also move to Core. Usage data will settle these rather than debate.

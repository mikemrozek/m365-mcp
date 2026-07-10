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
| `core.json` | M365 Core | 68 | Read + communicate daily driver — everyone gets this. Mail/calendar/chat/Teams read & send, people lookup, basic OneDrive read, search, and the 4 TSQ custom tools. |
| `docs-excel.json` | M365 Docs & Excel Authoring | 56 | Excel read/write, OneDrive/SharePoint file authoring, SharePoint lists, OneNote. |
| `collab-admin.json` | M365 Collaboration Admin | 72 | Teams/Group administration, online meetings & recordings, Planner, subscriptions, Teams message power-features. Smallest audience, most sensitive ops. |
| `mailbox-calendar.json` | M365 Mailbox & Calendar Management | 61 | Mail folders/rules/settings, calendar administration, non-default & shared calendars, contacts, To Do, delta sync. |

Every one of the 257 org-mode tools is assigned to exactly one tier (no overlap, no gaps).
A user in no advanced group sees only Core (68 schemas, ~55% fewer schema tokens than the
current 166-tool production connector). Advanced group members see Core **plus** their tier.

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

## Status: DRAFT for review

The Core / Advanced split below is a first draft from the plan's tier definitions, not yet
signed off. Moving a tool between tiers is a config edit + redeploy (minutes). Known
judgment calls to review:

- **Reactions / pins** (`set-*-reaction`, `pin-chat-message`, …) are in Collab Admin for
  leanness; they are common enough that they may belong in Core.
- **Shared mailbox** read/send is in Mailbox Management (assistant workflow); could be Core.
- **Non-default ("specific") calendar** tools are in Mailbox Management, not Core.
- **Planner** → Collab Admin (team boards); **To Do** → Mailbox Management (personal tasks).
- **Subscriptions / change-notification webhooks** → Collab Admin (sensitive integration surface).
- OneDrive/SharePoint: only basic OneDrive *read* is in Core; all authoring, SharePoint lists,
  and OneNote are in Docs & Excel Authoring.

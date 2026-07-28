# Handoff — M365 MCP connector packaging strategy (for review with claude.ai)

**Date:** 2026-07-27 · **Author:** Mike Mrozek · **Next milestone:** follow-up
discussion Tue **2026-07-28** (Scott) on "M365 connector packaging strategy."

## How to use this document

I'm re-evaluating **how to split up the M365 tool surface into right-sized packages**
after a July 27 call changed the direction. This doc gives you (claude.ai) the full
context with **no prior knowledge assumed**. I want your help pressure-testing the new
proposed architecture, nailing down terminology, and checking what's actually feasible on
the Claude/claude.ai platform today. The open questions I most want answered are at the end.

---

## 1. The problem (unchanged)

We run **Microsoft 365 as an MCP connector** (the open-source `softeria/ms-365-mcp-server`,
a TSQ fork) so Claude can act on Outlook mail/calendar, Teams chat/channels, SharePoint/
OneDrive files, Excel, and directory data via Microsoft Graph.

**Core pain:** too many tools = **token/context bloat**. Every conversation with the
connector enabled loads *all* the tool JSON schemas into context, inflating every prompt
and reducing effective working room, before the user has done anything.

**Real numbers (reconciled — the call quoted "~200" loosely):**
- **166 tools** — what the current *production* connector actually advertises today
  (gated by an `ENABLED_TOOLS` keyword regex).
- **257 tools** — the full org-mode surface the server *can* expose.
- Measured schema cost of the full manifest ≈ **~200K tokens** (chars/4 approximation);
  a 68-tool subset measured ≈ **~69K** (−66%). So the bloat is real and reducible.

The whole effort is about **shrinking the per-session tool set to what a given user/task
actually needs**, without losing capability for those who need more.

---

## 2. Where we were before the call (the "4-connector" plan)

Original plan (Scott-approved June 2026): **one server image, deployed as four separate
connectors**, each with a static allowlist and its own Entra group:

| Connector | ~Tools | Audience |
|---|---|---|
| **Core** | 115 | Everyone — read + communicate daily driver |
| **Docs & Excel Authoring** | 56 | Analysts/finance |
| **Collab Admin** | 59 | Team owners (sensitive ops) |
| **Mailbox & Calendar Mgmt** | 27 | Assistants/schedulers |

A user in no advanced group sees only Core; advanced users see Core **+ their tier**
(two connectors). Split by **capability/sensitivity**, delivered as **multiple connector
URLs**.

### What is already BUILT and still useful regardless of direction

1. **`TOOL_ALLOWLIST` mechanism** (shipped in the fork): an exact-match, per-deployment
   list of tool names that gates which tools register. Fails closed, wins over the legacy
   regex. This is the *static, per-deployment* version of "expose fewer tools."
2. **Per-tool / per-user usage logging** (built this week): the server now writes a
   structured `usage.log` (one JSON line per call: `tool`, `outcome`, and caller identity
   `oid`/`upn`/`tid` derived from the access-token JWT). **This is the empirical data that
   should define package boundaries** — see the focus-group action item below.
3. **Full tool→service classification**: every one of the 257 tools is mapped to a
   service area (mail, calendar, chat/Teams, files/SharePoint, Excel, directory, etc.).
   Reusable raw material for *any* packaging axis.

### Why it's now in question

On the call, **Scott said he's not confident the three advanced tiers (Docs/Excel,
Mailbox/Cal, Collab Admin) are the right cut**, and leaned toward *getting the connector
into users' hands and letting real usage define the boundaries* rather than deciding
top-down.

---

## 3. What changed on the July 27 call (the NEW direction)

**Edwin Lovett proposed a different architecture:** instead of multiple connector variants,
keep a **single M365 connector** and **gate tool exposure per session** so only the relevant
subset is injected into context. He framed it as "hooks/plugins" inspired by **Cloudflare's
plugin architecture** — Claude can reach the full connector, but what gets loaded is
restricted by the active plugin/hook. Stated benefits: less bloat, one thing to maintain,
and it stops Claude from "going crazy" trying to use out-of-scope tools.

- **Packaging axis is shifting from capability-tiers to purpose-built packages.** Edwin's
  examples were **app/task-focused**: "email-only," "calendar-focused," "collaboration-heavy"
  — i.e. divisions by **Microsoft app** (Outlook, Teams, SharePoint…) rather than by user
  level.
- **But "levels" may still be a dimension** — Mike also heard appetite for **beginner /
  intermediate / advanced** packages. So there are potentially **two orthogonal axes**
  (by app × by level) still on the table.
- **Scott approved exploring the hooks direction** as "more elegant and easier to maintain."
  Cory Mills validated it. **No final architecture was locked** — that's the July 28 topic.
- **"Claude plugins for distribution"** came up as the delivery/packaging vehicle.

### Mike's action items from the call
- **Flush out the hooks-based connector architecture; prep for the July 28 follow-up.**
- Complete connector implementation focused on **user-assignment logic**.
- **Run focus-group testing** with target user teams to define package boundaries from
  real usage (this is where the `usage.log` data feeds in).

---

## 4. Terminology to align FIRST (this is where the review should start)

The call used "hooks," "plugins," and "skills" loosely. In the actual Claude platform these
are distinct, and conflating them risks aiming at a mechanism that doesn't solve the token
problem. Please help validate/correct this framing:

- **MCP connector**: the remote server that advertises tools. When enabled, **all its
  advertised tool schemas load into context.** Context cost is driven by *how many tools
  are advertised/enabled in the session* — nothing else.
- **What actually reduces context bloat:** advertising/enabling **fewer tools** in a
  session. That happens either **server-side** (the server lists fewer tools — statically
  per deployment, or dynamically per user/session) or **client-side** (the connector's
  tools are toggled off for a user/group).
- **Claude Code "hooks":** a **local Claude Code (CLI) feature** — shell commands that fire
  on tool-call events and can allow/deny/modify a call. ⚠️ **Important:** hooks gate
  *execution*, not *exposure* — the tool's schema is still loaded in context. So if "hooks"
  means Claude Code hooks, they **do not reduce token bloat** (they *do* stop Claude from
  misusing out-of-scope tools — a real but different benefit Edwin also cited). They also
  don't apply to claude.ai web/Team users. **Open question: is "hooks" here a literal Claude
  Code feature, or a conceptual name for server-side/ client-side tool filtering?**
- **Claude Skills:** packaged instructions (progressive disclosure — only a name+description
  load until invoked). Great for reducing *instruction* bloat and steering Claude to the
  right workflow; a Skill does **not** by itself unload MCP tool schemas. Could be the
  user-facing "package" abstraction ("Outlook skill," "Teams skill") layered over a lean
  connector.
- **Claude Code Plugins:** distributable bundles (skills + commands + hooks + subagents +
  MCP configs) installed from marketplaces — a **distribution** mechanism, primarily for
  **Claude Code** users. For claude.ai Team users the analog is admin-managed connectors +
  skills.

**Net:** the "single connector + gate exposure per session" idea is sound *if* the gating
changes what tools are **advertised** to the session. The key is which real lever implements
that — and whether claude.ai supports it today.

---

## 5. Delivery mechanisms to evaluate (with feasibility flags)

| # | Mechanism | Reduces context? | Single connector? | Feasible today on claude.ai? |
|---|---|---|---|---|
| A | **Multiple connectors, static allowlists** (original plan) | Yes | No (2–4 URLs) | ✅ Yes — already built |
| B | **One connector, tool set varies by authenticated user/group** (server keys exposure off Entra identity at connect) | Yes | Yes | ❓ Needs per-user tool-list variation on one connector URL — **verify** |
| C | **One connector, per-session dynamic tool set** (user/"plugin" selects the active package; server advertises only those tools, MCP `tools/list_changed`) | Yes | Yes | ❓ Closest to Edwin's proposal — depends on claude.ai honoring dynamic/ selectable tool lists mid-session — **verify** |
| D | **Claude Skills as the package abstraction** over a lean connector | Partial (instruction bloat; steers tool use) | Yes | ✅ Skills exist — but doesn't unload MCP schemas by itself |
| E | **Client-side per-tool toggles** in the connector admin UI (per group) | Yes | Yes | ❓ Depends on current claude.ai connector admin granularity — **verify** |
| F | **Claude Code Plugins / marketplace** for distribution | Yes (bundle scopes tools) | N/A | ✅ For **Claude Code** users; not the claude.ai Team surface |

The three ❓ rows (B, C, E) are the **make-or-break feasibility questions** for the
single-connector direction — the original plan explicitly parked "per-user schema variation
on one connector URL" as a *future* capability, so we need to confirm whether the platform
can do it **now** before committing.

---

## 6. Packaging-axis options (independent of mechanism)

Whatever the delivery mechanism, we still must decide **how to group the tools**:

- **By Microsoft app / task** (Edwin's steer): Outlook (mail+calendar), Teams (chat+
  channels), Files (SharePoint/OneDrive), Excel, Directory/People. Intuitive to users;
  maps to "I need the email package."
- **By level**: beginner / intermediate / advanced — progressive capability.
- **By capability/sensitivity** (original plan): read-vs-write, admin-vs-user.
- **Hybrid**: app × level matrix, or a small "Core" everyone gets + opt-in app packages.
- **Data-driven / emergent** (Scott's lean): ship broad, watch `usage.log`, let real usage
  cluster into packages. The focus-group testing action item is exactly this.

These axes are **orthogonal** — the review should decide the *primary* axis and whether a
second axis layers on top.

---

## 7. Constraints & facts to respect

- **Don't disrupt existing users.** Prod connector `tsqm365mcp-app` (image `v0.88.2-tsq.9`,
  `ENABLED_TOOLS` regex, 166 tools) must keep running unchanged until any new design is
  proven. The `TOOL_ALLOWLIST` config is *inert* on prod unless explicitly set.
- **Usage logging is ephemeral as written** — `usage.log` is a file on the container FS
  (lost on restart). For the focus-group data to survive it needs a durable sink (stdout →
  Log Analytics, or a mounted volume). Small fix, but required before relying on the data.
- **User-assignment logic** (who gets which package) is an open build item and interacts
  directly with the mechanism choice (Entra group vs claude.ai connector assignment vs
  in-session selection).
- **Platform:** claude.ai Team/Enterprise connectors + Microsoft Entra (OAuth) for auth and
  group membership. Azure Container Apps hosts the server(s).

---

## 8. Questions for the claude.ai review (paste these)

1. **Terminology:** Is "hooks" (Edwin's proposal) the literal Claude Code hooks feature, or
   a general term for tool-exposure filtering? Given hooks gate *execution* not *context*,
   what's the right Claude-native mechanism to actually reduce the tool-schema token load?
2. **Feasibility (the crux):** On claude.ai **today**, can a **single connector** expose a
   **different set of tools per user/group (B) or per session (C)**? If yes, how (dynamic
   `tools/list`, connector-level tool toggles, identity-keyed server behavior)? If no,
   multiple connectors (A) may still be the only real option.
3. **Skills vs connector:** Should the "packages" be **Claude Skills** layered over one lean
   connector, or subsets of the connector itself? What does each do to context cost?
4. **Packaging axis:** by app, by level, by capability, or hybrid — what best matches how
   users actually think and keeps context lean? How should focus-group/`usage.log` data
   drive this rather than top-down guesses?
5. **Distribution:** what role (if any) do **Claude plugins/marketplace** play for our
   audience (claude.ai Team users vs any Claude Code users)?
6. **Migration:** given the "don't break prod" constraint, what's the least-disruptive path
   from today's single 166-tool connector to the chosen design?

---

## Appendix — status of built work

- `src/tool-allowlist.ts` — `TOOL_ALLOWLIST` exact-match gating (works, tested).
- `src/usage-log.ts` + `request-context.ts` + `lib/microsoft-auth.ts` — per-tool/per-user
  usage logging (built this week; build + smoke verified; **not yet committed/deployed**).
- `config/allowlists/{core,docs-excel,collab-admin,mailbox-calendar}.json` — the 4-tier
  classification (draft; now likely to be re-cut by a different axis, but the underlying
  tool→service mapping is reusable).
- Full plan of record: `docs/2026-0707-1502 m365-mcp-tiered-rollout-plan.md` (the
  now-superseded 4-connector plan — read for background, not as current direction).

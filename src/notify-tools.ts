import { z } from 'zod';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type GraphClient from './graph-client.js';
import logger from './logger.js';
import { withUsageLog } from './usage-log.js';
import type { RegisterHooks } from './reply-tracking-tools.js';

/**
 * notify-me (tsq.28): tell the signed-in user something, and nobody else.
 *
 * Why it exists. Scott (2026-10-13/14) wants a scheduled check that messages
 * him in Teams when a tracked reply arrives. A Claude scheduled task can run
 * the check unattended, but only with tools allowed to run without asking —
 * and a general send tool on "Always allow" lets any scheduled task message
 * anyone. This tool has no recipient parameter, so it is safe to allow.
 *
 * Why Teams goes through a Workflows webhook. Graph will not touch the "notes
 * to self" chat (2026-10-14: GET /chats/48:notes → 400 "Call made for a thread
 * which is not a ChatThread"), and a message the user posts as themselves
 * never notifies them anywhere. An alert has to come from another sender. A
 * Teams Workflows flow ("post to a chat when a webhook request is received")
 * is that sender: the user creates it once, and its posts arrive from
 * "Workflows" and ping. Only Power Automate hosts are accepted, so the tool
 * cannot be pointed at an arbitrary URL, and that path uses no Microsoft 365
 * permission at all.
 *
 * Email is to the user's own address only — a new message does notify — and is
 * the fallback when no webhook is given or the webhook fails.
 */

const MAX_TEXT = 4000;
const WEBHOOK_TIMEOUT_MS = 10_000;

/** Power Automate / Workflows trigger hosts. Anything else is refused. */
export function isWorkflowWebhook(raw: string): boolean {
  let url: URL;
  try {
    url = new URL(raw.trim());
  } catch {
    return false;
  }
  if (url.protocol !== 'https:') return false;
  const host = url.hostname.toLowerCase();
  return host.endsWith('.logic.azure.com') || host.endsWith('.api.powerplatform.com');
}

/** The Adaptive Card envelope the Workflows "webhook request" trigger expects. */
export function workflowPayload(title: string | undefined, text: string, link?: string) {
  const body: unknown[] = [];
  if (title) body.push({ type: 'TextBlock', text: title, weight: 'Bolder', wrap: true });
  body.push({ type: 'TextBlock', text, wrap: true });
  return {
    type: 'message',
    attachments: [
      {
        contentType: 'application/vnd.microsoft.card.adaptive',
        content: {
          $schema: 'http://adaptivecards.io/schemas/adaptive-card.json',
          type: 'AdaptiveCard',
          version: '1.4',
          body,
          ...(link ? { actions: [{ type: 'Action.OpenUrl', title: 'Open', url: link }] } : {}),
        },
      },
    ],
  };
}

function escapeHtml(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function jsonResult(payload: unknown, isError = false) {
  return {
    content: [{ type: 'text' as const, text: JSON.stringify(payload) }],
    ...(isError ? { isError: true } : {}),
  };
}

export function registerNotifyTools(
  server: McpServer,
  graphClient: GraphClient,
  hooks: RegisterHooks
): void {
  const { isToolEnabled, readOnly, push, fail, skip } = hooks;
  const name = 'notify-me';
  if (readOnly) {
    skip(name);
    return;
  }
  if (!isToolEnabled(name)) return;

  async function emailSelf(subject: string, text: string, link?: string) {
    const me = (await graphClient.makeRequest('/me?$select=mail,userPrincipalName')) as {
      mail?: string;
      userPrincipalName?: string;
    };
    const address = me?.mail || me?.userPrincipalName;
    if (!address) throw new Error('Could not determine your own email address.');
    const html =
      `<p>${escapeHtml(text).replace(/\n/g, '<br>')}</p>` +
      (link ? `<p><a href="${escapeHtml(link)}">${escapeHtml(link)}</a></p>` : '');
    await graphClient.makeRequest('/me/sendMail', {
      method: 'POST',
      body: JSON.stringify({
        message: {
          subject,
          body: { contentType: 'HTML', content: html },
          toRecipients: [{ emailAddress: { address } }],
        },
        saveToSentItems: false,
      }),
    });
    return address;
  }

  try {
    server.tool(
      name,
      'Notify ME — send the signed-in user a reminder or alert, and nobody else. For ' +
        'scheduled checks and long tasks: "tell me when someone replies", "alert me if ' +
        'anything is overdue". There is no recipient parameter: it can only reach the user, ' +
        'so it is safe to allow in unattended scheduled tasks.\n\n' +
        'Teams: pass teamsWebhookUrl — the URL of the user’s own Teams Workflows flow ' +
        '("post to a chat when a webhook request is received"); its posts arrive from ' +
        'Workflows and notify. Only Power Automate URLs are accepted. Without it, or if the ' +
        'webhook fails, the notice goes by email to the user’s own address.\n\n' +
        'Send one notice per check summarising everything new, not one per item. Do not ' +
        'call it when there is nothing to report.',
      {
        message: z.string().min(1).max(MAX_TEXT).describe('What to tell the user, plain text.'),
        title: z.string().max(200).optional().describe('Short heading, e.g. "2 replies came in".'),
        link: z
          .string()
          .url()
          .optional()
          .describe('Optional link to open, e.g. the email or chat webUrl.'),
        teamsWebhookUrl: z
          .string()
          .optional()
          .describe("The user's own Teams Workflows webhook URL. Omit to notify by email instead."),
      },
      {
        title: 'notify-me',
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: true,
      },
      async (params) =>
        withUsageLog(name, async () => {
          const title = params.title?.trim();
          const text = params.message.trim();
          const subject = `[Claude] ${title || text.slice(0, 80)}`;
          let teamsError: string | undefined;

          const webhook = params.teamsWebhookUrl?.trim();
          if (webhook) {
            if (!isWorkflowWebhook(webhook)) {
              return jsonResult(
                {
                  error: 'not_a_workflow_webhook',
                  message:
                    'teamsWebhookUrl must be a Teams Workflows / Power Automate URL ' +
                    '(*.logic.azure.com or *.api.powerplatform.com). Nothing was sent.',
                },
                true
              );
            }
            try {
              const controller = new AbortController();
              const timer = setTimeout(() => controller.abort(), WEBHOOK_TIMEOUT_MS);
              const res = await fetch(webhook, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(workflowPayload(title, text, params.link)),
                signal: controller.signal,
              }).finally(() => clearTimeout(timer));
              if (res.ok) {
                logger.info('notify-me: delivered to Teams via Workflows webhook');
                return jsonResult({ delivered: true, channel: 'teams' });
              }
              teamsError = `Workflows webhook answered HTTP ${res.status}`;
            } catch (error) {
              teamsError = `Workflows webhook failed: ${(error as Error).message}`;
            }
            logger.warn(`notify-me: ${teamsError}; falling back to email`);
          }

          try {
            const address = await emailSelf(subject, text, params.link);
            return jsonResult({
              delivered: true,
              channel: 'email',
              to: address,
              ...(teamsError ? { teamsError, note: 'Teams failed, so this went by email.' } : {}),
            });
          } catch (error) {
            return jsonResult(
              {
                error: 'notify_failed',
                message: `Could not notify you: ${(error as Error).message}`,
                ...(teamsError ? { teamsError } : {}),
              },
              true
            );
          }
        })
    );
    push(name);
  } catch (error) {
    fail(name, error as Error);
  }
}

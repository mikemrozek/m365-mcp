/**
 * Everything the weekly report needs from outside the process: a managed
 * identity token, the usage rows from Log Analytics, and a way to send mail.
 *
 * NO SECRET IS STORED ANYWHERE IN THIS PATH, and that is the design constraint
 * rather than a nicety.
 *
 * The connector's own app registration holds 35 permissions, every one of them
 * DELEGATED — which is what makes "the connector can never do more than the
 * signed-in user could do themselves" a true statement, and it is a statement
 * we have made to management in writing. A weekly report runs with nobody
 * signed in, so it cannot borrow that path, and adding an application
 * permission to that registration to make it work would quietly falsify the
 * claim for all 136 capabilities.
 *
 * So the report sends as a SEPARATE registration that holds exactly one
 * application permission (Mail.Send), narrowed further by an Exchange
 * Application Access Policy to a single mailbox. It authenticates by federating
 * to the job's managed identity — no client secret exists to leak, rotate, or
 * find in a log. Compromising the report gets an attacker the ability to send
 * mail as one mailbox, and nothing else: not the connector, not anyone's files.
 */

import logger from '../logger.js';
import type { UsageRow } from './analyze.js';

/** Exchanged for a real Graph token; the audience Entra reserves for federation. */
const FEDERATION_AUDIENCE = 'api://AzureADTokenExchange';
const LOG_ANALYTICS_RESOURCE = 'https://api.loganalytics.io';

async function postForm(url: string, form: Record<string, string>): Promise<unknown> {
  const response = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams(form).toString(),
  });
  const text = await response.text();
  if (!response.ok) {
    // Never echo the body of a token response — a failure body can still carry
    // fragments of the assertion that was sent (SEC-2026-001).
    throw new Error(`Token request to ${new URL(url).host} failed: ${response.status}`);
  }
  return JSON.parse(text);
}

/**
 * A token for the job's user-assigned managed identity.
 *
 * Container Apps injects IDENTITY_ENDPOINT / IDENTITY_HEADER; the classic IMDS
 * address is the fallback so this can also be exercised on a VM. `clientId` is
 * required whenever more than one identity is assigned, which is why it is not
 * optional here.
 */
export async function getManagedIdentityToken(resource: string, clientId: string): Promise<string> {
  const endpoint = process.env.IDENTITY_ENDPOINT;
  const header = process.env.IDENTITY_HEADER;

  const url = endpoint
    ? `${endpoint}?api-version=2019-08-01&resource=${encodeURIComponent(resource)}&client_id=${encodeURIComponent(clientId)}`
    : `http://169.254.169.254/metadata/identity/oauth2/token?api-version=2018-02-01&resource=${encodeURIComponent(resource)}&client_id=${encodeURIComponent(clientId)}`;

  const response = await fetch(url, {
    headers: endpoint && header ? { 'X-IDENTITY-HEADER': header } : { Metadata: 'true' },
  });
  if (!response.ok) {
    throw new Error(`Managed identity token request failed: ${response.status}`);
  }
  const body = (await response.json()) as { access_token?: string };
  if (!body.access_token) throw new Error('Managed identity returned no access_token.');
  return body.access_token;
}

/**
 * A Graph token for the reporting application, obtained without a secret: the
 * managed identity's token is presented as a client assertion, and Entra trusts
 * it because a federated identity credential on the reporting app names that
 * identity as its issuer and subject.
 */
export async function getGraphTokenViaFederation(config: {
  tenantId: string;
  clientId: string;
  managedIdentityClientId: string;
}): Promise<string> {
  const assertion = await getManagedIdentityToken(
    FEDERATION_AUDIENCE,
    config.managedIdentityClientId
  );
  const body = (await postForm(
    `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
    {
      client_id: config.clientId,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
      client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
      client_assertion: assertion,
    }
  )) as { access_token?: string };
  if (!body.access_token) throw new Error('Federated token exchange returned no access_token.');
  return body.access_token;
}

/**
 * Pulls the usage records. The KQL mirrors the ad-hoc queries the briefs have
 * been run by hand — same table, same marker, same fields — so a figure in the
 * emailed report and a figure someone checks by hand cannot disagree.
 */
export async function fetchUsageRows(config: {
  workspaceId: string;
  days: number;
  managedIdentityClientId: string;
}): Promise<UsageRow[]> {
  const token = await getManagedIdentityToken(
    LOG_ANALYTICS_RESOURCE,
    config.managedIdentityClientId
  );
  const query = `
    ContainerAppConsoleLogs_CL
    | where TimeGenerated >= ago(${config.days}d)
    | where Log_s has 'm365-usage'
    | extend d = parse_json(Log_s)
    | project ts = tostring(d.timestamp), tool = tostring(d.tool),
              outcome = tostring(d.outcome), upn = tostring(d.upn)
    | order by ts asc`;

  const response = await fetch(
    `${LOG_ANALYTICS_RESOURCE}/v1/workspaces/${encodeURIComponent(config.workspaceId)}/query`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({ query }),
    }
  );
  if (!response.ok) {
    throw new Error(`Log Analytics query failed: ${response.status} ${response.statusText}`);
  }

  const body = (await response.json()) as {
    tables?: { columns: { name: string }[]; rows: unknown[][] }[];
  };
  const table = body.tables?.[0];
  if (!table) return [];
  const index = (name: string) => table.columns.findIndex((c) => c.name === name);
  const [ts, tool, outcome, upn] = ['ts', 'tool', 'outcome', 'upn'].map(index);

  return table.rows.map((row) => ({
    ts: String(row[ts] ?? ''),
    tool: String(row[tool] ?? ''),
    outcome: String(row[outcome] ?? ''),
    upn: String(row[upn] ?? ''),
  }));
}

export async function sendReportMail(config: {
  token: string;
  senderUpn: string;
  recipients: string[];
  subject: string;
  html: string;
}): Promise<void> {
  const response = await fetch(
    `https://graph.microsoft.com/v1.0/users/${encodeURIComponent(config.senderUpn)}/sendMail`,
    {
      method: 'POST',
      headers: { Authorization: `Bearer ${config.token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify({
        message: {
          subject: config.subject,
          body: { contentType: 'HTML', content: config.html },
          toRecipients: config.recipients.map((address) => ({ emailAddress: { address } })),
        },
        saveToSentItems: true,
      }),
    }
  );
  if (!response.ok) {
    // Graph's own message is safe to surface here and is the difference between
    // "the access policy is wrong" and "the mailbox does not exist".
    throw new Error(`sendMail failed: ${response.status} ${await response.text()}`);
  }
  logger.info(`[WEEKLY REPORT] Sent to ${config.recipients.length} recipient(s).`);
}

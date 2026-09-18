FROM node:24-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm i

COPY . .
RUN npm run generate
RUN npm run build

FROM node:20-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package*.json ./

ENV NODE_ENV=production
RUN npm i --ignore-scripts --omit=dev

# SEC-2026-003: drop root. Dependencies are installed above as root so /app stays
# root-owned and read-only to the process, which is the point.
#
# HOME is set explicitly because the log directory is derived from os.homedir():
# with HOME unset it resolves through /etc/passwd instead, and a base-image change
# would move the logs silently. /home/node is created and owned by node:node in the
# official image, so both logger.ts and usage-log.ts can still mkdir under it.
#
# The /data/token-cache Azure Files mount is NOT writable by this user. That is
# survivable rather than overlooked: the MSAL cache is written only by the stdio
# device-code flow, and this deployment runs HTTP/OAuth mode, where getToken()
# returns the caller's own token before the cache path is reached. Both save paths
# also catch and log a write failure rather than throwing. Verify on the test
# connector before production all the same — see the tsq.24 runbook.
ENV HOME=/home/node
USER node

ENTRYPOINT ["node", "dist/index.js"]

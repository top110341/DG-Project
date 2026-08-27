# =====================================================================
# DG Projects Enterprise — Production Dockerfile
# =====================================================================
# Database (Turso/libSQL) and file storage (Vercel Blob) are both external
# services now, so this image has no local persistent state and needs no
# volume — it just needs the right env vars at runtime (see below).

FROM node:20-alpine

WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

ENV NODE_ENV=production
ENV PORT=3000

# Required at runtime (set these in your host's env/secrets, not here):
#   TURSO_DATABASE_URL, TURSO_AUTH_TOKEN  — from turso.tech
#   BLOB_READ_WRITE_TOKEN                 — from Vercel Blob
#   ALLOWED_ORIGIN                        — your deployed domain (optional)

EXPOSE 3000

CMD ["node", "server.js"]

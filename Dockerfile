# =====================================================================
# AG Projects Enterprise — Production Dockerfile
# Multi-stage build for Node.js 20 LTS with persistent SQLite storage
# =====================================================================

# ---- Build stage: compiles native deps (sqlite3) ----
FROM node:20-alpine AS builder

# python3/make/g++ are required in case sqlite3 falls back to compiling its
# native addon from source (no prebuilt binary for this platform/arch).
RUN apk add --no-cache python3 make g++ sqlite

WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev

COPY . .

# ---- Runtime stage: slim image, no build toolchain ----
FROM node:20-alpine AS runtime

RUN apk add --no-cache sqlite bash tzdata

WORKDIR /app
COPY --from=builder /app ./

# All persistent state (SQLite DB + uploaded files) lives under DATA_DIR so a
# single volume mount covers both and survives redeploys/restarts.
ENV NODE_ENV=production
ENV PORT=3000
ENV DATA_DIR=/app/data

RUN mkdir -p /app/data/uploads && chmod -R 777 /app/data

EXPOSE 3000

CMD ["node", "server.js"]

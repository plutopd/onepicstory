# Vibe Cut needs more than a Node runtime: the export endpoint shells out to
# ffmpeg, and burning Korean captions in needs a font with Hangul glyphs.
# Both are installed here so the deployed host matches what the code expects.
FROM node:20-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
      ffmpeg \
      fonts-nanum \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Dependencies first so a code-only change reuses the cached install layer.
COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY . .

ENV NODE_ENV=production
# server.js honours process.env.PORT, which the host assigns at runtime.
EXPOSE 8787

CMD ["node", "server.js"]

# Build stage
FROM node:20-slim AS build
WORKDIR /app
COPY package*.json ./
RUN npm ci
COPY tsconfig.json ./
COPY src ./src
RUN npm run build

# Runtime stage
FROM node:20-slim AS runtime
ENV NODE_ENV=production
ENV MCP_TRANSPORT=http
ENV PORT=8080
WORKDIR /app
COPY package*.json ./
RUN npm ci --omit=dev && npm cache clean --force
COPY --from=build /app/dist ./dist
EXPOSE 8080
# Container Apps / most PaaS set PORT; the server reads it.
CMD ["node", "dist/index.js"]

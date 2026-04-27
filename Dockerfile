# --- STAGE 1: Build Frontend ---
FROM node:20-alpine AS frontend-builder
WORKDIR /app/frontend
COPY frontend/package*.json ./
# Changed 'ci' to 'install' because package-lock.json might be missing
RUN npm install --ignore-scripts
COPY frontend/ ./
RUN npm run build

# --- STAGE 2: Build Backend ---
FROM node:20-alpine AS backend-builder
WORKDIR /app/backend
COPY backend/package*.json ./
# Updated --only=production to --omit=dev per npm 10+ standards
RUN npm install --omit=dev --ignore-scripts
COPY backend/ ./

# --- STAGE 3: Final Production Image ---
FROM node:20-alpine
WORKDIR /app

# SECURITY: Create a non-root user
RUN addgroup -S appgroup && adduser -S appuser -G appgroup
USER appuser

# Copy the backend files
COPY --from=backend-builder /app/backend ./backend
# Copy the frontend build artifacts
COPY --from=frontend-builder /app/frontend/build ./backend/public

EXPOSE 3000
WORKDIR /app/backend
CMD ["node", "index.js"]
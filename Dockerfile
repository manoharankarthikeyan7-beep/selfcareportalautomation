# --- STAGE 1: Build Frontend ---
FROM node:20-alpine AS frontend-builder
WORKDIR /app/frontend
COPY frontend/package*.json ./
RUN npm install --ignore-scripts
COPY frontend/ ./
RUN npm run build

# --- STAGE 2: Build Backend ---
FROM node:20-alpine AS backend-builder
WORKDIR /app/backend
COPY backend/package*.json ./
RUN npm install --omit=dev --ignore-scripts
COPY backend/ ./

# --- STAGE 3: Final Production Image ---
FROM node:20-alpine
WORKDIR /app

# SECURITY: Create a non-root user
RUN addgroup -S appgroup && adduser -S appuser -G appgroup

# Copy the backend files
COPY --from=backend-builder /app/backend ./backend
# Copy the frontend build artifacts
COPY --from=frontend-builder /app/frontend/build ./backend/public

# --- NEW: Copy SSL Certificates ---
# Ensure key.pem and cert.pem are in your local folder before building
COPY key.pem ./backend/key.pem
COPY cert.pem ./backend/cert.pem

# Set permissions for the non-root user
RUN chown -R appuser:appgroup /app/backend

USER appuser

# Updated to 8080 to match your "Secured Server" logs
EXPOSE 8080
WORKDIR /app/backend

CMD ["node", "index.js"]
# ---- Build: compile the React app into static files ----
FROM node:20-alpine AS build
WORKDIR /app
COPY package.json package-lock.json ./
RUN npm ci
COPY . .
RUN npm run build

# ---- Run: serve the static files with nginx ----
FROM nginx:1.27-alpine
COPY docker/nginx.conf /etc/nginx/conf.d/default.conf
COPY docker/40-env-config.sh /docker-entrypoint.d/40-env-config.sh
# Strip Windows line endings in case the script was checked out with CRLF.
RUN sed -i 's/\r$//' /docker-entrypoint.d/40-env-config.sh \
 && chmod +x /docker-entrypoint.d/40-env-config.sh
COPY --from=build /app/dist /usr/share/nginx/html
EXPOSE 80

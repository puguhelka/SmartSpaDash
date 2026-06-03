# SmartSpaDash — Node.js Express + SQLite (sql.js)
FROM node:20-alpine

WORKDIR /app

# Install dependencies
COPY package.json package-lock.json ./
RUN npm ci --production

# Copy app
COPY . .

# Expose port
EXPOSE 3000

# Start
CMD ["node", "server.js"]

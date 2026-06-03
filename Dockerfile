FROM node:20-alpine
WORKDIR /app
RUN mkdir -p /app/data && chmod 777 /app/data
COPY package.json package-lock.json ./
RUN npm ci --production
COPY . .
EXPOSE 3000
CMD ["node", "server.js"]

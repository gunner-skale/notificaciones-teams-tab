FROM node:24-alpine
WORKDIR /app
COPY server.js ./
COPY public ./public
ENV PORT=6090
EXPOSE 6090
USER node
CMD ["node", "server.js"]

const https = require('https');
const express = require('express');
const path = require('path');
const devCerts = require('office-addin-dev-certs');

async function main() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  const app = express();
  const webRoot = path.join(__dirname, 'web');

  app.use((req, res, next) => {
    res.setHeader('Cache-Control', 'no-store');
    next();
  });

  app.use(express.static(webRoot));

  https.createServer(httpsOptions, app).listen(3000, 'localhost', () => {
    console.log('Email Exporter is running at https://localhost:3000/taskpane.html');
    console.log('Keep this window open while testing the add-in.');
  });
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});

const http = require('http');
const fs = require('fs');
const path = require('path');

// Frontend UI port
const PORT = process.env.PORT || 3000;

const indexPath = path.join(__dirname, 'index.html');

const server = http.createServer((req, res) => {
  // Serve a single-page index for all routes
  fs.readFile(indexPath, (err, data) => {
    if (err) {
      res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.end('Error loading frontend index.html');
      return;
    }

    res.writeHead(200, { 'Content-Type': 'text/html' });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`Frontend running on http://localhost:${PORT}`);
});


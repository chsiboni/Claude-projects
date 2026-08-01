// Minimal zero-dependency static server for Railway.
// Serves index.html for the app; binds to the port Railway provides.
const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = process.env.PORT || 3000;
const ROOT = __dirname;

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".png": "image/png",
  ".json": "application/json; charset=utf-8",
};

const server = http.createServer((req, res) => {
  // Strip query string and normalize; default to index.html.
  let urlPath = decodeURIComponent((req.url || "/").split("?")[0]);
  if (urlPath === "/") urlPath = "/index.html";
  if (urlPath === "/snake") urlPath = "/snake.html";

  // Prevent path traversal.
  const filePath = path.join(ROOT, path.normalize(urlPath));
  if (!filePath.startsWith(ROOT)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      // Fall back to the app so any path shows the tool.
      fs.readFile(path.join(ROOT, "index.html"), (e2, home) => {
        if (e2) {
          res.writeHead(404);
          res.end("Not found");
          return;
        }
        res.writeHead(200, { "Content-Type": TYPES[".html"] });
        res.end(home);
      });
      return;
    }
    const type = TYPES[path.extname(filePath)] || "application/octet-stream";
    res.writeHead(200, { "Content-Type": type });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`עברית ישרה running on port ${PORT}`);
});

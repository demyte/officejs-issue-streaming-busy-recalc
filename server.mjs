import https from "node:https";
import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";

const root = new URL("./", import.meta.url);
const certificates = join(homedir(), ".office-addin-dev-certs");
const routes = new Map([
  ["/", ["index.html", "text/html"]],
  ["/index.html", ["index.html", "text/html"]],
  ["/functions.js", ["functions.js", "text/javascript"]],
  ["/taskpane.js", ["taskpane.js", "text/javascript"]],
  ["/functions.json", ["functions.json", "application/json"]],
  ["/manifest.xml", ["manifest.xml", "application/xml"]],
  ...[16, 32, 80].map(size => [`/icon-${size}.png`, [`icon-${size}.png`, "image/png"]]),
]);

let tls;
try {
  tls = {
    key: readFileSync(join(certificates, "localhost.key")),
    cert: readFileSync(join(certificates, "localhost.crt")),
  };
} catch {
  console.error("Local HTTPS certificate missing. Run npm run certs, then npm start.");
  process.exit(1);
}

https.createServer(tls, (request, response) => {
  const pathname = new URL(request.url, "https://localhost:3443").pathname;
  const route = routes.get(pathname);
  response.on("finish", () => {
    console.log(`${new Date().toISOString()} ${request.method} ${pathname} ${response.statusCode}`);
  });
  if (request.method !== "GET" || !route) {
    response.writeHead(404).end();
    return;
  }
  try {
    const body = readFileSync(new URL(route[0], root));
    response.writeHead(200, {
      "Content-Type": route[1],
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    }).end(body);
  } catch (error) {
    console.error(error.message);
    response.writeHead(500).end();
  }
}).on("error", error => {
  console.error(error.message);
  process.exitCode = 1;
}).listen(3443, "127.0.0.1", () => {
  console.log("Excel stream repro: https://localhost:3443");
});

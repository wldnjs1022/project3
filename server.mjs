import { createReadStream, existsSync, statSync } from "node:fs";
import { createServer } from "node:http";
import { extname, join, normalize } from "node:path";

const root = process.cwd();
const port = Number(process.env.PORT || 4173);
const types = { ".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".json": "application/json; charset=utf-8" };

createServer((request, response) => {
  const pathname = decodeURIComponent(new URL(request.url, `http://${request.headers.host}`).pathname);
  const requested = pathname === "/" ? "/index.html" : pathname;
  const filepath = normalize(join(root, requested));
  if (!filepath.startsWith(root) || !existsSync(filepath) || !statSync(filepath).isFile()) {
    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not found");
    return;
  }
  response.writeHead(200, { "Content-Type": types[extname(filepath)] || "application/octet-stream" });
  createReadStream(filepath).pipe(response);
}).listen(port, "127.0.0.1", () => {
  console.log(`맞춤표가 http://127.0.0.1:${port} 에서 실행 중입니다.`);
});

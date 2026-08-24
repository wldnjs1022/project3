import { cp, mkdir, rm } from "node:fs/promises";

const output = new URL("./dist/", import.meta.url);
await rm(output, { recursive: true, force: true });
await mkdir(output, { recursive: true });
for (const source of ["index.html", "app-v2.js", "styles.css", "fixes.css", "vendor"]) {
  await cp(new URL(`./${source}`, import.meta.url), new URL(`./dist/${source}`, import.meta.url), { recursive: true });
}
await mkdir(new URL("./dist/server/", import.meta.url), { recursive: true });
await mkdir(new URL("./dist/.openai/", import.meta.url), { recursive: true });
await cp(new URL("./server-entry.js", import.meta.url), new URL("./dist/server/index.js", import.meta.url));
await cp(new URL("./.openai/hosting.json", import.meta.url), new URL("./dist/.openai/hosting.json", import.meta.url));

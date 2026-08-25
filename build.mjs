import { cp, mkdir, rm } from "node:fs/promises";

const output = new URL("./dist/", import.meta.url);
await rm(output, { recursive: true, force: true });
await mkdir(output, { recursive: true });
for (const source of ["index.html", "app-v2.js", "styles.css", "fixes.css", "vendor"]) {
  await cp(new URL(`./${source}`, import.meta.url), new URL(`./dist/${source}`, import.meta.url), { recursive: true });
}

import { promisify } from "node:util";
import { execFile } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

const execFileAsync = promisify(execFile);

const chromePath = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";
const url = `file://${path.resolve("/Users/bala/Documents/Codex/Diet Plan/index.html")}`;
const outputPath = "/Users/bala/Documents/Codex/Diet Plan/output/weekly-diet-dashboard.pdf";

await fs.mkdir("/Users/bala/Documents/Codex/Diet Plan/output", { recursive: true });

await execFileAsync(chromePath, [
  "--headless=new",
  "--disable-gpu",
  "--allow-file-access-from-files",
  "--print-to-pdf-no-header",
  `--print-to-pdf=${outputPath}`,
  url,
]);

console.log(`Wrote ${outputPath}`);

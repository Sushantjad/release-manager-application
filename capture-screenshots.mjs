import { createRequire } from "node:module";
import { mkdir } from "node:fs/promises";
import { resolve } from "node:path";

const require = createRequire("C:/Users/Admin/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/.pnpm/playwright@1.60.0/node_modules/playwright/package.json");
const { chromium } = require("playwright");

const outputDir = resolve("screenshots");
await mkdir(outputDir, { recursive: true });

const browser = await chromium.launch({
  executablePath: "C:/Program Files/Google/Chrome/Application/chrome.exe",
  headless: true
});
const page = await browser.newPage({ viewport: { width: 1440, height: 950 } });

await page.goto("http://127.0.0.1:8080/index.html", { waitUntil: "networkidle" });
await page.screenshot({ path: resolve(outputDir, "01-dashboard.png"), fullPage: true });

await page.getByRole("button", { name: "Deployment" }).click();
await page.screenshot({ path: resolve(outputDir, "02-deployment.png"), fullPage: true });

await page.getByRole("button", { name: "Release Day" }).click();
await page.screenshot({ path: resolve(outputDir, "03-release-day-command-center.png"), fullPage: true });

await browser.close();

import { defineConfig } from "@playwright/test";

// Live runs (NYMFORM_E2E_LIVE=1) go to the real endpoint; behind an HTTPS proxy, Chromium needs
// to be told about it. Local test traffic never goes through the proxy.
const proxyServer = process.env.NYMFORM_E2E_LIVE ? process.env.HTTPS_PROXY || process.env.https_proxy : undefined;
// Another checkout can run its own suite at the same time on another port (NYMFORM_E2E_PORT). A
// busy port fails the run rather than testing whatever that server serves (see webServer below).
const port = Number(process.env.NYMFORM_E2E_PORT || 4173);

// End-to-end tests drive the built task pane (dist/) in Chromium with a fake Excel host
// (tests/e2e/fakeOffice.js) in place of Office.js, and a fake model endpoint unless a
// test says otherwise. Build first: `npm run build`.
export default defineConfig({
  testDir: "tests/e2e",
  testMatch: /.*\.spec\.ts/,
  timeout: 60_000,
  retries: 0,
  reporter: [["list"]],
  use: {
    baseURL: `http://127.0.0.1:${port}`,
    viewport: { width: 400, height: 900 },
    screenshot: "only-on-failure",
    ...(proxyServer ? { proxy: { server: proxyServer, bypass: "127.0.0.1,localhost" } } : {}),
  },
  webServer: {
    command: "node scripts/serve-dist.mjs",
    env: { PORT: String(port) },
    url: `http://127.0.0.1:${port}/taskpane.html`,
    // A server already on the port may be another checkout's: fail rather than test its dist/,
    // unless asked to reuse one (NYMFORM_E2E_REUSE=1).
    reuseExistingServer: process.env.NYMFORM_E2E_REUSE === "1",
  },
});

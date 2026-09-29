// The manifests and the dev server must agree on one local port, set in package.json
// ("config.dev_server_port"). A manifest pointing at another port loads nothing, and Excel shows an
// add-in error.
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const read = (name: string) => readFileSync(join(ROOT, name), "utf8");
const pkg = JSON.parse(read("package.json")) as { config: { dev_server_port: number } };
const PORT = pkg.config.dev_server_port;

/** Every localhost URL in a manifest, with its port. */
function localUrls(xml: string): { url: string; port: string }[] {
  return [...xml.matchAll(/https?:\/\/localhost(?::(\d+))?[^"<\s]*/g)].map((m) => ({ url: m[0], port: m[1] ?? "" }));
}

describe("manifests", () => {
  it("the dev server port is a plain number outside the ports common local tools use", () => {
    expect(Number.isInteger(PORT)).toBe(true);
    expect(PORT).toBeGreaterThan(1024);
    // Below macOS's temporary-port range, which starts at 49152.
    expect(PORT).toBeLessThan(49152);
    expect([3000, 3001, 4173, 4200, 5000, 5173, 8000, 8080, 8443, 9000]).not.toContain(PORT);
  });

  for (const name of ["manifest.dev.xml", "manifest.bench.xml"]) {
    it(`${name} points only at https://localhost on the dev server port`, () => {
      const urls = localUrls(read(name));
      expect(urls.length).toBeGreaterThan(0);
      for (const { url, port } of urls) {
        expect(url, url).toMatch(/^https:/);
        expect(port, url).toBe(String(PORT));
      }
    });
  }

  it("manifest.release.xml doesn't point at localhost", () => {
    expect(localUrls(read("manifest.release.xml"))).toEqual([]);
  });

  for (const name of ["manifest.dev.xml", "manifest.bench.xml", "manifest.release.xml"]) {
    it(`${name} names SkySpan as the publisher`, () => {
      expect(read(name)).toContain("<ProviderName>SkySpan</ProviderName>");
    });
  }

  it("no skyspanmarketing.com address is given as a contact (the security address moves to nymform.com)", () => {
    for (const name of ["README.md", "SECURITY.md", "docs/spec.md"]) expect(read(name), name).not.toMatch(/@skyspanmarketing\.com/u);
  });

  it("webpack serves on the same port", () => {
    const config = createRequire(import.meta.url)(join(ROOT, "webpack.config.js")) as { DEV_URL: string };
    expect(config.DEV_URL).toBe(`https://localhost:${PORT}/`);
  });
});

// SEC-6 — the Content-Security-Policy in index.html, kept honest.
//
// A CSP is easy to add and easy to silently break. Two failure modes matter,
// and this file catches both:
//
//  1. The inline theme bootstrap in index.html is admitted by a sha256 hash.
//     Edit that script by one character and the hash stops matching — the
//     browser then refuses to run it, the app paints in the wrong theme on
//     first load, and nothing in the build or the type-checker notices. The
//     first test recomputes the hash from the file and compares.
//
//  2. A directive gets loosened (or a needed one dropped) without anyone
//     weighing it. The rest of the tests pin each source list to the reason
//     it exists, so widening one means editing an assertion that says why.
import { describe, expect, it } from "vitest";
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const indexHtml = readFileSync(
  join(dirname(fileURLToPath(import.meta.url)), "..", "index.html"),
  "utf8",
);

/** The CSP's directives, as a map of name -> source list. */
function policy(): Map<string, string[]> {
  const meta = /<meta\s+http-equiv="Content-Security-Policy"\s+content="([\s\S]*?)"/.exec(indexHtml);
  if (!meta) {
    throw new Error("index.html has no Content-Security-Policy meta tag");
  }
  const directives = new Map<string, string[]>();
  for (const part of meta[1].split(";")) {
    const [name, ...sources] = part.trim().split(/\s+/).filter(Boolean);
    if (name) {
      directives.set(name, sources);
    }
  }
  return directives;
}

function sources(directive: string): string[] {
  const found = policy().get(directive);
  expect(found, `CSP is missing the ${directive} directive`).toBeDefined();
  return found ?? [];
}

describe("Content-Security-Policy", () => {
  it("admits every inline script in index.html by hash", () => {
    const inline = [...indexHtml.matchAll(/<script>([\s\S]*?)<\/script>/g)];
    // One today: the no-flash theme bootstrap. If a second appears it needs
    // its own hash, so assert the count rather than only the first.
    expect(inline).toHaveLength(1);

    const scriptSrc = sources("script-src");
    for (const [, body] of inline) {
      const hash = createHash("sha256").update(body, "utf8").digest("base64");
      expect(
        scriptSrc,
        `The inline <script> in index.html hashes to sha256-${hash}, which script-src does not list. ` +
          "The browser will refuse to run it (first paint flashes the wrong theme). " +
          "Update the hash in index.html's CSP to match the edited script.",
      ).toContain(`'sha256-${hash}'`);
    }
  });

  it("does not allow arbitrary inline or eval'd script", () => {
    const scriptSrc = sources("script-src");
    expect(scriptSrc).not.toContain("'unsafe-inline'");
    // pdf.js used to be the one library that needed eval; v6 removed that
    // path entirely (see the note in lib/pdfRaster.ts), so nothing in the app
    // requires it and the policy must not quietly regain it.
    expect(scriptSrc).not.toContain("'unsafe-eval'");
    expect(scriptSrc).toContain("'self'");
  });

  it("locks down the directives that have no legitimate use here", () => {
    expect(sources("default-src")).toEqual(["'self'"]);
    // Nothing in the app submits a form or rewrites the document base.
    expect(sources("base-uri")).toEqual(["'none'"]);
    expect(sources("form-action")).toEqual(["'none'"]);
    expect(sources("frame-src")).toEqual(["'none'"]);
  });

  it("allows the Google Fonts pair index.html actually loads, and nothing else remote", () => {
    // The stylesheet comes from fonts.googleapis.com, the face files it
    // references from fonts.gstatic.com — both are in index.html's <link>s.
    expect(sources("style-src")).toContain("https://fonts.googleapis.com");
    expect(sources("font-src")).toContain("https://fonts.gstatic.com");
    expect(sources("font-src")).toContain("'self'");
  });

  it("allows blob: where the app genuinely needs it, and only there", () => {
    // The PDF viewer is the browser's own, driven by an <object> over a
    // blob: URL (components/media/MediaPdfEmbed.vue), so object-src cannot
    // be 'none' here the way it usually would be.
    expect(sources("object-src")).toContain("blob:");
    // Downloads, image rasters and the media surfaces are all blob:/data:.
    expect(sources("img-src")).toEqual(expect.arrayContaining(["'self'", "data:", "blob:"]));
    expect(sources("media-src")).toEqual(expect.arrayContaining(["'self'", "data:", "blob:"]));
    // Both workers are same-origin module chunks; blob: covers bundlers that
    // inline a worker.
    expect(sources("worker-src")).toEqual(expect.arrayContaining(["'self'", "blob:"]));
  });

  it("permits the outbound https the open-from-URL feature is built on", () => {
    // /open?url= and Markdown's image resolution fetch documents the user
    // points at, so connect-src cannot be 'self'. openDocumentFromUrl.ts's
    // own allow-list (isFetchableUrl) is what narrows this to http(s).
    // ws:/wss: are Vite's dev-server HMR socket.
    const connect = sources("connect-src");
    expect(connect).toEqual(expect.arrayContaining(["'self'", "https:"]));
    expect(connect).not.toContain("*");
  });
});

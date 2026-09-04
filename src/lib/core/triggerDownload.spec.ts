// triggerBlobDownload is the single download mechanism, enforced as a test so
// it stays single.
//
// It did not stay single on its own. The helper was extracted so the
// object-URL lifecycle would exist once, and the two newer products
// (presentation, media) used it — but five older call sites in the text
// product's view layer had each grown their own copy of the same five
// statements: MergeRunDialog, useVersionHistory, useFileActions, useFonts and
// usePdfExport. Nothing failed, which is the problem: a change to how a
// download is delivered (the revoke delay, a filename rule, an added
// `rel`/`target`, a Safari workaround) had six places to land in and would
// have landed in one.
//
// So the invariant is checked rather than remembered: outside this module,
// nothing constructs an anchor and sets `download` on it.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const coreDir = dirname(fileURLToPath(import.meta.url));
const srcDir = join(coreDir, "..", "..");
const HELPER = join(coreDir, "triggerDownload.ts");

/** Every runtime source file under src/ — specs excluded, since a test may
 * legitimately stub or assert on the download mechanism. */
function runtimeFiles(dir: string, out: string[] = []): string[] {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) {
      runtimeFiles(path, out);
    } else if (
      (path.endsWith(".ts") || path.endsWith(".vue")) &&
      !path.endsWith(".spec.ts") &&
      !path.endsWith(".d.ts")
    ) {
      out.push(path);
    }
  }
  return out;
}

describe("the download mechanism exists exactly once", () => {
  it("no module outside core/triggerDownload.ts sets `download` on an anchor", () => {
    const offenders: string[] = [];
    for (const path of runtimeFiles(srcDir)) {
      if (path === HELPER) {
        continue;
      }
      const source = readFileSync(path, "utf8");
      // The tell of a hand-rolled download: assigning the `download`
      // attribute. Deliberately narrow — reading `.download`, or a property
      // named `download` in a data object, is not this.
      if (/\.download\s*=/.test(source)) {
        offenders.push(path.slice(srcDir.length));
      }
    }
    expect(
      offenders,
      "These modules hand-roll a file download. Call triggerBlobDownload " +
        "(src/lib/core/triggerDownload.ts) instead — it owns the anchor and the " +
        "object-URL revoke, and it exists so this mechanism is in one place.",
    ).toEqual([]);
  });

  it("still delivers a download through an anchor and revokes the URL", () => {
    // Guards the helper itself, since the test above only says where the
    // mechanism may live, not that it works.
    const source = readFileSync(HELPER, "utf8");
    expect(source).toContain("URL.createObjectURL");
    expect(source).toContain("URL.revokeObjectURL");
    expect(source).toMatch(/anchor\.download\s*=/);
  });
});

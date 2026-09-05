// Every `<input type="file">` in the app is a picker a visible button opens:
// "Open a file", "Image…", "Import styles…". None is meant to be reached on
// its own, and none carries a label of its own — the button beside it is the
// accessible name.
//
// Four different classes were being used to hide them (hidden-file-input,
// menu-image-input, visually-hidden, sr-only), and only the two backed by
// `display: none` actually left the tab order. The other seven used the
// screen-reader pattern, which is for text you DO want announced: clipped to a
// pixel but still focusable. Tabbing to the end of the home page landed on one
// of them — a control with no name, no visible focus ring and no apparent
// effect, because the file dialog only opens on click.
//
// A shared CSS class would not have prevented that: the next picker can still
// pick the wrong one. The attributes are what the browser actually reads, so
// this asserts those instead, the same way noDependencies.spec.ts pins the
// manifest and composableSize.spec.ts pins the composables.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const srcDir = dirname(fileURLToPath(import.meta.url));

function vueFiles(): string[] {
  const found: string[] = [];
  const walk = (dir: string) => {
    for (const name of readdirSync(dir)) {
      const path = join(dir, name);
      if (statSync(path).isDirectory()) {
        walk(path);
      } else if (name.endsWith(".vue")) {
        found.push(path);
      }
    }
  };
  walk(srcDir);
  return found.sort();
}

interface FileInput {
  file: string;
  testid: string;
  tag: string;
}

/** Every `<input type="file">` element in the app, with the file it lives in.
 * Matched as a whole tag so multi-line attribute lists are read together. */
function fileInputs(): FileInput[] {
  const found: FileInput[] = [];
  for (const path of vueFiles()) {
    const source = readFileSync(path, "utf8");
    for (const match of source.matchAll(/<input\b[^>]*>/g)) {
      const tag = match[0];
      if (!/\btype="file"/.test(tag)) {
        continue;
      }
      found.push({
        file: path.slice(srcDir.length + 1),
        testid: /data-testid="([^"]+)"/.exec(tag)?.[1] ?? "(no test id)",
        tag,
      });
    }
  }
  return found;
}

describe("hidden file inputs", () => {
  it("finds the pickers (a guard that scans nothing guards nothing)", () => {
    expect(fileInputs().length).toBeGreaterThanOrEqual(9);
  });

  it("keeps every picker out of the tab order", () => {
    const tabbable = fileInputs()
      .filter((input) => !/\btabindex="-1"/.test(input.tag))
      .map((input) => `${input.file} — ${input.testid}`);
    expect(
      tabbable,
      'These <input type="file"> elements can be tabbed to. A picker a button ' +
        'opens is not a stop of its own: add tabindex="-1".',
    ).toEqual([]);
  });

  it("keeps every picker out of the accessibility tree", () => {
    const announced = fileInputs()
      .filter((input) => !/\baria-hidden="true"/.test(input.tag))
      .map((input) => `${input.file} — ${input.testid}`);
    expect(
      announced,
      'These <input type="file"> elements are still announced, with no name of ' +
        "their own — the visible button that opens them is the accessible " +
        'name. Add aria-hidden="true".',
    ).toEqual([]);
  });
});

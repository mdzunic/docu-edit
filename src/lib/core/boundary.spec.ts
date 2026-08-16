// The core layer's one structural invariant, enforced as a test so it
// cannot silently erode as specs keep landing: nothing in src/lib/core may
// import from outside src/lib/core. No document model, no commands, no
// docx/odf readers, no Vue — a module that needs any of those is not core
// and belongs a layer up. This is what keeps the layer reusable by another
// product (e.g. Slides) without dragging the word processor in.
import { readdirSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";

const coreDir = dirname(fileURLToPath(import.meta.url));
const libDir = join(coreDir, "..");

/** Every runtime module directly in `dir` (specs excluded — they may import
 * vitest, node builtins and, for the round-trip suites, both formats). */
function runtimeModules(dir: string): string[] {
  return readdirSync(dir).filter(
    (name) => name.endsWith(".ts") && !name.endsWith(".spec.ts"),
  );
}

function coreModules(): string[] {
  return runtimeModules(coreDir);
}

/** Every import/export specifier in a module's source: static imports,
 * re-exports and dynamic import() calls alike. */
function importSpecifiers(source: string): string[] {
  const specifiers: string[] = [];
  const patterns = [
    /(?:^|\n)\s*import\s[^;]*?from\s*"([^"]+)"/g,
    /(?:^|\n)\s*import\s*"([^"]+)"/g,
    /(?:^|\n)\s*export\s[^;]*?from\s*"([^"]+)"/g,
    /import\(\s*"([^"]+)"\s*\)/g,
  ];
  for (const pattern of patterns) {
    for (const match of source.matchAll(pattern)) {
      specifiers.push(match[1]);
    }
  }
  return specifiers;
}

describe("core layer boundary", () => {
  it("has modules to check", () => {
    expect(coreModules().length).toBeGreaterThan(0);
  });

  it("no core module imports from outside src/lib/core", () => {
    const violations: string[] = [];
    for (const name of coreModules()) {
      const source = readFileSync(join(coreDir, name), "utf8");
      for (const specifier of importSpecifiers(source)) {
        // Allowed: sibling core modules ("./x"). Everything else — parent
        // paths ("../documentModel"), aliases ("@/…"), bare packages
        // ("vue") — couples core to a product or a framework.
        if (!specifier.startsWith("./") || specifier.startsWith("./..")) {
          violations.push(`${name} imports "${specifier}"`);
        }
      }
    }
    expect(violations).toEqual([]);
  });
});

// The rest of the layer rule, one level up from core: core → model layer
// (documentModel.ts and the command/query modules typed against it) →
// format readers/writers (docx/, odf/). A format module may import the
// model layer and core; the model layer may import core only; and the two
// format directories never import each other — anything both formats need
// is format-neutral by definition and belongs in the model layer
// (documentModel.ts's shared helpers, mediaRefs.ts) or in core
// (core/media.ts's content-type mapping, core/xml.ts's escapes).
describe("format layer boundaries", () => {
  /** Import specifiers in `name` (under `dir`) that resolve into a `docx/`
   * or `odf/` path segment named by `banned` — "../docx/x", "./odf/x" and
   * aliased "…/docx/x" forms alike. */
  function formatImportViolations(dir: string, name: string, banned: "docx" | "odf"): string[] {
    const source = readFileSync(join(dir, name), "utf8");
    const bannedSegment = new RegExp(`(^|/)${banned}/`);
    return importSpecifiers(source)
      .filter((specifier) => bannedSegment.test(specifier))
      .map((specifier) => `${name} imports "${specifier}"`);
  }

  // Runtime modules only, like coreModules(): the round-trip specs
  // (docx/roundTrip.spec.ts, odf/writer.spec.ts, …) deliberately read one
  // format and write the other — that is the point of a round-trip test.
  it("no odf/ module imports from docx/", () => {
    const odfDir = join(libDir, "odf");
    const violations = runtimeModules(odfDir).flatMap((name) => formatImportViolations(odfDir, name, "docx"));
    expect(violations).toEqual([]);
  });

  it("no docx/ module imports from odf/", () => {
    const docxDir = join(libDir, "docx");
    const violations = runtimeModules(docxDir).flatMap((name) => formatImportViolations(docxDir, name, "odf"));
    expect(violations).toEqual([]);
  });

  // The model layer must not import a format module — a reader/writer knows
  // the model, never the reverse. Enforced for an explicit list rather than
  // all of src/lib because src/lib also holds the format-orchestration
  // modules whose whole job is to call readers/writers: openDocument.ts and
  // the workers deliberately import both readers (they are the open path),
  // and merge.ts deliberately imports both writers (mail-merge output saves
  // a .docx/.odt package per record) — none of those are listed, and that
  // is not a violation. Everything listed here is pure model-shaped logic.
  const MODEL_LAYER_MODULES = [
    "documentModel.ts",
    "commands.ts",
    "textEditing.ts",
    "styles.ts",
    "equality.ts",
    "opaque.ts",
    "tables.ts",
    "images.ts",
    "mediaRefs.ts",
    "shapes.ts",
    "charts.ts",
    "links.ts",
    "comments.ts",
    "notes.ts",
    "lists.ts",
    "trackChanges.ts",
    "trackChangesDiff.ts",
    "versionHistory.ts",
    "clipboard.ts",
    "findReplace.ts",
  ];

  it("model-layer modules import no format module", () => {
    const violations = MODEL_LAYER_MODULES.flatMap((name) => [
      ...formatImportViolations(libDir, name, "docx"),
      ...formatImportViolations(libDir, name, "odf"),
    ]);
    expect(violations).toEqual([]);
  });
});

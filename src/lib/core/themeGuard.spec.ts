import { describe, it, expect } from "vitest";
import { readFileSync, readdirSync } from "node:fs";
import { join, relative } from "node:path";

// SPEC-077 AC-7's mechanical guard: `resolveColor`/`resolveFont` (with the
// `isThemeColor`/`isThemeFont`/`sameModelColor`/`sameModelFont` helpers) in
// src/lib/core/theme.ts are the ONLY sanctioned readers of a ModelColor /
// ModelFont reference — no other file may pick a scheme reference apart and
// interpret its fields itself.
//
// A test cannot type-check "reads `.theme` off a colour value" syntactically,
// so the rule is a pragmatic mechanical one, precise enough to catch a
// renderer doing `run.formatting.color.theme` while never firing on
// `props.theme`, `model.theme`, `card.theme` etc.:
//
//   Outside the allow-list below, over every src/**/*.{ts,vue} source
//   (comments stripped, *.spec.ts excluded — tests construct references):
//
//   RULE 1 — no property access to a ThemeColorRef transform field:
//     /\.(lumMod|lumOff|shade|tint)\b/
//     These four names are the reference's transform vocabulary; nothing in
//     the model uses them for anything else, so any access outside the
//     resolution/serialisation layer is a hand-rolled resolver.
//
//   RULE 2 — no `.theme` access off a colour-named expression:
//     /(color|colour|fill|outline|stroke|shading|shd|highlight)[a-z_$]*(\.value)?\??\.theme\b/i
//     i.e. `.theme` read from an identifier segment whose name marks it as a
//     colour value (`run.color.theme`, `solidFill?.theme`,
//     `strokeColor.value.theme`), which is exactly the shape of bypassing
//     isThemeColor/resolveColor. Non-colour `.theme` reads — a document
//     model's `model.theme`, a component's `props.theme`, the app-shell
//     preference `prefs.theme` — do not match.
//
// The allow-list is by role, not a frozen inventory:
//   - the resolution point itself and the reference *serialisers* (files whose
//     whole job is turning a ThemeColorRef into wire attributes and back);
//   - the presentation product (shipped under SPEC-094 with its own
//     reader/writer set under src/lib/presentation/**);
//   - picker/swatch UI, which legitimately BUILDS references from
//     THEME_SWATCH_VARIANTS (`variant.lumMod` → new ref) and stamps
//     `data-theme-slot` for tests.
const ALLOWED: RegExp[] = [
  // The single resolution point (defines and interprets the fields).
  /^lib\/core\/theme\.ts$/,
  // Core reference serialisers: themeOdf.ts, themeDrawingMl.ts and any future
  // core theme codec — they emit/parse the reference's wire form.
  /^lib\/core\/theme[A-Z][A-Za-z]*\.ts$/,
  // Format-layer theme parts (w:themeColor/w:themeTint attrs etc.).
  /^lib\/(docx|odf|md)\/theme[A-Za-z]*\.ts$/,
  // The presentation product's readers/writers (SPEC-094) serialise
  // references throughout pptx/odp.
  /^lib\/presentation\//,
  // Swatch-building UI: colour pickers, swatch rows, style popovers and the
  // theme gallery construct references from THEME_SWATCH_VARIANTS.
  /(ColorPicker|StylePopover|ThemeGallery|ThemeSwatch)[^/]*\.(vue|ts)$/,
];

const TRANSFORM_ACCESS = /\.(lumMod|lumOff|shade|tint)\b/;
const COLOUR_THEME_ACCESS = /(color|colour|fill|outline|stroke|shading|shd|highlight)[a-z_$]*(\.value)?\??\.theme\b/i;

const SRC = join(__dirname, "..", "..");

function stripComments(source: string): string {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, (c) => c.replace(/[^\n]/g, " "))
    .replace(/<!--[\s\S]*?-->/g, (c) => c.replace(/[^\n]/g, " "))
    .replace(/(^|[^:"'])\/\/[^\n]*/g, (m, lead: string) => lead);
}

function violations(source: string): string[] {
  const out: string[] = [];
  const lines = stripComments(source).split("\n");
  lines.forEach((line, i) => {
    if (TRANSFORM_ACCESS.test(line)) {
      out.push(`line ${i + 1}: transform-field access (${line.trim()})`);
    }
    if (COLOUR_THEME_ACCESS.test(line)) {
      out.push(`line ${i + 1}: .theme read off a colour value (${line.trim()})`);
    }
  });
  return out;
}

function sourceFiles(dir: string = SRC, acc: string[] = []): string[] {
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) {
      sourceFiles(full, acc);
    } else if (
      (entry.name.endsWith(".ts") || entry.name.endsWith(".vue")) &&
      !entry.name.endsWith(".spec.ts") &&
      !entry.name.endsWith(".d.ts")
    ) {
      acc.push(full);
    }
  }
  return acc;
}

describe("only core/theme.ts and the reference serialisers read ThemeColorRef fields (SPEC-077 AC-7)", () => {
  it("no file outside the allow-list reads .theme/.lumMod/.lumOff/.shade/.tint off a colour value", () => {
    const bad: string[] = [];
    for (const file of sourceFiles()) {
      const rel = relative(SRC, file).replace(/\\/g, "/");
      if (ALLOWED.some((re) => re.test(rel))) {
        continue;
      }
      for (const v of violations(readFileSync(file, "utf8"))) {
        bad.push(`  ${rel} ${v}`);
      }
    }
    expect(
      bad,
      `ThemeColorRef fields read outside resolveColor/isThemeColor (route through src/lib/core/theme.ts):\n${bad.join("\n")}`,
    ).toEqual([]);
  });

  // Self-test: the scanner is actually capable of failing — it flags the
  // shapes the rule targets and stays quiet on the innocent look-alikes.
  it("flags a renderer reading .theme off a run colour", () => {
    expect(violations(`const css = run.formatting.color.theme;`)).toHaveLength(1);
    expect(violations(`const hex = table.cell.shadingFill?.theme ?? "";`)).toHaveLength(1);
    expect(violations(`const slot = strokeColor.value.theme;`)).toHaveLength(1);
  });

  it("flags a hand-rolled transform resolver", () => {
    expect(violations(`const l = ref.lumMod ?? 1;`)).toHaveLength(1);
    expect(violations(`if (c.shade !== undefined) darken(c.shade);`).length).toBeGreaterThan(0);
  });

  it("does not flag non-colour .theme reads or comments", () => {
    expect(violations(`const t = props.theme;`)).toEqual([]);
    expect(violations(`model.theme = next;`)).toEqual([]);
    expect(violations(`const pick = card.theme && prefs.theme;`)).toEqual([]);
    expect(violations(`// a comment about run.color.theme and ref.lumMod`)).toEqual([]);
  });
});

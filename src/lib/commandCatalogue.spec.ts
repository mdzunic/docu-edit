import { readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import {
  COMMANDS,
  COMMON_COMMAND_IDS,
  loadRecentCommands,
  pushRecentCommand,
  RECENT_COMMANDS_LIMIT,
  searchCommands,
} from "./commandCatalogue";

const here = dirname(fileURLToPath(import.meta.url));

/** Extracts the `["id", …]` tuple keys declared inside one array block of a
 * source file — the same source-reading discipline DocumentView.size.spec.ts
 * uses, so the catalogue cannot drift from the live registry without a test
 * turning red. */
function tupleKeysInBlock(source: string, openMarker: string): string[] {
  const start = source.indexOf(openMarker);
  if (start === -1) {
    throw new Error(`marker not found: ${openMarker}`);
  }
  const rest = source.slice(start + openMarker.length);
  // The block ends at "]);" (a Map literal) or a bare "];" (a command array),
  // whichever comes first.
  const ends = [rest.indexOf("]);"), rest.indexOf("];")].filter((index) => index !== -1);
  const block = ends.length === 0 ? rest : rest.slice(0, Math.min(...ends));
  return [...block.matchAll(/\[\s*"([a-zA-Z][\w.-]*)"\s*,/g)].map((match) => match[1]);
}

/** The live `editorCommands` keys, read from source: the inline Map in
 * DocumentView.vue plus the command lists spread in from useThemeControls,
 * useFileActions (File -> Export) and usePageSetupDialogs (the ruler
 * toggles). */
const COMMAND_LIST_MARKER = "const commands: Array<[string, () => void]> = [";

function liveRegistryIds(): string[] {
  const view = readFileSync(join(here, "../views/DocumentView.vue"), "utf8");
  const theme = readFileSync(join(here, "../views/document/useThemeControls.ts"), "utf8");
  const fileActions = readFileSync(join(here, "../views/document/useFileActions.ts"), "utf8");
  const pageSetup = readFileSync(join(here, "../views/document/usePageSetupDialogs.ts"), "utf8");
  const drawingArrange = readFileSync(join(here, "../views/document/useDrawingArrange.ts"), "utf8");
  const tableTools = readFileSync(join(here, "../views/document/useTableTools.ts"), "utf8");
  const inline = tupleKeysInBlock(view, "const editorCommands = new Map([");
  const themeBlock = theme.slice(theme.indexOf("themeCommands"));
  const themeIds = [...themeBlock.matchAll(/\[\s*"([a-zA-Z][\w.-]*)"\s*,/g)].map((m) => m[1]);
  const exportIds = tupleKeysInBlock(fileActions, COMMAND_LIST_MARKER);
  const rulerIds = tupleKeysInBlock(pageSetup, COMMAND_LIST_MARKER);
  // SPEC-072 drawing align/distribute command ids (useDrawingArrange).
  const arrangeIds = tupleKeysInBlock(drawingArrange, COMMAND_LIST_MARKER);
  // SPEC-076 Table menu command ids (useTableTools).
  const tableIds = tupleKeysInBlock(tableTools, COMMAND_LIST_MARKER);
  return [...new Set([...inline, ...themeIds, ...exportIds, ...rulerIds, ...arrangeIds, ...tableIds])];
}

describe("COMMANDS completeness (SPEC-062 AC-31, AC-32)", () => {
  it("has a document-scope entry for exactly every editorCommands id", () => {
    const registry = new Set(liveRegistryIds());
    const catalogue = new Set(COMMANDS.filter((c) => c.scope === "document").map((c) => c.id));
    // Both directions: a new registry id without a catalogue entry fails here,
    // and so does a catalogue entry for an id that no longer dispatches.
    const missingFromCatalogue = [...registry].filter((id) => !catalogue.has(id));
    const staleInCatalogue = [...catalogue].filter((id) => !registry.has(id));
    expect(missingFromCatalogue).toEqual([]);
    expect(staleInCatalogue).toEqual([]);
  });

  it("gives every command a label and a non-empty menu path", () => {
    for (const command of COMMANDS) {
      expect(command.label.length).toBeGreaterThan(0);
      expect(command.path.length).toBeGreaterThan(0);
    }
  });

  it("has no duplicate ids", () => {
    const ids = COMMANDS.map((c) => c.id);
    expect(new Set(ids).size).toBe(ids.length);
  });
});

describe("searchCommands ranking (SPEC-062)", () => {
  const find = (query: string, id: string) =>
    searchCommands(query).find((match) => match.meta.id === id);

  it("returns every command in catalogue order for an empty query", () => {
    const results = searchCommands("");
    expect(results.map((r) => r.meta.id)).toEqual(COMMANDS.map((c) => c.id));
  });

  it("ranks an exact label match first", () => {
    const results = searchCommands("bold");
    expect(results[0].meta.id).toBe("format.bold");
  });

  it("ranks a label prefix above a mid-label substring", () => {
    const results = searchCommands("under");
    const underline = results.findIndex((r) => r.meta.id === "format.underline");
    expect(underline).toBe(0);
  });

  it("matches by subsequence, not just substring (ftnt → Footnote)", () => {
    const results = searchCommands("ftnt");
    expect(results.some((r) => r.meta.id === "insert.footnote")).toBe(true);
  });

  it("matches across label and path (instab → Insert table)", () => {
    const results = searchCommands("instab");
    expect(results.some((r) => r.meta.id === "insert.table-dialog")).toBe(true);
  });

  it("finds a command by an alias with no label highlight", () => {
    const a11y = find("a11y", "tools.accessibility");
    expect(a11y).toBeDefined();
    // "a11y" cannot be a subsequence of "Check accessibility" (no digits); it
    // matches via the alias, so there is nothing in the label to highlight.
    expect(a11y?.ranges).toEqual([]);
  });

  it("is case- and accent-insensitive", () => {
    expect(searchCommands("CENTRE").some((r) => r.meta.id === "paragraph.alignCenter")).toBe(true);
    expect(searchCommands("centré").some((r) => r.meta.id === "paragraph.alignCenter")).toBe(true);
  });

  it("highlights the matched characters in the label", () => {
    const match = find("clear", "format.clearFormatting");
    expect(match?.ranges).toEqual([[0, 5]]);
  });

  it("breaks ties toward more-recently-used", () => {
    const withoutRecent = searchCommands("heading", COMMANDS);
    const withRecent = searchCommands("heading", COMMANDS, ["style.heading3"]);
    // All the heading style rows score identically (alias/subsequence); the
    // recent one is pulled ahead of where alphabetical order would put it.
    const plainIndex = withoutRecent.findIndex((r) => r.meta.id === "style.heading3");
    const recentIndex = withRecent.findIndex((r) => r.meta.id === "style.heading3");
    expect(recentIndex).toBeLessThanOrEqual(plainIndex);
    expect(withRecent[0].meta.id).toBe("style.heading3");
  });

  it("only searches the given command list (library scope isolation)", () => {
    const libraryOnly = COMMANDS.filter((c) => c.scope === "library");
    const results = searchCommands("new", libraryOnly);
    expect(results.every((r) => r.meta.scope === "library")).toBe(true);
    expect(results.some((r) => r.meta.id === "library.new-document")).toBe(true);
  });
});

describe("recent commands persistence (SPEC-062)", () => {
  beforeEach(() => localStorage.clear());
  afterEach(() => localStorage.clear());

  it("round-trips through localStorage, most recent first", () => {
    let recent: string[] = [];
    recent = pushRecentCommand("edit.undo", recent);
    recent = pushRecentCommand("format.bold", recent);
    expect(recent).toEqual(["format.bold", "edit.undo"]);
    expect(loadRecentCommands()).toEqual(["format.bold", "edit.undo"]);
  });

  it("de-duplicates and moves a repeat to the front", () => {
    const recent = pushRecentCommand("format.bold", ["edit.undo", "format.bold"]);
    expect(recent).toEqual(["format.bold", "edit.undo"]);
  });

  it("caps the list at the limit", () => {
    let recent: string[] = [];
    for (let i = 0; i < RECENT_COMMANDS_LIMIT + 5; i++) {
      recent = pushRecentCommand(`id-${i}`, recent);
    }
    expect(recent.length).toBe(RECENT_COMMANDS_LIMIT);
  });

  it("drops ids no longer in the catalogue on read", () => {
    localStorage.setItem("docu-edit:recent-commands", JSON.stringify(["format.bold", "gone.forever"]));
    expect(loadRecentCommands()).toEqual(["format.bold"]);
  });

  it("returns [] for malformed storage", () => {
    localStorage.setItem("docu-edit:recent-commands", "{not json");
    expect(loadRecentCommands()).toEqual([]);
  });

  it("exposes common command ids that all exist in the catalogue", () => {
    const ids = new Set(COMMANDS.map((c) => c.id));
    for (const id of COMMON_COMMAND_IDS) {
      expect(ids.has(id)).toBe(true);
    }
  });
});

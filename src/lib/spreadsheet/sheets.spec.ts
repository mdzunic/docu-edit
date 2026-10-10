// SPEC-121 "Sheet commands": the blank workbook, the sheet-name rules and
// their exact messages, Insert's and Duplicate's naming, every command's
// result (order, active sheet, alsoRecalc), the flags (`structureDirty`,
// `dirty`, `copiedFrom`), copy-on-write (no command mutates a pre-existing
// Sheet, Row or Cell — the inputs are deep-frozen to prove it), the copy's
// own retained source at command time (xlsx closure with its fix-ups, ods
// ids and embedded objects), the cross-format copy (`stripRetention`) and
// what it leaves out (`unconvertibleContent`).
import { afterEach, describe, expect, it, vi } from "vitest";
import { base64ToBytes, bytesToBase64 } from "../core/packageParts";
import { DEFAULT_THEME, type ModelColor } from "../core/theme";
import { parseRelationships, relsPathFor, resolvePartPath } from "../core/xml";
import { getCell, parseA1 } from "./address";
import {
  buildMiniBookOdsBytes,
  buildSpec119FormulasXlsxBytes,
  buildSpec121LinksOdsBytes,
  buildSpec121LinksXlsxBytes,
  buildSpecMacrosXlsmBytes,
  SPEC121_LINKS_XLSX,
} from "./fixtures/build";
import { cellKey } from "./formula/dependencies";
import { formulaForTarget } from "./formula/dialect";
import { renameSheetInFormula } from "./formula/references";
import { DEFAULT_SPREADSHEET_TITLE, type Cell, type CellValue, type Row, type Sheet, type SpreadsheetFormat, type SpreadsheetModel } from "./model";
import { readSpreadsheetBytes, SPREADSHEET_READER_VERSION } from "./open";
import {
  addSheet,
  cannotDuplicateReason,
  createBlankWorkbook,
  deleteSheet,
  duplicateSheet,
  isLastVisibleSheet,
  moveSheet,
  nextSheetName,
  renameSheet,
  setSheetHidden,
  setTabColor,
  sheetHasContent,
  stripRetention,
  UNCONVERTIBLE_CONTENT,
  unconvertibleContent,
  validateSheetName,
} from "./sheets";

// --- messages, as "Sheet names" and the AC spell them -----------------------------------

const BLANK = "A sheet name can't be blank";
const TOO_LONG = "A sheet name can be at most 31 characters";
const CHARACTERS = "A sheet name can't contain any of these characters: [ ] : * ? / \\";
const APOSTROPHE = "A sheet name can't begin or end with an apostrophe";
const RESERVED = '"History" is reserved — choose another name';
const taken = (name: string): string => `There's already a sheet called "${name}"`;
const LAST_VISIBLE = "A workbook must keep at least one visible sheet";
const FOREIGN_COPY = "This sheet type can't be copied yet";
const UNSAFE_COPY = "This sheet has content that can't be copied yet";

const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

// --- hand-built workbooks ----------------------------------------------------------------

/** A cell spec: a literal value, or a formula cell (its cached value null unless given). */
type CellSpec = CellValue | { formula: string; value?: CellValue; formulaUnparsed?: true };

const f = (formula: string, value: CellValue = null): CellSpec => ({ formula, value });

/** A grid sheet from A1-addressed cells, with no source. */
function gridSheet(name: string, cells: Record<string, CellSpec> = {}, extra: Partial<Sheet> = {}): Sheet {
  const rows: Row[] = [];
  const entries = Object.entries(cells)
    .map(([a1, spec]) => ({ ref: parseA1(a1)!, spec }))
    .toSorted((a, b) => a.ref.row - b.ref.row || a.ref.col - b.ref.col);
  for (const { ref, spec } of entries) {
    const cell: Cell =
      spec !== null && typeof spec === "object" && "formula" in spec
        ? { col: ref.col, value: spec.value ?? null, formula: spec.formula, ...(spec.formulaUnparsed ? { formulaUnparsed: true as const } : {}) }
        : { col: ref.col, value: spec };
    const last = rows.at(-1);
    if (last?.index === ref.row) {
      last.cells.push(cell);
    } else {
      rows.push({ index: ref.row, cells: [cell] });
    }
  }
  return {
    id: crypto.randomUUID(),
    name,
    visibility: "visible",
    kind: "grid",
    rows,
    columns: [],
    merges: [],
    defaultColWidthPx: 64,
    defaultRowHeightPt: 15,
    ...extra,
  };
}

function foreignSheet(name: string, partName: string, extra: Partial<Sheet> = {}): Sheet {
  return { ...gridSheet(name), kind: "foreign", partName, ...extra };
}

function book(sheets: Sheet[], activeSheetIndex = 0, format: SpreadsheetFormat = "xlsx"): SpreadsheetModel {
  return {
    kind: "spreadsheet",
    id: crypto.randomUUID(),
    title: "Hand built",
    format,
    readerVersion: SPREADSHEET_READER_VERSION,
    sheets,
    activeSheetIndex,
    dateSystem: "1900",
    opaque: [],
    styles: [{}],
    defaultStyleId: 0,
  };
}

const names = (model: SpreadsheetModel): string[] => model.sheets.map((sheet) => sheet.name);

function sheetNamed(model: SpreadsheetModel, name: string): Sheet {
  const sheet = model.sheets.find((candidate) => candidate.name === name);
  expect(sheet, `sheet "${name}"`).toBeDefined();
  return sheet!;
}

function cellAt(sheet: Sheet, a1: string): Cell | undefined {
  const ref = parseA1(a1)!;
  return getCell(sheet, ref.row, ref.col);
}

const formulaAt = (sheet: Sheet, a1: string): string | undefined => cellAt(sheet, a1)?.formula;

function keyAt(sheet: Sheet, a1: string): string {
  const ref = parseA1(a1)!;
  return cellKey(sheet.id, ref.row, ref.col);
}

/** The message `run` throws (fails the test when it does not throw). */
function thrownMessage(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return (error as Error).message;
  }
  throw new Error("expected the command to throw");
}

/** Freezes a value and everything it holds. */
function deepFreeze<T>(value: T): T {
  if (value !== null && typeof value === "object" && !Object.isFrozen(value)) {
    Object.freeze(value);
    for (const child of Object.values(value)) {
      deepFreeze(child);
    }
  }
  return value;
}

/** Deep-freezes everything a command must replace rather than mutate: the
 * sheet list and every Sheet, Row and Cell in it, `opaque` and its parts,
 * `styles`, the ODS envelope. The model object itself stays writable (a
 * command assigns its fields). */
function freezeContents(model: SpreadsheetModel): SpreadsheetModel {
  deepFreeze(model.sheets);
  deepFreeze(model.opaque);
  deepFreeze(model.styles);
  deepFreeze(model.odsContentEnvelope);
  return model;
}

// --- fixtures and their parts ------------------------------------------------------------

function once(build: () => Promise<Uint8Array>): () => Promise<Uint8Array> {
  let bytes: Promise<Uint8Array> | undefined;
  return () => (bytes ??= build());
}
const linksXlsxBytes = once(buildSpec121LinksXlsxBytes);
const linksOdsBytes = once(buildSpec121LinksOdsBytes);
const formulasXlsxBytes = once(buildSpec119FormulasXlsxBytes);
const macrosXlsmBytes = once(buildSpecMacrosXlsmBytes);
const miniOdsBytes = once(buildMiniBookOdsBytes);

const openLinksXlsx = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await linksXlsxBytes(), "xlsx", "Links");
const openLinksOds = async (): Promise<SpreadsheetModel> => readSpreadsheetBytes(await linksOdsBytes(), "ods", "Links");

const decoder = new TextDecoder();
const encoder = new TextEncoder();

function partText(model: SpreadsheetModel, name: string): string | undefined {
  const part = model.opaque.find((candidate) => candidate.name === name);
  return part === undefined ? undefined : decoder.decode(base64ToBytes(part.data));
}

function partData(model: SpreadsheetModel, name: string): string | undefined {
  return model.opaque.find((candidate) => candidate.name === name)?.data;
}

/** Sets (or adds) a part's text in a NEW opaque array. */
function setPart(model: SpreadsheetModel, name: string, text: string): void {
  const data = bytesToBase64(encoder.encode(text));
  model.opaque = model.opaque.some((part) => part.name === name)
    ? model.opaque.map((part) => (part.name === name ? { name, data } : part))
    : [...model.opaque, { name, data }];
}

function editPart(model: SpreadsheetModel, name: string, edit: (text: string) => string): void {
  const text = partText(model, name);
  expect(text, name).toBeDefined();
  setPart(model, name, edit(text!));
}

const addRelationship = (xml: string, relationship: string): string => xml.replace("</Relationships>", `${relationship}</Relationships>`);

/** A part's relationships, each target resolved to a part path (external targets as written). */
function resolvedRels(model: SpreadsheetModel, part: string): { id: string; type: string; target: string }[] {
  const xml = partText(model, relsPathFor(part));
  expect(xml, relsPathFor(part)).toBeDefined();
  return [...parseRelationships(xml!).values()]
    .map((rel) => ({ id: rel.id, type: rel.type, target: rel.targetMode === "External" ? rel.target : resolvePartPath(part, rel.target) }))
    .toSorted((a, b) => a.id.localeCompare(b.id));
}

/** The value of `attribute` on the first `<local …>` start tag of `xml`. */
function attributeOf(xml: string, local: string, attribute: string): string | undefined {
  const tag = new RegExp(`<(?:\\w+:)?${local}\\b[^>]*>`).exec(xml)?.[0];
  return tag === undefined ? undefined : new RegExp(`\\s${attribute}="([^"]*)"`).exec(tag)?.[1];
}

/** The manifest's full paths with their media types. */
function manifestEntries(model: SpreadsheetModel): Map<string, string> {
  const xml = partText(model, "META-INF/manifest.xml") ?? "";
  const entries = new Map<string, string>();
  for (const match of xml.matchAll(/<manifest:file-entry\b[^>]*>/g)) {
    const path = /manifest:full-path="([^"]*)"/.exec(match[0])?.[1];
    const type = /manifest:media-type="([^"]*)"/.exec(match[0])?.[1];
    if (path !== undefined) {
      entries.set(path, type ?? "");
    }
  }
  return entries;
}

/** A minimal ODS `<table:table>` text around `inner`. */
const odsTable = (inner: string): string => `<table:table table:name="T" table:style-name="ta1">${inner}<table:table-column/></table:table>`;

/** The labels in UNCONVERTIBLE_CONTENT order, restricted to `labels`. */
const inTableOrder = (labels: string[]): string[] => UNCONVERTIBLE_CONTENT.map((row) => row.label).filter((label) => labels.includes(label));

const unconvertibleRow = (label: string): (typeof UNCONVERTIBLE_CONTENT)[number] => UNCONVERTIBLE_CONTENT.find((candidate) => candidate.label === label)!;

afterEach(() => {
  vi.restoreAllMocks();
});

// =========================================================================================

describe("createBlankWorkbook", () => {
  it('returns exactly "The blank workbook": one visible, dirty grid sheet "Sheet1", xlsx from birth, the default style and theme', () => {
    const model = createBlankWorkbook();
    expect(model).toStrictEqual({
      kind: "spreadsheet",
      id: expect.stringMatching(UUID),
      title: "Untitled spreadsheet",
      format: "xlsx",
      readerVersion: SPREADSHEET_READER_VERSION,
      sheets: [
        {
          id: expect.stringMatching(UUID),
          name: "Sheet1",
          visibility: "visible",
          kind: "grid",
          rows: [],
          columns: [],
          merges: [],
          defaultColWidthPx: 64,
          defaultRowHeightPt: 15,
          dirty: true,
        },
      ],
      activeSheetIndex: 0,
      dateSystem: "1900",
      opaque: [],
      styles: [{ font: { family: { theme: "minor" }, sizePt: 11, color: { theme: "dk1" } }, namedStyle: "Normal" }],
      defaultStyleId: 0,
      theme: structuredClone(DEFAULT_THEME),
    });
  });

  it("titles it DEFAULT_SPREADSHEET_TITLE, and carries no properties, structureDirty or sheet source", () => {
    const model = createBlankWorkbook();
    expect(DEFAULT_SPREADSHEET_TITLE).toBe("Untitled spreadsheet");
    expect(model.title).toBe(DEFAULT_SPREADSHEET_TITLE);
    expect("properties" in model).toBe(false);
    expect("structureDirty" in model).toBe(false);
    expect("partName" in model.sheets[0]!).toBe(false);
    expect("sourceXml" in model.sheets[0]!).toBe(false);
    expect("copiedFrom" in model.sheets[0]!).toBe(false);
  });

  it("gives the theme as a copy of DEFAULT_THEME, never the shared object", () => {
    const model = createBlankWorkbook();
    expect(model.theme).toEqual(DEFAULT_THEME);
    expect(model.theme).not.toBe(DEFAULT_THEME);
    expect(model.theme!.colors).not.toBe(DEFAULT_THEME.colors);
    expect(model.theme!.fonts).not.toBe(DEFAULT_THEME.fonts);
  });

  it("mints fresh, distinct ids on every call", () => {
    const a = createBlankWorkbook();
    const b = createBlankWorkbook();
    expect(a.id).not.toBe(b.id);
    expect(a.sheets[0]!.id).not.toBe(b.sheets[0]!.id);
    expect(a.sheets[0]!.id).not.toBe(a.id);
  });

  it("is plain JSON", () => {
    const model = createBlankWorkbook();
    expect(JSON.parse(JSON.stringify(model))).toStrictEqual(model);
  });
});

// =========================================================================================

describe("validateSheetName", () => {
  const model = (): SpreadsheetModel =>
    book([
      gridSheet("Summary"),
      gridSheet("Data"),
      gridSheet("Archive", {}, { visibility: "hidden" }),
      gridSheet("Vault", {}, { visibility: "veryHidden" }),
      gridSheet("STRASSE"),
    ]);

  it("accepts ordinary names (null = valid), an apostrophe inside included", () => {
    for (const name of ["Budget", "O'Brien", "Q1 (Draft)", "My History", "History 2", "Sheet 9", "a"]) {
      expect(validateSheetName(model(), name), name).toBeNull();
    }
  });

  it("rejects a name that is empty after trimming", () => {
    for (const name of ["", " ", "   ", "\t"]) {
      expect(validateSheetName(model(), name), JSON.stringify(name)).toBe(BLANK);
    }
  });

  it("allows 31 characters and rejects 32", () => {
    expect(validateSheetName(model(), "x".repeat(31))).toBeNull();
    expect(validateSheetName(model(), "x".repeat(32))).toBe(TOO_LONG);
    expect(validateSheetName(model(), "y".repeat(40))).toBe(TOO_LONG);
  });

  it("rejects each of [ ] : * ? / \\ wherever it appears", () => {
    for (const character of ["[", "]", ":", "*", "?", "/", "\\"]) {
      expect(validateSheetName(model(), `Q${character}1`), character).toBe(CHARACTERS);
      expect(validateSheetName(model(), `${character}Q1`), character).toBe(CHARACTERS);
      expect(validateSheetName(model(), `Q1${character}`), character).toBe(CHARACTERS);
    }
  });

  it("rejects a leading or trailing apostrophe", () => {
    for (const name of ["'Budget", "Budget'", "'Budget'", "'"]) {
      expect(validateSheetName(model(), name), name).toBe(APOSTROPHE);
    }
  });

  it('reserves "History" in any case', () => {
    for (const name of ["History", "history", "HISTORY", "hIsToRy"]) {
      expect(validateSheetName(model(), name), name).toBe(RESERVED);
    }
  });

  it("rejects another sheet's name case-insensitively, hidden and very hidden sheets included", () => {
    expect(validateSheetName(model(), "Data")).toBe(taken("Data"));
    expect(validateSheetName(model(), "data")).toBe(taken("data"));
    expect(validateSheetName(model(), "ARCHIVE")).toBe(taken("ARCHIVE"));
    expect(validateSheetName(model(), "vault")).toBe(taken("vault"));
  });

  it("compares by toUpperCase (\"straße\" clashes with \"STRASSE\")", () => {
    expect(validateSheetName(model(), "straße")).toBe(taken("straße"));
  });

  it("ignores the sheet being renamed, so a case-only change of its own name is valid", () => {
    expect(validateSheetName(model(), "DATA", 1)).toBeNull();
    expect(validateSheetName(model(), "Data", 1)).toBeNull();
    expect(validateSheetName(model(), "summary", 1)).toBe(taken("summary"));
    expect(validateSheetName(model(), "data", 0)).toBe(taken("data"));
  });

  it("reports the first failing rule only", () => {
    expect(validateSheetName(model(), " ".repeat(40))).toBe(BLANK);
    expect(validateSheetName(model(), `[${"a".repeat(40)}`)).toBe(TOO_LONG);
    expect(validateSheetName(model(), "'a:b'")).toBe(CHARACTERS);
    expect(validateSheetName(model(), "Data*")).toBe(CHARACTERS);
    expect(validateSheetName(model(), "'History'")).toBe(APOSTROPHE);
    // A "History" read from a file is kept as read; validating its name hits the reserved rule first.
    const withHistory = book([gridSheet("History"), gridSheet("Other")]);
    expect(validateSheetName(withHistory, "history")).toBe(RESERVED);
  });
});

// =========================================================================================

describe("nextSheetName / addSheet", () => {
  it("names Sheet1 when no sheet is called Sheet<number>", () => {
    expect(nextSheetName(book([gridSheet("Summary"), gridSheet("Data")]))).toBe("Sheet1");
    expect(nextSheetName(book([gridSheet("Sheet 5"), gridSheet("Sheet5a"), gridSheet("MySheet9"), gridSheet("Sheet")]))).toBe("Sheet1");
  });

  it("takes one more than the largest number, not the first gap", () => {
    expect(nextSheetName(book([gridSheet("Sheet1"), gridSheet("Sheet3")]))).toBe("Sheet4");
    expect(nextSheetName(book([gridSheet("Sheet10"), gridSheet("Sheet2")]))).toBe("Sheet11");
  });

  it("counts names case-insensitively, hidden and very hidden sheets included", () => {
    expect(nextSheetName(book([gridSheet("Data"), gridSheet("sheet2"), gridSheet("SHEET7", {}, { visibility: "hidden" })]))).toBe("Sheet8");
    expect(nextSheetName(book([gridSheet("Sheet3"), gridSheet("Sheet12", {}, { visibility: "veryHidden" })]))).toBe("Sheet13");
  });

  it("inserts an empty grid sheet right after afterIndex, active, born dirty with no source", () => {
    const model = book([gridSheet("Sheet1"), gridSheet("Data"), gridSheet("Sheet3")], 1);
    const result = addSheet(model, 1);
    expect(names(model)).toEqual(["Sheet1", "Data", "Sheet4", "Sheet3"]);
    expect(result.activeIndex).toBe(2);
    expect(model.activeSheetIndex).toBe(2);
    const added = model.sheets[2]!;
    expect(added).toMatchObject({ name: "Sheet4", visibility: "visible", kind: "grid", rows: [], columns: [], merges: [], dirty: true });
    expect(added.id).toMatch(UUID);
    expect(new Set(model.sheets.map((sheet) => sheet.id)).size).toBe(4);
    expect("partName" in added || "sourceXml" in added || "copiedFrom" in added).toBe(false);
    expect(sheetHasContent(model, 2)).toBe(false);
  });

  it("inserts at the end after the last sheet and at the start after index -1", () => {
    const model = book([gridSheet("A"), gridSheet("B")]);
    addSheet(model, 1);
    expect(names(model)).toEqual(["A", "B", "Sheet1"]);
    addSheet(model, -1);
    expect(names(model)).toEqual(["Sheet2", "A", "B", "Sheet1"]);
    expect(model.activeSheetIndex).toBe(0);
  });

  it("replaces model.sheets rather than splicing it, keeps every other sheet object, and sets structureDirty", () => {
    const model = freezeContents(book([gridSheet("A", { A1: 1 }), gridSheet("B")]));
    const before = model.sheets;
    const [a, b] = before;
    addSheet(model, 0);
    expect(model.sheets).not.toBe(before);
    expect(before.map((sheet) => sheet.name)).toEqual(["A", "B"]);
    expect(model.sheets[0]).toBe(a);
    expect(model.sheets[2]).toBe(b);
    expect(model.structureDirty).toBe(true);
  });

  it("leaves an opened workbook's retained package alone: the new sheet's part is the writer's", async () => {
    const model = await openLinksXlsx();
    const opaque = model.opaque;
    addSheet(model, 0);
    expect(model.sheets[1]!.name).toBe("Sheet1");
    expect(model.opaque).toBe(opaque);
    expect(model.sheets.filter((sheet) => sheet.partName !== undefined)).toHaveLength(5);
  });

  it("lists in alsoRecalc the cells naming the new sheet and the 3-D spans the insertion falls inside", () => {
    const calc = gridSheet("Calc", {
      A1: f("Sheet1!A1*2"),
      A2: f("SUM(Jan:Mar!B2)"),
      A3: f("Jan!B2"),
      A4: f("SUM(Feb:Mar!B2)"),
      A5: f("1+1"),
      A6: f("SHEET1!B1"),
    });
    const model = book([gridSheet("Jan"), gridSheet("Feb"), gridSheet("Mar"), calc]);
    const { alsoRecalc } = addSheet(model, 0);
    expect(model.sheets[1]!.name).toBe("Sheet1");
    expect(alsoRecalc).toEqual(expect.arrayContaining([keyAt(calc, "A1"), keyAt(calc, "A2"), keyAt(calc, "A6")]));
    expect(alsoRecalc).not.toContain(keyAt(calc, "A3"));
    expect(alsoRecalc).not.toContain(keyAt(calc, "A4"));
    expect(alsoRecalc).not.toContain(keyAt(calc, "A5"));
  });

  it("lists no 3-D span the insertion falls outside", () => {
    const calc = gridSheet("Calc", { A1: f("SUM(Jan:Feb!B2)") });
    const model = book([gridSheet("Jan"), gridSheet("Feb"), calc]);
    expect(addSheet(model, 2).alsoRecalc).toEqual([]);
  });
});

// =========================================================================================

describe("renameSheet", () => {
  /** Summary's formulas name Data and Lookup; Data names itself; hidden Archive names Data. */
  function linkedBook(): SpreadsheetModel {
    return book([
      gridSheet("Summary", {
        A1: "Code",
        B1: "Target",
        A2: 103,
        B2: f("VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)", 360),
        C2: f("Data!B2*2", 200),
        D2: f("SUM(Data:Lookup!A1)", 101),
        E2: f("Lookup!B2", 240),
        F2: f("Facts!A1"),
        G2: { formula: "_xlfn.FUTURE(Data!A1)", value: 1, formulaUnparsed: true },
      }),
      gridSheet("Data", { A1: "Code", B2: 100, C2: f("Data!B2+1", 101), A5: 104 }),
      gridSheet("Lookup", { A1: 101, B2: 240 }),
      gridSheet("Archive", { A1: f("Data!B2", 100) }, { visibility: "hidden" }),
      gridSheet("Vault", { A1: "Secret" }, { visibility: "veryHidden" }),
    ]);
  }

  it("rewrites every formula naming the old name, in every sheet, hidden ones included", () => {
    const model = linkedBook();
    renameSheet(model, 1, "Facts");
    expect(names(model)).toEqual(["Summary", "Facts", "Lookup", "Archive", "Vault"]);
    const summary = model.sheets[0]!;
    expect(formulaAt(summary, "C2")).toBe("Facts!B2*2");
    expect(formulaAt(summary, "D2")).toBe("SUM(Facts:Lookup!A1)");
    expect(formulaAt(summary, "B2")).toBe("VLOOKUP(A2,Lookup!$A$1:$B$5,2,FALSE)");
    expect(formulaAt(summary, "E2")).toBe("Lookup!B2");
    expect(formulaAt(model.sheets[1]!, "C2")).toBe("Facts!B2+1");
    expect(formulaAt(model.sheets[3]!, "A1")).toBe("Facts!B2");
  });

  it("follows renameSheetInFormula, quoting the new name as SPEC-119 prints it", () => {
    for (const next of ["My Sheet", "O'Brien", "2026"]) {
      const before = linkedBook();
      const model = linkedBook();
      renameSheet(model, 1, next);
      before.sheets.forEach((sheet, index) => {
        for (const row of sheet.rows) {
          for (const cell of row.cells) {
            if (cell.formula === undefined || cell.formulaUnparsed) {
              continue;
            }
            const after = getCell(model.sheets[index]!, row.index, cell.col)!;
            expect(after.formula, `${next}: ${cell.formula}`).toBe(renameSheetInFormula(cell.formula, "Data", next));
          }
        }
      });
    }
    const model = linkedBook();
    renameSheet(model, 1, "My Sheet");
    expect(formulaAt(model.sheets[0]!, "C2")).toBe("'My Sheet'!B2*2");
  });

  it("builds new Cell, Row and rows only where a formula changed; every untouched row, cell and sheet keeps its identity", () => {
    const model = freezeContents(linkedBook());
    const [summary, data, lookup, archive, vault] = model.sheets;
    renameSheet(model, 1, "Facts");
    const next = model.sheets[0]!;
    expect(next).not.toBe(summary);
    expect(next.rows).not.toBe(summary!.rows);
    expect(next.rows[0]).toBe(summary!.rows[0]);
    expect(next.rows[1]).not.toBe(summary!.rows[1]);
    expect(cellAt(next, "A2")).toBe(cellAt(summary!, "A2"));
    expect(cellAt(next, "B2")).toBe(cellAt(summary!, "B2"));
    expect(cellAt(next, "E2")).toBe(cellAt(summary!, "E2"));
    expect(cellAt(next, "C2")).not.toBe(cellAt(summary!, "C2"));
    expect(model.sheets[1]!.rows[0]).toBe(data!.rows[0]);
    expect(model.sheets[2]).toBe(lookup);
    expect(model.sheets[3]).not.toBe(archive);
    expect(model.sheets[4]).toBe(vault);
  });

  it("marks dirty exactly the sheets whose formulas changed", () => {
    const model = linkedBook();
    renameSheet(model, 1, "Facts");
    expect(model.sheets.map((sheet) => sheet.dirty === true)).toEqual([true, true, false, true, false]);
  });

  it("replaces the renamed sheet with a new object and leaves the old one as it was; a rename alone does not dirty it", () => {
    const model = linkedBook();
    const lookup = model.sheets[2]!;
    renameSheet(model, 2, "Ref");
    expect(model.sheets[2]).not.toBe(lookup);
    expect(lookup.name).toBe("Lookup");
    expect(model.sheets[2]).toEqual({ ...lookup, name: "Ref" });
    expect(model.sheets[2]!.dirty).toBeUndefined();
    expect(model.structureDirty).toBe(true);
  });

  it("keeps the active sheet", () => {
    const model = linkedBook();
    model.activeSheetIndex = 2;
    expect(renameSheet(model, 1, "Facts").activeIndex).toBe(2);
    expect(model.activeSheetIndex).toBe(2);
  });

  it("accepts a case-only change of the sheet's own name and rewrites the references with it", () => {
    const model = linkedBook();
    renameSheet(model, 1, "DATA");
    expect(model.sheets[1]!.name).toBe("DATA");
    expect(formulaAt(model.sheets[0]!, "C2")).toBe("DATA!B2*2");
    expect(model.structureDirty).toBe(true);
  });

  it("changes nothing for the unchanged name: same sheet list, same objects, no structureDirty", () => {
    const model = freezeContents(linkedBook());
    const sheets = model.sheets;
    const result = renameSheet(model, 1, "Data");
    expect(result).toEqual({ activeIndex: 0, alsoRecalc: [] });
    expect(model.sheets).toBe(sheets);
    expect(model.structureDirty).toBeUndefined();
  });

  it("changes nothing for the unchanged name of a sheet whose file-given name breaks the rules (kept as read)", () => {
    const long = "A long sheet name from LibreOffice Calc";
    expect(long.length).toBeGreaterThan(31);
    const model = book([gridSheet(long), gridSheet("Other", { A1: f(`'${long}'!A1`) })], 0, "ods");
    const sheets = model.sheets;
    expect(() => renameSheet(model, 0, long)).not.toThrow();
    expect(model.sheets).toBe(sheets);
    expect(model.structureDirty).toBeUndefined();
  });

  it("throws the exact message for an invalid name and changes nothing", () => {
    const cases: [string, string][] = [
      ["", BLANK],
      ["x".repeat(32), TOO_LONG],
      ["Q1/Q2", CHARACTERS],
      ["'Facts", APOSTROPHE],
      ["history", RESERVED],
      ["summary", taken("summary")],
      ["Vault", taken("Vault")],
    ];
    for (const [name, message] of cases) {
      const model = freezeContents(linkedBook());
      const sheets = model.sheets;
      expect(thrownMessage(() => renameSheet(model, 1, name)), name).toBe(message);
      expect(model.sheets).toBe(sheets);
      expect(model.structureDirty).toBeUndefined();
    }
  });

  it("leaves formulaUnparsed wire text untouched", () => {
    const model = linkedBook();
    const unparsed = cellAt(model.sheets[0]!, "G2");
    renameSheet(model, 1, "Facts");
    expect(cellAt(model.sheets[0]!, "G2")).toBe(unparsed);
    expect(unparsed!.formula).toBe("_xlfn.FUTURE(Data!A1)");
  });

  it("lists in alsoRecalc the cells naming the new name, and not the ones it does not concern", () => {
    const model = linkedBook();
    const summary = model.sheets[0]!;
    const { alsoRecalc } = renameSheet(model, 1, "Facts");
    expect(alsoRecalc).toContain(keyAt(summary, "F2"));
    expect(alsoRecalc).not.toContain(keyAt(summary, "B2"));
    expect(alsoRecalc).not.toContain(keyAt(summary, "E2"));
  });

  it("on spec121-links.xlsx: Data → Facts gives Summary!C2 =Facts!B2*2, and the retained package stays source-time", async () => {
    const model = await openLinksXlsx();
    const opaque = model.opaque;
    const partName = model.sheets[1]!.partName;
    renameSheet(model, 1, "Facts");
    const summary = sheetNamed(model, "Summary");
    expect(formulaAt(summary, "C2")).toBe("Facts!B2*2");
    expect(formulaAt(summary, "D2")).toBe("SUM(Facts:Lookup!A1)");
    expect(model.sheets[1]!.partName).toBe(partName);
    expect(model.sheets[1]!.copiedFrom).toBeUndefined();
    expect(model.opaque).toBe(opaque);
  });
});

// =========================================================================================

describe("duplicateSheet", () => {
  describe("names", () => {
    const copyName = (sheetNames: string[], index: number): string => {
      const model = book(sheetNames.map((name) => gridSheet(name)));
      duplicateSheet(model, index);
      return model.sheets[index + 1]!.name;
    };

    it("names the copy <base> (<n>) with the smallest free n ≥ 2", () => {
      expect(copyName(["Data"], 0)).toBe("Data (2)");
      expect(copyName(["Data", "Data (2)"], 0)).toBe("Data (3)");
      expect(copyName(["Data", "Data (3)"], 0)).toBe("Data (2)");
    });

    it("strips one trailing ` (<digits>)` from the source name to find the base", () => {
      expect(copyName(["Data", "Data (2)"], 1)).toBe("Data (3)");
      expect(copyName(["Budget (7)"], 0)).toBe("Budget (2)");
      expect(copyName(["Q1 (Draft)"], 0)).toBe("Q1 (Draft) (2)");
      expect(copyName(["Plan(2)"], 0)).toBe("Plan(2) (2)");
    });

    it("treats a name as taken case-insensitively", () => {
      expect(copyName(["Data", "DATA (2)"], 0)).toBe("Data (3)");
    });

    it("truncates the base so the name fits 31 characters", () => {
      const long = "R".repeat(31);
      expect(copyName([long], 0)).toBe(`${"R".repeat(27)} (2)`);
      const crowded = [long, ...[2, 3, 4, 5, 6, 7, 8, 9].map((n) => `${"R".repeat(27)} (${n})`)];
      expect(copyName(crowded, 0)).toBe(`${"R".repeat(26)} (10)`);
      expect(copyName(["A long sheet name from LibreOffice Calc"], 0)).toBe("A long sheet name from Libr (2)");
    });
  });

  it("inserts the copy right after its source and makes it active", () => {
    const model = book([gridSheet("A"), gridSheet("B", { A1: 1 }), gridSheet("C")]);
    const result = duplicateSheet(model, 1);
    expect(names(model)).toEqual(["A", "B", "B (2)", "C"]);
    expect(result.activeIndex).toBe(2);
    expect(model.activeSheetIndex).toBe(2);
    expect(model.structureDirty).toBe(true);
    expect(model.sheets[2]!.id).toMatch(UUID);
    expect(model.sheets[2]!.id).not.toBe(model.sheets[1]!.id);
  });

  it("shares the source's rows by reference when nothing in the copy names the source", () => {
    const model = book([gridSheet("Data", { A1: 1, A2: f("A1*2", 2), B3: f("Lookup!A1") }), gridSheet("Lookup")]);
    const source = model.sheets[0]!;
    duplicateSheet(model, 0);
    expect(model.sheets[1]!.rows).toBe(source.rows);
  });

  it("clones columns, merges, frozen panes, tab colour and view state: deep-equal, never the same objects", () => {
    const source = gridSheet(
      "Data",
      { A1: 1 },
      {
        columns: [{ min: 0, max: 1, widthPx: 120 }],
        merges: [{ top: 0, left: 0, bottom: 1, right: 1 }],
        frozen: { rows: 1, cols: 2 },
        tabColor: { theme: "accent1", lumMod: 0.75 },
        selection: { active: { row: 3, col: 2 }, ranges: [{ top: 3, left: 2, bottom: 4, right: 2 }], activeRangeIndex: 0 },
        topLeftCell: { row: 2, col: 1 },
        zoomPercent: 150,
        showGridLines: false,
        defaultColWidthPx: 80,
        defaultRowHeightPt: 18,
      },
    );
    const model = book([source]);
    duplicateSheet(model, 0);
    const copy = model.sheets[1]!;
    for (const field of ["columns", "merges", "frozen", "tabColor", "selection", "topLeftCell"] as const) {
      expect(copy[field], field).toEqual(source[field]);
      expect(copy[field], field).not.toBe(source[field]);
    }
    expect(copy.columns[0]).not.toBe(source.columns[0]);
    expect(copy.merges[0]).not.toBe(source.merges[0]);
    expect(copy).toMatchObject({ zoomPercent: 150, showGridLines: false, defaultColWidthPx: 80, defaultRowHeightPt: 18, visibility: "visible", kind: "grid" });
  });

  it("rewrites qualified references to the source inside the copy (which makes it dirty) and leaves the source untouched", () => {
    const model = freezeContents(
      book([
        gridSheet("Data", { A1: f("Data!B2*2"), A2: f("B2+1"), A3: f("Lookup!A1"), B2: 5, A9: 9 }),
        gridSheet("Lookup"),
        gridSheet("Summary", { A1: f("Data!B2") }),
      ]),
    );
    const [source, , summary] = model.sheets;
    duplicateSheet(model, 0);
    const copy = model.sheets[1]!;
    expect(copy.name).toBe("Data (2)");
    expect(formulaAt(copy, "A1")).toBe("'Data (2)'!B2*2");
    expect(formulaAt(copy, "A2")).toBe("B2+1");
    expect(formulaAt(copy, "A3")).toBe("Lookup!A1");
    expect(copy.dirty).toBe(true);
    expect(copy.rows).not.toBe(source!.rows);
    expect(copy.rows.at(-1)).toBe(source!.rows.at(-1));
    expect(model.sheets[0]).toBe(source);
    expect(formulaAt(source!, "A1")).toBe("Data!B2*2");
    expect(source!.dirty).toBeUndefined();
    expect(model.sheets[3]).toBe(summary);
  });

  it("lists in alsoRecalc the 3-D spans the copy lands in and the cells naming the copy", () => {
    const summary = gridSheet("Summary", { A1: f("SUM(Data:Lookup!A1)"), A2: f("'Data (2)'!A1"), A3: f("Data!A1"), A4: f("SUM(Lookup:Last!A1)") });
    const model = book([summary, gridSheet("Data"), gridSheet("Lookup"), gridSheet("Last")]);
    const { alsoRecalc } = duplicateSheet(model, 1);
    expect(alsoRecalc).toEqual(expect.arrayContaining([keyAt(summary, "A1"), keyAt(summary, "A2")]));
    expect(alsoRecalc).not.toContain(keyAt(summary, "A3"));
    expect(alsoRecalc).not.toContain(keyAt(summary, "A4"));
  });

  it("copies a sheet with no source as a dirty sheet with no copiedFrom and no source", () => {
    const model = createBlankWorkbook();
    duplicateSheet(model, 0);
    const copy = model.sheets[1]!;
    expect(copy.name).toBe("Sheet1 (2)");
    expect(copy.dirty).toBe(true);
    expect("copiedFrom" in copy || "partName" in copy || "sourceXml" in copy).toBe(false);
    expect(model.opaque).toEqual([]);
  });

  describe("copiedFrom (xlsx)", () => {
    it("is the origin's source name", async () => {
      const model = await openLinksXlsx();
      duplicateSheet(model, 1);
      expect(model.sheets[2]).toMatchObject({ name: "Data (2)", copiedFrom: "Data" });
      expect(model.sheets[1]!.copiedFrom).toBeUndefined();
    });

    it("stays the SOURCE name when the origin was renamed first", async () => {
      const model = await openLinksXlsx();
      renameSheet(model, 1, "Facts");
      duplicateSheet(model, 1);
      expect(model.sheets[2]).toMatchObject({ name: "Facts (2)", copiedFrom: "Data" });
    });

    it("is the origin's copiedFrom when duplicating a copy", async () => {
      const model = await openLinksXlsx();
      duplicateSheet(model, 1);
      renameSheet(model, 2, "Copy");
      duplicateSheet(model, 2);
      expect(model.sheets[3]).toMatchObject({ name: "Copy (2)", copiedFrom: "Data" });
      expect(model.sheets[3]!.partName).not.toBe(model.sheets[2]!.partName);
    });

    it("leaves the copy of a clean sheet clean, and copies a dirty origin's dirty flag", async () => {
      const model = await openLinksXlsx();
      duplicateSheet(model, 2);
      expect(model.sheets[3]!.name).toBe("Lookup (2)");
      expect(model.sheets[3]!.dirty).toBeUndefined();
      model.sheets = model.sheets.map((sheet, index) => (index === 1 ? { ...sheet, dirty: true } : sheet));
      duplicateSheet(model, 1);
      expect(model.sheets[2]).toMatchObject({ name: "Data (2)", dirty: true, copiedFrom: "Data" });
    });
  });

  it("copiedFrom is the origin's table:name (ods), also after a rename", async () => {
    const model = await openLinksOds();
    renameSheet(model, 1, "Facts");
    duplicateSheet(model, 1);
    expect(model.sheets[2]).toMatchObject({ name: "Facts (2)", copiedFrom: "Data" });
  });

  describe("refusals", () => {
    it('refuses a foreign sheet: "This sheet type can\'t be copied yet"', () => {
      const model = freezeContents(book([gridSheet("Data"), foreignSheet("Chart1", "xl/chartsheets/sheet1.xml")]));
      const sheets = model.sheets;
      expect(cannotDuplicateReason(model, 1)).toBe(FOREIGN_COPY);
      expect(thrownMessage(() => duplicateSheet(model, 1))).toBe(FOREIGN_COPY);
      expect(model.sheets).toBe(sheets);
      expect(model.structureDirty).toBeUndefined();
      expect(cannotDuplicateReason(model, 0)).toBeNull();
    });

    it("accepts the safe relationship types: drawing (chart, styles, colours), comments, VML, threaded comment, table, printer settings, external hyperlink", async () => {
      const model = await openLinksXlsx();
      expect(cannotDuplicateReason(model, 1)).toBeNull();
      editPart(model, "xl/worksheets/_rels/sheet2.xml.rels", (xml) =>
        addRelationship(
          xml,
          `<Relationship Id="rId8" Type="${REL}/hyperlink" Target="https://example.com/ledger" TargetMode="External"/>` +
            `<Relationship Id="rId9" Type="${REL}/printerSettings" Target="../printerSettings/printerSettings1.bin"/>`,
        ),
      );
      expect(cannotDuplicateReason(model, 1)).toBeNull();
    });

    for (const [label, type, target] of [
      ["an OLE object", `${REL}/oleObject`, "../embeddings/oleObject1.bin"],
      ["a form control", `${REL}/ctrlProp`, "../ctrlProps/ctrlProp1.xml"],
      ["a slicer", "http://schemas.microsoft.com/office/2007/relationships/slicer", "../slicers/slicer1.xml"],
      ["an internal hyperlink relationship", `${REL}/hyperlink`, "../other.xml"],
    ] as const) {
      it(`refuses an xlsx sheet with ${label}: "This sheet has content that can't be copied yet"`, async () => {
        const model = await openLinksXlsx();
        editPart(model, "xl/worksheets/_rels/sheet2.xml.rels", (xml) => addRelationship(xml, `<Relationship Id="rId9" Type="${type}" Target="${target}"/>`));
        freezeContents(model);
        const sheets = model.sheets;
        const opaque = model.opaque;
        expect(cannotDuplicateReason(model, 1)).toBe(UNSAFE_COPY);
        expect(thrownMessage(() => duplicateSheet(model, 1))).toBe(UNSAFE_COPY);
        expect(model.sheets).toBe(sheets);
        expect(model.opaque).toBe(opaque);
        expect(cannotDuplicateReason(model, 2)).toBeNull();
      });
    }

    it("refuses an xlsx sheet whose drawing carries something other than charts, chart styles/colours and images (SmartArt)", async () => {
      const model = await openLinksXlsx();
      editPart(model, "xl/drawings/_rels/drawing1.xml.rels", (xml) =>
        addRelationship(xml, `<Relationship Id="rId5" Type="${REL}/diagramData" Target="../diagrams/data1.xml"/>`),
      );
      setPart(model, "xl/diagrams/data1.xml", `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"/>`);
      expect(cannotDuplicateReason(model, 1)).toBe(UNSAFE_COPY);
    });
  });

  describe("xlsx: the copy's own retained source", () => {
    const COPY_PARTS = [
      "xl/worksheets/sheet6.xml",
      "xl/worksheets/_rels/sheet6.xml.rels",
      "xl/drawings/drawing2.xml",
      "xl/drawings/_rels/drawing2.xml.rels",
      "xl/charts/chart2.xml",
      "xl/charts/_rels/chart2.xml.rels",
      "xl/charts/style2.xml",
      "xl/charts/colors2.xml",
      "xl/tables/table2.xml",
      "xl/drawings/vmlDrawing2.vml",
      "xl/comments2.xml",
      "xl/threadedComments/threadedComment2.xml",
    ];

    it("gets the next free xl/worksheets/sheet<N>.xml holding the origin's bytes", async () => {
      const model = await openLinksXlsx();
      duplicateSheet(model, 1);
      const copy = model.sheets[2]!;
      expect(copy.partName).toBe("xl/worksheets/sheet6.xml");
      expect(model.sheets[1]!.partName).toBe("xl/worksheets/sheet2.xml");
      expect(partData(model, "xl/worksheets/sheet6.xml")).toBe(partData(model, "xl/worksheets/sheet2.xml"));
    });

    it("copies the closure into a NEW opaque array under next-free names; the old array is untouched and shared parts are not copied", async () => {
      const model = await openLinksXlsx();
      freezeContents(model);
      const before = model.opaque;
      const beforeNames = before.map((part) => part.name);
      duplicateSheet(model, 1);
      expect(model.opaque).not.toBe(before);
      expect(before.map((part) => part.name)).toEqual(beforeNames);
      for (const part of before) {
        expect(partData(model, part.name), part.name).toBe(part.data);
      }
      const added = model.opaque.map((part) => part.name).filter((name) => !beforeNames.includes(name));
      expect(added.toSorted()).toEqual(COPY_PARTS.toSorted());
      expect(model.opaque.filter((part) => part.name.startsWith("xl/persons/")).map((part) => part.name)).toEqual(["xl/persons/person.xml"]);
    });

    it("adds no Override, workbook relationship or sheetId at command time (those are the writer's)", async () => {
      const model = await openLinksXlsx();
      const workbookParts = ["[Content_Types].xml", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "docProps/app.xml"];
      const before = workbookParts.map((name) => partData(model, name));
      duplicateSheet(model, 1);
      expect(workbookParts.map((name) => partData(model, name))).toEqual(before);
    });

    it("retargets the copied .rels at the copies, keeping ids and types", async () => {
      const model = await openLinksXlsx();
      duplicateSheet(model, 1);
      const R = SPEC121_LINKS_XLSX.dataRelationships;
      expect(resolvedRels(model, "xl/worksheets/sheet6.xml")).toEqual([
        { id: "rId1", type: R.rId1.type, target: "xl/drawings/drawing2.xml" },
        { id: "rId2", type: R.rId2.type, target: "xl/drawings/vmlDrawing2.vml" },
        { id: "rId3", type: R.rId3.type, target: "xl/tables/table2.xml" },
        { id: "rId4", type: R.rId4.type, target: "xl/comments2.xml" },
        { id: "rId5", type: R.rId5.type, target: "xl/threadedComments/threadedComment2.xml" },
      ]);
      expect(resolvedRels(model, "xl/drawings/drawing2.xml").map((rel) => rel.target)).toEqual(["xl/charts/chart2.xml"]);
      expect(resolvedRels(model, "xl/charts/chart2.xml").map((rel) => rel.target).toSorted()).toEqual(["xl/charts/colors2.xml", "xl/charts/style2.xml"]);
      // The origin's own closure is untouched.
      expect(resolvedRels(model, "xl/worksheets/sheet2.xml").map((rel) => rel.target)).toContain("xl/drawings/drawing1.xml");
      // Parts that need no fix-up are byte copies (their sheet names stay source-time).
      for (const [copy, origin] of [
        ["xl/drawings/drawing2.xml", "xl/drawings/drawing1.xml"],
        ["xl/charts/chart2.xml", "xl/charts/chart1.xml"],
        ["xl/charts/style2.xml", "xl/charts/style1.xml"],
        ["xl/charts/colors2.xml", "xl/charts/colors1.xml"],
      ]) {
        expect(partData(model, copy!), copy).toBe(partData(model, origin!));
      }
    });

    it("gives a copied table id = max + 1 and name/displayName <name>_<n>", async () => {
      const model = await openLinksXlsx();
      const original = partText(model, "xl/tables/table1.xml");
      duplicateSheet(model, 1);
      const table = partText(model, "xl/tables/table2.xml")!;
      expect(attributeOf(table, "table", "id")).toBe("2");
      expect(attributeOf(table, "table", "name")).toBe("Table1_2");
      expect(attributeOf(table, "table", "displayName")).toBe("Table1_2");
      expect(attributeOf(table, "table", "ref")).toBe("A1:B5");
      expect(partText(model, "xl/tables/table1.xml")).toBe(original);
      // A second copy takes the next id and suffix.
      duplicateSheet(model, 1);
      const second = partText(model, "xl/tables/table3.xml")!;
      expect(attributeOf(second, "table", "id")).toBe("3");
      expect(attributeOf(second, "table", "displayName")).toBe("Table1_3");
    });

    it("takes the table id over every table part and the name unique among tables and defined names", async () => {
      const model = await openLinksXlsx();
      setPart(
        model,
        "xl/tables/table5.xml",
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="7" name="Other" displayName="Other" ref="A1:A2"/>`,
      );
      editPart(model, "xl/workbook.xml", (xml) => xml.replace("</definedNames>", `<definedName name="Table1_2">Lookup!$A$1</definedName></definedNames>`));
      duplicateSheet(model, 1);
      const table = partText(model, "xl/tables/table6.xml")!;
      expect(attributeOf(table, "table", "id")).toBe("8");
      expect(attributeOf(table, "table", "displayName")).toBe("Table1_3");
      expect(attributeOf(table, "table", "name")).toBe("Table1_3");
    });

    it("renames structured references to the copied table inside the copy only", async () => {
      const model = await openLinksXlsx();
      const total: Cell = { col: 2, value: 1000, formula: "SUM(Table1[Amount])" };
      model.sheets = model.sheets.map((sheet, index) =>
        index === 1 ? { ...sheet, rows: [...sheet.rows, { index: 6, cells: [total] }] } : index === 0 ? { ...sheet, rows: [...sheet.rows, { index: 6, cells: [{ ...total }] }] } : sheet,
      );
      editPart(model, "xl/worksheets/sheet2.xml", (xml) => xml.replace("</sheetData>", `<row r="7"><c r="C7"><f>SUM(Table1[Amount])</f><v>1000</v></c></row></sheetData>`));
      duplicateSheet(model, 1);
      expect(formulaAt(model.sheets[2]!, "C7")).toBe("SUM(Table1_2[Amount])");
      expect(formulaAt(model.sheets[1]!, "C7")).toBe("SUM(Table1[Amount])");
      expect(formulaAt(model.sheets[0]!, "C7")).toBe("SUM(Table1[Amount])");
      expect(partText(model, "xl/worksheets/sheet6.xml")).toContain("<f>SUM(Table1_2[Amount])</f>");
      expect(partText(model, "xl/worksheets/sheet2.xml")).toContain("<f>SUM(Table1[Amount])</f>");
    });

    it("moves a copied VML part to o:idmap data = max + 1 and rebases its _x0000_s ids by 1024 per block", async () => {
      const model = await openLinksXlsx();
      const original = partText(model, "xl/drawings/vmlDrawing1.vml");
      duplicateSheet(model, 1);
      const vml = partText(model, "xl/drawings/vmlDrawing2.vml")!;
      expect(/<o:idmap\b[^>]*\bdata="([^"]*)"/.exec(vml)?.[1]).toBe("2");
      expect(vml).toContain(`id="_x0000_s2049"`);
      expect(vml).toContain(`o:spid="_x0000_s2049"`);
      expect(vml).not.toContain("_x0000_s1025");
      expect(vml).toContain(`type="#_x0000_t202"`);
      expect(partText(model, "xl/drawings/vmlDrawing1.vml")).toBe(original);
    });

    it("takes the VML block over every VML part of the package", async () => {
      const model = await openLinksXlsx();
      setPart(
        model,
        "xl/drawings/vmlDrawing7.vml",
        `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="3"/></o:shapelayout></xml>`,
      );
      duplicateSheet(model, 1);
      const vml = partText(model, "xl/drawings/vmlDrawing8.vml")!;
      expect(/<o:idmap\b[^>]*\bdata="([^"]*)"/.exec(vml)?.[1]).toBe("4");
      expect(vml).toContain(`o:spid="_x0000_s4097"`);
    });

    it("re-mints the threaded comment's id and the matching legacy xr:uid, v4-shaped, keeping the person", async () => {
      const model = await openLinksXlsx();
      duplicateSheet(model, 1);
      const threaded = partText(model, "xl/threadedComments/threadedComment2.xml")!;
      const id = attributeOf(threaded, "threadedComment", "id")!;
      expect(id).not.toBe(SPEC121_LINKS_XLSX.threadedComment.id);
      expect(id).toMatch(/^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/);
      expect(attributeOf(threaded, "threadedComment", "personId")).toBe(SPEC121_LINKS_XLSX.threadedComment.personId);
      expect(attributeOf(partText(model, "xl/comments2.xml")!, "comment", "xr:uid")).toBe(id);
      expect(attributeOf(partText(model, "xl/threadedComments/threadedComment1.xml")!, "threadedComment", "id")).toBe(SPEC121_LINKS_XLSX.threadedComment.id);
      expect(attributeOf(partText(model, "xl/comments1.xml")!, "comment", "xr:uid")).toBe(SPEC121_LINKS_XLSX.comment.uid);
    });

    it("mints the same ids from the same inputs (copy sheet id + original id), and other ids for another copy", async () => {
      const copyParts = async (copyId: string): Promise<string[]> => {
        const model = await openLinksXlsx();
        vi.spyOn(crypto, "randomUUID").mockReturnValue(copyId as ReturnType<typeof crypto.randomUUID>);
        duplicateSheet(model, 1);
        vi.restoreAllMocks();
        expect(model.sheets[2]!.id).toBe(copyId);
        return COPY_PARTS.map((name) => partText(model, name) ?? "");
      };
      const a = await copyParts("0c0ffee0-1111-4222-8333-444455556666");
      const b = await copyParts("0c0ffee0-1111-4222-8333-444455556666");
      const c = await copyParts("7e57ab1e-9999-4888-a777-666655554444");
      expect(b).toEqual(a);
      const threadedId = (texts: string[]): string | undefined => attributeOf(texts[COPY_PARTS.indexOf("xl/threadedComments/threadedComment2.xml")]!, "threadedComment", "id");
      expect(threadedId(c)).not.toBe(threadedId(a));
    });

    it("shares a part also reachable from elsewhere (an image another sheet's drawing shows) and copies one only the origin reaches", async () => {
      const IMAGE = "xl/media/image1.png";
      const withImage = async (shared: boolean): Promise<SpreadsheetModel> => {
        const model = await openLinksXlsx();
        setPart(model, IMAGE, "\u0089PNG stand-in");
        editPart(model, "xl/drawings/_rels/drawing1.xml.rels", (xml) => addRelationship(xml, `<Relationship Id="rId7" Type="${REL}/image" Target="../media/image1.png"/>`));
        if (shared) {
          setPart(model, "xl/drawings/drawing3.xml", `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`);
          setPart(
            model,
            "xl/drawings/_rels/drawing3.xml.rels",
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/image" Target="../media/image1.png"/></Relationships>`,
          );
          setPart(
            model,
            "xl/worksheets/_rels/sheet3.xml.rels",
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/drawing" Target="../drawings/drawing3.xml"/></Relationships>`,
          );
        }
        duplicateSheet(model, 1);
        return model;
      };
      const shared = await withImage(true);
      expect(shared.opaque.filter((part) => part.name.startsWith("xl/media/")).map((part) => part.name)).toEqual([IMAGE]);
      const copyDrawing = resolvedRels(shared, "xl/worksheets/sheet6.xml").find((rel) => rel.id === "rId1")!.target;
      expect(copyDrawing).toBe("xl/drawings/drawing4.xml");
      expect(resolvedRels(shared, copyDrawing).map((rel) => rel.target).toSorted()).toEqual(["xl/charts/chart2.xml", IMAGE]);

      const owned = await withImage(false);
      expect(owned.opaque.filter((part) => part.name.startsWith("xl/media/")).map((part) => part.name).toSorted()).toEqual([IMAGE, "xl/media/image2.png"]);
      expect(partData(owned, "xl/media/image2.png")).toBe(partData(owned, IMAGE));
      expect(resolvedRels(owned, "xl/drawings/drawing2.xml").map((rel) => rel.target).toSorted()).toEqual(["xl/charts/chart2.xml", "xl/media/image2.png"]);
    });
  });

  describe("ods: the copy's own retained source", () => {
    it("is the origin's table text with every xml:id/draw:id suffixed -c<k> consistently and its embedded object retargeted; table:name stays source-time", async () => {
      const model = await openLinksOds();
      const origin = model.sheets[1]!;
      duplicateSheet(model, 1);
      const copy = model.sheets[2]!;
      const suffix = /\sxml:id="id1(-c\d+)"/.exec(copy.sourceXml!)?.[1];
      expect(suffix).toMatch(/^-c\d+$/);
      expect(copy.sourceXml).toBe(
        origin
          .sourceXml!.replaceAll(' xml:id="id1"', ` xml:id="id1${suffix}"`)
          .replaceAll(' draw:id="id1"', ` draw:id="id1${suffix}"`)
          .replaceAll('xlink:href="./Object 1"', 'xlink:href="./Object 2"')
          .replaceAll('xlink:href="./ObjectReplacements/Object 1"', 'xlink:href="./ObjectReplacements/Object 2"'),
      );
      expect(model.sheets[1]).toBe(origin);
      // A second copy gets ids no other text declares.
      duplicateSheet(model, 1);
      const second = model.sheets[2]!.sourceXml!;
      const secondId = /\sxml:id="([^"]*)"/.exec(second)?.[1];
      expect(secondId).toMatch(/^id1-c\d+$/);
      expect(secondId).not.toBe(`id1${suffix}`);
      expect(/\sdraw:id="([^"]*)"/.exec(second)?.[1]).toBe(secondId);
    });

    it("suffixes the references to a renamed id along with its declaration", () => {
      const sourceXml =
        `<table:table table:name="Shapes" table:style-name="ta1"><table:shapes>` +
        `<draw:rect xml:id="s1" draw:id="s1"/><draw:rect xml:id="s2" draw:id="s2"/>` +
        `<draw:connector draw:start-shape="s1" draw:end-shape="s2" draw:name="s3"/></table:shapes>` +
        `<table:table-column/><table:table-row><table:table-cell/></table:table-row></table:table>`;
      const model = book([gridSheet("Shapes", {}, { sourceXml })], 0, "ods");
      model.odsContentEnvelope = { prefix: "<office:document-content><office:body><office:spreadsheet>", suffix: "</office:spreadsheet></office:body></office:document-content>" };
      duplicateSheet(model, 0);
      const copy = model.sheets[1]!.sourceXml!;
      const k = /xml:id="s1(-c\d+)"/.exec(copy)?.[1];
      expect(k).toMatch(/^-c\d+$/);
      expect(copy).toBe(
        sourceXml
          .replace(`xml:id="s1" draw:id="s1"`, `xml:id="s1${k}" draw:id="s1${k}"`)
          .replace(`xml:id="s2" draw:id="s2"`, `xml:id="s2${k}" draw:id="s2${k}"`)
          .replace(`draw:start-shape="s1" draw:end-shape="s2"`, `draw:start-shape="s1${k}" draw:end-shape="s2${k}"`),
      );
    });

    it("copies the embedded object to the next free Object N with the manifest spliced in a NEW manifest part", async () => {
      const model = await openLinksOds();
      freezeContents(model);
      const before = model.opaque;
      const manifestBefore = model.opaque.find((part) => part.name === "META-INF/manifest.xml")!;
      const manifestText = partText(model, "META-INF/manifest.xml");
      duplicateSheet(model, 1);
      expect(model.opaque).not.toBe(before);
      for (const [copy, origin] of [
        ["Object 2/content.xml", "Object 1/content.xml"],
        ["Object 2/styles.xml", "Object 1/styles.xml"],
        ["Object 2/meta.xml", "Object 1/meta.xml"],
        ["ObjectReplacements/Object 2", "ObjectReplacements/Object 1"],
      ]) {
        expect(partData(model, copy!), copy).toBeDefined();
        expect(partData(model, copy!), copy).toBe(partData(model, origin!));
      }
      const manifest = model.opaque.find((part) => part.name === "META-INF/manifest.xml")!;
      expect(manifest).not.toBe(manifestBefore);
      expect(manifestBefore.data).toBe(bytesToBase64(encoder.encode(manifestText!)));
      const entries = manifestEntries(model);
      for (const path of ["Object 1/content.xml", "Object 1/styles.xml", "Object 1/meta.xml", "Object 1/", "ObjectReplacements/Object 1"]) {
        const copyPath = path.replace("Object 1", "Object 2");
        expect(entries.get(copyPath), copyPath).toBeDefined();
        expect(entries.get(copyPath), copyPath).toBe(entries.get(path));
      }
      // Everything else in the package is as it was.
      for (const part of before) {
        if (part.name !== "META-INF/manifest.xml") {
          expect(partData(model, part.name), part.name).toBe(part.data);
        }
      }
    });
  });
});

// =========================================================================================

describe("deleteSheet", () => {
  /** One, very hidden Secret, Two, hidden Stash, Three. */
  const hiddenBook = (active: number): SpreadsheetModel =>
    book(
      [
        gridSheet("One", { A1: 1 }),
        gridSheet("Secret", { A1: 2 }, { visibility: "veryHidden" }),
        gridSheet("Two", { A1: 3 }),
        gridSheet("Stash", { A1: 4 }, { visibility: "hidden" }),
        gridSheet("Three", { A1: 5 }),
      ],
      active,
    );

  it("turns every reference to the deleted sheet into #REF! (spec121-links: delete Lookup)", async () => {
    const model = await openLinksXlsx();
    deleteSheet(model, 2);
    expect(names(model)).toEqual(["Summary", "Data", "Archive", "Vault"]);
    const summary = model.sheets[0]!;
    expect(formulaAt(summary, "B2")).toBe("VLOOKUP(A2,#REF!,2,FALSE)");
    expect(formulaAt(summary, "E2")).toBe("#REF!");
    expect(formulaAt(summary, "D2")).toBe("SUM(Data!A1)");
    expect(formulaAt(summary, "C2")).toBe("Data!B2*2");
    expect(summary.dirty).toBe(true);
    expect(model.structureDirty).toBe(true);
  });

  it("moves a deleted 3-D endpoint inward", () => {
    const build = (): SpreadsheetModel => book([gridSheet("Jan"), gridSheet("Feb"), gridSheet("Mar"), gridSheet("Calc", { A1: f("SUM(Jan:Mar!B2)"), A2: f("SUM(Mar:Jan!B2)") })]);
    const lastGone = build();
    deleteSheet(lastGone, 2);
    expect(formulaAt(lastGone.sheets[2]!, "A1")).toBe("SUM(Jan:Feb!B2)");
    const firstGone = build();
    deleteSheet(firstGone, 0);
    expect(formulaAt(firstGone.sheets[2]!, "A1")).toBe("SUM(Feb:Mar!B2)");
  });

  it("leaves a 3-D span's text alone when the deleted sheet was inside it, and lists the cell in alsoRecalc", () => {
    const calc = gridSheet("Calc", { A1: f("SUM(Jan:Mar!B2)"), A2: f("Jan!B2"), A3: f("SUM(Mar:Calc!B2)") });
    const model = freezeContents(book([gridSheet("Jan"), gridSheet("Feb"), gridSheet("Mar"), calc]));
    const { alsoRecalc } = deleteSheet(model, 1);
    expect(model.sheets[2]).toBe(calc);
    expect(formulaAt(model.sheets[2]!, "A1")).toBe("SUM(Jan:Mar!B2)");
    expect(alsoRecalc).toContain(keyAt(calc, "A1"));
    expect(alsoRecalc).not.toContain(keyAt(calc, "A2"));
    expect(alsoRecalc).not.toContain(keyAt(calc, "A3"));
  });

  it("activates the next visible sheet to the RIGHT of a deleted active sheet", () => {
    const model = hiddenBook(0);
    const result = deleteSheet(model, 0);
    expect(names(model)).toEqual(["Secret", "Two", "Stash", "Three"]);
    expect(result.activeIndex).toBe(1);
    expect(model.activeSheetIndex).toBe(1);
  });

  it("…else the one to its LEFT", () => {
    const model = hiddenBook(4);
    expect(deleteSheet(model, 4).activeIndex).toBe(2);
    expect(model.sheets[model.activeSheetIndex]!.name).toBe("Two");
    const links = book([gridSheet("Summary"), gridSheet("Data"), gridSheet("Lookup"), gridSheet("Archive", {}, { visibility: "hidden" })], 2);
    expect(deleteSheet(links, 2).activeIndex).toBe(1);
  });

  it("keeps the active sheet when deleting another", () => {
    const before = hiddenBook(4);
    expect(deleteSheet(before, 0).activeIndex).toBe(3);
    expect(before.sheets[before.activeSheetIndex]!.name).toBe("Three");
    const after = hiddenBook(2);
    expect(deleteSheet(after, 4).activeIndex).toBe(2);
    expect(after.sheets[after.activeSheetIndex]!.name).toBe("Two");
  });

  it('refuses the last visible sheet: "A workbook must keep at least one visible sheet"', () => {
    const model = freezeContents(book([gridSheet("Only"), gridSheet("Stash", {}, { visibility: "hidden" }), gridSheet("Secret", {}, { visibility: "veryHidden" })]));
    const sheets = model.sheets;
    expect(isLastVisibleSheet(model, 0)).toBe(true);
    expect(isLastVisibleSheet(model, 1)).toBe(false);
    expect(isLastVisibleSheet(model, 2)).toBe(false);
    expect(thrownMessage(() => deleteSheet(model, 0))).toBe(LAST_VISIBLE);
    expect(model.sheets).toBe(sheets);
    expect(model.structureDirty).toBeUndefined();
    deleteSheet(model, 1);
    expect(names(model)).toEqual(["Only", "Secret"]);
    expect(isLastVisibleSheet(book([gridSheet("A"), gridSheet("B")]), 0)).toBe(false);
  });

  it("deleting an original leaves opaque alone (xlsx and ods)", async () => {
    const xlsx = await openLinksXlsx();
    const xlsxOpaque = xlsx.opaque;
    deleteSheet(xlsx, 1);
    deleteSheet(xlsx, 1);
    expect(xlsx.opaque).toBe(xlsxOpaque);
    const ods = await openLinksOds();
    const odsOpaque = ods.opaque;
    deleteSheet(ods, 1);
    expect(ods.opaque).toBe(odsOpaque);
  });

  it("deleting an xlsx copy removes from opaque exactly the parts its duplication added", async () => {
    const model = await openLinksXlsx();
    const original = model.opaque.map((part) => ({ name: part.name, data: part.data }));
    duplicateSheet(model, 1);
    expect(model.opaque.length).toBeGreaterThan(original.length);
    deleteSheet(model, 2);
    expect(model.opaque.map((part) => ({ name: part.name, data: part.data }))).toEqual(original);
    expect(names(model)).toEqual(SPEC121_LINKS_XLSX.sheetNames);
  });

  it("deleting an ods copy drops its embedded objects and their manifest entries", async () => {
    const model = await openLinksOds();
    duplicateSheet(model, 1);
    deleteSheet(model, 2);
    expect(model.opaque.filter((part) => part.name.includes("Object 2"))).toEqual([]);
    const entries = manifestEntries(model);
    expect([...entries.keys()].filter((path) => path.includes("Object 2"))).toEqual([]);
    for (const path of ["Object 1/content.xml", "Object 1/", "ObjectReplacements/Object 1"]) {
      expect(entries.has(path), path).toBe(true);
    }
    expect(partData(model, "Object 1/content.xml")).toBeDefined();
  });

  it("keeps every untouched sheet object; rewritten sheets are new and dirty", async () => {
    const model = freezeContents(await openLinksXlsx());
    const [summary, data, , archive, vault] = model.sheets;
    deleteSheet(model, 2);
    expect(model.sheets[0]).not.toBe(summary);
    expect(model.sheets[0]!.rows[0]).toBe(summary!.rows[0]);
    expect(model.sheets[1]).toBe(data);
    expect(model.sheets[2]).toBe(archive);
    expect(model.sheets[3]).toBe(vault);
  });
});

// =========================================================================================

describe("sheetHasContent", () => {
  it("is true for stored rows and for a foreign sheet, false for an empty new sheet", () => {
    const model = book([gridSheet("Cells", { B2: 1 }), foreignSheet("Chart1", "xl/chartsheets/sheet1.xml"), gridSheet("Empty")]);
    expect(sheetHasContent(model, 0)).toBe(true);
    expect(sheetHasContent(model, 1)).toBe(true);
    expect(sheetHasContent(model, 2)).toBe(false);
    addSheet(model, 2);
    expect(sheetHasContent(model, 3)).toBe(false);
    expect(sheetHasContent(createBlankWorkbook(), 0)).toBe(false);
  });

  it("is true for an xlsx sheet whose part has a .rels, even with no rows", async () => {
    const model = await openLinksXlsx();
    model.sheets = model.sheets.map((sheet) => ({ ...sheet, rows: [] }));
    expect(sheetHasContent(model, 1)).toBe(true);
    expect(sheetHasContent(model, 2)).toBe(false);
  });

  it("is true for an ods table holding <table:shapes or <office:annotation, even with no rows", () => {
    const model = book(
      [
        gridSheet("Shapes", {}, { sourceXml: odsTable(`<table:shapes><draw:frame/></table:shapes>`) }),
        gridSheet("Notes", {}, { sourceXml: odsTable(`<table:table-row><table:table-cell><office:annotation><text:p>Hi</text:p></office:annotation></table:table-cell></table:table-row>`) }),
        gridSheet("Plain", {}, { sourceXml: odsTable(`<table:table-row><table:table-cell/></table:table-row>`) }),
      ],
      0,
      "ods",
    );
    expect(sheetHasContent(model, 0)).toBe(true);
    expect(sheetHasContent(model, 1)).toBe(true);
    expect(sheetHasContent(model, 2)).toBe(false);
  });
});

// =========================================================================================

describe("moveSheet", () => {
  const abcd = (): SpreadsheetModel => book(["A", "B", "C", "D"].map((name) => gridSheet(name)));

  it("puts the sheet at index `to` of the resulting order and keeps it active", () => {
    const right = abcd();
    const result = moveSheet(right, 0, 2);
    expect(names(right)).toEqual(["B", "C", "A", "D"]);
    expect(result.activeIndex).toBe(2);
    expect(right.activeSheetIndex).toBe(2);
    const left = abcd();
    left.activeSheetIndex = 3;
    expect(moveSheet(left, 3, 0).activeIndex).toBe(0);
    expect(names(left)).toEqual(["D", "A", "B", "C"]);
    expect(left.activeSheetIndex).toBe(0);
  });

  it("replaces model.sheets with the same sheet objects reordered and sets structureDirty", () => {
    const model = freezeContents(abcd());
    const before = model.sheets;
    const [a, b, c, d] = before;
    moveSheet(model, 1, 3);
    expect(model.sheets).not.toBe(before);
    expect(before.map((sheet) => sheet.name)).toEqual(["A", "B", "C", "D"]);
    expect(model.sheets).toEqual([a, c, d, b]);
    model.sheets.forEach((sheet, index) => expect(sheet).toBe([a, c, d, b][index]));
    expect(model.structureDirty).toBe(true);
  });

  it("changes nothing for an unchanged move", () => {
    const model = freezeContents(abcd());
    model.activeSheetIndex = 1;
    const before = model.sheets;
    expect(moveSheet(model, 1, 1)).toEqual({ activeIndex: 1, alsoRecalc: [] });
    expect(model.sheets).toBe(before);
    expect(model.structureDirty).toBeUndefined();
    expect(model.activeSheetIndex).toBe(1);
  });

  it("lists in alsoRecalc the 3-D cells whose span gains or loses a sheet", () => {
    const build = (): { model: SpreadsheetModel; calc: Sheet } => {
      const calc = gridSheet("Calc", { A1: f("SUM(Jan:Mar!B2)"), A2: f("Jan!B2"), A3: f("SUM(Jan:Apr!B2)") });
      return { model: book([gridSheet("Jan"), gridSheet("Feb"), gridSheet("Mar"), gridSheet("Apr"), calc]), calc };
    };
    const gains = build();
    const gained = moveSheet(gains.model, 3, 1).alsoRecalc;
    expect(gained).toEqual(expect.arrayContaining([keyAt(gains.calc, "A1"), keyAt(gains.calc, "A3")]));
    expect(gained).not.toContain(keyAt(gains.calc, "A2"));

    const loses = build();
    const lost = moveSheet(loses.model, 1, 4).alsoRecalc;
    expect(lost).toEqual(expect.arrayContaining([keyAt(loses.calc, "A1"), keyAt(loses.calc, "A3")]));
    expect(lost).not.toContain(keyAt(loses.calc, "A2"));

    const outside = build();
    expect(moveSheet(outside.model, 4, 0).alsoRecalc).toEqual([]);
  });
});

// =========================================================================================

describe("setSheetHidden", () => {
  it("hides a sheet in a new sheet object and keeps the active sheet when it stays visible", () => {
    const model = freezeContents(book([gridSheet("A"), gridSheet("B"), gridSheet("C")]));
    const b = model.sheets[1]!;
    const result = setSheetHidden(model, [1], true);
    expect(model.sheets[1]).not.toBe(b);
    expect(model.sheets[1]).toEqual({ ...b, visibility: "hidden" });
    expect(b.visibility).toBe("visible");
    expect(result.activeIndex).toBe(0);
    expect(model.activeSheetIndex).toBe(0);
    expect(model.structureDirty).toBe(true);
  });

  it("moves the active sheet off a hidden one as delete does: next visible to the right, else left", () => {
    const middle = book([gridSheet("A"), gridSheet("B"), gridSheet("C")], 1);
    expect(setSheetHidden(middle, [1], true).activeIndex).toBe(2);
    expect(middle.activeSheetIndex).toBe(2);
    const last = book([gridSheet("A"), gridSheet("B"), gridSheet("C")], 2);
    expect(setSheetHidden(last, [2], true).activeIndex).toBe(1);
    const skipping = book(
      [gridSheet("One"), gridSheet("Secret", {}, { visibility: "veryHidden" }), gridSheet("Two"), gridSheet("Stash", {}, { visibility: "hidden" }), gridSheet("Three")],
      0,
    );
    expect(setSheetHidden(skipping, [0], true).activeIndex).toBe(2);
    const several = book(["A", "B", "C", "D"].map((name) => gridSheet(name)));
    expect(setSheetHidden(several, [0, 1], true).activeIndex).toBe(2);
  });

  it('refuses to hide the last visible sheet: "A workbook must keep at least one visible sheet"', () => {
    const one = freezeContents(book([gridSheet("A"), gridSheet("B", {}, { visibility: "hidden" })]));
    const sheets = one.sheets;
    expect(thrownMessage(() => setSheetHidden(one, [0], true))).toBe(LAST_VISIBLE);
    expect(one.sheets).toBe(sheets);
    expect(one.structureDirty).toBeUndefined();
    const both = book([gridSheet("A"), gridSheet("B")]);
    expect(thrownMessage(() => setSheetHidden(both, [0, 1], true))).toBe(LAST_VISIBLE);
  });

  it("unhides the chosen sheets, the last of them active, and leaves an unchosen very hidden sheet as it is", () => {
    const model = freezeContents(
      book([
        gridSheet("Summary"),
        gridSheet("Data"),
        gridSheet("Archive", {}, { visibility: "hidden" }),
        gridSheet("Vault", {}, { visibility: "veryHidden" }),
        gridSheet("Old", {}, { visibility: "hidden" }),
      ]),
    );
    const vault = model.sheets[3];
    const result = setSheetHidden(model, [2, 4], false);
    expect(model.sheets.map((sheet) => sheet.visibility)).toEqual(["visible", "visible", "visible", "veryHidden", "visible"]);
    expect(model.sheets[3]).toBe(vault);
    expect(result.activeIndex).toBe(4);
    expect(model.activeSheetIndex).toBe(4);
    expect(model.structureDirty).toBe(true);
  });
});

// =========================================================================================

describe("setTabColor", () => {
  it("sets a theme colour as a reference, replacing the sheet object, and sets structureDirty", () => {
    const model = freezeContents(book([gridSheet("A"), gridSheet("B")]));
    const [a, b] = model.sheets;
    const color: ModelColor = { theme: "accent2", lumMod: 0.6, lumOff: 0.4 };
    setTabColor(model, 1, color);
    expect(model.sheets[1]!.tabColor).toEqual({ theme: "accent2", lumMod: 0.6, lumOff: 0.4 });
    expect(model.sheets[1]).not.toBe(b);
    expect(b!.tabColor).toBeUndefined();
    expect(model.sheets[0]).toBe(a);
    expect(model.structureDirty).toBe(true);
    expect(model.sheets[1]!.dirty).toBeUndefined();
  });

  it("sets a hex colour and clears a colour", () => {
    const model = book([gridSheet("A", {}, { tabColor: { theme: "accent1" } })]);
    setTabColor(model, 0, "#ff0000");
    expect(model.sheets[0]!.tabColor).toBe("#ff0000");
    setTabColor(model, 0, undefined);
    expect("tabColor" in model.sheets[0]!).toBe(false);
    expect(model.structureDirty).toBe(true);
  });

  it("changes nothing for an unchanged colour, equal by value: same objects, no structureDirty", () => {
    const model = freezeContents(book([gridSheet("A", {}, { tabColor: { theme: "accent1", lumMod: 0.75 } }), gridSheet("B", {}, { tabColor: "#00ff00" }), gridSheet("C")]));
    const sheets = model.sheets;
    setTabColor(model, 0, { theme: "accent1", lumMod: 0.75 });
    setTabColor(model, 1, "#00ff00");
    setTabColor(model, 2, undefined);
    expect(model.sheets).toBe(sheets);
    expect(model.structureDirty).toBeUndefined();
  });

  it("treats spec121-links' read theme colour, re-applied by value, as unchanged", async () => {
    const model = await openLinksXlsx();
    const sheets = model.sheets;
    setTabColor(model, 1, structuredClone(model.sheets[1]!.tabColor));
    expect(model.sheets).toBe(sheets);
    expect(model.structureDirty).toBeUndefined();
  });
});

// =========================================================================================

describe("every command", () => {
  type Command = { label: string; run: (model: SpreadsheetModel) => unknown };
  const commands: Command[] = [
    { label: "addSheet", run: (model) => addSheet(model, 0) },
    { label: "renameSheet", run: (model) => renameSheet(model, 1, "Facts") },
    { label: "renameSheet (case only)", run: (model) => renameSheet(model, 2, "LOOKUP") },
    { label: "duplicateSheet", run: (model) => duplicateSheet(model, 1) },
    { label: "deleteSheet", run: (model) => deleteSheet(model, 2) },
    { label: "deleteSheet (active)", run: (model) => deleteSheet(model, 0) },
    { label: "moveSheet", run: (model) => moveSheet(model, 2, 0) },
    { label: "setSheetHidden (hide)", run: (model) => setSheetHidden(model, [0], true) },
    { label: "setSheetHidden (unhide)", run: (model) => setSheetHidden(model, [3], false) },
    { label: "setTabColor (set)", run: (model) => setTabColor(model, 0, { theme: "accent1" }) },
    { label: "setTabColor (clear)", run: (model) => setTabColor(model, 1, undefined) },
  ];

  for (const format of ["xlsx", "ods"] as const) {
    for (const command of commands) {
      it(`${command.label} (${format}) sets structureDirty and mutates no pre-existing Sheet, Row or Cell`, async () => {
        const model = format === "xlsx" ? await openLinksXlsx() : await openLinksOds();
        const snapshot = structuredClone(model.sheets);
        const sheets = freezeContents(model).sheets;
        expect(() => command.run(model)).not.toThrow();
        expect(model.structureDirty).toBe(true);
        expect(sheets).toEqual(snapshot);
        expect(model.sheets[model.activeSheetIndex]!.visibility).toBe("visible");
      });
    }
  }

  it("except setTabColor with an unchanged colour", async () => {
    const model = freezeContents(await openLinksXlsx());
    setTabColor(model, 0, undefined);
    expect(model.structureDirty).toBeUndefined();
  });
});

// =========================================================================================

describe("stripRetention", () => {
  it("drops the retained package and every edit flag: no opaque, envelope, structureDirty, packageVariant; no partName/sourceXml/copiedFrom/dirty", async () => {
    const xlsx = await openLinksXlsx();
    duplicateSheet(xlsx, 1);
    renameSheet(xlsx, 2, "Facts");
    addSheet(xlsx, 0);
    const fromXlsx = stripRetention({ ...xlsx, format: "ods" });
    expect(fromXlsx.format).toBe("ods");
    expect(fromXlsx.opaque).toEqual([]);
    expect("structureDirty" in fromXlsx).toBe(false);
    for (const sheet of fromXlsx.sheets) {
      expect(["partName", "sourceXml", "copiedFrom", "dirty"].filter((field) => field in sheet), sheet.name).toEqual([]);
    }
    expect(names(fromXlsx)).toEqual(names(xlsx));

    const ods = await openLinksOds();
    duplicateSheet(ods, 1);
    const fromOds = stripRetention({ ...ods, format: "xlsx" });
    expect(fromOds.opaque).toEqual([]);
    expect("odsContentEnvelope" in fromOds).toBe(false);
    for (const sheet of fromOds.sheets) {
      expect(["partName", "sourceXml", "copiedFrom", "dirty"].filter((field) => field in sheet), sheet.name).toEqual([]);
    }

    const xlsm = await readSpreadsheetBytes(await macrosXlsmBytes(), "xlsx", "Macros");
    expect(xlsm.packageVariant).toBe("macroEnabled");
    expect("packageVariant" in stripRetention({ ...xlsm, format: "ods" })).toBe(false);
  });

  it("removes foreign sheets and keeps the active sheet when it survives", () => {
    const model = book([gridSheet("Data", { A1: 1 }), foreignSheet("Chart1", "xl/chartsheets/sheet1.xml"), gridSheet("Lookup", { A1: 2 })], 2);
    const stripped = stripRetention({ ...model, format: "ods" });
    expect(names(stripped)).toEqual(["Data", "Lookup"]);
    expect(stripped.activeSheetIndex).toBe(1);
  });

  it("re-points the active index at a visible sheet when the active sheet was foreign", () => {
    const model = book(
      [gridSheet("Stash", {}, { visibility: "hidden" }), foreignSheet("Chart1", "xl/chartsheets/sheet1.xml"), gridSheet("Data", { A1: 1 })],
      1,
    );
    const stripped = stripRetention({ ...model, format: "ods" });
    expect(names(stripped)).toEqual(["Stash", "Data"]);
    expect(stripped.sheets[stripped.activeSheetIndex]!.visibility).toBe("visible");
    expect(stripped.sheets[stripped.activeSheetIndex]!.name).toBe("Data");
  });

  it("still yields a visible grid sheet with no retention flags when every sheet was foreign", () => {
    const model = book([foreignSheet("Chart1", "xl/chartsheets/sheet1.xml"), foreignSheet("Chart2", "xl/chartsheets/sheet2.xml")]);
    const stripped = stripRetention({ ...model, format: "ods" });
    expect(stripped.sheets.length).toBeGreaterThan(0);
    for (const sheet of stripped.sheets) {
      expect(sheet.kind).toBe("grid");
      expect(["partName", "sourceXml", "copiedFrom", "dirty"].filter((field) => field in sheet), sheet.name).toEqual([]);
    }
    expect(stripped.sheets[stripped.activeSheetIndex]!.visibility).toBe("visible");
  });

  it("strips sourceXf and sourceStyleName from every style and keeps the rest", async () => {
    for (const model of [await openLinksXlsx(), await openLinksOds()]) {
      expect(model.styles.some((style) => style.sourceXf !== undefined || style.sourceStyleName !== undefined)).toBe(true);
      const stripped = stripRetention({ ...model, format: model.format === "xlsx" ? "ods" : "xlsx" });
      expect(stripped.styles).toEqual(
        model.styles.map((style) => {
          const { sourceXf: _xf, sourceStyleName: _name, ...rest } = style;
          return rest;
        }),
      );
      expect(stripped.styles.every((style) => !("sourceXf" in style) && !("sourceStyleName" in style))).toBe(true);
    }
  });

  it("crosses every formula cell through formulaForTarget: unparsed and data-table formulas drop to their cached value, a dynamic array to .ods loses cm/dynamicArray", async () => {
    const model = await readSpreadsheetBytes(await formulasXlsxBytes(), "xlsx", "Formulas");
    const calc = sheetNamed(model, "Calc");
    expect(cellAt(calc, "S2")!.formulaUnparsed).toBe(true);
    expect(cellAt(calc, "L3")!.dataTableXml).toBeDefined();
    expect(cellAt(calc, "H2")!.cm).toBeDefined();
    const stripped = stripRetention({ ...model, format: "ods" });
    const crossed = sheetNamed(stripped, "Calc");

    const s2 = cellAt(crossed, "S2")!;
    expect(s2.value).toBe(3);
    expect(["formula", "formulaUnparsed"].filter((field) => field in s2)).toEqual([]);
    const l3 = cellAt(crossed, "L3")!;
    expect(l3.value).toBe(10);
    expect(["formula", "dataTableXml"].filter((field) => field in l3)).toEqual([]);
    const h2 = cellAt(crossed, "H2")!;
    expect(h2.formula).toBe("SORT(A2:A6)");
    expect(h2.arrayRange).toEqual({ top: 1, left: 7, bottom: 5, right: 7 });
    expect(["cm", "dynamicArray"].filter((field) => field in h2)).toEqual([]);

    for (const row of calc.rows) {
      for (const cell of row.cells) {
        expect(getCell(crossed, row.index, cell.col)).toEqual(formulaForTarget(cell, "ods").cell);
      }
    }
    // The legacy array formula crosses as it is.
    expect(cellAt(crossed, "E2")).toBe(cellAt(calc, "E2"));
  });

  it("shares rows by reference wherever no cell changes", async () => {
    const model = await readSpreadsheetBytes(await formulasXlsxBytes(), "xlsx", "Formulas");
    const stripped = stripRetention({ ...model, format: "ods" });
    model.sheets.forEach((sheet, index) => {
      const crossed = stripped.sheets[index]!;
      if (sheet.name === "Calc") {
        expect(crossed.rows).not.toBe(sheet.rows);
        crossed.rows.forEach((row, at) => {
          const changed = row.cells.some((cell, col) => cell !== sheet.rows[at]!.cells[col]);
          expect(row === sheet.rows[at], `row ${row.index}`).toBe(!changed);
        });
        expect(crossed.rows[0]).toBe(sheet.rows[0]);
      } else {
        expect(crossed.rows, sheet.name).toBe(sheet.rows);
      }
    });
    const links = await openLinksXlsx();
    const linksStripped = stripRetention({ ...links, format: "ods" });
    links.sheets.forEach((sheet, index) => expect(linksStripped.sheets[index]!.rows, sheet.name).toBe(sheet.rows));
  });

  it("is pure: the open model deep-equals a structuredClone taken before", async () => {
    for (const model of [await openLinksXlsx(), await openLinksOds(), await readSpreadsheetBytes(await formulasXlsxBytes(), "xlsx", "Formulas")]) {
      if (model.sheets.length > 1) {
        duplicateSheet(model, 1);
      }
      const before = structuredClone(model);
      freezeContents(model);
      stripRetention({ ...model, format: model.format === "xlsx" ? "ods" : "xlsx" });
      expect(model).toEqual(before);
    }
  });
});

// =========================================================================================

describe("unconvertibleContent", () => {
  const LABELS = [
    "chart sheets",
    "macros",
    "images and shapes",
    "charts",
    "comments",
    "conditional formatting",
    "data validation",
    "named ranges",
    "tables",
    "filters",
    "hyperlinks",
    "pivot tables",
    "page setup",
    "untranslatable formulas",
    "custom date origin",
  ];

  it("is the spec's table, in its order, with markers on the sides it names", () => {
    expect(UNCONVERTIBLE_CONTENT.map((row) => row.label)).toEqual(LABELS);
    expect(unconvertibleRow("chart sheets").ods).toBeUndefined();
    expect(unconvertibleRow("custom date origin").xlsx).toBeUndefined();
    expect(unconvertibleRow("custom date origin").ods).toBeDefined();
    for (const label of LABELS.filter((name) => name !== "chart sheets" && name !== "custom date origin")) {
      expect(unconvertibleRow(label).xlsx, label).toBeDefined();
      expect(unconvertibleRow(label).ods, label).toBeDefined();
    }
  });

  it("is empty for the model's own format", async () => {
    expect(unconvertibleContent(await openLinksXlsx(), "xlsx")).toEqual([]);
    expect(unconvertibleContent(await openLinksOds(), "ods")).toEqual([]);
    expect(unconvertibleContent(await readSpreadsheetBytes(await macrosXlsmBytes(), "xlsx", "Macros"), "xlsx")).toEqual([]);
    expect(unconvertibleContent(createBlankWorkbook(), "xlsx")).toEqual([]);
  });

  it("is empty for a blank workbook crossing to .ods", () => {
    expect(unconvertibleContent(createBlankWorkbook(), "ods")).toEqual([]);
  });

  it("names macros for an .xlsm", async () => {
    const labels = unconvertibleContent(await readSpreadsheetBytes(await macrosXlsmBytes(), "xlsx", "Macros"), "ods");
    expect(labels).toContain("macros");
    expect(labels).toEqual(inTableOrder(labels));
  });

  it("names untranslatable formulas for spec119's unparsed and data-table formulas", async () => {
    const model = await readSpreadsheetBytes(await formulasXlsxBytes(), "xlsx", "Formulas");
    expect(unconvertibleContent(model, "ods")).toEqual(["untranslatable formulas"]);
  });

  it("lists spec121-links.xlsx's structures in table order (a chart's drawing is the charts row, as the AC's toast \"charts, comments\" reads)", async () => {
    const labels = unconvertibleContent(await openLinksXlsx(), "ods");
    expect(labels).toEqual(["charts", "comments", "conditional formatting", "data validation", "named ranges", "tables", "hyperlinks"]);
  });

  it("lists spec121-links.ods's structures in table order", async () => {
    const labels = unconvertibleContent(await openLinksOds(), "xlsx");
    expect(labels).toEqual(["charts", "comments", "conditional formatting", "data validation", "named ranges", "tables", "hyperlinks"]);
  });

  it("names images and shapes for a worksheet drawing holding a picture, and for an ods frame that is not an object's", async () => {
    const xlsx = await openLinksXlsx();
    editPart(xlsx, "xl/drawings/drawing1.xml", (xml) =>
      xml.replace(
        "</xdr:wsDr>",
        `<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
          `<xdr:ext cx="100" cy="100"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="3" name="Picture 2"/><xdr:cNvPicPr/></xdr:nvPicPr>` +
          `<xdr:blipFill/><xdr:spPr/></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>`,
      ),
    );
    const xlsxLabels = unconvertibleContent(xlsx, "ods");
    expect(xlsxLabels.slice(0, 2)).toEqual(["images and shapes", "charts"]);

    const ods = await openLinksOds();
    ods.sheets = ods.sheets.map((sheet, index) =>
      index === 1
        ? { ...sheet, sourceXml: sheet.sourceXml!.replace("</table:shapes>", `<draw:frame draw:name="Image 1"><draw:image xlink:href="Pictures/1.png"/></draw:frame></table:shapes>`) }
        : sheet,
    );
    const odsLabels = unconvertibleContent(ods, "xlsx");
    expect(odsLabels.slice(0, 2)).toEqual(["images and shapes", "charts"]);
  });

  it("ignores _xlnm names for named ranges", async () => {
    const model = await openLinksXlsx();
    editPart(model, "xl/workbook.xml", (xml) => xml.replace(/<definedName name="(?!_xlnm\.)[^"]*"[^>]*>[^<]*<\/definedName>/g, ""));
    expect(partText(model, "xl/workbook.xml")).toContain("_xlnm.Print_Area");
    expect(unconvertibleContent(model, "ods")).not.toContain("named ranges");
  });

  it("names chart sheets for a foreign chartsheet", () => {
    const model = book([gridSheet("Data"), foreignSheet("Chart1", "xl/chartsheets/sheet1.xml")]);
    expect(unconvertibleContent(model, "ods")).toEqual(["chart sheets"]);
  });

  it("names the custom date origin for an .ods with odsNullDate crossing to .xlsx, last", async () => {
    const model = await readSpreadsheetBytes(await miniOdsBytes(), "ods", "Mini");
    expect(model.odsNullDate).toBeUndefined();
    expect(unconvertibleContent(model, "xlsx")).not.toContain("custom date origin");
    const labels = unconvertibleContent({ ...model, odsNullDate: "1899-12-31" }, "xlsx");
    expect(labels.at(-1)).toBe("custom date origin");
    expect(labels).toEqual(inTableOrder(labels));
  });
});

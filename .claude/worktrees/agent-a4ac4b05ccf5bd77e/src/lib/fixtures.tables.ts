// The SPEC-076 table fixture models (AC-75), shared by the generator that
// writes the committed packages (scripts/build-spec076-fixtures.mjs) and the
// unit spec that guards them (fixtures.tables.spec.ts). Bytes and nothing
// else — no filesystem access, no side effects on import — so importing it
// can never rewrite a fixture the e2e suite is reading.
//
// Four documents, each built in both .docx and .odt:
//
//   1. a four-page table with a repeating header  (AC-1..AC-11)
//   2. a banded table                              (AC-22..AC-29, AC-66..AC-69)
//   3. a table sorted by two columns               (AC-38..AC-48)
//   4. a tab-separated list                        (AC-52..AC-60)
import { writeDocx } from "./docx/writer";
import { writeOdt } from "./odf/writer";
import { createDocument } from "./documentModel";
import { builtInTableStyle } from "./tableStyles";
import { sortTableRows } from "./tableSort";
import type { DocumentModel, NamedStyle, Paragraph, Table, TableCell, TableRow } from "./documentModel";

/** Pinned so a regenerated fixture's dcterms:modified / dc:date does not
 * become whichever day someone re-ran the generator — the same rule
 * fixtures.borders.ts and fixtures.fields.ts follow. */
export const FIXED_MODIFIED = new Date(Date.UTC(2026, 8, 3, 9, 0, 0));

/** How many body rows the tall fixture carries. Sized so the table runs over
 * four pages at the product's default A4 geometry — the pagination engine
 * measures the real rendered rows, so this is a generous floor rather than an
 * exact page count. */
export const TALL_TABLE_BODY_ROWS = 120;

/** The gallery style the banded fixture applies (AC-31's "Accent"). */
export const BANDED_STYLE_ID = "TableStyleAccent";

function paragraph(text: string): Paragraph {
  return { type: "paragraph", runs: text.length > 0 ? [{ text, formatting: {} }] : [] };
}

function cell(text: string): TableCell {
  return { blocks: [paragraph(text)] };
}

function row(texts: string[], isHeader = false): TableRow {
  const built: TableRow = { cells: texts.map(cell) };
  if (isHeader) {
    built.isHeader = true;
  }
  return built;
}

function widths(columns: number): number[] {
  // A4 at the product's 2.54 cm margins, split evenly and snapped the same
  // way lib/tables.ts's equalColumnWidthsCm does.
  const total = 21 - 2.54 * 2;
  return Array.from({ length: columns }, () => Math.round((total / columns) * 10000) / 10000);
}

/** 1 — a table whose single header row repeats onto every continuation. */
export function tallHeaderTable(): Table {
  return {
    type: "table",
    columnWidthsCm: widths(3),
    rows: [
      row(["Item", "Region", "Units"], true),
      ...Array.from({ length: TALL_TABLE_BODY_ROWS }, (_unused, index) =>
        row([`Item ${index + 1}`, index % 2 === 0 ? "North" : "South", String((index * 7) % 97)]),
      ),
    ],
  };
}

/** 2 — a table carrying the Accent look: a header row plus banded rows. */
export function bandedTable(): Table {
  return {
    type: "table",
    columnWidthsCm: widths(3),
    styleId: BANDED_STYLE_ID,
    look: { firstRow: true, bandedRows: true },
    rows: [
      row(["Quarter", "Revenue", "Change"], true),
      row(["Q1", "120", "+4%"]),
      row(["Q2", "138", "+15%"]),
      row(["Q3", "131", "-5%"]),
      row(["Q4", "164", "+25%"]),
    ],
  };
}

/** The keys fixture 3 is sorted by: region ascending as text, then units
 * descending as a number (AC-38, AC-46). */
export const SORTED_TABLE_KEYS = [
  { columnIndex: 0, direction: "asc" as const, type: "text" as const },
  { columnIndex: 1, direction: "desc" as const, type: "number" as const },
];

/** 3 — a table already sorted by two columns, so reopening it proves the
 * order the sort produced survives a round trip. */
export function sortedTable(): Table {
  const unsorted: Table = {
    type: "table",
    columnWidthsCm: widths(3),
    rows: [
      row(["Region", "Units", "Owner"], true),
      row(["South", "12", "Ada"]),
      row(["North", "40", "Grace"]),
      row(["South", "7", "Alan"]),
      row(["North", "40", "Edsger"]),
      row(["East", "3", "Barbara"]),
    ],
  };
  return sortTableRows(unsorted, SORTED_TABLE_KEYS, true);
}

/** The tab-separated paragraphs fixture 4 holds, in document order. Every
 * paragraph has the same field count and no field contains a tab, so this is
 * exactly the "simple case" the round-trip test uses (AC-65). */
export const TAB_SEPARATED_LINES = [
  "Name\tRole\tCity",
  "Ada\tAnalyst\tLondon",
  "Grace\tAdmiral\tArlington",
  "Alan\tLogician\tWilmslow",
];

function model(title: string, blocks: (Paragraph | Table)[], styles?: NamedStyle[]): DocumentModel {
  const built = createDocument(title);
  built.title = title;
  built.sections = [{ blocks }];
  if (styles) {
    built.styles = styles;
  }
  return built;
}

/** The four fixture models, keyed by the name each package takes. */
export function tableFixtureModels(): { key: string; title: string; model: DocumentModel }[] {
  const accent = builtInTableStyle(BANDED_STYLE_ID);
  return [
    {
      key: "tables-repeating-header",
      title: "Repeating header",
      model: model("Repeating header", [
        paragraph("A table whose header row repeats on every page."),
        tallHeaderTable(),
        paragraph(""),
      ]),
    },
    {
      key: "tables-banded",
      title: "Banded table",
      model: model(
        "Banded table",
        [paragraph("A banded table."), bandedTable(), paragraph("")],
        accent ? [accent] : undefined,
      ),
    },
    {
      key: "tables-sorted",
      title: "Sorted table",
      model: model("Sorted table", [paragraph("Sorted by region, then units."), sortedTable(), paragraph("")]),
    },
    {
      key: "tables-tab-separated",
      title: "Tab-separated list",
      model: model("Tab-separated list", [
        paragraph("A tab-separated list, ready to convert."),
        ...TAB_SEPARATED_LINES.map(paragraph),
        paragraph(""),
      ]),
    },
  ];
}

/** Every fixture as package bytes, in both formats (AC-75). */
export async function buildTableFixtures(): Promise<{ name: string; bytes: Uint8Array }[]> {
  const built: { name: string; bytes: Uint8Array }[] = [];
  for (const fixture of tableFixtureModels()) {
    const docx = await writeDocx(fixture.model, { modified: FIXED_MODIFIED });
    const odt = await writeOdt(fixture.model, { modified: FIXED_MODIFIED });
    built.push({ name: `${fixture.key}.docx`, bytes: docx.bytes });
    built.push({ name: `${fixture.key}.odt`, bytes: odt.bytes });
  }
  return built;
}

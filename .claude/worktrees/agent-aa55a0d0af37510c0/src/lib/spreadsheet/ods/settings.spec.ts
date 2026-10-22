// SPEC-117 "Reading ODS sheet layout" — settings.xml: each per-table item
// (split modes and positions → frozen, the scroll position, ShowGrid,
// ShowZeroValues, ZoomType/ZoomValue, the cursor → selection), split mode 1
// ignored, the view-level fallback, ActiveTable, and the LibreOffice-shaped
// fixture's settings as the reader applies them.
import { describe, expect, it } from "vitest";
import { XmlParseError } from "../../core/xml";
import { parseA1 } from "../address";
import { buildSpec117StylesOdsBytes, SPEC117_STYLES_ODS } from "../fixtures/build";
import { parseOds } from "./reader";
import { readOdsSettings, type OdsTableView } from "./settings";

const item = (name: string, value: string | number | boolean, type = "int"): string =>
  `<config:config-item config:name="${name}" config:type="${type}">${value}</config:config-item>`;

/** settings.xml with one view: its own items and per-table entries. */
function settingsXml(tables: Record<string, string>, viewItems = ""): string {
  const entries = Object.entries(tables)
    .map(([name, items]) => `<config:config-item-map-entry config:name="${name}">${items}</config:config-item-map-entry>`)
    .join("");
  return (
    `<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
    `xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings>` +
    `<config:config-item-set config:name="ooo:view-settings">${item("VisibleAreaTop", 0)}` +
    `<config:config-item-map-indexed config:name="Views"><config:config-item-map-entry>${item("ViewId", "view1", "string")}` +
    `<config:config-item-map-named config:name="Tables">${entries}</config:config-item-map-named>${viewItems}` +
    `</config:config-item-map-entry></config:config-item-map-indexed></config:config-item-set></office:settings></office:document-settings>`
  );
}

function view(items: string, viewItems = ""): OdsTableView {
  return readOdsSettings(settingsXml({ T: items }, viewItems)).tables.get("T")!;
}

describe("per-table items", () => {
  it("split mode 2 on both axes → frozen rows and columns; the scrolled pane's first cell → topLeftCell", () => {
    expect(
      view(
        item("HorizontalSplitMode", 2, "short") +
          item("VerticalSplitMode", 2, "short") +
          item("HorizontalSplitPosition", 1) +
          item("VerticalSplitPosition", 2) +
          item("PositionLeft", 0) +
          item("PositionRight", 3) +
          item("PositionTop", 0) +
          item("PositionBottom", 40),
      ),
    ).toEqual({ frozen: { rows: 2, cols: 1 }, topLeftCell: { row: 40, col: 3 } });
  });

  it("one axis frozen: the other reads its own pane's position", () => {
    expect(
      view(item("VerticalSplitMode", 2, "short") + item("VerticalSplitPosition", 3) + item("PositionLeft", 5) + item("PositionRight", 9) + item("PositionBottom", 3)),
    ).toEqual({ frozen: { rows: 3, cols: 0 }, topLeftCell: { row: 3, col: 5 } });
  });

  it("the scrolled pane never starts inside the frozen band", () => {
    expect(
      view(
        item("HorizontalSplitMode", 2, "short") +
          item("VerticalSplitMode", 2, "short") +
          item("HorizontalSplitPosition", 2) +
          item("VerticalSplitPosition", 4) +
          item("PositionRight", 0) +
          item("PositionBottom", 0),
      ),
    ).toEqual({ frozen: { rows: 4, cols: 2 }, topLeftCell: { row: 4, col: 2 } });
  });

  it("split mode 1 (a split, not a freeze) is ignored: no frozen, the left/top pane's position", () => {
    expect(
      view(
        item("HorizontalSplitMode", 1, "short") +
          item("HorizontalSplitPosition", 3216) +
          item("VerticalSplitMode", 1, "short") +
          item("VerticalSplitPosition", 1200) +
          item("PositionLeft", 0) +
          item("PositionRight", 4) +
          item("PositionTop", 0) +
          item("PositionBottom", 7),
      ),
    ).toEqual({});
  });

  it("an unscrolled, unfrozen table has no topLeftCell; positions are clamped to the grid", () => {
    expect(view(item("PositionLeft", 0) + item("PositionTop", 0))).toEqual({});
    expect(view(item("PositionLeft", 99999) + item("PositionTop", 9999999))).toEqual({ topLeftCell: { row: 1_048_575, col: 16_383 } });
  });

  it("ShowGrid false → showGridLines: false; ShowZeroValues false → showZeros: false", () => {
    expect(view(item("ShowGrid", false, "boolean"))).toEqual({ showGridLines: false });
    expect(view(item("ShowGrid", true, "boolean"))).toEqual({});
    expect(view(item("ShowZeroValues", false, "boolean"))).toEqual({ showZeros: false });
  });

  it("ZoomType 0 + ZoomValue → zoomPercent, clamped to 10–400, 100 → absent; other zoom types are left alone", () => {
    expect(view(item("ZoomType", 0, "short") + item("ZoomValue", 85))).toEqual({ zoomPercent: 85 });
    expect(view(item("ZoomValue", 85))).toEqual({ zoomPercent: 85 });
    expect(view(item("ZoomType", 0, "short") + item("ZoomValue", 100))).toEqual({});
    expect(view(item("ZoomValue", 5))).toEqual({ zoomPercent: 10 });
    expect(view(item("ZoomValue", 900))).toEqual({ zoomPercent: 400 });
    expect(view(item("ZoomType", 1, "short") + item("ZoomValue", 85))).toEqual({});
  });

  it("CursorPositionX/Y → a one-cell selection", () => {
    expect(view(item("CursorPositionX", 1) + item("CursorPositionY", 2))).toEqual({
      selection: { active: { row: 2, col: 1 }, ranges: [{ top: 2, left: 1, bottom: 2, right: 1 }], activeRangeIndex: 0 },
    });
    expect(view(item("CursorPositionX", 0) + item("CursorPositionY", 0)).selection?.active).toEqual({ row: 0, col: 0 });
    expect(view("").selection).toBeUndefined();
  });
});

describe("the view-level fallback", () => {
  it("a table's entry lacking an item takes the view's item of the same name", () => {
    const viewItems = item("ZoomType", 0, "short") + item("ZoomValue", 150) + item("ShowGrid", false, "boolean");
    expect(view(item("ZoomValue", 85), viewItems)).toEqual({ zoomPercent: 85, showGridLines: false });
    expect(view(item("ShowGrid", true, "boolean"), viewItems)).toEqual({ zoomPercent: 150 });
  });

  it("a table with no entry gets the view-level items alone", () => {
    const settings = readOdsSettings(settingsXml({ T: item("ZoomValue", 85) }, item("ZoomValue", 120) + item("ShowZeroValues", false, "boolean")));
    expect(settings.defaultView).toEqual({ zoomPercent: 120, showZeros: false });
  });
});

describe("the active table and odd parts", () => {
  it("reads ActiveTable; none → undefined", () => {
    expect(readOdsSettings(settingsXml({}, item("ActiveTable", "Plain", "string"))).activeTable).toBe("Plain");
    expect(readOdsSettings(settingsXml({})).activeTable).toBeUndefined();
  });

  it("an empty settings part has no tables; a malformed one or a DTD is refused", () => {
    const empty = readOdsSettings(`<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>`);
    expect(empty.tables.size).toBe(0);
    expect(empty.defaultView).toEqual({});
    expect(() => readOdsSettings("<office:document-settings")).toThrow(XmlParseError);
    expect(() => readOdsSettings(`<!DOCTYPE x [<!ENTITY a "b">]><x/>`)).toThrow(XmlParseError);
  });
});

describe("spec117-styles.ods (LibreOffice-shaped)", () => {
  it("applies each table's settings: frozen, top-left, zoom and cursor on Styled; no grid on Plain; Split not frozen", async () => {
    const parsed = await parseOds(await buildSpec117StylesOdsBytes());
    const [styled, plain, split] = parsed.sheets;
    const facts = SPEC117_STYLES_ODS.styled;
    expect(parsed.activeSheetIndex).toBe(SPEC117_STYLES_ODS.tableNames.indexOf(SPEC117_STYLES_ODS.activeTable));
    expect(styled!.frozen).toEqual(facts.frozen);
    const topLeft = parseA1(facts.topLeftCell)!;
    expect(styled!.topLeftCell).toEqual({ row: topLeft.row, col: topLeft.col });
    expect(styled!.zoomPercent).toBe(facts.zoomPercent);
    const cursor = parseA1(facts.cursor)!;
    expect(styled!.selection).toEqual({
      active: { row: cursor.row, col: cursor.col },
      ranges: [{ top: cursor.row, left: cursor.col, bottom: cursor.row, right: cursor.col }],
      activeRangeIndex: 0,
    });
    expect(styled!.showGridLines).toBeUndefined();
    expect(plain!.showGridLines).toBe(SPEC117_STYLES_ODS.plain.showGridLines);
    expect(plain!.zoomPercent).toBeUndefined();
    expect(split!.frozen).toBeUndefined();
    expect(split!.topLeftCell).toBeUndefined();
    for (const sheet of parsed.sheets) {
      expect(JSON.parse(JSON.stringify(sheet))).toEqual(sheet);
    }
  });
});

/** settings.xml with only `ooo:configuration-settings` items. */
const configuration = (items: string): string =>
  `<office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ` +
  `xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><office:settings>` +
  `<config:config-item-set config:name="ooo:configuration-settings">${items}</config:config-item-set></office:settings></office:document-settings>`;

describe("AutoCalculate (SPEC-120)", () => {
  it("reads ooo:configuration-settings' AutoCalculate, absent when the file does not say", () => {
    expect(readOdsSettings(configuration(item("AutoCalculate", true, "boolean"))).autoCalculate).toBe(true);
    expect(readOdsSettings(configuration(item("AutoCalculate", false, "boolean"))).autoCalculate).toBe(false);
    expect(readOdsSettings(configuration(item("SaveVersionOnClose", false, "boolean"))).autoCalculate).toBeUndefined();
  });
});

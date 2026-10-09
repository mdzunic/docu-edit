// settings.xml (SPEC-117 "Reading ODS sheet layout"): the view settings an
// .ods workbook was saved with — the active table, and per table the frozen
// panes, the scroll position, gridlines, zero display, zoom and the cell
// cursor. The part is small, so it is parsed whole (core/xml.ts's parseXml,
// behind parseOdsPart's DTD refusal).
//
// Where LibreOffice keeps them: `ooo:view-settings` → the `Views` indexed
// map → its first entry (the view). The view's `Tables` named map holds one
// entry per table name; the view's own config items of the same names are
// the fallback for every item a table's entry lacks (and all a table with
// no entry gets — `defaultView`).
// - `HorizontalSplitMode` = 2 → `frozen.cols = HorizontalSplitPosition`;
//   `VerticalSplitMode` = 2 → `frozen.rows = VerticalSplitPosition`. Mode 1
//   is a split, not a freeze (its positions are lengths, not counts): the
//   normal view is shown and the split stays in the retained part.
// - `topLeftCell`: per axis, `PositionRight` (columns frozen) or
//   `PositionLeft`, and `PositionBottom` (rows frozen) or `PositionTop` —
//   the first column and row of the pane that scrolls (never inside the
//   frozen band); A1 → absent.
// - `ShowGrid` false → `showGridLines: false`; `ShowZeroValues` false →
//   `showZeros: false` (LibreOffice keeps it on the view, so it applies to
//   every table through the fallback).
// - `ZoomType` 0 (a percentage; absent reads as 0) + `ZoomValue` →
//   `zoomPercent`, clamped to 10–400, 100 → absent. The other zoom types
//   (whole page, page width) are not a percentage and are left alone.
// - `CursorPositionX`/`Y` → `selection`: that one cell, active.
// Counts and positions are clamped into the 1,048,576 × 16,384 grid.
import type { CellRef } from "../address";
import { MAX_COLS, MAX_ROWS } from "../model";
import type { SheetSelection } from "../selection";
import { parseOdsPart } from "./styles";

const CONFIG_NAMESPACE = "urn:oasis:names:tc:opendocument:xmlns:config:1.0";

/** The view fields settings.xml gives one table — exactly the Sheet fields
 * of the same names, each present only when the file sets it. */
export interface OdsTableView {
  frozen?: { rows: number; cols: number };
  topLeftCell?: CellRef;
  showGridLines?: false;
  showZeros?: false;
  zoomPercent?: number;
  selection?: SheetSelection;
}

export interface OdsSettings {
  /** The view's `ActiveTable` (a table name). */
  activeTable?: string;
  /** Each table's view fields, by table name. */
  tables: Map<string, OdsTableView>;
  /** The view-level items alone — what a table with no entry of its own
   * shows. */
  defaultView: OdsTableView;
}

type Items = Map<string, string>;

function configName(el: Element): string | null {
  return el.getAttributeNS(CONFIG_NAMESPACE, "name");
}

function configChildren(parent: Element, local: string): Element[] {
  return Array.from(parent.children).filter((child) => child.namespaceURI === CONFIG_NAMESPACE && child.localName === local);
}

/** An entry's own `config:config-item` children, by name. */
function itemsOf(entry: Element | undefined): Items {
  const items: Items = new Map();
  if (entry) {
    for (const item of configChildren(entry, "config-item")) {
      const name = configName(item);
      if (name !== null && !items.has(name)) {
        items.set(name, (item.textContent ?? "").trim());
      }
    }
  }
  return items;
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(Math.max(value, min), max);
}

/** One table's view fields from its items, the view's as the fallback. */
export function odsTableView(own: Items, view: Items): OdsTableView {
  const item = (name: string): string | undefined => own.get(name) ?? view.get(name);
  const int = (name: string): number | undefined => {
    const text = item(name);
    const value = text === undefined || text === "" ? Number.NaN : Number(text);
    return Number.isInteger(value) ? value : undefined;
  };
  const result: OdsTableView = {};

  const colsFrozen = int("HorizontalSplitMode") === 2;
  const rowsFrozen = int("VerticalSplitMode") === 2;
  const cols = colsFrozen ? clamp(int("HorizontalSplitPosition") ?? 0, 0, MAX_COLS) : 0;
  const rows = rowsFrozen ? clamp(int("VerticalSplitPosition") ?? 0, 0, MAX_ROWS) : 0;
  if (rows > 0 || cols > 0) {
    result.frozen = { rows, cols };
  }

  // The scrolled pane never starts inside the frozen band.
  const left = clamp(Math.max(int(colsFrozen ? "PositionRight" : "PositionLeft") ?? 0, cols), 0, MAX_COLS - 1);
  const top = clamp(Math.max(int(rowsFrozen ? "PositionBottom" : "PositionTop") ?? 0, rows), 0, MAX_ROWS - 1);
  if (left > 0 || top > 0) {
    result.topLeftCell = { row: top, col: left };
  }

  if (item("ShowGrid") === "false") {
    result.showGridLines = false;
  }
  if (item("ShowZeroValues") === "false") {
    result.showZeros = false;
  }

  const zoom = int("ZoomValue");
  if ((int("ZoomType") ?? 0) === 0 && zoom !== undefined) {
    const percent = clamp(zoom, 10, 400);
    if (percent !== 100) {
      result.zoomPercent = percent;
    }
  }

  const cursorX = int("CursorPositionX");
  const cursorY = int("CursorPositionY");
  if (cursorX !== undefined || cursorY !== undefined) {
    const active = { row: clamp(cursorY ?? 0, 0, MAX_ROWS - 1), col: clamp(cursorX ?? 0, 0, MAX_COLS - 1) };
    result.selection = {
      active,
      ranges: [{ top: active.row, left: active.col, bottom: active.row, right: active.col }],
      activeRangeIndex: 0,
    };
  }
  return result;
}

/** Reads settings.xml. Throws XmlParseError for a malformed part (or one
 * declaring a DTD). */
export function readOdsSettings(xml: string): OdsSettings {
  const doc = parseOdsPart(xml);
  const settings: OdsSettings = { tables: new Map(), defaultView: {} };

  // The first ActiveTable item in document order (where SPEC-116 found it).
  for (const item of Array.from(doc.getElementsByTagNameNS(CONFIG_NAMESPACE, "config-item"))) {
    if (configName(item) === "ActiveTable") {
      settings.activeTable = item.textContent ?? "";
      break;
    }
  }

  // The view: the map entry holding the Tables map, else the first entry of
  // the Views map.
  const tablesMap = Array.from(doc.getElementsByTagNameNS(CONFIG_NAMESPACE, "config-item-map-named")).find(
    (map) => configName(map) === "Tables",
  );
  const parent = tablesMap?.parentElement;
  const viewsMap = Array.from(doc.getElementsByTagNameNS(CONFIG_NAMESPACE, "config-item-map-indexed")).find(
    (map) => configName(map) === "Views",
  );
  const viewEntry =
    parent?.namespaceURI === CONFIG_NAMESPACE && parent.localName === "config-item-map-entry"
      ? parent
      : viewsMap
        ? configChildren(viewsMap, "config-item-map-entry")[0]
        : undefined;
  const view = itemsOf(viewEntry);
  settings.defaultView = odsTableView(new Map(), view);
  for (const entry of tablesMap ? configChildren(tablesMap, "config-item-map-entry") : []) {
    const name = configName(entry);
    if (name !== null && !settings.tables.has(name)) {
      settings.tables.set(name, odsTableView(itemsOf(entry), view));
    }
  }
  return settings;
}

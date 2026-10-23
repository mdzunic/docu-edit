// The style book an .ods table is scanned with in SPEC-137's phase 3
// (ods/phases.ts), where DOMParser does not exist. ods/styles.ts's
// OdsStyleBook holds DOM elements (the data styles) and flattens on demand,
// so it stays on the main thread: the shell phase projects it into plain data
// (OdsStyleBook.scanProjection — every style any part defines, and what a
// name no part defines reads as), and this book answers every question the
// table scan, the table reader and the conditional-format reader ask exactly
// as OdsStyleBook would.
//
// The cell style table grows as OdsStyleBook's does: each table-cell style
// name gets ONE entry, in first-reference order, and an implied format a
// derived one (ImpliedFormatStyles). That growing half lives in the scan
// context (OdsScanBookState) — the context is posted to the worker once per
// read and every job of the read scans with the same object, in sheet order —
// so a later sheet reuses an earlier sheet's entries exactly as one book
// would, and each job's appended entries are the ones it added.
import type { Dxf } from "../conditionalFormats";
import { ImpliedFormatStyles } from "../impliedFormats";
import type { CellFont, CellStyle } from "../model";
import type { OdsScanStyles } from "./reader";
import { encodeOdsStyleName, type OdsCellStyleEntry, type OdsRowHeight, type OdsStyleMap, type OdsStyleProjection, type OdsTableStyle } from "./styles";

/** The growing half of the scan's book, carried in the scan context. */
export interface OdsScanBookState {
  /** The cell style table so far (a copy of the shell's, index 0 "Default"). */
  styles: CellStyle[];
  /** ImpliedFormatStyles' derived-style record. */
  implied: Record<string, number>;
  /** Each table-cell style name met (but "Default", always 0) → its id, in
   * first-reference order. */
  cellIds: [string, number][];
  /** The ids of the entries that name a data style (impliedStyleId keeps them). */
  formatted: number[];
}

/** The projection's lists as lookups — built once per projection object. */
interface Lookups {
  cellStyles: Map<string, OdsCellStyleEntry>;
  textStyles: Map<string, { modelled: boolean; font?: CellFont }>;
  tableStyles: Map<string, OdsTableStyle>;
  columnWidths: Map<string, number | null>;
  rowHeights: Map<string, OdsRowHeight>;
  mapped: Map<string, OdsStyleMap[]>;
  mappedNames: string[];
  conditional: Map<string, Dxf>;
  displayNames: Map<string, string>;
}

const lookupsByProjection = new WeakMap<OdsStyleProjection, Lookups>();

function lookupsOf(projection: OdsStyleProjection): Lookups {
  let lookups = lookupsByProjection.get(projection);
  if (!lookups) {
    const displayNames = new Map<string, string>();
    for (const [display, name] of projection.displayNames) {
      if (!displayNames.has(display)) {
        displayNames.set(display, name); // the first in document order, as the book's search finds it
      }
    }
    lookups = {
      cellStyles: new Map(projection.cellStyles),
      textStyles: new Map(projection.textStyles),
      tableStyles: new Map(projection.tableStyles),
      columnWidths: new Map(projection.columnWidths),
      rowHeights: new Map(projection.rowHeights),
      mapped: new Map(projection.mapped),
      mappedNames: projection.mapped.map(([name]) => name),
      conditional: new Map(projection.conditional),
      displayNames,
    };
    lookupsByProjection.set(projection, lookups);
  }
  return lookups;
}

/** OdsStyleBook's answers from its projection (see the module header). */
export class OdsScanBook implements OdsScanStyles {
  readonly styles: CellStyle[];
  private readonly projection: OdsStyleProjection;
  private readonly state: OdsScanBookState;
  private readonly lookups: Lookups;
  private readonly cellIds: Map<string, number>;
  private readonly formatted: Set<number>;
  private readonly implied: ImpliedFormatStyles;

  constructor(projection: OdsStyleProjection, state: OdsScanBookState) {
    this.projection = projection;
    this.state = state;
    this.styles = state.styles;
    this.lookups = lookupsOf(projection);
    this.cellIds = new Map([["Default", 0], ...state.cellIds]);
    this.formatted = new Set(state.formatted);
    if (projection.defaultFormatted) {
      this.formatted.add(0);
    }
    this.implied = new ImpliedFormatStyles(state.styles, state.implied);
  }

  cellStyleId(name: string): number {
    let id = this.cellIds.get(name);
    if (id === undefined) {
      id = this.styles.length;
      const defined = this.lookups.cellStyles.get(name);
      const entry = defined ?? this.projection.undefinedCellStyle;
      // A defined entry is used once per read (its name is then registered);
      // the undefined one is every undefined name's, so each gets a copy.
      const style = defined ? defined.style : { ...structuredClone(entry.style), sourceStyleName: name };
      if (entry.formatted) {
        this.formatted.add(id);
        this.state.formatted.push(id);
      }
      this.styles.push(style);
      this.cellIds.set(name, id);
      this.state.cellIds.push([name, id]);
    }
    return id;
  }

  impliedStyleId(baseId: number, code: string): number {
    return this.formatted.has(baseId) ? baseId : this.implied.styleFor(baseId, code);
  }

  columnWidthPx(name: string | undefined): number | undefined {
    const width = name !== undefined && this.lookups.columnWidths.has(name) ? this.lookups.columnWidths.get(name)! : this.projection.undefinedColumnWidth;
    return width ?? undefined;
  }

  rowHeight(name: string | undefined): OdsRowHeight {
    return (name === undefined ? undefined : this.lookups.rowHeights.get(name)) ?? this.projection.undefinedRowHeight;
  }

  textFont(name: string): CellFont | undefined {
    return this.lookups.textStyles.get(name)?.font;
  }

  textStyleModelled(name: string): boolean {
    return this.lookups.textStyles.get(name)?.modelled ?? true;
  }

  tableStyle(name: string | undefined): OdsTableStyle {
    return (name === undefined ? undefined : this.lookups.tableStyles.get(name)) ?? this.projection.undefinedTableStyle;
  }

  cellStyleMaps(name: string): readonly OdsStyleMap[] | undefined {
    return this.lookups.mapped.get(name);
  }

  mappedCellStyles(): readonly string[] {
    return this.lookups.mappedNames;
  }

  conditionalStyle(name: string): { name: string; dxf: Dxf } | null {
    const { conditional, displayNames } = this.lookups;
    let resolved: string | undefined;
    if (conditional.has(name)) {
      resolved = name;
    } else if (conditional.has(encodeOdsStyleName(name))) {
      resolved = encodeOdsStyleName(name);
    } else {
      resolved = displayNames.get(name);
    }
    const dxf = resolved === undefined ? undefined : conditional.get(resolved);
    return resolved === undefined || dxf === undefined ? null : { name: resolved, dxf: structuredClone(dxf) };
  }
}

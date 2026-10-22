// Excel's default VML note shape (SPEC-135 "xlsx wire mapping", VML) — the
// golden every generated note and thread-placeholder shape follows, byte for
// byte as Excel 365 writes one: CRLF line breaks, `fillcolor="#ffffe1"`,
// `visibility:hidden` unless shown (then `<x:Visible/>` after `<x:Column>`),
// the anchor `{c+1}, 15, {max(r−1,0)}, 10, {c+3}, 15, {max(r−1,0)+4}, 4`
// (`col, dx, row, dy, col2, dx2, row2, dy2`, 0-based, offsets in px) and the
// id `_x0000_s{idmap·1024+n}`, `z-index:n`. The shape's style carries its
// box (`width:…pt;height:…pt`, 108 × 59.25 pt by default — 144 × 79 px)
// and its position (`margin-left`/`margin-top`, the anchor's start in pt).
// A box other than the default moves the anchor's far corner so the shape
// spans it (a moved note keeps its size, not its custom position).
//
// Also here: the `_x0000_t202` shapetype a note shape types itself with and
// the skeleton of a VML part minted for a sheet's first note.
import type { CellRef } from "../address";
import type { GridMetrics } from "../geometry";

export const VML_NOTE_SHAPETYPE_ID = "_x0000_t202";

/** The note shapetype, as Excel writes it ahead of the first note shape. */
export const VML_NOTE_SHAPETYPE = [
  `<v:shapetype id="${VML_NOTE_SHAPETYPE_ID}" coordsize="21600,21600" o:spt="202"`,
  `  path="m,l,21600r21600,l21600,xe">`,
  `  <v:stroke joinstyle="miter"/>`,
  `  <v:path gradientshapeok="t" o:connecttype="rect"/>`,
  ` </v:shapetype>`,
].join("\r\n");

/** Excel's default note box in pt (= NOTE_DEFAULT_BOX, 144 × 79 px). */
export const DEFAULT_NOTE_SIZE_PT = { width: 108, height: 59.25 } as const;

const PX_TO_PT = 0.75;

/** The anchor's offsets into its start cell, in px. */
const START_DX = 15;
const START_DY = 10;

/** `_x0000_s{idmap·1024+n}`. */
export function vmlShapeId(idmap: number, n: number): string {
  return `_x0000_s${idmap * 1024 + n}`;
}

/** Excel's default anchor for a note on `ref`. */
export function defaultNoteAnchor(ref: CellRef): number[] {
  const top = Math.max(ref.row - 1, 0);
  return [ref.col + 1, START_DX, top, START_DY, ref.col + 3, 15, top + 4, 4];
}

/** A length in pt as VML writes it: at most two decimals, no trailing zeros. */
export function vmlPt(value: number): string {
  return `${Number((Math.round(value * 100) / 100).toFixed(2))}pt`;
}

/** What a generated note shape needs. */
export interface VmlNoteShape {
  shapeId: string;
  ref: CellRef;
  visible: boolean;
  widthPt: number;
  heightPt: number;
  marginLeftPt: number;
  marginTopPt: number;
  zIndex: number;
  anchor: readonly number[];
}

/** One note shape, Excel's bytes: `<v:shape …>` through ` </v:shape>`. */
export function vmlNoteShapeXml(shape: VmlNoteShape): string {
  return [
    `<v:shape id="${shape.shapeId}" type="#${VML_NOTE_SHAPETYPE_ID}" style='position:absolute;`,
    `  margin-left:${vmlPt(shape.marginLeftPt)};margin-top:${vmlPt(shape.marginTopPt)};width:${vmlPt(shape.widthPt)};height:${vmlPt(shape.heightPt)};z-index:${shape.zIndex};`,
    `  visibility:${shape.visible ? "visible" : "hidden"}' fillcolor="#ffffe1" o:insetmode="auto">`,
    `  <v:fill color2="#ffffe1"/>`,
    `  <v:shadow color="black" obscured="t"/>`,
    `  <v:path o:connecttype="none"/>`,
    `  <v:textbox style='mso-direction-alt:auto'>`,
    `   <div style='text-align:left'></div>`,
    `  </v:textbox>`,
    `  <x:ClientData ObjectType="Note">`,
    `   <x:MoveWithCells/>`,
    `   <x:SizeWithCells/>`,
    `   <x:Anchor>`,
    `    ${shape.anchor.join(", ")}</x:Anchor>`,
    `   <x:AutoFill>False</x:AutoFill>`,
    `   <x:Row>${shape.ref.row}</x:Row>`,
    `   <x:Column>${shape.ref.col}</x:Column>`,
    ...(shape.visible ? [`   <x:Visible/>`] : []),
    `  </x:ClientData>`,
    ` </v:shape>`,
  ].join("\r\n");
}

/** The shape for a note on `ref` at the default placement: shape number `n`
 * of a part whose `o:idmap` is `idmap`, its box `box` px (default 144 × 79),
 * positioned through the sheet's `metrics` (Excel's default 64 × 20 px grid
 * when absent). */
export function vmlNoteShapeFor(
  ref: CellRef,
  options: { idmap: number; n: number; visible: boolean; box?: { widthPx: number; heightPx: number }; metrics?: Pick<GridMetrics, "colLeft" | "rowTop" | "colAt" | "rowAt"> },
): string {
  const metrics = options.metrics ?? DEFAULT_GRID;
  const anchor = defaultNoteAnchor(ref);
  const left = metrics.colLeft(anchor[0]!) + START_DX;
  const top = metrics.rowTop(anchor[2]!) + START_DY;
  const custom = options.box !== undefined && (options.box.widthPx !== 144 || options.box.heightPx !== 79);
  if (custom) {
    // The far corner where the box ends, so Excel (which sizes by the
    // anchor) shows the note at its own size.
    const right = left + options.box!.widthPx;
    const bottom = top + options.box!.heightPx;
    const col2 = metrics.colAt(right);
    const row2 = metrics.rowAt(bottom);
    anchor[4] = col2;
    anchor[5] = Math.round(right - metrics.colLeft(col2));
    anchor[6] = row2;
    anchor[7] = Math.round(bottom - metrics.rowTop(row2));
  }
  return vmlNoteShapeXml({
    shapeId: vmlShapeId(options.idmap, options.n),
    ref,
    visible: options.visible,
    widthPt: custom ? options.box!.widthPx * PX_TO_PT : DEFAULT_NOTE_SIZE_PT.width,
    heightPt: custom ? options.box!.heightPx * PX_TO_PT : DEFAULT_NOTE_SIZE_PT.height,
    marginLeftPt: left * PX_TO_PT,
    marginTopPt: top * PX_TO_PT,
    zIndex: options.n,
    anchor,
  });
}

/** Excel's default grid: 64 px columns, 20 px rows. */
const DEFAULT_GRID: Pick<GridMetrics, "colLeft" | "rowTop" | "colAt" | "rowAt"> = {
  colLeft: (c) => c * 64,
  rowTop: (r) => r * 20,
  colAt: (x) => Math.floor(x / 64),
  rowAt: (y) => Math.floor(y / 20),
};

/** A VML part minted for a sheet's first note: the root, the shape layout
 * with `o:idmap data`, the note shapetype, then `shapes` (each a
 * vmlNoteShapeXml text). */
export function vmlDrawingXml(idmap: number, shapes: readonly string[]): string {
  return (
    [
      `<xml xmlns:v="urn:schemas-microsoft-com:vml"`,
      ` xmlns:o="urn:schemas-microsoft-com:office:office"`,
      ` xmlns:x="urn:schemas-microsoft-com:office:excel">`,
      ` <o:shapelayout v:ext="edit">`,
      `  <o:idmap v:ext="edit" data="${idmap}"/>`,
      ` </o:shapelayout>`,
    ].join("\r\n") +
    VML_NOTE_SHAPETYPE +
    shapes.join("") +
    `</xml>`
  );
}

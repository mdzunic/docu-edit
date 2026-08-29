// Library thumbnails (SPEC-067): a lightweight rendered preview of a
// document's first page, produced from the model alone. The product's real
// page renderer (DocPage.vue + usePagination) is DOM-measurement driven and
// would have to mount the whole document off-screen — far too heavy to run
// per row of a 200-document library. Instead we extract the first page's
// leading blocks as a small list of styled lines — paragraphs as text,
// images as their resolved data: URI scaled the way the page scales them —
// and let the card render them as a scaled-down paper miniature. It is a
// genuine preview of the first page's content, cheap enough to schedule
// through the idle scheduler.
//
// Lines are flow-ordered, but the drawings the page positions by their own
// offsets are NOT: a run of them (with the connectors between them) becomes
// one `ThumbnailDiagram` laid out at its real relative geometry, so a
// flowchart previews as a flowchart rather than as a column of loose boxes.
// Where that run sits *down the page* still follows block order — resolving
// a paragraph-relative anchor needs the layout this module deliberately does
// not run.
//
// This module is the pure, framework-free part: it builds and caches thumbnail
// data. The IntersectionObserver gating and idle scheduling live in the
// LibraryThumbnail component, so only visible cards ever call in here.
import type {
  Block,
  Chart,
  Connector,
  DocumentModel,
  Image,
  Section,
  Shape,
  Table,
  TableCell,
} from "./documentModel";
import {
  isChart,
  isConnector,
  isImage,
  isShape,
  isTable,
  sectionContentWidthCm,
  sectionPageSize,
} from "./documentModel";
import type { ShapeVisual } from "./core/shapes";
import {
  buildConnectorPaint,
  connectorPaintBounds,
  scaleRoutedPath,
  type ConnectorPaint,
} from "./core/connectorPaint";
import { translateRoutedPath } from "./core/connectors";
import { mediaDataUri } from "./core/media";
import { cmToPx, pxToCm } from "./core/units";
import { routeConnectorInSection } from "./connectors";
import { blockText } from "./textEditing";

/** An image on a thumbnail preview line, resolved for rendering. */
export interface ThumbnailImage {
  /** The media part as a `data:` URI (core/media's mediaDataUri — the same
   * resolve-only-for-rendering rule as the page canvas). */
  src: string;
  /** Display width as a fraction (0..1] of the section's content width, so
   * the miniature sizes the image exactly like the page does. */
  widthFraction: number;
  /** Display aspect ratio (widthCm / heightCm). The page stretches the
   * bitmap to its display box, so the miniature must render this ratio, not
   * the bitmap's natural one. */
  aspectRatio: number;
}

/** A table on a thumbnail preview line: its leading cells' text, laid out
 * with the table's own column proportions, so the miniature shows the same
 * grid the page does. */
export interface ThumbnailTable {
  /** Up to THUMBNAIL_TABLE_MAX_ROWS × THUMBNAIL_TABLE_MAX_COLUMNS of cell
   * text, in grid order (covered cells of a merge read as ""). */
  rows: string[][];
  /** How many leading rows are header rows (rendered emphasised). */
  headerRows: number;
  /** Column widths as fractions of the shown columns' total (sums to ~1). */
  columnFractions: number[];
  /** Table width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
}

/** A chart on a thumbnail preview line: the chart block itself (deep-cloned,
 * so the cache never aliases the live model), which the card paints small
 * through the same core/chartCanvas painter the page uses. */
export interface ThumbnailChart {
  chart: Chart;
  /** Display width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
  /** Display aspect ratio (widthCm / heightCm). */
  aspectRatio: number;
}

/** A shape or text box on a thumbnail preview line: the paintable visual
 * subset for the card's canvas (the same core/shapeCanvas painter the page
 * uses) plus the text frame's leading lines, which the page renders as DOM
 * — so a text box previews as its text, not as an empty rectangle. */
export interface ThumbnailShape {
  /** The ShapeVisual to paint — a stripped clone: `unsupportedGeometryXml`
   * is reduced to "" when present (only its presence drives the dashed
   * placeholder), so the cache never holds the preserved XML. */
  shape: ShapeVisual;
  /** The shape's box in model px — the canvas backing-store size. */
  widthPx: number;
  heightPx: number;
  /** Leading non-empty paragraphs of the shape's text frame, capped. */
  textLines: string[];
  /** The text frame's vertical alignment ("top" when there is no frame). */
  vAlign: "top" | "middle" | "bottom";
  /** Display width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
  /** Display aspect ratio (width / height, degenerate axes clamped). */
  aspectRatio: number;
}

/** A LONE connector on a thumbnail preview line (SPEC-070) — one with no
 * neighbouring drawing to be positioned against, so there is no diagram to
 * put it in. It previews as the line it paints: routed from the model
 * exactly as the page routes it, then framed by its own painted bounds.
 * A connector beside drawings goes into a ThumbnailDiagram instead. */
export interface ThumbnailConnector {
  /** Paint data in the miniature's own px box (0,0 at its top-left). */
  paint: ConnectorPaint;
  /** The box the paint sits in — the SVG's viewBox. */
  widthPx: number;
  heightPx: number;
  /** Display width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
  /** Display aspect ratio (width / height). */
  aspectRatio: number;
}

/** One drawing inside a diagram line, placed by its own offsets. The
 * placement is in fractions of the diagram's box, so the card can size the
 * whole diagram however it likes and every member stays where it belongs. */
export interface ThumbnailDiagramItem {
  leftFraction: number;
  topFraction: number;
  widthFraction: number;
  heightFraction: number;
  /** Exactly one is present. The payload is the same one a stand-alone line
   * carries, so the card paints it with the same code; its own
   * `widthFraction`/`aspectRatio` are unused here — the placement above
   * replaces them. */
  shape?: ThumbnailShape;
  chart?: ThumbnailChart;
}

/** A diagram on a thumbnail preview line: a run of offset-positioned
 * drawings and the connectors between them, laid out at their real relative
 * positions instead of stacked one per line.
 *
 * This is the only placement the model alone can support, and it is the one
 * that matters: `shapeDisplayStyle` positions a non-inline Shape/Chart
 * absolutely at `offsetXCm`/`offsetYCm` from its anchor, so drawings sharing
 * an anchor share an origin — the same assumption `routeConnectorInSection`
 * already makes. Where a drawing sits *down the page* depends on where its
 * anchor paragraph landed, which needs the layout the miniature deliberately
 * does not run, so the diagram as a whole still takes its turn in flow. */
export interface ThumbnailDiagram {
  items: ThumbnailDiagramItem[];
  /** Connectors routed in the diagram box's own px (the SVG's viewBox). */
  connectors: ConnectorPaint[];
  /** The box in model px — the items' and connectors' coordinate basis. */
  widthPx: number;
  heightPx: number;
  /** Indent from the content box's left edge, as a fraction of the section's
   * content width — the diagram's own horizontal place on the page. */
  leftFraction: number;
  /** Display width as a fraction (0..1] of the section's content width. */
  widthFraction: number;
  /** Display aspect ratio (width / height). */
  aspectRatio: number;
}

/** One line of a thumbnail preview: a text line, an image, a table, a chart,
 * a shape, a connector or a diagram. Exactly one of the optional members is
 * present on a non-text line; `text` is then "" and `level` 0. All optional
 * members are absent on every text line, so pre-image thumbnails and their
 * tests keep their exact shape. */
export interface ThumbnailLine {
  text: string;
  /** 0 for body text, 1..6 for a heading of that level — the card scales the
   * type so a document's shape (a big title over paragraphs) reads at a
   * glance. */
  level: number;
  image?: ThumbnailImage;
  table?: ThumbnailTable;
  chart?: ThumbnailChart;
  shape?: ThumbnailShape;
  connector?: ThumbnailConnector;
  diagram?: ThumbnailDiagram;
}

export interface Thumbnail {
  lines: ThumbnailLine[];
  /** True when the model could not be previewed (corrupt, or nothing
   * renderable). The card then shows a format-badge placeholder, never an
   * error. */
  failed: boolean;
}

/** How many leading lines a first-page preview shows before it is visually
 * "full". More than this never fits the miniature. An image consumes several
 * line slots — its share of the page height — so a picture-heavy first page
 * fills the budget (and stops the walk, and its data-URI decodes) as fast as
 * it fills the miniature. */
export const THUMBNAIL_MAX_LINES = 16;

/** Paper aspect ratio (A4 portrait, width / height) the placeholder and the
 * rendered thumbnail both use, so nothing reflows when the real one arrives. */
export const THUMBNAIL_ASPECT_RATIO = 210 / 297;

/** How much of a table the miniature shows: enough rows and columns to read
 * its shape, never the whole grid. */
export const THUMBNAIL_TABLE_MAX_ROWS = 6;
export const THUMBNAIL_TABLE_MAX_COLUMNS = 6;

function isParagraph(block: Block): block is Extract<Block, { type: "paragraph" }> {
  return block.type === "paragraph";
}

/** Resolves an image block to its thumbnail line, or null when it cannot be
 * previewed (media part missing/empty, degenerate display size) — a bad
 * image is skipped, never fails the whole thumbnail. */
function thumbnailImage(
  model: DocumentModel,
  sectionIndex: number,
  block: Image,
): ThumbnailImage | null {
  const part = (model.media ?? []).find((media) => media.mediaId === block.mediaId);
  if (!part || part.data.length === 0 || !(block.widthCm > 0) || !(block.heightCm > 0)) {
    return null;
  }
  const src = mediaDataUri(part);
  return {
    src,
    widthFraction: drawingWidthFraction(model, sectionIndex, block),
    aspectRatio: block.widthCm / block.heightCm,
  };
}

/** How many of the THUMBNAIL_MAX_LINES slots a drawing (image/chart/shape)
 * consumes: its share of the page height, so the budget still means "about
 * one miniature's worth". */
function drawingLineCost(
  model: DocumentModel,
  sectionIndex: number,
  block: { heightCm: number },
): number {
  const pageHeightCm = sectionPageSize(model, sectionIndex).heightCm;
  const share = pageHeightCm > 0 ? block.heightCm / pageHeightCm : 0;
  const cost = Math.round(share * THUMBNAIL_MAX_LINES);
  return Math.min(THUMBNAIL_MAX_LINES, Math.max(1, Number.isFinite(cost) ? cost : 1));
}

/** A drawing's display width as a fraction of the section's content width,
 * clamped so it stays visible and never overflows the miniature. */
function drawingWidthFraction(
  model: DocumentModel,
  sectionIndex: number,
  block: { widthCm: number },
): number {
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  return Math.min(1, Math.max(0.02, block.widthCm / contentWidthCm));
}

/** A table cell's text for the miniature — its paragraphs joined, trimmed
 * and capped (the card ellipsises anyway; the cap just keeps cached
 * thumbnails small). */
function cellPreviewText(cell: TableCell): string {
  return cell.blocks
    .map((block) => blockText(block))
    .join(" ")
    .trim()
    .slice(0, 60);
}

/** Resolves a table block to its thumbnail line, or null for a table with
 * nothing to show. */
function thumbnailTable(
  model: DocumentModel,
  sectionIndex: number,
  block: Table,
): ThumbnailTable | null {
  const rows = block.rows.slice(0, THUMBNAIL_TABLE_MAX_ROWS);
  if (rows.length === 0) {
    return null;
  }
  const columnCount = Math.min(
    THUMBNAIL_TABLE_MAX_COLUMNS,
    Math.max(1, block.columnWidthsCm.length),
  );
  const widthsCm = block.columnWidthsCm.slice(0, columnCount);
  const totalShownCm = widthsCm.reduce((sum, cm) => sum + Math.max(0, cm), 0);
  const columnFractions =
    totalShownCm > 0
      ? widthsCm.map((cm) => Math.max(0, cm) / totalShownCm)
      : widthsCm.map(() => 1 / columnCount);
  const totalCm = block.columnWidthsCm.reduce((sum, cm) => sum + Math.max(0, cm), 0);
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  const widthFraction =
    block.widthMode === "percent" && typeof block.widthPercent === "number"
      ? Math.min(1, Math.max(0.05, block.widthPercent / 100))
      : Math.min(1, Math.max(0.05, totalCm > 0 ? totalCm / contentWidthCm : 1));
  let headerRows = 0;
  while (headerRows < rows.length && rows[headerRows].isHeader) {
    headerRows += 1;
  }
  return {
    rows: rows.map((row) => row.cells.slice(0, columnCount).map(cellPreviewText)),
    headerRows,
    columnFractions,
    widthFraction,
  };
}

/** How many text-frame lines a shape miniature shows. */
const SHAPE_TEXT_MAX_LINES = 4;

/** Resolves a shape/text-box block to its thumbnail line, or null for a
 * degenerate one. A line/arrow's zero-ish cross axis is clamped so its
 * ratio and cost stay finite. */
function thumbnailShape(
  model: DocumentModel,
  sectionIndex: number,
  block: Shape,
): ThumbnailShape | null {
  const widthCm = Math.max(block.widthCm, 0.05);
  const heightCm = Math.max(block.heightCm, 0.05);
  if (!(widthCm > 0) || !(heightCm > 0)) {
    return null;
  }
  const visual = JSON.parse(
    JSON.stringify({
      geometry: block.geometry,
      fill: block.fill,
      outline: block.outline,
      rotationDeg: block.rotationDeg,
      adjustments: block.adjustments,
      unsupportedGeometryXml: block.unsupportedGeometryXml === undefined ? undefined : "",
    }),
  ) as ShapeVisual;
  const textLines = (block.text?.blocks ?? [])
    .filter(isParagraph)
    .map((paragraph) => blockText(paragraph).trim())
    .filter((text) => text.length > 0)
    .slice(0, SHAPE_TEXT_MAX_LINES)
    .map((text) => text.slice(0, 60));
  return {
    shape: visual,
    widthPx: Math.max(1, Math.round(cmToPx(widthCm))),
    heightPx: Math.max(1, Math.round(cmToPx(heightCm))),
    textLines,
    vAlign: block.text?.vAlign ?? "top",
    widthFraction: drawingWidthFraction(model, sectionIndex, block),
    aspectRatio: widthCm / heightCm,
  };
}

/** Resolves a connector block to its thumbnail line, or null when it routes
 * to nothing paintable. The routing is the model-level `routeConnectorInSection`
 * — the same resolver the page's DOM measurement mirrors — so an attached
 * connector previews where its shapes put it, in cm, scaled to px and framed
 * by the bounds it actually paints (arrowheads included). */
function thumbnailConnector(
  model: DocumentModel,
  sectionIndex: number,
  section: Section,
  block: Connector,
): ThumbnailConnector | null {
  const options = {
    outline: block.outline,
    startArrow: block.startArrow,
    endArrow: block.endArrow ?? ("triangle" as const),
  };
  const routed = scaleRoutedPath(routeConnectorInSection(section, block), cmToPx(1));
  const bounds = connectorPaintBounds(routed, options);
  const widthPx = bounds.right - bounds.left;
  const heightPx = bounds.bottom - bounds.top;
  if (!(widthPx > 0) || !(heightPx > 0)) {
    return null;
  }
  const paint = buildConnectorPaint(
    "connector",
    translateRoutedPath(routed, -bounds.left, -bounds.top),
    options,
  );
  return {
    paint,
    widthPx,
    heightPx,
    widthFraction: connectorWidthFraction(model, sectionIndex, widthPx, heightPx),
    aspectRatio: widthPx / heightPx,
  };
}

/** The share of a page's height one connector line may take in the
 * miniature. */
const CONNECTOR_MAX_HEIGHT_SHARE = 0.25;

/** A connector's display width fraction. Unlike a shape's, a connector's box
 * is derived from where it routes, so an elbow's vertical leg gives a tall,
 * nearly empty box — and the miniature stacks drawings in flow rather than at
 * their page positions, so one of those pushes everything after it off the
 * card. Scaling the whole line down (aspect ratio intact) until it fits
 * CONNECTOR_MAX_HEIGHT_SHARE of the page keeps the rest of the preview. */
function connectorWidthFraction(
  model: DocumentModel,
  sectionIndex: number,
  widthPx: number,
  heightPx: number,
): number {
  const widthFraction = drawingWidthFraction(model, sectionIndex, { widthCm: pxToCm(widthPx) });
  const pageHeightCm = sectionPageSize(model, sectionIndex).heightCm;
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  if (!(pageHeightCm > 0) || !(contentWidthCm > 0) || !(widthPx > 0)) {
    return widthFraction;
  }
  const shownHeightCm = ((widthFraction * contentWidthCm) / widthPx) * heightPx;
  const heightShare = shownHeightCm / pageHeightCm;
  if (heightShare <= CONNECTOR_MAX_HEIGHT_SHARE) {
    return widthFraction;
  }
  return Math.max(0.02, (widthFraction * CONNECTOR_MAX_HEIGHT_SHARE) / heightShare);
}

/** How many drawings one diagram line renders. A hundred-shape page still
 * has to be a cheap miniature; past this the run is closed and any further
 * members start a second diagram. */
export const THUMBNAIL_DIAGRAM_MAX_ITEMS = 24;

/** The blocks a diagram is made of: the drawings the page positions by their
 * offsets (a non-inline Shape or Chart — `imageDisplayStyle` positions no
 * Image that way, so images stay in flow) plus the connectors between them. */
type DiagramMember = Shape | Chart | Connector;

/** Whether `block` can join a diagram run. */
function isDiagramMember(block: Block): block is DiagramMember {
  if (isConnector(block)) {
    return true;
  }
  return (isShape(block) || isChart(block)) && block.wrap !== "inline";
}

/** A diagram member's box in the shared offset space, in cm. */
function memberBoxCm(member: Shape | Chart): { x: number; y: number; w: number; h: number } {
  return {
    x: member.offsetXCm ?? 0,
    y: member.offsetYCm ?? 0,
    w: Math.max(member.widthCm, 0.05),
    h: Math.max(member.heightCm, 0.05),
  };
}

/** Lays a run of drawings and connectors out at their real relative
 * positions. Null when nothing in the run has a box worth showing — the
 * caller then falls back to one line per member. */
function thumbnailDiagram(
  model: DocumentModel,
  sectionIndex: number,
  section: Section,
  run: DiagramMember[],
): ThumbnailDiagram | null {
  const drawings = run.filter((member): member is Shape | Chart => !isConnector(member));
  const routes = run
    .filter(isConnector)
    .map((connector) => ({
      connector,
      path: routeConnectorInSection(section, connector),
      options: {
        outline: connector.outline,
        startArrow: connector.startArrow,
        endArrow: connector.endArrow ?? ("triangle" as const),
      },
    }));

  // The union of every member's painted extent, in the shared cm space.
  let minX = Infinity;
  let minY = Infinity;
  let maxX = -Infinity;
  let maxY = -Infinity;
  for (const drawing of drawings) {
    const box = memberBoxCm(drawing);
    minX = Math.min(minX, box.x);
    minY = Math.min(minY, box.y);
    maxX = Math.max(maxX, box.x + box.w);
    maxY = Math.max(maxY, box.y + box.h);
  }
  for (const route of routes) {
    // The connector's painted bounds are computed in px (arrowheads scale
    // with the stroke width, which is in points), then read back as cm.
    const bounds = connectorPaintBounds(scaleRoutedPath(route.path, cmToPx(1)), route.options);
    minX = Math.min(minX, pxToCm(bounds.left));
    minY = Math.min(minY, pxToCm(bounds.top));
    maxX = Math.max(maxX, pxToCm(bounds.right));
    maxY = Math.max(maxY, pxToCm(bounds.bottom));
  }
  const boxWidthCm = maxX - minX;
  const boxHeightCm = maxY - minY;
  if (!(boxWidthCm > 0) || !(boxHeightCm > 0)) {
    return null;
  }

  const items: ThumbnailDiagramItem[] = [];
  for (const drawing of drawings) {
    const box = memberBoxCm(drawing);
    const placement = {
      leftFraction: (box.x - minX) / boxWidthCm,
      topFraction: (box.y - minY) / boxHeightCm,
      widthFraction: box.w / boxWidthCm,
      heightFraction: box.h / boxHeightCm,
    };
    if (isChart(drawing)) {
      const chart = thumbnailChart(model, sectionIndex, drawing);
      if (chart) {
        items.push({ ...placement, chart });
      }
      continue;
    }
    const shape = thumbnailShape(model, sectionIndex, drawing);
    if (shape) {
      items.push({ ...placement, shape });
    }
  }

  const widthPx = cmToPx(boxWidthCm);
  const heightPx = cmToPx(boxHeightCm);
  const connectors = routes.map((route, index) =>
    buildConnectorPaint(
      `connector-${index}`,
      translateRoutedPath(scaleRoutedPath(route.path, cmToPx(1)), -cmToPx(minX), -cmToPx(minY)),
      route.options,
    ),
  );
  if (items.length === 0 && connectors.length === 0) {
    return null;
  }

  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  const widthFraction = drawingWidthFraction(model, sectionIndex, { widthCm: boxWidthCm });
  const leftFraction =
    contentWidthCm > 0
      ? Math.min(Math.max(0, minX / contentWidthCm), Math.max(0, 1 - widthFraction))
      : 0;
  return {
    items,
    connectors,
    widthPx,
    heightPx,
    leftFraction,
    widthFraction,
    aspectRatio: boxWidthCm / boxHeightCm,
  };
}

/** How tall a diagram is shown, in cm — its clamped display width carried
 * through its own aspect ratio. The line budget charges this, so a page-tall
 * flowchart fills the miniature and stops the walk, exactly as a page-tall
 * image already does. */
function diagramShownHeightCm(
  model: DocumentModel,
  sectionIndex: number,
  diagram: ThumbnailDiagram,
): number {
  const contentWidthCm = sectionContentWidthCm(model, sectionIndex);
  return diagram.aspectRatio > 0 ? (diagram.widthFraction * contentWidthCm) / diagram.aspectRatio : 0;
}

/** Resolves a chart block to its thumbnail line (a deep clone — the cache
 * must never alias the live model), or null for a degenerate one. */
function thumbnailChart(
  model: DocumentModel,
  sectionIndex: number,
  block: Chart,
): ThumbnailChart | null {
  if (!(block.widthCm > 0) || !(block.heightCm > 0)) {
    return null;
  }
  return {
    chart: JSON.parse(JSON.stringify(block)) as Chart,
    widthFraction: drawingWidthFraction(model, sectionIndex, block),
    aspectRatio: block.widthCm / block.heightCm,
  };
}

/** Extracts the leading styled lines of a document's first page. Walks blocks
 * in order — paragraphs as text (with heading level), images as data: URIs,
 * tables as capped text grids, charts and shapes as paintable clones for the
 * card's canvas — until the preview is full; whatever cannot be previewed
 * (a broken image, a degenerate drawing) is skipped rather than failing the
 * whole thumbnail. */
export function buildThumbnail(model: DocumentModel): Thumbnail {
  try {
    const lines: ThumbnailLine[] = [];
    let used = 0;
    const sections = model.sections ?? [];
    for (let sectionIndex = 0; sectionIndex < sections.length; sectionIndex += 1) {
      const section = sections[sectionIndex];
      // Offset-positioned drawings and their connectors are gathered into a
      // run and laid out together (see ThumbnailDiagram) rather than stacked
      // one per line — a flowchart previewed as a column of loose boxes with
      // its arrows dangling underneath before this.
      let run: DiagramMember[] = [];

      /** Emits the pending run: a diagram when there is relative geometry to
       * show, otherwise the single member's own line (unchanged). */
      const flushRun = () => {
        if (run.length === 0) {
          return;
        }
        const members = run;
        run = [];
        if (members.length > 1) {
          const diagram = thumbnailDiagram(model, sectionIndex, section, members);
          if (diagram) {
            lines.push({ text: "", level: 0, diagram });
            used += drawingLineCost(model, sectionIndex, {
              heightCm: diagramShownHeightCm(model, sectionIndex, diagram),
            });
            return;
          }
        }
        // One member, or a run with no showable box: one line each, exactly
        // as before diagrams existed.
        for (const member of members) {
          if (isConnector(member)) {
            const connector = thumbnailConnector(model, sectionIndex, section, member);
            if (connector) {
              lines.push({ text: "", level: 0, connector });
              used += drawingLineCost(model, sectionIndex, { heightCm: pxToCm(connector.heightPx) });
            }
            continue;
          }
          if (isChart(member)) {
            const chart = thumbnailChart(model, sectionIndex, member);
            if (chart) {
              lines.push({ text: "", level: 0, chart });
              used += drawingLineCost(model, sectionIndex, member);
            }
            continue;
          }
          const shape = thumbnailShape(model, sectionIndex, member);
          if (shape) {
            lines.push({ text: "", level: 0, shape });
            used += drawingLineCost(model, sectionIndex, member);
          }
        }
      };

      for (const block of section.blocks ?? []) {
        if (isDiagramMember(block)) {
          if (run.length >= THUMBNAIL_DIAGRAM_MAX_ITEMS) {
            flushRun();
          }
          run.push(block);
          continue;
        }
        // The anchor paragraphs an insert leaves between drawings carry no
        // text and must not split the diagram they anchor.
        if (run.length > 0 && isParagraph(block) && blockText(block).trim().length === 0) {
          continue;
        }
        flushRun();
        if (used >= THUMBNAIL_MAX_LINES) {
          return { lines, failed: false };
        }
        if (isImage(block)) {
          const image = thumbnailImage(model, sectionIndex, block);
          if (image) {
            lines.push({ text: "", level: 0, image });
            used += drawingLineCost(model, sectionIndex, block);
          }
          continue;
        }
        if (isTable(block)) {
          const table = thumbnailTable(model, sectionIndex, block);
          if (table) {
            lines.push({ text: "", level: 0, table });
            used += table.rows.length;
          }
          continue;
        }
        if (isChart(block)) {
          const chart = thumbnailChart(model, sectionIndex, block);
          if (chart) {
            lines.push({ text: "", level: 0, chart });
            used += drawingLineCost(model, sectionIndex, block);
          }
          continue;
        }
        if (isShape(block)) {
          const shape = thumbnailShape(model, sectionIndex, block);
          if (shape) {
            lines.push({ text: "", level: 0, shape });
            used += drawingLineCost(model, sectionIndex, block);
          }
          continue;
        }
        if (!isParagraph(block)) {
          continue;
        }
        const text = blockText(block).trim();
        if (text.length === 0) {
          continue;
        }
        lines.push({ text, level: block.headingLevel ?? 0 });
        used += 1;
      }
      flushRun();
      if (used >= THUMBNAIL_MAX_LINES) {
        break;
      }
    }
    return { lines, failed: false };
  } catch {
    return { lines: [], failed: true };
  }
}

// Cache keyed by document id, storing the modified timestamp it was built for
// (SPEC-067): a thumbnail is produced once per edit, not once per view. A
// re-render for the same id overwrites the entry, so a stale timestamp is
// discarded automatically. The cache is purely in-memory and never persisted,
// so it costs no IndexedDB quota — quota pressure "evicts" it for free — and
// clearing it is always safe.
interface CacheEntry {
  lastModified: number;
  thumbnail: Thumbnail;
}

const cache = new Map<string, CacheEntry>();

/** The cached thumbnail for a document at a given modified time, or undefined
 * when none exists or the cached one is for an older edit (and must be
 * re-rendered). */
export function getCachedThumbnail(
  id: string,
  lastModified: number,
): Thumbnail | undefined {
  const entry = cache.get(id);
  if (entry && entry.lastModified === lastModified) {
    return entry.thumbnail;
  }
  return undefined;
}

/** Builds (or returns the cached) thumbnail for a document, keyed by id and
 * modified time. Overwrites any stale entry for the same id. */
export function renderThumbnail(
  id: string,
  lastModified: number,
  model: DocumentModel,
): Thumbnail {
  const cached = getCachedThumbnail(id, lastModified);
  if (cached) {
    return cached;
  }
  const thumbnail = buildThumbnail(model);
  cache.set(id, { lastModified, thumbnail });
  return thumbnail;
}

/** Drops one document's cached thumbnail. */
export function evictThumbnail(id: string): void {
  cache.delete(id);
}

/** Drops every cached thumbnail. Safe at any time — the cache is derived data. */
export function clearThumbnailCache(): void {
  cache.clear();
}

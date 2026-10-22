// SPEC-070: the view-layer half of connectors — the two-click draw mode, the
// connection-site hints, the page-level overlay geometry (measured live from
// the DOM so a shape drag re-routes with no commit), the endpoint/midpoint
// drag handles, and the formatting commands. The pure routing and the model
// commands live in lib/core/connectors.ts and lib/connectors.ts; this file
// only measures the DOM and drives them.
//
// The measurement descends into `.doc-shape` (NOT the zero-height
// `.doc-shape-anchor-floating` span that is the block's `[data-section]
// [data-block]` element) so a connector routes to the actually-painted box —
// the same pattern useShapes uses for its own handles.
import { computed, ref, shallowRef, type Ref } from "vue";
import type { EditorContext } from "./context";
import {
  asConnectorKind,
  documentThemeOf,
  isChart,
  isConnector,
  isImage,
  isShape,
  type Connector,
  type ConnectorKind,
  type DocumentModel,
} from "../../lib/documentModel";
import { resolveColor } from "../../lib/core/theme";
import type { ArrowheadKind, ArrowheadSize, ShapeOutline } from "../../lib/core/shapes";
import {
  clampSite,
  connectionSitePoint,
  connectionSites,
  nearestSite,
  routeConnector,
  routedPathToSvg,
  segmentPolygonCrossing,
  translateRoutedPath,
  type ConnectorBox,
  type ConnectorEndpoint,
  type RoutedPath,
} from "../../lib/core/connectors";
import {
  buildConnectorPaint,
  connectorPaintBounds,
  type ConnectorPaint,
  type ConnectorPaintOptions,
} from "../../lib/core/connectorPaint";
import type { Point } from "../../lib/core/shapeGeometry";
import {
  connectorAt,
  deleteConnector,
  insertConnector,
  setConnectorAdjustment,
  setConnectorAltText,
  setConnectorArrowheads,
  setConnectorEndpoint,
  setConnectorKind,
  setConnectorOutline,
  setConnectorWrap,
  type ConnectorLocation,
  type NewConnector,
} from "../../lib/connectors";
import { reorderShapeZ, type ZOrderCommand } from "../../lib/shapes";
import { measureImageGeometry } from "../../lib/caretGeometry";
import { caretSelection } from "../../lib/commands";
import type { ConnectorRender, SiteHint } from "../../lib/connectorOverlay";

export type { ConnectorRender, SiteHint } from "../../lib/connectorOverlay";

const PX_PER_CM = 96 / 2.54;

/** The one empty paint array every connector-free page is handed. */
const NO_CONNECTOR_PAINT: ConnectorPaint[] = [];

/** A free connector endpoint's page-relative px point. The model stores
 * `start`/`end` relative to the connector's OFFSET origin (offsetXCm/
 * offsetYCm), while the measured DOM origin is the connector's anchor (where
 * section-space 0,0 paints), so the offset must be added back — the exact
 * inverse of the section-space point setConnectorEndpoint stored. Omitting it
 * snapped a frozen/dropped endpoint up by the connector's own offset (the
 * AC-11 delete-freeze regression; also AC-6's free end). */
function freeEndpointPx(origin: Point, connector: Connector, which: "start" | "end"): Point {
  const free = which === "start" ? connector.start : connector.end;
  const offX = connector.offsetXCm ?? 0;
  const offY = connector.offsetYCm ?? 0;
  return { x: origin.x + (offX + free.xCm) * PX_PER_CM, y: origin.y + (offY + free.yCm) * PX_PER_CM };
}

export interface ConnectorsDeps extends EditorContext {}

/** The box's four corners (sites 4,5,6,7) — honouring rotation, so a
 * centre-site attachment clips against the rotated outline and re-routes
 * when the shape turns (AC-9), not against an axis-aligned rectangle. */
function rectPolygon(box: ConnectorBox): Point[] {
  return [4, 5, 6, 7].map((corner) => connectionSitePoint(box, corner));
}

function midpointOf(points: Point[], kind: ConnectorKind): Point {
  if (points.length === 0) {
    return { x: 0, y: 0 };
  }
  if (kind === "curved") {
    const a = points[0];
    const b = points[points.length - 1];
    return { x: (a.x + b.x) / 2, y: (a.y + b.y) / 2 };
  }
  // The middle segment's midpoint (elbow), or the whole segment (straight).
  const midIndex = Math.floor((points.length - 1) / 2);
  const p = points[midIndex];
  const q = points[Math.min(points.length - 1, midIndex + 1)];
  return { x: (p.x + q.x) / 2, y: (p.y + q.y) / 2 };
}

function straightPreviewD(a: Point, b: Point): string {
  return `M ${a.x} ${a.y} L ${b.x} ${b.y}`;
}

function projectFraction(a: Point, b: Point, p: Point): number {
  // Fraction along the dominant axis of the endpoints — matches the elbow
  // crossing's own placement.
  if (Math.abs(b.x - a.x) >= Math.abs(b.y - a.y)) {
    const denom = b.x - a.x || 1;
    return Math.min(1, Math.max(0, (p.x - a.x) / denom));
  }
  const denom = b.y - a.y || 1;
  return Math.min(1, Math.max(0, (p.y - a.y) / denom));
}

export function useConnectors(deps: ConnectorsDeps) {
  const { doc, contentRoot, pageRef, editable, currentSelection, applyEdit, getCaret, zoom } = deps;

  const drawKind = ref<ConnectorKind | null>(null);
  const drawFrom = ref<{ point: Point; attachment?: NewConnector["startAttachment"] } | null>(null);
  const drawPointer = ref<Point | null>(null);
  const hoverSites = ref<SiteHint | null>(null);
  const selected = ref<ConnectorLocation | null>(null);
  const renders = ref<ConnectorRender[]>([]);
  // The painted connectors, split per realized page and expressed in that
  // page's own px (ConnectorPaintLayer renders them inside the page box, so
  // they reach the print preview, the sheet and the PDF). A shallowRef
  // guarded by a signature: remeasuring on every scroll tick must not
  // re-render the page list when nothing actually moved.
  const pageLayers = shallowRef<ConnectorPaint[][]>([]);
  let pageLayerSignature = "";

  let endpointDrag: { location: ConnectorLocation; which: "start" | "end" } | null = null;

  // --- DOM measurement (page-relative px) -----------------------------------

  /** The scale the page stack is ACTUALLY rendered at — the measured box over
   * the layout box.
   *
   * Not `zoom.value`: that is the *requested* zoom, and the two disagree
   * exactly when a mode neutralises the `.pages` transform without changing
   * the zoom value, which is what the PDF export and printing do
   * (`is-exporting-pdf`, App.vue). Dividing untransformed rects by a 0.75
   * zoom scaled every exported connector by 1/0.75 and slid it off the
   * shapes it joins. Measuring both boxes is true in either state. */
  function renderScale(): number {
    // `pageRef` is typed Element (see context.ts); clientWidth is on Element
    // and, unlike getBoundingClientRect, is the untransformed layout width.
    const pageEl = pageRef.value;
    const layoutWidth = pageEl?.clientWidth ?? 0;
    if (!pageEl || layoutWidth <= 0) {
      return zoom.value || 1;
    }
    const scale = pageEl.getBoundingClientRect().width / layoutWidth;
    return scale > 0 ? scale : zoom.value || 1;
  }

  function pagePointFromClient(clientX: number, clientY: number): Point | null {
    const pageEl = pageRef.value;
    if (!pageEl) {
      return null;
    }
    const rect = pageEl.getBoundingClientRect();
    const z = renderScale();
    return { x: (clientX - rect.left) / z, y: (clientY - rect.top) / z };
  }

  /** The painted box of a drawing block (descends into `.doc-shape`), in
   * page-relative px. Null when it is not in the DOM. */
  function measureDrawingBox(location: ConnectorLocation): ConnectorBox | null {
    const root = contentRoot.value;
    const pageEl = pageRef.value;
    if (!root || !pageEl) {
      return null;
    }
    const el =
      root.querySelector(
        `.doc-shape-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-shape`,
      ) ??
      root.querySelector(
        `.doc-image-anchor[data-section="${location.section}"][data-block="${location.block}"] .doc-image, [data-section="${location.section}"][data-block="${location.block}"] .doc-shape`,
      );
    if (!el) {
      return null;
    }
    const geo = measureImageGeometry(pageEl, el, renderScale());
    if (!geo) {
      return null;
    }
    // The `.doc-shape` proxy box stays axis-aligned (DocShapeBlock keeps it
    // unrotated so every measurement/hit-test goes through an upright box), so
    // the box's px w/h are the shape's true unrotated size. Carry the model
    // rotation through so connectionSites rotates the sites about the centre —
    // otherwise a rotated shape's connectors never re-route (AC-9).
    return { x: geo.left, y: geo.top, w: geo.width, h: geo.height, rotationDeg: drawingRotationAt(location) };
  }

  /** The model rotation (whole degrees CW) of the drawing at `location`, or 0
   * when it is absent or not a rotatable drawing. */
  function drawingRotationAt(location: ConnectorLocation): number {
    const block = doc.value?.sections[location.section]?.blocks[location.block];
    if (block && (isShape(block) || isImage(block) || isChart(block))) {
      return block.rotationDeg ?? 0;
    }
    return 0;
  }

  /** The page-relative px top-left of a connector's own anchor box (the
   * origin its free endpoints are measured from). */
  function measureConnectorOrigin(location: ConnectorLocation): Point | null {
    const root = contentRoot.value;
    const pageEl = pageRef.value;
    if (!root || !pageEl) {
      return null;
    }
    const el = root.querySelector(
      `.doc-connector-anchor[data-section="${location.section}"][data-block="${location.block}"]`,
    );
    if (!el) {
      return null;
    }
    const pageRect = pageEl.getBoundingClientRect();
    const rect = el.getBoundingClientRect();
    const z = renderScale();
    return { x: (rect.left - pageRect.left) / z, y: (rect.top - pageRect.top) / z };
  }

  /** Resolves one endpoint of a connector to page-relative px, plus its
   * outward normal when attached. `otherPx` clips a centre attachment. */
  function resolveEndpointPx(
    location: ConnectorLocation,
    connector: Connector,
    which: "start" | "end",
    otherPx: Point,
  ): ConnectorEndpoint | null {
    const attachment = which === "start" ? connector.startAttachment : connector.endAttachment;
    if (!attachment) {
      const origin = measureConnectorOrigin(location);
      if (!origin) {
        return null;
      }
      return { point: freeEndpointPx(origin, connector, which) };
    }
    const box = measureDrawingBox({ section: location.section, block: attachment.target.block });
    if (!box) {
      // Target not in DOM (or a stale index): fall back to the free position.
      const origin = measureConnectorOrigin(location);
      return origin ? { point: freeEndpointPx(origin, connector, which) } : null;
    }
    const site = clampSite(attachment.site);
    if (site === 8) {
      const centre = connectionSitePoint(box, 8);
      const crossing = segmentPolygonCrossing(centre, otherPx, rectPolygon(box));
      if (crossing) {
        const nx = crossing.x - centre.x;
        const ny = crossing.y - centre.y;
        const len = Math.hypot(nx, ny) || 1;
        return { point: crossing, normal: { x: nx / len, y: ny / len } };
      }
      return { point: centre };
    }
    const sites = connectionSites(box);
    return { point: sites[site].point, normal: sites[site].normal };
  }

  /** Every realized `.page`'s box, in the same unscaled page-stack px the
   * connector geometry is measured in — the frame ConnectorPaintLayer's
   * page-local coordinates are derived from. */
  function measurePageBoxes(): Array<{ left: number; top: number; right: number; bottom: number }> {
    const pageEl = pageRef.value;
    if (!pageEl) {
      return [];
    }
    const pageRect = pageEl.getBoundingClientRect();
    const z = renderScale();
    return Array.from(pageEl.querySelectorAll(".page")).map((el) => {
      const rect = el.getBoundingClientRect();
      return {
        left: (rect.left - pageRect.left) / z,
        top: (rect.top - pageRect.top) / z,
        right: (rect.right - pageRect.left) / z,
        bottom: (rect.bottom - pageRect.top) / z,
      };
    });
  }

  /** Re-measures every connector and rebuilds the overlay geometry. Cheap
   * enough to run on every mousemove of a shape drag (AC-7/8/9). */
  function remeasureConnectors() {
    const model = doc.value;
    if (!model) {
      renders.value = [];
      publishPageLayers([]);
      return;
    }
    const next: ConnectorRender[] = [];
    // What each connector paints, collected here and split across pages only
    // once we know there IS one: measurePageBoxes() reads a rect off every
    // realized page, and remeasuring runs on every scroll tick — a 200-page
    // document with no connectors must not pay 200 forced layouts for it.
    const painted: Array<{ key: string; path: RoutedPath; options: ConnectorPaintOptions }> = [];
    model.sections.forEach((section, sectionIndex) => {
      section.blocks.forEach((block, blockIndex) => {
        if (!isConnector(block)) {
          return;
        }
        const location = { section: sectionIndex, block: blockIndex };
        // Two passes so a centre attachment can clip toward the far endpoint.
        const originGuess = measureConnectorOrigin(location) ?? { x: 0, y: 0 };
        const rawA = block.startAttachment
          ? measureDrawingBox({ section: sectionIndex, block: block.startAttachment.target.block })
          : null;
        const rawB = block.endAttachment
          ? measureDrawingBox({ section: sectionIndex, block: block.endAttachment.target.block })
          : null;
        const guessA = rawA ? connectionSitePoint(rawA, clampSite(block.startAttachment!.site)) : freeEndpointPx(originGuess, block, "start");
        const guessB = rawB ? connectionSitePoint(rawB, clampSite(block.endAttachment!.site)) : freeEndpointPx(originGuess, block, "end");
        const a = resolveEndpointPx(location, block, "start", guessB);
        const b = resolveEndpointPx(location, block, "end", a?.point ?? guessA);
        const a2 = resolveEndpointPx(location, block, "start", b?.point ?? guessB);
        if (!a2 || !b) {
          return;
        }
        const path = routeConnector(block.kind, a2, b, block.adjustments?.[0] ?? 0.5);
        const mid = midpointOf(path.points, block.kind);
        const key = `${sectionIndex}:${blockIndex}`;
        const paintOptions: ConnectorPaintOptions = {
          outline: block.outline,
          startArrow: block.startArrow,
          endArrow: block.endArrow ?? "triangle",
        };
        painted.push({ key, path, options: paintOptions });
        next.push({
          location,
          key,
          kind: block.kind,
          d: routedPathToSvg(path),
          selected: !!selected.value && selected.value.section === sectionIndex && selected.value.block === blockIndex,
          outline: block.outline,
          startArrow: block.startArrow,
          endArrow: block.endArrow ?? "triangle",
          a: a2.point,
          b: b.point,
          mid: block.kind === "straight" ? undefined : mid,
        });
      });
    });
    renders.value = next;
    publishPageLayers(painted.length > 0 ? splitAcrossPages(painted) : []);
  }

  /** Paints each connector onto every page it overlaps, in that page's own
   * frame: a connector whose shapes straddle a page break then shows its own
   * portion on each sheet rather than being clipped away on one of them. */
  function splitAcrossPages(
    painted: Array<{ key: string; path: RoutedPath; options: ConnectorPaintOptions }>,
  ): ConnectorPaint[][] {
    const pageBoxes = measurePageBoxes();
    const layers: ConnectorPaint[][] = pageBoxes.map(() => []);
    for (const { key, path, options } of painted) {
      const bounds = connectorPaintBounds(path, options);
      pageBoxes.forEach((box, pageIndex) => {
        if (bounds.right < box.left || bounds.left > box.right || bounds.bottom < box.top || bounds.top > box.bottom) {
          return;
        }
        layers[pageIndex].push(
          buildConnectorPaint(key, translateRoutedPath(path, -box.left, -box.top), options),
        );
      });
    }
    return layers;
  }

  /** Publishes the per-page paint data only when it actually changed. The
   * measurers run on every scroll tick, and page-local geometry is
   * scroll-invariant — without this guard each tick would hand DocPage a
   * fresh array and re-render the whole page list. */
  function publishPageLayers(layers: ConnectorPaint[][]) {
    const signature = layers
      .map((page) =>
        page
          .map((paint) => `${paint.key}|${paint.d}|${paint.color}|${paint.widthPx}|${paint.dash}|${paint.heads.map((head) => head.d).join("~")}`)
          .join(";"),
      )
      .join("/");
    if (signature === pageLayerSignature) {
      return;
    }
    pageLayerSignature = signature;
    pageLayers.value = layers;
  }

  // --- hit testing a drawing under the pointer (for draw / drag) ------------

  /** The topmost drawing whose painted box contains the client point, with
   * its measured box, or null. */
  function drawingAtClient(clientX: number, clientY: number): { location: ConnectorLocation; box: ConnectorBox } | null {
    const root = contentRoot.value;
    if (!root) {
      return null;
    }
    const els = Array.from(root.querySelectorAll(".doc-shape-anchor[data-section][data-block] .doc-shape"));
    for (let i = els.length - 1; i >= 0; i--) {
      const el = els[i];
      const rect = el.getBoundingClientRect();
      if (clientX >= rect.left && clientX <= rect.right && clientY >= rect.top && clientY <= rect.bottom) {
        const anchor = el.closest("[data-section][data-block]") as HTMLElement | null;
        if (!anchor) {
          continue;
        }
        const section = Number(anchor.getAttribute("data-section"));
        const block = Number(anchor.getAttribute("data-block"));
        const box = measureDrawingBox({ section, block });
        if (box) {
          return { location: { section, block }, box };
        }
      }
    }
    return null;
  }

  /** Updates the connection-site hint for the shape under the pointer (shown
   * while drawing or dragging an endpoint), or clears it. */
  function updateHoverSites(clientX: number, clientY: number) {
    const hit = drawingAtClient(clientX, clientY);
    if (!hit) {
      hoverSites.value = null;
      return;
    }
    hoverSites.value = { location: hit.location, sites: connectionSites(hit.box).map((s) => s.point) };
  }

  // --- two-click draw mode (AC-3..6) ----------------------------------------

  const connectorDrawKind = computed(() => drawKind.value);

  function armConnectorDraw(kind: ConnectorKind | string) {
    drawKind.value = asConnectorKind(String(kind));
    drawFrom.value = null;
    drawPointer.value = null;
    selected.value = null;
  }

  function disarmConnectorDraw() {
    drawKind.value = null;
    drawFrom.value = null;
    drawPointer.value = null;
    hoverSites.value = null;
  }

  /** Resolves a click to either a shape connection site (attach) or a free
   * page point (AC-5/AC-6). */
  function resolveDrawTarget(clientX: number, clientY: number): { point: Point; attachment?: NewConnector["startAttachment"] } | null {
    const page = pagePointFromClient(clientX, clientY);
    if (!page) {
      return null;
    }
    const hit = drawingAtClient(clientX, clientY);
    if (hit) {
      const site = nearestSite(hit.box, page);
      return { point: connectionSitePoint(hit.box, site), attachment: { target: { block: hit.location.block }, site } };
    }
    return { point: page };
  }

  /** Handles a canvas mousedown while a connector kind is armed. Returns true
   * when it consumed the event (draw mode). */
  function handleConnectorMouseDown(event: MouseEvent): boolean {
    if (!drawKind.value || !editable.value) {
      return false;
    }
    event.preventDefault();
    event.stopPropagation();
    const target = resolveDrawTarget(event.clientX, event.clientY);
    if (!target) {
      return true;
    }
    if (!drawFrom.value) {
      drawFrom.value = target;
      drawPointer.value = target.point;
      return true;
    }
    commitDraw(drawFrom.value, target);
    return true;
  }

  function handleConnectorMouseMove(event: MouseEvent) {
    if (drawKind.value) {
      const page = pagePointFromClient(event.clientX, event.clientY);
      if (page) {
        drawPointer.value = page;
      }
      updateHoverSites(event.clientX, event.clientY);
      return;
    }
    if (endpointDrag) {
      updateHoverSites(event.clientX, event.clientY);
    }
  }

  function commitDraw(
    from: { point: Point; attachment?: NewConnector["startAttachment"] },
    to: { point: Point; attachment?: NewConnector["startAttachment"] },
  ) {
    const model = doc.value;
    if (!model) {
      disarmConnectorDraw();
      return;
    }
    const kind = drawKind.value ?? "straight";
    const anchor = anchorTopLeftPx();
    const aCm = { xCm: from.point.x / PX_PER_CM, yCm: from.point.y / PX_PER_CM };
    const bCm = { xCm: to.point.x / PX_PER_CM, yCm: to.point.y / PX_PER_CM };
    const anchorCm = { xCm: anchor.x / PX_PER_CM, yCm: anchor.y / PX_PER_CM };
    const leftCm = Math.min(aCm.xCm, bCm.xCm);
    const topCm = Math.min(aCm.yCm, bCm.yCm);
    const spec: NewConnector = {
      kind,
      widthCm: Math.abs(bCm.xCm - aCm.xCm),
      heightCm: Math.abs(bCm.yCm - aCm.yCm),
      offsetXCm: leftCm - anchorCm.xCm,
      offsetYCm: topCm - anchorCm.yCm,
      start: { xCm: aCm.xCm - leftCm, yCm: aCm.yCm - topCm },
      end: { xCm: bCm.xCm - leftCm, yCm: bCm.yCm - topCm },
      startAttachment: from.attachment,
      endAttachment: to.attachment,
    };
    const before = currentCaretSelection();
    applyEdit(insertConnector(model, before, spec), before, "insert connector");
    disarmConnectorDraw();
    void Promise.resolve().then(remeasureConnectors);
  }

  function anchorTopLeftPx(): Point {
    const caret = getCaret();
    const root = contentRoot.value;
    const pageEl = pageRef.value;
    if (!caret || !root || !pageEl) {
      return { x: 0, y: 0 };
    }
    const el = root.querySelector(`[data-section="${caret.section}"][data-block="${caret.block}"]`);
    if (!el) {
      return { x: 0, y: 0 };
    }
    const pageRect = pageEl.getBoundingClientRect();
    const rect = el.getBoundingClientRect();
    const z = renderScale();
    return { x: (rect.left - pageRect.left) / z, y: (rect.top - pageRect.top) / z };
  }

  function currentCaretSelection() {
    const root = contentRoot.value as HTMLElement | null;
    return currentSelection(root) ?? caretSelection(getCaret() ?? { section: 0, block: 0, offset: 0 });
  }

  // --- selection + endpoint / midpoint handle dragging (AC-15, AC-16) -------

  const selectedConnector = computed(() => {
    const model = doc.value;
    return model && selected.value ? connectorAt(model, selected.value) : null;
  });

  /** ConnectorToolbar's display props: the model connector plus its line
   * colour resolved to a hex for the native colour input (SPEC-077 AC-7) —
   * the stored value, reference or literal, is what setOutline's spread
   * re-emits, so a width/dash edit never flattens a theme reference. */
  const connectorToolbarProps = computed((): Record<string, unknown> => {
    const connector = selectedConnector.value;
    if (!connector) {
      return {}; // never rendered: DocumentView's v-if guards on selectedConnector
    }
    const color = connector.outline?.color;
    return {
      connector,
      lineColorHex: color ? resolveColor(color, documentThemeOf(doc.value)) : "",
    };
  });

  function selectConnector(location: ConnectorLocation) {
    selected.value = location;
    remeasureConnectors();
  }

  function clearConnectorSelection() {
    if (selected.value) {
      selected.value = null;
      remeasureConnectors();
    }
  }

  function startEndpointDrag(location: ConnectorLocation, which: "start" | "end", event: MouseEvent) {
    event.preventDefault();
    event.stopPropagation();
    selected.value = location;
    endpointDrag = { location, which };
    window.addEventListener("mousemove", handleEndpointMove);
    window.addEventListener("mouseup", finishEndpointDrag);
  }

  function handleEndpointMove(event: MouseEvent) {
    if (!endpointDrag) {
      return;
    }
    updateHoverSites(event.clientX, event.clientY);
    // Live-preview the dragged endpoint by moving it to the pointer.
    const page = pagePointFromClient(event.clientX, event.clientY);
    if (!page) {
      return;
    }
    const render = renders.value.find(
      (r) => r.location.section === endpointDrag!.location.section && r.location.block === endpointDrag!.location.block,
    );
    if (render) {
      if (endpointDrag.which === "start") {
        render.a = page;
      } else {
        render.b = page;
      }
      render.d = straightPreviewD(render.a, render.b);
    }
  }

  function finishEndpointDrag(event: MouseEvent) {
    window.removeEventListener("mousemove", handleEndpointMove);
    window.removeEventListener("mouseup", finishEndpointDrag);
    const drag = endpointDrag;
    endpointDrag = null;
    hoverSites.value = null;
    if (!drag) {
      return;
    }
    const model = doc.value;
    if (!model) {
      return;
    }
    const page = pagePointFromClient(event.clientX, event.clientY);
    const hit = drawingAtClient(event.clientX, event.clientY);
    const before = currentCaretSelection();
    if (hit) {
      // Attach / re-attach to the nearest site of the dropped-on shape (AC-15).
      const site = nearestSite(hit.box, page ?? { x: 0, y: 0 });
      applyEdit(
        setConnectorEndpoint(model, before, drag.location, drag.which, {
          attachment: { target: { block: hit.location.block }, site },
        }),
        before,
        "re-attach connector",
      );
    } else if (page) {
      // Detach and leave free at the dropped point. setConnectorEndpoint's
      // `point` is section-space cm (offset + local): the connector anchor
      // origin is where section-space (0,0) paints, so subtract it.
      const origin = measureConnectorOrigin(drag.location) ?? { x: 0, y: 0 };
      applyEdit(
        setConnectorEndpoint(model, before, drag.location, drag.which, {
          point: { xCm: (page.x - origin.x) / PX_PER_CM, yCm: (page.y - origin.y) / PX_PER_CM },
        }),
        before,
        "detach connector",
      );
    }
    void Promise.resolve().then(remeasureConnectors);
  }

  // --- midpoint adjustment drag (AC-16) -------------------------------------

  let adjustDrag: { location: ConnectorLocation; a: Point; b: Point } | null = null;

  function startAdjustDrag(location: ConnectorLocation, event: MouseEvent) {
    event.preventDefault();
    event.stopPropagation();
    const render = renders.value.find((r) => r.location.section === location.section && r.location.block === location.block);
    if (!render) {
      return;
    }
    selected.value = location;
    adjustDrag = { location, a: render.a, b: render.b };
    window.addEventListener("mousemove", handleAdjustMove);
    window.addEventListener("mouseup", finishAdjustDrag);
  }

  function handleAdjustMove(event: MouseEvent) {
    if (!adjustDrag) {
      return;
    }
    const page = pagePointFromClient(event.clientX, event.clientY);
    const model = doc.value;
    if (!page || !model) {
      return;
    }
    const fraction = projectFraction(adjustDrag.a, adjustDrag.b, page);
    const before = currentCaretSelection();
    applyEdit(setConnectorAdjustment(model, before, adjustDrag.location, fraction), before, "adjust connector");
    void Promise.resolve().then(remeasureConnectors);
  }

  function finishAdjustDrag() {
    window.removeEventListener("mousemove", handleAdjustMove);
    window.removeEventListener("mouseup", finishAdjustDrag);
    adjustDrag = null;
  }

  // --- formatting commands (AC-17..21) --------------------------------------

  function applyConnectorEdit(mutate: (model: DocumentModel, before: ReturnType<typeof currentCaretSelection>) => ReturnType<typeof insertConnector>, label: string) {
    const model = doc.value;
    const location = selected.value;
    if (!model || !location) {
      return;
    }
    const before = currentCaretSelection();
    applyEdit(mutate(model, before), before, label);
    void Promise.resolve().then(remeasureConnectors);
  }

  function setKind(kind: ConnectorKind | string) {
    const location = selected.value;
    if (!location) {
      return;
    }
    applyConnectorEdit((model, before) => setConnectorKind(model, before, location, asConnectorKind(String(kind))), "connector kind");
  }

  function setOutline(outline: ShapeOutline) {
    const location = selected.value;
    if (!location) {
      return;
    }
    applyConnectorEdit((model, before) => setConnectorOutline(model, before, location, outline), "connector outline");
  }

  function setArrowheads(ends: {
    startArrow?: ArrowheadKind;
    endArrow?: ArrowheadKind;
    startArrowSize?: ArrowheadSize;
    endArrowSize?: ArrowheadSize;
  }) {
    const location = selected.value;
    if (!location) {
      return;
    }
    applyConnectorEdit((model, before) => setConnectorArrowheads(model, before, location, ends), "connector arrowheads");
  }

  function setWrap(wrap: string) {
    const location = selected.value;
    if (!location) {
      return;
    }
    applyConnectorEdit((model, before) => setConnectorWrap(model, before, location, wrap as Connector["wrap"]), "connector wrap");
  }

  function setAltText(text: string) {
    const location = selected.value;
    if (!location) {
      return;
    }
    applyConnectorEdit((model, before) => setConnectorAltText(model, before, location, text), "connector alt text");
  }

  function setZOrder(command: ZOrderCommand | string) {
    const location = selected.value;
    if (!location) {
      return;
    }
    applyConnectorEdit((model, before) => reorderShapeZ(model, before, location, command as ZOrderCommand), "connector z-order");
  }

  /** The formatting toolbar's position — centred above the selected
   * connector's endpoints, in the overlay's page-relative px frame. */
  const connectorToolbarStyle = computed(() => {
    const location = selected.value;
    const render = location
      ? renders.value.find((r) => r.location.section === location.section && r.location.block === location.block)
      : undefined;
    if (!render) {
      return { display: "none" };
    }
    const midX = (render.a.x + render.b.x) / 2;
    const topY = Math.min(render.a.y, render.b.y);
    return { left: `${midX}px`, top: `${Math.max(0, topY - 46)}px` };
  });

  /** This page's painted connectors — always the SAME empty array when it has
   * none, so a page without connectors never sees its prop change identity
   * and never re-renders because a connector moved on another page (or
   * because the page count changed). */
  function connectorsForPage(pageIndex: number): ConnectorPaint[] {
    const layer = pageLayers.value[pageIndex];
    return layer && layer.length > 0 ? layer : NO_CONNECTOR_PAINT;
  }

  function removeSelectedConnector() {
    const location = selected.value;
    if (!location) {
      return;
    }
    const model = doc.value;
    if (!model) {
      return;
    }
    const before = currentCaretSelection();
    applyEdit(deleteConnector(model, before, location), before, "delete connector");
    selected.value = null;
    void Promise.resolve().then(remeasureConnectors);
  }

  return {
    // draw mode
    connectorDrawKind,
    armConnectorDraw,
    disarmConnectorDraw,
    handleConnectorMouseDown,
    handleConnectorMouseMove,
    // overlay geometry
    connectorRenders: renders as Ref<ConnectorRender[]>,
    /** One page's painted connectors, for DocPage (SPEC-034). */
    connectorsForPage,
    hoverSites,
    drawFrom,
    drawPointer,
    remeasureConnectors,
    // selection + handles
    selected,
    selectedConnector,
    connectorToolbarProps,
    selectConnector,
    clearConnectorSelection,
    startEndpointDrag,
    startAdjustDrag,
    // formatting
    setKind,
    setOutline,
    setArrowheads,
    setWrap,
    setAltText,
    setZOrder,
    connectorToolbarStyle,
    removeSelectedConnector,
  };
}

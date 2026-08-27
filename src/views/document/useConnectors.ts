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
import { computed, ref, type Ref } from "vue";
import type { EditorContext } from "./context";
import {
  asConnectorKind,
  isConnector,
  type Connector,
  type ConnectorKind,
  type DocumentModel,
} from "../../lib/documentModel";
import type { ArrowheadKind, ShapeOutline } from "../../lib/core/shapes";
import {
  clampSite,
  connectionSitePoint,
  connectionSites,
  nearestSite,
  routeConnector,
  routedPathToSvg,
  segmentPolygonCrossing,
  type ConnectorBox,
  type ConnectorEndpoint,
} from "../../lib/core/connectors";
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

export interface ConnectorsDeps extends EditorContext {}

export function useConnectors(deps: ConnectorsDeps) {
  const { doc, contentRoot, pageRef, editable, currentSelection, applyEdit, getCaret, zoom } = deps;

  const drawKind = ref<ConnectorKind | null>(null);
  const drawFrom = ref<{ point: Point; attachment?: NewConnector["startAttachment"] } | null>(null);
  const drawPointer = ref<Point | null>(null);
  const hoverSites = ref<SiteHint | null>(null);
  const selected = ref<ConnectorLocation | null>(null);
  const renders = ref<ConnectorRender[]>([]);

  let endpointDrag: { location: ConnectorLocation; which: "start" | "end" } | null = null;

  // --- DOM measurement (page-relative px) -----------------------------------

  function pagePointFromClient(clientX: number, clientY: number): Point | null {
    const pageEl = pageRef.value;
    if (!pageEl) {
      return null;
    }
    const rect = pageEl.getBoundingClientRect();
    const z = zoom.value || 1;
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
    const geo = measureImageGeometry(pageEl, el, zoom.value);
    if (!geo) {
      return null;
    }
    return { x: geo.left, y: geo.top, w: geo.width, h: geo.height };
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
    const z = zoom.value || 1;
    return { x: (rect.left - pageRect.left) / z, y: (rect.top - pageRect.top) / z };
  }

  function rectPolygon(box: ConnectorBox): Point[] {
    return [
      { x: box.x, y: box.y },
      { x: box.x + box.w, y: box.y },
      { x: box.x + box.w, y: box.y + box.h },
      { x: box.x, y: box.y + box.h },
    ];
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
      const free = which === "start" ? connector.start : connector.end;
      return { point: { x: origin.x + free.xCm * PX_PER_CM, y: origin.y + free.yCm * PX_PER_CM } };
    }
    const box = measureDrawingBox({ section: location.section, block: attachment.target.block });
    if (!box) {
      // Target not in DOM (or a stale index): fall back to the free position.
      const origin = measureConnectorOrigin(location);
      const free = which === "start" ? connector.start : connector.end;
      return origin ? { point: { x: origin.x + free.xCm * PX_PER_CM, y: origin.y + free.yCm * PX_PER_CM } } : null;
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

  /** Re-measures every connector and rebuilds the overlay geometry. Cheap
   * enough to run on every mousemove of a shape drag (AC-7/8/9). */
  function remeasureConnectors() {
    const model = doc.value;
    if (!model) {
      renders.value = [];
      return;
    }
    const next: ConnectorRender[] = [];
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
        const guessA = rawA ? connectionSitePoint(rawA, clampSite(block.startAttachment!.site)) : { x: originGuess.x + block.start.xCm * PX_PER_CM, y: originGuess.y + block.start.yCm * PX_PER_CM };
        const guessB = rawB ? connectionSitePoint(rawB, clampSite(block.endAttachment!.site)) : { x: originGuess.x + block.end.xCm * PX_PER_CM, y: originGuess.y + block.end.yCm * PX_PER_CM };
        const a = resolveEndpointPx(location, block, "start", guessB);
        const b = resolveEndpointPx(location, block, "end", a?.point ?? guessA);
        const a2 = resolveEndpointPx(location, block, "start", b?.point ?? guessB);
        if (!a2 || !b) {
          return;
        }
        const path = routeConnector(block.kind, a2, b, block.adjustments?.[0] ?? 0.5);
        const mid = midpointOf(path.points, block.kind);
        next.push({
          location,
          key: `${sectionIndex}:${blockIndex}`,
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
    const z = zoom.value || 1;
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

  function straightPreviewD(a: Point, b: Point): string {
    return `M ${a.x} ${a.y} L ${b.x} ${b.y}`;
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

  function setArrowheads(ends: { startArrow?: ArrowheadKind; endArrow?: ArrowheadKind }) {
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
    hoverSites,
    drawFrom,
    drawPointer,
    remeasureConnectors,
    // selection + handles
    selected,
    selectedConnector,
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

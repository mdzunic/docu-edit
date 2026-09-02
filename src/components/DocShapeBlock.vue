<script setup lang="ts">
import { inject, nextTick, onMounted, onUpdated, ref, useTemplateRef, watch } from "vue";
import { asShape } from "../lib/documentModel";
import { asListDefinitions } from "../lib/core/lists";
import { drawingWrapSpacerStyle, shapeDisplayStyle, shapeTextRectStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { asTrackChangesDisplayMode } from "../lib/documentModel";
import { paintShapeIntoCanvas } from "../lib/core/shapeCanvas";
import { noShapeMediaSrc, shapeFillIsVisible } from "../lib/core/shapes";
import { shapeClipPath } from "../lib/core/shapeGeometry";
import { shapeHasVisibleText } from "../lib/shapes";
import { shapeTextRenderEntries } from "../lib/pageRender";
import { cmToPx, pxToCm } from "../lib/core/units";
import DocBlockEntries from "./DocBlockEntries.vue";

// One shape/text-box block on the page (SPEC-051). Runtime prop
// declarations, not defineProps<{...}>() generics — see DocImageBlock.vue's
// comment on the same espree constraint. Paints geometry/fill/outline on a
// canvas (CLAUDE.md's "HTML Canvas for complex shapes") and owns one
// absolutely positioned, transparent DOM proxy over its bounds — the
// element every e2e assertion, keyboard nudge and accessibility scan goes
// through ("Rendering and test surface"). The proxy's editable text is real
// DOM, not canvas-painted, since it has to be editable.
defineOptions({ inheritAttrs: false });

const props = defineProps({
  /** The Shape block (documentModel.Shape). */
  block: { type: Object, required: true },
  section: { type: Number, required: true },
  blockIndex: { type: Number, required: true },
  selected: { type: Boolean, default: false },
  /** This drawing's 0-based index in document order — data-testid="drawing-<n>". */
  drawingIndex: { type: Number, default: 0 },
  editingText: { type: Boolean, default: false },
  displayMode: { type: String, default: "all" },
  /** The document's multilevel list definitions (SPEC-032) — a shape's own
   * list paragraphs (AC-18) resolve defined markers/indents from the same
   * catalogue the body does. */
  listDefinitions: { type: Array, default: () => [] },
});

defineEmits([
  "select",
  "edit-text",
  "text-beforeinput",
  "text-keydown",
  "text-compositionstart",
  "text-compositionend",
]);

const canvasEl = useTemplateRef("canvasEl");

// SPEC-063 AC-20: the presentation-layer zoom, injected from DocumentView so
// the canvas backing store is painted at devicePixelRatio * zoom and the
// .pages transform scales it without blurring. Defaults to a static 1 when
// rendered outside the editor (unit tests, print-only trees).
const documentZoom = inject("documentZoom", ref(1));

// SPEC-073 picture fills: DocumentView (via useShapes) provides a mediaId →
// data-URI resolver; outside the editor there is none and a picture fill
// paints nothing. Decoded images are cached per mediaId; a load completion
// triggers a repaint, so the fill appears as soon as the bitmap is ready.
const shapeMediaSrc = inject("shapeMediaSrc", noShapeMediaSrc());
const pictureCache = new Map();

function resolvePicture(mediaId = "") {
  let img = pictureCache.get(mediaId);
  if (!img) {
    const src = typeof shapeMediaSrc === "function" ? shapeMediaSrc(mediaId) : "";
    if (!src) {
      return null;
    }
    img = new Image();
    img.onload = repaint;
    img.src = src;
    pictureCache.set(mediaId, img);
  }
  return img.complete && img.naturalWidth > 0 ? img : null;
}

function repaint() {
  const shape = asShape(props.block);
  const canvas = canvasEl.value;
  if (!shape || !canvas) {
    return;
  }
  const widthPx = Math.max(1, Math.round(cmToPx(shape.widthCm)));
  const heightPx = Math.max(1, Math.round(cmToPx(Math.max(shape.heightCm, 0.05))));
  paintShapeIntoCanvas(canvas, shape, widthPx, heightPx, documentZoom.value, { pictures: resolvePicture });
}

onMounted(repaint);
onUpdated(repaint);
watch(() => props.block, repaint, { deep: true });
// Re-paint the backing store when the zoom changes so it stays crisp (AC-20).
watch(documentZoom, repaint);

const textFrameEl = useTemplateRef("textFrameEl");
/** Whether the shape's text is currently clipped past its box (AC-21,
 * AC-22) — measured directly off the text frame's own scrollHeight vs
 * clientHeight, since this component is the only place that DOM exists.
 * Re-measured whenever the block (text or size) changes and after layout
 * settles post-paint, via nextTick. */
const overflowing = ref(false);

function measureOverflow() {
  const el = textFrameEl.value;
  overflowing.value = !!el && el.scrollHeight - el.clientHeight > 1;
}

onMounted(measureOverflow);
onUpdated(measureOverflow);
watch(
  () => props.block,
  () => void nextTick().then(measureOverflow),
  { deep: true },
);

const floatingWrap = ["square", "topAndBottom", "behind", "infront"];

// The text-column width the wrap spacer needs (drawingWrapSpacerStyle):
// measured off the anchor's own block container, so a shape inside a
// SPEC-030 column wraps within that column — the same measure-in-component
// pattern as measureOverflow above.
const anchorEl = useTemplateRef("anchorEl");
const columnWidthCm = ref(0);

function measureColumnWidth() {
  const parent = anchorEl.value instanceof HTMLElement ? anchorEl.value.parentElement : null;
  const cm = parent ? pxToCm(parent.clientWidth) : 0;
  if (Math.abs(cm - columnWidthCm.value) > 0.01) {
    columnWidthCm.value = cm;
  }
}

onMounted(measureColumnWidth);
onUpdated(measureColumnWidth);

/** The in-flow float that reserves this shape's text room (square /
 * topAndBottom) — null for every other wrap mode. */
function wrapSpacerStyle() {
  return drawingWrapSpacerStyle(asShape(props.block), columnWidthCm.value);
}

/** Restricts the proxy's own hit-testing to the shape's painted region
 * ("Selection and hit testing", AC-11) so a click never selects by bounding
 * box: a filled shape clips to its geometry (a click outside the clip falls
 * through natively to whatever is beneath — the page background, or a
 * shape further down in z-order); an unfilled shape with no visible text
 * has nothing in its interior to click at all, so the whole proxy goes
 * `pointer-events: none` and passes every click through — DocumentView's
 * shapeAtClientPoint recovers the outline-tolerance click (and the
 * dblclick-to-edit gesture) that this can no longer catch natively. A
 * text-frame click still lands on `.doc-shape-text`, which sets its own
 * `pointer-events: auto` regardless of the proxy's (see the scoped style
 * below — pointer-events inherits, so that override is required, not
 * cosmetic). */
function proxyStyle() {
  const shape = asShape(props.block);
  const clip = shapeClipPath(shape.geometry, shapeFillIsVisible(shape.fill), shape.path);
  // The selection frame reads its angle from here (see the stylesheet): the
  // proxy itself must stay unrotated, since it is the element every
  // measurement and hit test goes through. Untyped object literal per the
  // espree constraint this file's header notes — a type annotation here is
  // a parse error for the .vue linter.
  const style = { "--shape-rotation": `${shape.rotationDeg ?? 0}deg` };
  return clip ? { ...style, clipPath: clip } : style;
}

/** Whether the proxy has nothing in its interior to click at all — an
 * unfilled shape with no visible text (proxyStyle()'s own comment). Applied
 * as a class, not an inline style, so `pointer-events` stays a plain CSS
 * value Vue's typed `:style` binding never has to widen. */
function proxyPassthrough() {
  const shape = asShape(props.block);
  return !shapeFillIsVisible(shape.fill) && !shapeHasVisibleText(shape);
}

/** The shape's text frame as DocBlockEntries entries (AC-17..21): real,
 * editable DOM inside the proxy — paragraphs (with lists) and tables, the
 * same Block content body text accepts. Rendered whether or not this shape
 * is the one actively being edited; only `editingText` toggles whether it's
 * `contenteditable`. */
function shapeTextEntries() {
  const shape = asShape(props.block);
  if (!shape.text) {
    return [];
  }
  return shapeTextRenderEntries(shape, asListDefinitions(props.listDefinitions));
}

/** The geometry's own text rectangle (AC-11), or an empty object when the
 * geometry fills its box — the scoped `.doc-shape-text { inset: 0 }` then
 * applies unchanged. */
function textRectStyle() {
  return shapeTextRectStyle(asShape(props.block)) ?? {};
}
</script>

<template>
  <span
    ref="anchorEl"
    class="doc-shape-anchor"
    :class="{ 'doc-shape-anchor-floating': floatingWrap.includes(block.wrap) }"
    :data-section="section"
    :data-block="blockIndex"
  >
    <!-- The wrap spacer: an invisible float in the TEXT flow that the
         surrounding lines yield to (square / topAndBottom). The shape span
         below stays an absolute overlay; this is what makes its wrap mode
         real rather than decorative. -->
    <span
      v-if="wrapSpacerStyle() && !trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-shape-wrap-spacer"
      data-testid="shape-wrap-spacer"
      aria-hidden="true"
      :style="wrapSpacerStyle()"
    />
    <span
      v-show="!trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-shape"
      :class="{ 'doc-shape-selected': selected, 'doc-shape-passthrough': proxyPassthrough() }"
      :style="shapeDisplayStyle(asShape(block))"
    >
      <canvas
        ref="canvasEl"
        class="doc-shape-canvas"
        aria-hidden="true"
      />
      <span
        class="doc-shape-proxy"
        :class="{ 'doc-shape-proxy-passthrough': proxyPassthrough() }"
        :style="proxyStyle()"
        :data-testid="`drawing-${drawingIndex}`"
        :role="block.text ? 'group' : 'img'"
        :aria-label="block.altText || ''"
        tabindex="0"
        @click.stop="(event) => $emit('select', event.shiftKey)"
        @dblclick.stop="$emit('edit-text')"
      >
        <!-- data-canvas-ignore (same as header/footer bands and note areas):
             this frame's paragraphs carry synthetic section-0 addressing that
             must never shadow the body's own blocks in domCaret resolution;
             the shape's own editing uses this frame as its root (SPEC-051). -->
        <div
          v-if="block.text"
          ref="textFrameEl"
          class="doc-shape-text"
          data-canvas-ignore="1"
          :style="textRectStyle()"
          :class="`doc-shape-valign-${block.text.vAlign}`"
          data-testid="shape-text"
          :data-shape-text-editable="editingText ? '1' : undefined"
          :contenteditable="editingText"
          spellcheck="false"
          @beforeinput.stop="editingText ? $emit('text-beforeinput', $event) : undefined"
          @keydown.stop="editingText ? $emit('text-keydown', $event) : undefined"
          @compositionstart.stop="editingText ? $emit('text-compositionstart') : undefined"
          @compositionend.stop="editingText ? $emit('text-compositionend', $event) : undefined"
        >
          <DocBlockEntries :entries="shapeTextEntries()" />
        </div>
        <span
          v-if="selected && overflowing"
          class="doc-shape-overflow print-exclude"
          data-testid="shape-overflow-indicator"
          title="Text does not fit"
        >&hellip;</span>
      </span>
    </span>
  </span>
</template>

<style scoped>
.doc-shape-anchor {
  display: inline-block;
  line-height: 0;
}

/* SPEC-078's zero-footprint floating anchor: hoistFloatingDrawingEntries
   renders it immediately before its anchor paragraph, so this box's
   top-left IS the paragraph's top-left — the origin the .doc-shape child's
   `left`/`top` (offsetXCm/offsetYCm) measure from, the same origin every
   command measures (the old inline-block anchor added a baseline's worth
   of drift — the SPEC-051 "lands ~14px below" bug).

   A zero-height BLOCK, not an inline-block and not position:absolute, for
   two reasons. An inline-block is its own block formatting context, which
   would trap the wrap spacer's float inside a zero-size box where no text
   could yield to it — a block lets the float escape into the page's own
   text flow. And a block box is immune to floats (only line boxes yield to
   them), so the anchor's top-left stays pinned at the paragraph's top-left
   even when its own spacer float would have pushed an inline anchor
   sideways. */
.doc-shape-anchor-floating {
  display: block;
  height: 0;
  position: relative;
}

.doc-shape-wrap-spacer {
  pointer-events: none;
}

.doc-shape {
  display: block;
  /* The containing block for the absolutely positioned canvas below in EVERY
     wrap mode. Floating wraps override this with an inline position:absolute
     (shapeDisplayStyle) — same containment; but an inline-with-text shape
     used to leave this box static, so the canvas escaped to the nearest
     positioned ancestor and painted the geometry at the page's top-left
     corner, nowhere near the proxy the handles are drawn from. */
  position: relative;
  cursor: pointer;
}

.doc-shape-canvas {
  display: block;
  position: absolute;
  top: 0;
  left: 0;
  pointer-events: none;
}

.doc-shape-proxy {
  display: block;
  position: relative;
  width: 100%;
  height: 100%;
  background: transparent;
  outline: none;
}

/* The selection frame turns with the shape, exactly as its resize and
   rotation handles do (useShapes.ts's shapeHandleStyle) — an upright box
   drawn round a shape lying at 30° traces an edge that is not there. It
   rides on a pseudo-element rather than on the proxy's own outline because
   the proxy is the measured, hit-tested and e2e-asserted surface: a
   transform on it would turn every getBoundingClientRect into the bounding
   box of a rotated rectangle. A descendant's transform leaves its parent's
   own border box alone. The -3px inset and 2px border reproduce the
   `outline: 2px` at `outline-offset: 1px` this replaced. */
.doc-shape-selected .doc-shape-proxy::after {
  content: "";
  position: absolute;
  inset: -3px;
  border: 2px solid var(--accent);
  transform: rotate(var(--shape-rotation, 0deg));
  pointer-events: none;
}

/* An unfilled, textless shape's proxy passes every click through to
   whatever's beneath (AC-11) — DocumentView's shapeAtClientPoint recovers
   the outline-tolerance click and dblclick-to-edit gesture natively lost
   here (proxyPassthrough()'s own comment). The `.doc-shape` box itself must
   pass through too: it is the proxy's parent and, left pointer-events: auto,
   swallows every interior click so the recovery (which alone runs the
   outline-precise hit test) only fires for a click that misses the box
   entirely — leaving an open Freeform/Curve/Scribble stroke selectable only
   at its extreme vertices, never along its body (SPEC-071 AC-24). A passthrough
   shape has no text frame (proxyPassthrough requires no visible text), so
   nothing inside needs to stay clickable. */
.doc-shape-proxy-passthrough,
.doc-shape-passthrough {
  pointer-events: none;
}

.doc-shape-text {
  display: flex;
  flex-direction: column;
  position: absolute;
  inset: 0;
  padding: 0.1cm 0.2cm;
  overflow: hidden;
  white-space: pre-wrap;
  word-wrap: break-word;
  cursor: text;
  /* .doc-shape-anchor resets line-height: 0 (so its own inline-block line
     box adds no gutter) — that inherits all the way down into paragraphs
     with no explicit lineSpacing (renderFormatting's paragraphStyle only
     sets an inline line-height when the model has one), collapsing every
     line to zero height. Reset back to normal here so shape text lays out
     like body text; an explicit per-paragraph lineHeight still overrides
     this via its own inline style regardless. */
  line-height: normal;
  /* pointer-events inherits — this keeps the text frame clickable even when
     the parent proxy goes pointer-events: none for an unfilled, textless
     shape's *other* siblings' geometry (proxyStyle()'s own comment). */
  pointer-events: auto;
}

.doc-shape-valign-top {
  justify-content: flex-start;
}

.doc-shape-valign-middle {
  justify-content: center;
}

.doc-shape-valign-bottom {
  justify-content: flex-end;
}

.doc-shape-overflow {
  position: absolute;
  right: 2px;
  bottom: 0;
  background: var(--danger);
  color: var(--text-on-brand);
  font-size: 10px;
  line-height: 1;
  padding: 1px 3px;
  border-radius: 2px;
  pointer-events: none;
}
</style>

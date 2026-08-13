<script setup lang="ts">
import { onMounted, onUpdated, useTemplateRef, watch } from "vue";
import { asListDefinitions, asShape } from "../lib/documentModel";
import { shapeDisplayStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { asTrackChangesDisplayMode } from "../lib/documentModel";
import { paintShape } from "../lib/shapeCanvas";
import { shapeClipPath } from "../lib/shapeGeometry";
import { shapeHasVisibleText } from "../lib/shapes";
import { shapeTextRenderEntries } from "../lib/pageRender";
import { cmToPx } from "../lib/images";
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
  /** Whether this shape's text is currently clipped past its box (AC-21,
   * AC-22) — measured by the caller, which has access to the rendered text
   * frame's scrollHeight. */
  overflowing: { type: Boolean, default: false },
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

function repaint() {
  const shape = asShape(props.block);
  const canvas = canvasEl.value;
  if (!shape || !canvas) {
    return;
  }
  const widthPx = Math.max(1, Math.round(cmToPx(shape.widthCm)));
  const heightPx = Math.max(1, Math.round(cmToPx(Math.max(shape.heightCm, 0.05))));
  const dpr = typeof window !== "undefined" ? window.devicePixelRatio || 1 : 1;
  canvas.width = widthPx * dpr;
  canvas.height = heightPx * dpr;
  canvas.style.width = `${widthPx}px`;
  canvas.style.height = `${heightPx}px`;
  const ctx = canvas.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  ctx.clearRect(0, 0, widthPx, heightPx);
  paintShape(ctx, shape, widthPx, heightPx);
}

onMounted(repaint);
onUpdated(repaint);
watch(() => props.block, repaint, { deep: true });

const floatingWrap = ["square", "topAndBottom", "behind", "infront"];

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
  const clip = shapeClipPath(shape.geometry, !!shape.fill);
  return clip ? { clipPath: clip } : {};
}

/** Whether the proxy has nothing in its interior to click at all — an
 * unfilled shape with no visible text (proxyStyle()'s own comment). Applied
 * as a class, not an inline style, so `pointer-events` stays a plain CSS
 * value Vue's typed `:style` binding never has to widen. */
function proxyPassthrough() {
  const shape = asShape(props.block);
  return !shape.fill && !shapeHasVisibleText(shape);
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
</script>

<template>
  <span
    class="doc-shape-anchor"
    :class="{ 'doc-shape-anchor-floating': floatingWrap.includes(block.wrap) }"
    :data-section="section"
    :data-block="blockIndex"
  >
    <span
      v-show="!trackChangeBlockHidden(block.changes, asTrackChangesDisplayMode(displayMode))"
      class="doc-shape"
      :class="{ 'doc-shape-selected': selected }"
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
          class="doc-shape-text"
          data-canvas-ignore="1"
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

.doc-shape-anchor-floating {
  position: relative;
}

.doc-shape {
  display: block;
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

.doc-shape-selected .doc-shape-proxy {
  outline: 2px solid #1a73e8;
  outline-offset: 1px;
}

/* An unfilled, textless shape's proxy passes every click through to
   whatever's beneath (AC-11) — DocumentView's shapeAtClientPoint recovers
   the outline-tolerance click and dblclick-to-edit gesture natively lost
   here (proxyPassthrough()'s own comment). */
.doc-shape-proxy-passthrough {
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
  background: #d93025;
  color: #fff;
  font-size: 10px;
  line-height: 1;
  padding: 1px 3px;
  border-radius: 2px;
  pointer-events: none;
}
</style>

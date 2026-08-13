<script setup lang="ts">
import { onMounted, onUpdated, useTemplateRef, watch } from "vue";
import { asShape } from "../lib/documentModel";
import { shapeDisplayStyle, trackChangeBlockHidden } from "../lib/renderFormatting";
import { asTrackChangesDisplayMode } from "../lib/documentModel";
import { paintShape } from "../lib/shapeCanvas";
import { shapeClipPath } from "../lib/shapeGeometry";
import { cmToPx } from "../lib/images";

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
});

const emit = defineEmits(["select", "edit-text", "text-input"]);

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

/** Restricts the proxy's own hit-testing to the shape's painted region for
 * a filled shape ("Selection and hit testing", AC-11) — a click outside the
 * clip falls through to whatever is beneath (the page background, or a
 * shape further down in z-order). Text-frame clicks still land on the text
 * span itself, which sits inside the proxy and is unaffected by its
 * parent's clip. */
function proxyStyle() {
  const shape = asShape(props.block);
  const clip = shapeClipPath(shape.geometry, !!shape.fill);
  return clip ? { clipPath: clip } : {};
}

function onTextInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLElement) {
    emit("text-input", target.innerHTML);
  }
}

/** A plain-text preview of the shape's text frame — every paragraph's runs
 * concatenated, one paragraph per line. Not yet the full rich-text render
 * (see the implement summary): character formatting inside a shape's text
 * currently displays unstyled. */
function shapeTextPreview() {
  const shape = asShape(props.block);
  if (!shape.text) {
    return "";
  }
  return shape.text.blocks
    .map((paragraphBlock) => {
      const paragraph = paragraphBlock.type === "paragraph" ? paragraphBlock : null;
      return paragraph ? paragraph.runs.map((run) => run.text).join("") : "";
    })
    .join("\n");
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
        :style="proxyStyle()"
        :data-testid="`drawing-${drawingIndex}`"
        :role="block.text ? 'group' : 'img'"
        :aria-label="block.altText || ''"
        tabindex="0"
        @click.stop="$emit('select')"
        @dblclick.stop="$emit('edit-text')"
      >
        <span
          v-if="block.text"
          class="doc-shape-text"
          :class="`doc-shape-valign-${block.text.vAlign}`"
          data-testid="shape-text"
          :contenteditable="editingText"
          @input="onTextInput"
        >{{ shapeTextPreview() }}</span>
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

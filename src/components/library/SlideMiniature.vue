<script setup lang="ts">
// A presentation's first-slide miniature for the library (card and quick
// look): the slide surface at its natural size, scaled to the host's width,
// with a lightweight renderer over the same pure helpers the editor's stage
// uses (slideRender) — text with its resolved runs, shapes and charts through
// the shared canvas painters, images from the deck's own media. Deliberately
// NOT SlideFrameView: no selection, editing, animation or overlay machinery —
// paint only. Every colour here is document content (the deck's own values),
// carried inline; the <style> block stays token-clean.
// Runtime prop declarations — the .vue ESLint parser rejects TS-only syntax
// (see AppShell.vue).
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import {
  chartVisualOf,
  imageSrc,
  isEmptyTextPlaceholder,
  isForeignVisible,
  markParagraphList,
  markedParagraphs,
  mediaResolver,
  shapeVisualOf,
  surfaceBackground,
  textFrameFill,
} from "../presentation/slideRender";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import { paintChart } from "../../lib/core/chartCanvas";
import { cmToPx } from "../../lib/core/units";
import { resolveThemeColor } from "../../composables/useTheme";

const props = defineProps({
  /** The PresentationModel whose first slide previews. */
  presentation: { type: Object, required: true },
});

const slide = computed(() => props.presentation.slides[0] ?? { frames: [] });
const naturalWidth = computed(() => Math.max(1, cmToPx(props.presentation.slideWidthCm)));
const naturalHeight = computed(() => Math.max(1, cmToPx(props.presentation.slideHeightCm)));
const aspect = computed(() => naturalWidth.value / naturalHeight.value);
const resolveMedia = computed(() => mediaResolver(props.presentation.media ?? []));

// Group children carry slide-space coordinates already, so the paint list is
// the frames flattened in z-order.
const paintList = computed(() => {
  // Object([]) types the list `any` — the espree/StyleValue constraint this
  // file's style helpers note.
  const out = Object([]);
  const walk = (frames = Object([])) => {
    for (const frame of frames) {
      if (frame.kind === "group") {
        walk(frame.children ?? []);
      } else {
        out.push(frame);
      }
    }
  };
  walk(slide.value.frames ?? []);
  return out;
});

// The surface renders at natural px and scales to the host width — the same
// transform approach as the editor's filmstrip, so text stays proportionate.
const hostRef = useTemplateRef("host");
const scale = ref(0);
let observer = { disconnect() {} };

function measure() {
  const width = hostRef.value?.clientWidth ?? 0;
  if (width > 0) {
    scale.value = width / naturalWidth.value;
  }
}

onMounted(() => {
  measure();
  if (typeof ResizeObserver === "function" && hostRef.value) {
    const resizeObserver = new ResizeObserver(() => measure());
    resizeObserver.observe(hostRef.value);
    observer = resizeObserver;
  }
});

onBeforeUnmount(() => observer.disconnect());

// Style objects are built on Object() (any) — vue's StyleValue rejects the
// inferred literal types and annotations are barred in .vue scripts.
function frameStyle(frame = Object()) {
  const style = Object();
  style.position = "absolute";
  style.left = `${cmToPx(frame.xCm)}px`;
  style.top = `${cmToPx(frame.yCm)}px`;
  style.width = `${cmToPx(frame.widthCm)}px`;
  style.height = `${cmToPx(frame.heightCm)}px`;
  if (frame.rotationDeg) {
    style.transform = `rotate(${frame.rotationDeg}deg)`;
  }
  return style;
}

function textStyle(frame = Object()) {
  const style = framePositionOnly(frame);
  style.background = textFrameFill(frame);
  if (frame.stroke) {
    style.border = `${(frame.stroke.widthPt * 96) / 72}px solid ${frame.stroke.color}`;
  }
  return style;
}

function framePositionOnly(frame = Object()) {
  return frameStyle(frame);
}

function runStyle(run = Object()) {
  const style = Object();
  style.color = run.color || "#000000";
  style.fontSize = `${((run.sizePt || 18) * 96) / 72}px`;
  if (run.bold) {
    style.fontWeight = "700";
  }
  if (run.italic) {
    style.fontStyle = "italic";
  }
  if (run.underline) {
    style.textDecoration = "underline";
  }
  if (run.fontFamily) {
    style.fontFamily = run.fontFamily;
  }
  return style;
}

function paragraphStyle(entry = Object()) {
  const style = Object();
  style.textAlign = entry.para.align || "left";
  style.paddingInlineStart = `${(entry.para.level || 0) * 1.2}em`;
  return style;
}

/** Function ref for a shape frame's canvas — the shared core painter, guarded
 * for hosts without a real 2D context (jsdom). */
function paintShapeFrame(el = Object(), frame = Object()) {
  if (!(el instanceof HTMLCanvasElement)) {
    return;
  }
  const visual = shapeVisualOf(frame);
  if (!visual) {
    return;
  }
  try {
    paintShapeIntoCanvas(el, visual, cmToPx(frame.widthCm), cmToPx(frame.heightCm));
  } catch {
    // Paint is best-effort; the miniature must never break the library.
  }
}

/** Function ref for a chart frame's canvas — the shared chart painter with
 * the theme's chart chrome, as the document thumbnails paint. */
function paintChartFrame(el = Object(), frame = Object()) {
  if (!(el instanceof HTMLCanvasElement)) {
    return;
  }
  const visual = chartVisualOf(frame);
  if (!visual) {
    return;
  }
  const widthPx = Math.max(1, Math.round(cmToPx(frame.widthCm)));
  const heightPx = Math.max(1, Math.round(cmToPx(frame.heightCm)));
  el.width = widthPx;
  el.height = heightPx;
  try {
    const ctx = el.getContext("2d");
    if (!ctx) {
      return;
    }
    ctx.clearRect(0, 0, widthPx, heightPx);
    paintChart(ctx, visual, widthPx, heightPx, {
      axis: resolveThemeColor("--chart-axis"),
      gridline: resolveThemeColor("--chart-gridline"),
      text: resolveThemeColor("--chart-text"),
    });
  } catch {
    // Best-effort, as above.
  }
}

/** A table cell's paragraphs with their bullet markers. */
function markParagraphsOf(cell = Object()) {
  return markParagraphList(Array.isArray(cell.paragraphs) ? cell.paragraphs : []);
}

// Model-content greys for the frames the miniature cannot paint faithfully
// (table grids, foreign objects) — inline data, not chrome styling.
const TABLE_LINE = "#b9b6b4";
const FOREIGN_FILL = "#e9e7e6";
</script>

<template>
  <div
    ref="host"
    class="slide-miniature"
    data-testid="slide-miniature"
    :style="{ aspectRatio: String(aspect) }"
  >
    <div
      class="slide-surface"
      :style="{
        width: `${naturalWidth}px`,
        height: `${naturalHeight}px`,
        transform: `scale(${scale})`,
        background: surfaceBackground(slide),
      }"
    >
      <template
        v-for="frame in paintList"
        :key="frame.id"
      >
        <!-- Text: resolved runs, markers, alignment. Empty placeholders show
             nothing — the miniature never paints the editor's prompts. -->
        <div
          v-if="frame.kind === 'text' && !isEmptyTextPlaceholder(frame)"
          class="mini-text"
          :style="textStyle(frame)"
        >
          <div
            v-for="(entry, index) in markedParagraphs(frame)"
            :key="index"
            class="mini-paragraph"
            :style="paragraphStyle(entry)"
          >
            <span
              v-if="entry.marker"
              class="mini-marker"
            >{{ entry.marker }}</span>
            <span
              v-for="(run, runIndex) in entry.para.runs"
              :key="runIndex"
              :style="runStyle(run)"
            >{{ run.text }}</span>
          </div>
        </div>

        <div
          v-else-if="frame.kind === 'shape'"
          class="mini-shape"
          :style="framePositionOnly(frame)"
        >
          <canvas :ref="(el) => paintShapeFrame(el, frame)" />
          <div
            v-if="frame.paragraphs && frame.paragraphs.length"
            class="mini-shape-text"
          >
            <div
              v-for="(entry, index) in markedParagraphs(frame)"
              :key="index"
              :style="paragraphStyle(entry)"
            >
              <span
                v-for="(run, runIndex) in entry.para.runs"
                :key="runIndex"
                :style="runStyle(run)"
              >{{ run.text }}</span>
            </div>
          </div>
        </div>

        <img
          v-else-if="frame.kind === 'image'"
          class="mini-image"
          :src="imageSrc(frame, resolveMedia)"
          :style="framePositionOnly(frame)"
          alt=""
        >

        <canvas
          v-else-if="frame.kind === 'chart'"
          class="mini-chart"
          :ref="(el) => paintChartFrame(el, frame)"
          :style="framePositionOnly(frame)"
        />

        <!-- Tables preview as their grid; cell text at run size. -->
        <div
          v-else-if="frame.kind === 'table'"
          class="mini-table"
          :style="{ ...framePositionOnly(frame), borderColor: TABLE_LINE }"
        >
          <div
            v-for="(row, rowIndex) in frame.rows"
            :key="rowIndex"
            class="mini-table-row"
          >
            <div
              v-for="(cell, cellIndex) in row.cells"
              :key="cellIndex"
              class="mini-table-cell"
              :style="{
                borderColor: TABLE_LINE,
                background:
                  cell.fill && cell.fill.type === 'solid' ? cell.fill.color : 'transparent',
              }"
            >
              <template v-if="!cell.covered">
                <div
                  v-for="(entry, index) in markParagraphsOf(cell)"
                  :key="index"
                >
                  <span
                    v-for="(run, runIndex) in entry.para.runs"
                    :key="runIndex"
                    :style="runStyle(run)"
                  >{{ run.text }}</span>
                </div>
              </template>
            </div>
          </div>
        </div>

        <div
          v-else-if="frame.kind === 'foreign' && isForeignVisible(frame)"
          class="mini-foreign"
          :style="{ ...framePositionOnly(frame), background: FOREIGN_FILL }"
        />
      </template>
    </div>
  </div>
</template>

<style scoped>
.slide-miniature {
  position: relative;
  width: 100%;
  overflow: hidden;
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.slide-surface {
  position: absolute;
  top: 0;
  left: 0;
  transform-origin: top left;
  overflow: hidden;
}

.mini-text,
.mini-shape,
.mini-table,
.mini-foreign {
  overflow: hidden;
}

.mini-paragraph {
  line-height: var(--line-tight);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.mini-marker {
  margin-inline-end: 0.35em;
}

.mini-shape canvas {
  position: absolute;
  top: 0;
  left: 0;
}

.mini-shape-text {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  justify-content: center;
  text-align: center;
  overflow: hidden;
}

.mini-image {
  object-fit: contain;
}

.mini-table {
  display: flex;
  flex-direction: column;
  border-width: var(--border-width);
  border-style: solid;
}

.mini-table-row {
  display: flex;
  flex: 1 1 0;
  min-height: 0;
}

.mini-table-cell {
  flex: 1 1 0;
  min-width: 0;
  overflow: hidden;
  border-width: var(--border-width);
  border-style: solid;
  line-height: var(--line-tight);
}
</style>

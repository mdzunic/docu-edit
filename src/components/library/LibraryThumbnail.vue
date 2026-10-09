<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import { getDocument, getMedia, getPresentation } from "../../lib/documentStore";
import SlideMiniature from "./SlideMiniature.vue";
import { runWhenIdle } from "../../lib/core/idleScheduler";
import {
  getCachedThumbnail,
  renderThumbnail,
  THUMBNAIL_ASPECT_RATIO,
} from "../../lib/libraryThumbnail";
import { kindOfFormat } from "../../lib/documentKinds";
import {
  cacheMediaThumbnail,
  getCachedMediaThumbnail,
  MEDIA_THUMBNAIL_MAX_EDGE_PX,
  mediaImageDataUri,
  pdfCardPlaceholder,
  thumbnailRasterSize,
} from "../../lib/media/thumbnail";
import { mediaSourceBytes } from "../../lib/media/download";
import { pdfRasterSupported, rasterisePdfFirstPage } from "../../lib/pdfRaster";
import { formatFileSize, formatBadgeLabel } from "../../lib/format";
import { paintChart } from "../../lib/core/chartCanvas";
import { paintShape } from "../../lib/core/shapeCanvas";
import { rotatedBounds } from "../../lib/core/shapeGeometry";
import { cmToPx } from "../../lib/core/units";
import { resolveThemeColor, useTheme } from "../../composables/useTheme";

// A document's first-page thumbnail (SPEC-067). It NEVER blocks the library:
// the card renders immediately with a paper-shaped skeleton, and the real
// preview fills in only once the card scrolls into view (IntersectionObserver
// gating, AC-67/68/71) and the browser is idle (SPEC-047 idle scheduler). A
// cached thumbnail (produced once per edit, keyed by id + lastModified) shows
// on the next mount without re-rendering. Runtime prop declarations — the .vue
// ESLint parser rejects TS-only syntax (see AppShell.vue).
const props = defineProps({
  id: { type: String, required: true },
  lastModified: { type: Number, required: true },
  format: { type: String, required: true },
});

const rootRef = useTemplateRef("root");
const thumbnail = ref(getCachedThumbnail(props.id, props.lastModified) ?? null);

// --- media cards (SPEC-107) ------------------------------------------------
// An image card shows a REAL downscaled raster of the image (built once per
// record under the same IntersectionObserver + idle gating as text
// thumbnails, cached by id + lastModified, capped at
// MEDIA_THUMBNAIL_MAX_EDGE_PX). A PDF card shows a real raster of its first
// page too (lib/media/pdfRaster — pdf.js, the one sanctioned dependency),
// falling back to the paper-shaped placeholder (real aspect ratio, page
// count, size) wherever rasterisation is unavailable or fails.
const isMediaEntry = computed(() => kindOfFormat(props.format) === "media");

// --- presentation cards ----------------------------------------------------
// A deck previews as its real first slide (SlideMiniature) under the same
// observe-then-idle gating; the model renders directly — no PNG cache, the
// miniature is light.
const isPresentationEntry = computed(() => kindOfFormat(props.format) === "presentation");
// Seeded from a miss so the ref's inferred type is PresentationModel | null
// (an annotation is barred in .vue scripts).
const presentationModel = ref(getPresentation("") ?? null);
const presentationAspect = computed(() => {
  const model = presentationModel.value;
  if (model && model.slideWidthCm > 0 && model.slideHeightCm > 0) {
    return model.slideWidthCm / model.slideHeightCm;
  }
  return 16 / 9;
});
const mediaRaster = ref(getCachedMediaThumbnail(props.id, props.lastModified) ?? "");
// Seeded null-ish so the inferred type carries the placeholder shape.
const pdfInfo = ref({ aspectRatio: 0, pageCount: 0, pageSize: "" });
const mediaByteLength = ref(0);

function produceMediaThumbnail() {
  const model = getMedia(props.id);
  if (!model) {
    thumbnail.value = { lines: [], failed: true };
    return;
  }
  mediaByteLength.value = model.byteLength;
  if (model.format === "pdf") {
    if (!model.pdf) {
      thumbnail.value = { lines: [], failed: true };
      return;
    }
    // The aspect-true paper placeholder (SPEC-107: the first page's real
    // aspect ratio, page count and page size) shows at once. A real
    // first-page raster (pdf.js, loaded lazily) replaces it where the host
    // can rasterise; elsewhere — and on any parse/render failure — the
    // placeholder stays: the card degrades, it never lies.
    pdfInfo.value = pdfCardPlaceholder(model.pdf);
    if (!pdfRasterSupported()) {
      return;
    }
    rasterisePdfFirstPage(mediaSourceBytes(model), MEDIA_THUMBNAIL_MAX_EDGE_PX)
      .then((dataUrl) => {
        if (unmounted) {
          return;
        }
        cacheMediaThumbnail(props.id, props.lastModified, dataUrl);
        mediaRaster.value = dataUrl;
      })
      .catch(() => {
        // The placeholder already shows.
      });
    return;
  }
  const src = mediaImageDataUri(model);
  if (!src || typeof Image !== "function") {
    thumbnail.value = { lines: [], failed: true };
    return;
  }
  const image = new Image();
  image.addEventListener("load", () => {
    if (unmounted) {
      return;
    }
    try {
      const size = thumbnailRasterSize(
        image.naturalWidth || model.image?.widthPx || 1,
        image.naturalHeight || model.image?.heightPx || 1,
      );
      const canvas = document.createElement("canvas");
      canvas.width = size.widthPx;
      canvas.height = size.heightPx;
      const ctx = canvas.getContext("2d");
      if (!ctx) {
        throw new Error("no 2d context");
      }
      ctx.drawImage(image, 0, 0, size.widthPx, size.heightPx);
      const dataUrl = canvas.toDataURL("image/png");
      cacheMediaThumbnail(props.id, props.lastModified, dataUrl);
      mediaRaster.value = dataUrl;
    } catch {
      // jsdom / a broken image: the honest fallback is the format badge.
      thumbnail.value = { lines: [], failed: true };
    }
  });
  image.addEventListener("error", () => {
    if (!unmounted) {
      thumbnail.value = { lines: [], failed: true };
    }
  });
  image.src = src;
}

// A disconnectable stub as the initial value gives `observer` an inferred
// { disconnect(): void } type without a TS annotation (barred in .vue scripts)
// — a real IntersectionObserver is assignable to it.
let observer = { disconnect() {} };

// runWhenIdle is not itself cancellable, so a scheduled render can fire up to
// 500ms after the card has scrolled out and unmounted — during whatever route
// the user navigated to next. This guard drops that orphaned work so the
// library never spends idle time rendering a thumbnail for a gone card.
let unmounted = false;

function produce() {
  if (thumbnail.value || mediaRaster.value || presentationModel.value) {
    return;
  }
  runWhenIdle(() => {
    if (unmounted || thumbnail.value || mediaRaster.value || presentationModel.value) {
      return;
    }
    if (isMediaEntry.value) {
      produceMediaThumbnail();
      return;
    }
    if (isPresentationEntry.value) {
      const presentation = getPresentation(props.id);
      if (presentation) {
        presentationModel.value = presentation;
      } else {
        thumbnail.value = { lines: [], failed: true };
      }
      return;
    }
    const model = getDocument(props.id);
    thumbnail.value = model
      ? renderThumbnail(props.id, props.lastModified, model)
      : { lines: [], failed: true };
  }, 500);
}

onMounted(() => {
  // A PDF placeholder is derived synchronously from the stored summary, but
  // still goes through the same observe-then-idle path as everything else.
  if (thumbnail.value || mediaRaster.value) {
    return;
  }
  // No IntersectionObserver (jsdom, older engines): render on the next idle
  // tick rather than never — correctness over the perf optimisation.
  if (typeof IntersectionObserver !== "function" || !rootRef.value) {
    produce();
    return;
  }
  const io = new IntersectionObserver((entries) => {
    if (entries.some((entry) => entry.isIntersecting)) {
      io.disconnect();
      produce();
    }
  });
  io.observe(rootRef.value);
  observer = io;
});

onBeforeUnmount(() => {
  unmounted = true;
  observer.disconnect();
});

// A theme switch must repaint chart canvases with freshly resolved chrome
// colours — the canvases carry the theme in their :key, so flipping it
// remounts them and re-runs the paint ref callback below.
const { resolvedTheme } = useTheme();

/** Function ref for a chart line's canvas: paints the (cloned) chart block
 * through the same core/chartCanvas painter the page uses, at the chart's
 * model px size — the canvas's CSS box (a fraction of the miniature) then
 * scales it down crisply. jsdom has no 2D context, so a null ctx is a no-op. */
// Default expressions (not annotations) give the params inferred types —
// the espree constraint AppShell.vue documents; Object() types them `any`,
// wide enough for the template ref's Element | component | null.
function paintThumbnailChart(el = Object(), chart = Object()) {
  if (!(el instanceof HTMLCanvasElement)) {
    return;
  }
  const widthPx = Math.max(1, Math.round(cmToPx(chart.widthCm)));
  const heightPx = Math.max(1, Math.round(cmToPx(Math.max(chart.heightCm, 0.05))));
  el.width = widthPx;
  el.height = heightPx;
  const ctx = el.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.clearRect(0, 0, widthPx, heightPx);
  paintChart(ctx, chart, widthPx, heightPx, {
    axis: resolveThemeColor("--chart-axis"),
    gridline: resolveThemeColor("--chart-gridline"),
    text: resolveThemeColor("--chart-text"),
  });
}

/** grid-template-columns for a table line — the table's own column
 * proportions as fr units. */
function tableColumns(table = Object()) {
  return table.columnFractions.map((fraction = 0) => `${fraction.toFixed(4)}fr`).join(" ");
}

/** Function ref for a shape line's canvas: paints the stripped ShapeVisual
 * through the same core/shapeCanvas painter the page uses, at the shape's
 * model px, on a surface grown to the rotated bounds (mirroring
 * paintShapeIntoCanvas) so a rotated shape's corners are not clipped. The
 * canvas is percent-positioned against the shape-box wrapper, which the
 * miniature sizes like the page does. Colours are document data (plus the
 * painter's own placeholder grey), so no theme resolve is needed. */
function paintThumbnailShape(el = Object(), shapeLine = Object()) {
  if (!(el instanceof HTMLCanvasElement)) {
    return;
  }
  const widthPx = shapeLine.widthPx;
  const heightPx = shapeLine.heightPx;
  const bounds = rotatedBounds(widthPx, heightPx, shapeLine.shape.rotationDeg ?? 0);
  const surfaceW = Math.max(1, Math.ceil(bounds.width));
  const surfaceH = Math.max(1, Math.ceil(bounds.height));
  el.width = surfaceW;
  el.height = surfaceH;
  el.style.width = `${(surfaceW / widthPx) * 100}%`;
  el.style.left = `${(-bounds.padX / widthPx) * 100}%`;
  el.style.top = `${(-bounds.padY / heightPx) * 100}%`;
  el.style.aspectRatio = String(surfaceW / surfaceH);
  const ctx = el.getContext("2d");
  if (!ctx) {
    return;
  }
  ctx.clearRect(0, 0, surfaceW, surfaceH);
  ctx.translate(bounds.padX, bounds.padY);
  paintShape(ctx, shapeLine.shape, widthPx, heightPx);
}
</script>

<template>
  <div
    ref="root"
    class="thumbnail"
    :style="{
      aspectRatio: isPresentationEntry ? String(presentationAspect) : String(THUMBNAIL_ASPECT_RATIO),
    }"
    data-testid="row-thumbnail"
    aria-hidden="true"
  >
    <!-- A presentation card: the deck's real first slide. -->
    <div
      v-if="presentationModel"
      class="thumbnail-deck"
    >
      <SlideMiniature :presentation="presentationModel" />
    </div>
    <!-- A media image card: the real downscaled raster (SPEC-107). -->
    <div
      v-else-if="mediaRaster"
      class="thumbnail-media"
    >
      <img
        class="thumbnail-media-image"
        data-testid="media-thumbnail"
        :src="mediaRaster"
        alt=""
        decoding="async"
      >
    </div>
    <!-- A PDF card: a paper-shaped placeholder at the first page's REAL
         aspect ratio, with the page count and size — shown until (and
         unless) the pdf.js raster above arrives. -->
    <div
      v-else-if="pdfInfo.pageCount > 0"
      class="thumbnail-pdf"
      data-testid="pdf-placeholder"
    >
      <div
        class="thumbnail-pdf-paper"
        data-testid="pdf-paper"
        :style="{ '--paper-aspect': String(pdfInfo.aspectRatio) }"
      >
        <span class="thumbnail-pdf-pages">{{ pdfInfo.pageCount }} {{ pdfInfo.pageCount === 1 ? "page" : "pages" }}</span>
        <span class="thumbnail-pdf-size">{{ pdfInfo.pageSize }}</span>
        <span
          v-if="mediaByteLength > 0"
          class="thumbnail-pdf-size"
        >{{ formatFileSize(mediaByteLength) }}</span>
      </div>
    </div>
    <div
      v-else-if="thumbnail && !thumbnail.failed && thumbnail.lines.length > 0"
      class="thumbnail-page"
    >
      <template
        v-for="(line, index) in thumbnail.lines"
        :key="index"
      >
        <img
          v-if="line.image"
          class="thumbnail-image"
          :src="line.image.src"
          :style="{
            width: `${line.image.widthFraction * 100}%`,
            aspectRatio: String(line.image.aspectRatio),
          }"
          alt=""
          loading="lazy"
          decoding="async"
        >
        <div
          v-else-if="line.table"
          class="thumbnail-table"
          :style="{
            width: `${line.table.widthFraction * 100}%`,
            gridTemplateColumns: tableColumns(line.table),
          }"
        >
          <template
            v-for="(row, rowIndex) in line.table.rows"
            :key="rowIndex"
          >
            <span
              v-for="(cell, cellIndex) in row"
              :key="cellIndex"
              class="thumbnail-cell"
              :class="{ 'is-table-header': rowIndex < line.table.headerRows }"
            >{{ cell }}</span>
          </template>
        </div>
        <div
          v-else-if="line.shape"
          class="thumbnail-shape"
          :style="{
            width: `${line.shape.widthFraction * 100}%`,
            aspectRatio: String(line.shape.aspectRatio),
          }"
        >
          <canvas
            :ref="(el) => paintThumbnailShape(el, line.shape)"
            class="thumbnail-shape-canvas"
          />
          <div
            v-if="line.shape.textLines.length > 0"
            class="thumbnail-shape-text"
            :class="`valign-${line.shape.vAlign}`"
          >
            <p
              v-for="(shapeText, textIndex) in line.shape.textLines"
              :key="textIndex"
              class="thumbnail-line is-body"
            >
              {{ shapeText }}
            </p>
          </div>
        </div>
        <div
          v-else-if="line.diagram"
          class="thumbnail-diagram"
          :style="{
            width: `${line.diagram.widthFraction * 100}%`,
            marginLeft: `${line.diagram.leftFraction * 100}%`,
            aspectRatio: String(line.diagram.aspectRatio),
          }"
        >
          <div
            v-for="(item, itemIndex) in line.diagram.items"
            :key="itemIndex"
            class="thumbnail-diagram-item"
            :style="{
              left: `${item.leftFraction * 100}%`,
              top: `${item.topFraction * 100}%`,
              width: `${item.widthFraction * 100}%`,
              height: `${item.heightFraction * 100}%`,
            }"
          >
            <canvas
              v-if="item.chart"
              :key="`chart-${itemIndex}-${resolvedTheme}`"
              :ref="(el) => paintThumbnailChart(el, item.chart && item.chart.chart)"
              class="thumbnail-chart thumbnail-diagram-fill"
            />
            <template v-else-if="item.shape">
              <canvas
                :ref="(el) => paintThumbnailShape(el, item.shape)"
                class="thumbnail-shape-canvas"
              />
              <div
                v-if="item.shape.textLines.length > 0"
                class="thumbnail-shape-text"
                :class="`valign-${item.shape.vAlign}`"
              >
                <p
                  v-for="(shapeText, textIndex) in item.shape.textLines"
                  :key="textIndex"
                  class="thumbnail-line is-body"
                >
                  {{ shapeText }}
                </p>
              </div>
            </template>
          </div>
          <svg
            v-if="line.diagram.connectors.length > 0"
            class="thumbnail-connector thumbnail-diagram-connectors"
            :viewBox="`0 0 ${line.diagram.widthPx} ${line.diagram.heightPx}`"
            preserveAspectRatio="none"
          >
            <g
              v-for="paint in line.diagram.connectors"
              :key="paint.key"
            >
              <path
                :d="paint.d"
                fill="none"
                :stroke="paint.color"
                :stroke-width="paint.widthPx"
                :stroke-dasharray="paint.dash || undefined"
              />
              <path
                v-for="(head, headIndex) in paint.heads"
                :key="headIndex"
                :d="head.d"
                :fill="head.filled ? paint.color : 'none'"
                :stroke="head.filled ? 'none' : paint.color"
                :stroke-width="head.filled ? undefined : paint.widthPx"
              />
            </g>
          </svg>
        </div>
        <svg
          v-else-if="line.connector"
          class="thumbnail-connector"
          :viewBox="`0 0 ${line.connector.widthPx} ${line.connector.heightPx}`"
          preserveAspectRatio="none"
          :style="{
            width: `${line.connector.widthFraction * 100}%`,
            aspectRatio: String(line.connector.aspectRatio),
          }"
        >
          <path
            :d="line.connector.paint.d"
            fill="none"
            :stroke="line.connector.paint.color"
            :stroke-width="line.connector.paint.widthPx"
            :stroke-dasharray="line.connector.paint.dash || undefined"
          />
          <path
            v-for="(head, headIndex) in line.connector.paint.heads"
            :key="headIndex"
            :d="head.d"
            :fill="head.filled ? line.connector.paint.color : 'none'"
            :stroke="head.filled ? 'none' : line.connector.paint.color"
            :stroke-width="head.filled ? undefined : line.connector.paint.widthPx"
          />
        </svg>
        <canvas
          v-else-if="line.chart"
          :key="`chart-${index}-${resolvedTheme}`"
          :ref="(el) => paintThumbnailChart(el, line.chart && line.chart.chart)"
          class="thumbnail-chart"
          :style="{
            width: `${line.chart.widthFraction * 100}%`,
            aspectRatio: String(line.chart.aspectRatio),
          }"
        />
        <p
          v-else
          class="thumbnail-line"
          :class="line.level > 0 ? `is-heading level-${Math.min(line.level, 3)}` : 'is-body'"
        >
          {{ line.text }}
        </p>
      </template>
    </div>
    <div
      v-else-if="thumbnail"
      class="thumbnail-fallback"
    >
      <span class="thumbnail-badge">{{ formatBadgeLabel(props.format) }}</span>
    </div>
    <div
      v-else
      class="thumbnail-skeleton"
    >
      <span class="skeleton-line" />
      <span class="skeleton-line" />
      <span class="skeleton-line short" />
    </div>
  </div>
</template>

<style scoped>
.thumbnail {
  /* relative anchors .thumbnail-deck's absolute inset:0 — without it the
     deck's containing block escapes to the nearest positioned ancestor
     (the sticky quick-look pane / the relative card well) and the slide
     paints displaced outside its box. */
  position: relative;
  width: 100%;
  overflow: hidden;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.thumbnail-page {
  height: 100%;
  padding: var(--space-3);
  display: flex;
  flex-direction: column;
  gap: calc(var(--space-1) / 2);
}

/* Font sizes are a fraction of the smallest type token (calc is treated as an
   opaque token by the design-system guard) — the miniature needs sizes far
   below any standalone token. */
.thumbnail-line {
  margin: 0;
  color: var(--text-primary);
  font-size: calc(var(--text-sm) * 0.5);
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}

/* An image line: sized by the extractor as a fraction of the page's content
   width (with the page's display aspect ratio), so the miniature scales it
   exactly like the page does. object-fit fill mirrors the page renderer,
   which stretches the bitmap to its display box. */
.thumbnail-image {
  display: block;
  flex: none;
  max-width: 100%;
  object-fit: fill;
}

/* A table line: the leading cells laid out on the table's own column
   proportions. The 1px gap over the container's background draws the
   hairline grid; cells paint the paper colour back on top. */
.thumbnail-table {
  flex: none;
  display: grid;
  gap: var(--border-width);
  padding: var(--border-width);
  background: var(--border-subtle);
}

.thumbnail-cell {
  min-width: 0;
  min-height: calc(var(--text-sm) * 0.68);
  padding: 0 calc(var(--space-1) / 2);
  background: var(--surface-base);
  color: var(--text-secondary);
  font-size: calc(var(--text-sm) * 0.45);
  line-height: calc(var(--text-sm) * 0.68);
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
}

.thumbnail-cell.is-table-header {
  color: var(--text-primary);
  font-weight: var(--weight-bold);
}

/* A chart line: painted by paintThumbnailChart at the chart's model px into
   the backing store; this CSS box scales it down like the page miniature. */
.thumbnail-chart {
  display: block;
  flex: none;
  max-width: 100%;
}

/* A diagram line (SPEC-070): a run of offset-positioned drawings and the
   connectors between them, at their real relative positions. The box is the
   run's union bounding box, sized (and indented) like the page sizes it;
   every member is percent-placed inside it, so the whole diagram scales as
   one. */
.thumbnail-diagram {
  position: relative;
  flex: none;
  max-width: 100%;
}

.thumbnail-diagram-item {
  position: absolute;
}

/* A chart inside a diagram fills its placed box rather than taking a width
   fraction of the miniature's column. */
.thumbnail-diagram-fill {
  width: 100%;
  height: 100%;
}

/* The connectors paint over the drawings they join, matching the page (see
   ConnectorPaintLayer's own z-index note). */
.thumbnail-diagram-connectors {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
}

/* A connector line (SPEC-070): the routed line and its arrowheads, drawn at
   the connector's own painted size and scaled into the miniature by the
   viewBox — the one drawing kind the page paints as SVG rather than into a
   canvas, previewed the same way here. */
.thumbnail-connector {
  display: block;
  flex: none;
  max-width: 100%;
  overflow: visible;
}

/* A shape line: the wrapper is the shape's box (sized like the page sizes
   it); the canvas inside is percent-positioned by paintThumbnailShape so a
   rotated shape's grown surface still lines up, and the text frame's lines
   overlay the geometry the way the page's DOM text does. */
.thumbnail-shape {
  position: relative;
  flex: none;
  max-width: 100%;
}

.thumbnail-shape-canvas {
  position: absolute;
  display: block;
}

.thumbnail-shape-text {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  padding: calc(var(--space-1) / 2);
  overflow: hidden;
}

.thumbnail-shape-text.valign-middle {
  justify-content: center;
}

.thumbnail-shape-text.valign-bottom {
  justify-content: flex-end;
}

.thumbnail-line.is-body {
  color: var(--text-secondary);
}

.thumbnail-line.is-heading {
  font-weight: var(--weight-bold);
}

.thumbnail-line.level-1 {
  font-size: calc(var(--text-sm) * 0.85);
}

.thumbnail-line.level-2 {
  font-size: calc(var(--text-sm) * 0.72);
}

.thumbnail-line.level-3 {
  font-size: calc(var(--text-sm) * 0.6);
}

.thumbnail-deck {
  position: absolute;
  inset: 0;
  display: grid;
  align-items: center;
}

.thumbnail-fallback {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-raised);
}

/* A media image thumbnail: the downscaled raster, letterboxed inside the
   card's fixed paper box. */
.thumbnail-media {
  height: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--surface-raised);
}

.thumbnail-media-image {
  max-width: 100%;
  max-height: 100%;
  object-fit: contain;
}

/* The PDF placeholder: an honest paper shape at the first page's real
   aspect ratio (SPEC-107), shown wherever no first-page raster is (yet).
   The well fills the thumbnail (absolute, like the deck, so its padding
   stays inside the box) and is a size container, so the paper is sized from
   the WELL: the largest box at the page's ratio that fits. Sized from its
   content instead, the labels' widths and the paper's automatic minimum
   size overrode aspect-ratio and pulled the paper off the page's shape. */
.thumbnail-pdf {
  position: absolute;
  inset: 0;
  container-type: size;
  display: grid;
  place-items: center;
  background: var(--surface-raised);
  padding: var(--space-3);
}

.thumbnail-pdf-paper {
  box-sizing: border-box;
  width: min(100cqw, 100cqh * var(--paper-aspect));
  height: min(100cqh, 100cqw / var(--paper-aspect));
  /* The labels wrap and clip inside the paper; they never size it. */
  min-width: 0;
  min-height: 0;
  overflow: hidden;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: safe center;
  gap: var(--space-1);
  padding: var(--space-2);
  text-align: center;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-1);
}

.thumbnail-pdf-pages {
  max-width: 100%;
  overflow-wrap: anywhere;
  color: var(--text-primary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.thumbnail-pdf-size {
  max-width: 100%;
  overflow-wrap: anywhere;
  color: var(--text-secondary);
  font-size: calc(var(--text-sm) * 0.85);
}

.thumbnail-badge {
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.thumbnail-skeleton {
  height: 100%;
  padding: var(--space-4);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  justify-content: flex-start;
  background: var(--surface-raised);
}

.skeleton-line {
  height: var(--space-2);
  border-radius: var(--radius-pill);
  background: var(--surface-sunken);
}

.skeleton-line.short {
  width: 60%;
}
</style>

<script setup lang="ts">
// The assembly sheet's preview (SPEC-108): the resolved sheet drawn to scale
// — paper outline, margin guides, the placed item's box and the cropped-away
// region shaded — with the page size and placed size beneath it in the
// user's measurement unit. An SVG whose user units ARE centimetres, so every
// number drawn is the geometry's own: core/sheetLayout's r3-rounded values,
// the same ones the writer receives. It reads only the geometry, the
// placement and (for the picture) a display URL — never the source bytes.
// Runtime prop declarations (espree lints .vue scripts).
import { computed, useId } from "vue";
import { asMeasurementUnit, formatLength } from "../../lib/core/units";

const props = defineProps({
  /** core/sheetLayout's SheetGeometry, or null while there is none. */
  geometry: { type: Object, default: null },
  /** core/sheetLayout's SheetPlacement, or null while there is none. */
  placement: { type: Object, default: null },
  /** The SPEC-064 measurement unit the sizes are stated in. */
  unit: { type: String, default: "cm" },
  /** The paper colour: "#rrggbb" or "transparent" (drawn as a checkerboard). */
  background: { type: String, default: "#ffffff" },
  /** A display URL for the item's picture, or "" to draw its box alone. */
  imageUrl: { type: String, default: "" },
});

const id = useId();
const clipId = `${id}-visible`;
const checkerId = `${id}-checker`;

const sheet = computed(() => props.geometry);
const placed = computed(() => props.placement);

// The view box spans the sheet and, for a cropping fit, the parts of the item
// that overflow it, so the cut-away region is visible beside the paper.
const viewBox = computed(() => {
  const g = sheet.value;
  if (!g) {
    return "0 0 1 1";
  }
  let left = 0;
  let top = 0;
  let right = g.widthCm;
  let bottom = g.heightCm;
  const item = placed.value ? placed.value.item : null;
  if (item) {
    left = Math.min(left, item.xCm);
    top = Math.min(top, item.yCm);
    right = Math.max(right, item.xCm + item.widthCm);
    bottom = Math.max(bottom, item.yCm + item.heightCm);
  }
  const pad = Math.max(right - left, bottom - top) * 0.03;
  return `${left - pad} ${top - pad} ${right - left + 2 * pad} ${bottom - top + 2 * pad}`;
});

/** The cropped-away region: the item's whole extent minus its visible box,
 * one even-odd path. */
const croppedPath = computed(() => {
  const p = placed.value;
  if (!p || !p.cropped) {
    return "";
  }
  const rect = (x = 0, y = 0, w = 0, h = 0) => `M${x} ${y}h${w}v${h}h${-w}Z`;
  return (
    rect(p.item.xCm, p.item.yCm, p.item.widthCm, p.item.heightCm) +
    rect(p.xCm, p.yCm, p.widthCm, p.heightCm)
  );
});

const unitName = computed(() => asMeasurementUnit(props.unit));

function size(widthCm = 0, heightCm = 0) {
  return `${formatLength(widthCm, unitName.value)} × ${formatLength(heightCm, unitName.value)} ${unitName.value}`;
}

const pageSize = computed(() => (sheet.value ? size(sheet.value.widthCm, sheet.value.heightCm) : ""));
const placedSize = computed(() => (placed.value ? size(placed.value.widthCm, placed.value.heightCm) : ""));
const paperFill = computed(() => (props.background === "transparent" ? `url(#${checkerId})` : props.background));
const label = computed(() =>
  pageSize.value ? `Sheet preview: page ${pageSize.value}, placed ${placedSize.value}` : "Sheet preview",
);
</script>

<template>
  <figure
    class="sheet-preview"
    data-testid="sheet-preview"
  >
    <svg
      class="sheet-preview-canvas"
      :viewBox="viewBox"
      preserveAspectRatio="xMidYMid meet"
      role="img"
      :aria-label="label"
    >
      <defs>
        <pattern
          :id="checkerId"
          width="1"
          height="1"
          patternUnits="userSpaceOnUse"
        >
          <rect
            class="checker-light"
            width="1"
            height="1"
          />
          <rect
            class="checker-dark"
            width="0.5"
            height="0.5"
          />
          <rect
            class="checker-dark"
            x="0.5"
            y="0.5"
            width="0.5"
            height="0.5"
          />
        </pattern>
        <clipPath
          v-if="placed"
          :id="clipId"
        >
          <rect
            :x="placed.xCm"
            :y="placed.yCm"
            :width="placed.widthCm"
            :height="placed.heightCm"
          />
        </clipPath>
      </defs>
      <template v-if="sheet">
        <rect
          class="sheet-paper"
          data-testid="sheet-preview-paper"
          x="0"
          y="0"
          :width="sheet.widthCm"
          :height="sheet.heightCm"
          :fill="paperFill"
        />
        <rect
          class="sheet-margins"
          data-testid="sheet-preview-margins"
          :x="sheet.content.xCm"
          :y="sheet.content.yCm"
          :width="sheet.content.widthCm"
          :height="sheet.content.heightCm"
        />
      </template>
      <template v-if="placed">
        <template v-if="props.imageUrl">
          <image
            v-if="placed.cropped"
            class="sheet-item-overflow"
            :href="props.imageUrl"
            :x="placed.item.xCm"
            :y="placed.item.yCm"
            :width="placed.item.widthCm"
            :height="placed.item.heightCm"
            preserveAspectRatio="none"
          />
          <image
            :href="props.imageUrl"
            :x="placed.item.xCm"
            :y="placed.item.yCm"
            :width="placed.item.widthCm"
            :height="placed.item.heightCm"
            preserveAspectRatio="none"
            :clip-path="`url(#${clipId})`"
          />
        </template>
        <path
          v-if="croppedPath"
          class="sheet-cropped"
          data-testid="sheet-preview-cropped"
          :d="croppedPath"
          fill-rule="evenodd"
        />
        <rect
          class="sheet-placed"
          data-testid="sheet-preview-placed"
          :x="placed.xCm"
          :y="placed.yCm"
          :width="placed.widthCm"
          :height="placed.heightCm"
        />
      </template>
    </svg>
    <figcaption class="sheet-preview-caption">
      <span data-testid="sheet-preview-page-size">Page {{ pageSize }}</span>
      <span data-testid="sheet-preview-placed-size">Placed {{ placedSize }}</span>
    </figcaption>
  </figure>
</template>

<style scoped>
.sheet-preview {
  margin: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-2);
}

.sheet-preview-canvas {
  width: 100%;
  height: 16rem;
  background: var(--surface-sunken);
  border-radius: var(--radius-md);
}

.sheet-paper {
  stroke: var(--border-strong);
  stroke-width: 1;
  vector-effect: non-scaling-stroke;
}

.sheet-margins {
  fill: none;
  stroke: var(--accent);
  stroke-width: 1;
  stroke-dasharray: 4 3;
  vector-effect: non-scaling-stroke;
}

.sheet-placed {
  fill: none;
  stroke: var(--accent);
  stroke-width: 2;
  vector-effect: non-scaling-stroke;
}

.sheet-item-overflow {
  opacity: 0.35;
}

.sheet-cropped {
  fill: var(--text-secondary);
  fill-opacity: 0.35;
  stroke: var(--text-secondary);
  stroke-width: 1;
  stroke-dasharray: 2 2;
  vector-effect: non-scaling-stroke;
}

.checker-light {
  fill: var(--surface-base);
}

.checker-dark {
  fill: var(--border-subtle);
}

.sheet-preview-caption {
  display: flex;
  gap: var(--space-4);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}
</style>

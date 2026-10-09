<script setup lang="ts">
// The sheet paper (SPEC-117 "Rendering", "Dark mode"): the cell layer of the
// grid, in layers bottom to top — gridlines (one line per visible row and
// column boundary, absent when the sheet hides them) → cells (fill and text,
// merged ranges as one element, overflowing text raised over its empty
// neighbours) → borders (one element per resolved edge). SPEC-059's rule
// applies: the page is not a surface, so its white, its #d4d4d4 gridlines
// and every cell colour (the file's own, from the model) are literal in
// every app theme, and the dim-page filter reaches it through `.sheet-paper`
// (App.vue). This file is the one tokens.spec.ts exempts.
//
// `bands` is useSheetCells' BandItem[] — the frozen rows, then the scrolled
// rows — each split into the frozen-column and scrolled-column regions. A
// cell's family is the resolved name; it is wrapped in SPEC-048's fontStack
// here. Each cell element's id, role, data-ref and ARIA come from
// cellAttributes — none on a merge's paint-only copy in a further frozen
// pane. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import { fontStack } from "../../lib/fonts/substitution";
import { asBandItems, cellAttributes } from "../../views/spreadsheet/sheetLayerProps";

const props = defineProps({
  /** The bands to draw (useSheetCells' BandItem[]). */
  bands: { type: Array, default: () => [] },
  /** Whether the sheet shows gridlines (`showGridLines`). */
  gridLines: { type: Boolean, default: true },
  /** The frozen columns' width, unscaled px (0 without a freeze). */
  frozenWidth: { type: Number, default: 0 },
});

const items = computed(() => asBandItems(props.bands));

/** The text box's style, its family wrapped in the substitution stack. */
function textStyle(style = {}) {
  const family = Reflect.get(style, "font-family");
  return typeof family === "string" && family.length > 0 ? { ...style, "font-family": fontStack(family) } : style;
}

function px(value = 0) {
  return `${value}px`;
}
</script>

<template>
  <div class="sheet-paper">
    <div
      v-for="band in items"
      :key="band.kind"
      class="sheet-band"
      :style="{ top: px(band.top), height: px(band.height) }"
    >
      <template v-if="props.gridLines">
        <div
          class="sheet-gridlines"
          :style="{ left: '0px', width: px(props.frozenWidth) }"
        >
          <div
            v-for="y in band.gridRows"
            :key="`h${y}`"
            class="sheet-gridline is-horizontal"
            :style="{ top: px(y) }"
          />
          <div
            v-for="x in band.gridColsFrozen"
            :key="`v${x}`"
            class="sheet-gridline is-vertical"
            :style="{ left: px(x) }"
          />
        </div>
        <div
          class="sheet-gridlines"
          :style="{ left: px(props.frozenWidth), right: '0px' }"
        >
          <div
            v-for="y in band.gridRows"
            :key="`h${y}`"
            class="sheet-gridline is-horizontal"
            :style="{ top: px(y) }"
          />
          <div
            v-for="x in band.gridColsScrolled"
            :key="`v${x}`"
            class="sheet-gridline is-vertical"
            :style="{ left: px(x) }"
          />
        </div>
      </template>
      <div
        v-for="row in band.rows"
        :key="row.index"
        class="grid-row"
        role="row"
        :aria-rowindex="row.index + 2"
        :style="{ top: px(row.top), height: px(row.height) }"
      >
        <div
          v-for="region in [{ key: 'frozen', cells: row.frozen, left: 0, width: props.frozenWidth }, { key: 'scrolled', cells: row.scrolled, left: props.frozenWidth, width: null }]"
          :key="region.key"
          class="sheet-region"
          :style="region.width === null ? { left: px(region.left), right: '0px' } : { left: px(region.left), width: px(region.width) }"
        >
          <div
            v-for="cell in region.cells"
            :key="cell.key"
            class="sheet-cell"
            :class="[cell.kindClass, { 'is-merged': cell.merged, 'is-raised': cell.raised, 'is-paint-only': cell.paintOnly }]"
            v-bind="cellAttributes(cell)"
            :style="cell.boxStyle"
          >
            <div
              class="sheet-cell-text"
              :style="textStyle(cell.textStyle)"
            >
              <span
                class="sheet-cell-content"
                :style="cell.contentStyle"
              ><template v-if="cell.runs"><span
                v-for="(run, i) in cell.runs"
                :key="i"
                :style="textStyle(run.style)"
              >{{ run.text }}</span></template><template v-else>{{ cell.text }}</template></span>
            </div>
            <svg
              v-if="cell.diagonal"
              class="sheet-cell-diagonal"
              aria-hidden="true"
              :viewBox="`0 0 ${cell.width} ${cell.height}`"
              preserveAspectRatio="none"
            >
              <line
                v-if="cell.diagonal.down"
                x1="0"
                y1="0"
                :x2="cell.width"
                :y2="cell.height"
                :stroke="cell.diagonal.color"
                :stroke-width="cell.diagonal.width"
                :stroke-dasharray="cell.diagonal.dasharray || undefined"
                :stroke-opacity="cell.diagonal.opacity ?? undefined"
              />
              <line
                v-if="cell.diagonal.up"
                x1="0"
                :y1="cell.height"
                :x2="cell.width"
                y2="0"
                :stroke="cell.diagonal.color"
                :stroke-width="cell.diagonal.width"
                :stroke-dasharray="cell.diagonal.dasharray || undefined"
                :stroke-opacity="cell.diagonal.opacity ?? undefined"
              />
            </svg>
          </div>
        </div>
      </div>
      <div
        class="sheet-borders"
        :style="{ left: '0px', width: px(props.frozenWidth) }"
      >
        <div
          v-for="border in band.bordersFrozen"
          :key="border.key"
          class="sheet-border"
          :style="border.style"
        />
      </div>
      <div
        class="sheet-borders"
        :style="{ left: px(props.frozenWidth), right: '0px' }"
      >
        <div
          v-for="border in band.bordersScrolled"
          :key="border.key"
          class="sheet-border"
          :style="border.style"
        />
      </div>
    </div>
  </div>
</template>

<style scoped>
/* The paper: literal white with #d4d4d4 gridlines in every app theme — the
   file's colours are painted on it, and App.vue's dim-page rule dims it as it
   dims a text page. */
.sheet-paper {
  position: absolute;
  overflow: hidden;
  background: #ffffff;
  color: #000000;
}

.sheet-band {
  position: absolute;
  left: 0;
  right: 0;
  overflow: hidden;
}

.sheet-gridlines,
.sheet-borders {
  position: absolute;
  top: 0;
  bottom: 0;
  overflow: hidden;
  pointer-events: none;
}

.sheet-borders {
  z-index: 2;
}

.sheet-gridline {
  position: absolute;
  background: #d4d4d4;
}

.sheet-gridline.is-horizontal {
  left: 0;
  right: 0;
  height: 1px;
}

.sheet-gridline.is-vertical {
  top: 0;
  bottom: 0;
  width: 1px;
}

.grid-row {
  position: absolute;
  left: 0;
  right: 0;
}

/* A column region of one row: clips horizontally at the freeze line, never
   vertically (a merged range runs on below its anchor row). */
.sheet-region {
  position: absolute;
  top: 0;
  height: 100%;
  clip-path: inset(-100000px 0 -100000px 0);
}

.sheet-cell {
  position: absolute;
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  box-sizing: border-box;
  font-family: Calibri, Carlito, sans-serif;
  font-size: 11pt;
  line-height: 1.2;
}

.sheet-cell.is-raised {
  z-index: 1;
}

.sheet-cell-text {
  position: absolute;
  top: 0;
  bottom: 0;
  display: flex;
  flex-direction: column;
  justify-content: inherit;
  box-sizing: border-box;
  overflow: hidden;
  white-space: pre;
}

.sheet-cell-content {
  display: block;
}

.sheet-cell-diagonal {
  position: absolute;
  inset: 0;
  width: 100%;
  height: 100%;
  pointer-events: none;
}

.sheet-border {
  position: absolute;
  box-sizing: content-box;
  width: 0;
  height: 0;
}
</style>

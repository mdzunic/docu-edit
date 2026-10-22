<script setup lang="ts">
// One printed sheet page (SPEC-136): what sheetPrintLayout.ts laid out —
// header and footer bands, the headings, the cells (`sheet-print-cell`, each
// carrying its resolved style as CSS, its text clipped to its box; a link as
// `a[data-doc-link]`), and the drawings and charts of the page's band
// through SheetDrawingView (charts painted into `canvas.sheet-print-chart`).
// The page is paper: white, 1:1, its scale laid out rather than transformed.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import SheetDrawingView from "./SheetDrawingView.vue";
import { fontStack } from "../../lib/core/fontSubstitution";
import { asPrintPageLayout, bandsOf, cssPx, PRINT_SIDES } from "../../views/spreadsheet/sheetPrintLayout";

const props = defineProps({
  /** A PrintPageLayout. */
  layout: { type: Object, required: true },
  /** "Page N of M". */
  label: { type: String, required: true },
  /** The page's index in the job (the export finds it by it). */
  pageIndex: { type: Number, required: true },
  /** The `@page` name the page prints under ("" for none). */
  pageName: { type: String, default: "" },
  /** The default font family. */
  fontFamily: { type: String, default: "" },
  theme: { type: Object, default: null },
  imageSrc: { type: Function, default: undefined },
});

const l = computed(() => asPrintPageLayout(props.layout));
const pageStyle = computed(() => ({
  width: cssPx(l.value.widthPx),
  height: cssPx(l.value.heightPx),
  background: "#ffffff",
  color: "#000000",
  "font-family": props.fontFamily ? fontStack(props.fontFamily) : undefined,
  "font-size": `${l.value.baseFontPt}pt`,
  ...(props.pageName ? { page: props.pageName } : {}),
}));

function cellStyle(cell = l.value.cells[0]) {
  return { ...cell.box, left: cssPx(cell.left), top: cssPx(cell.top), width: cssPx(cell.width), height: cssPx(cell.height) };
}

const bands = computed(() => bandsOf(l.value));

function bandStyle(item = bands.value[0]) {
  return { [item.edge]: cssPx(item.band.offset), left: cssPx(item.band.insetLeft), right: cssPx(item.band.insetRight) };
}

const bandBox = computed(() => ({
  left: cssPx(l.value.bandBox.left),
  top: cssPx(l.value.bandBox.top),
  width: cssPx(l.value.bandBox.width),
  height: cssPx(l.value.bandBox.height),
}));
</script>

<template>
  <div
    class="sheet-print-page"
    data-testid="sheet-print-page"
    role="group"
    :aria-label="label"
    :data-page-index="pageIndex"
    :style="pageStyle"
  >
    <div
      v-for="item in bands"
      :key="item.testid"
      class="sheet-print-band"
      :class="`is-${item.edge}`"
      :data-testid="item.testid"
      :style="bandStyle(item)"
    >
      <span
        v-for="side in PRINT_SIDES"
        :key="side"
        class="sheet-print-section"
        :class="`is-${side}`"
      ><span
        v-for="(run, index) in item.band[side]"
        :key="index"
        :style="run.css"
      >{{ run.text }}</span></span>
    </div>
    <div
      v-for="cell in l.cells"
      :key="cell.key"
      class="sheet-print-cell"
      :class="{ 'sheet-print-heading': cell.heading, 'is-spill': cell.spill }"
      :data-testid="cell.heading ? 'sheet-print-heading' : 'sheet-print-cell'"
      :data-ref="cell.ref || undefined"
      :data-rows="cell.rowSpan"
      :data-cols="cell.colSpan"
      :style="cellStyle(cell)"
    >
      <a
        v-if="cell.href"
        class="sheet-print-text"
        data-doc-link=""
        :href="cell.href"
        :style="cell.text"
      >{{ cell.content }}</a>
      <span
        v-else-if="cell.content"
        class="sheet-print-text"
        :style="cell.text"
      >{{ cell.content }}</span>
    </div>
    <div
      v-if="l.drawings.length > 0"
      class="sheet-print-drawings"
      :style="bandBox"
    >
      <div
        v-for="item in l.drawings"
        :key="item.drawing.id"
        class="sheet-print-drawing"
        :style="{ left: cssPx(item.left), top: cssPx(item.top), width: cssPx(item.width * l.scale), height: cssPx(item.height * l.scale) }"
      >
        <SheetDrawingView
          :content="item.drawing.content"
          :width="item.width"
          :height="item.height"
          :zoom="l.scale"
          :theme="theme"
          :image-src="imageSrc"
          chart-canvas-class="sheet-print-chart"
        />
      </div>
    </div>
  </div>
</template>

<style scoped>
.sheet-print-page {
  position: relative;
  overflow: hidden;
  box-sizing: border-box;
  flex: none;
}

.sheet-print-band {
  position: absolute;
  display: grid;
  grid-template-columns: 1fr auto 1fr;
  white-space: pre;
}

.sheet-print-section.is-center {
  text-align: center;
}

.sheet-print-section.is-right {
  text-align: right;
}

.sheet-print-cell {
  position: absolute;
  box-sizing: border-box;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.sheet-print-text {
  display: block;
  overflow: hidden;
  color: inherit;
  text-decoration: inherit;
}

a.sheet-print-text {
  text-decoration: underline;
}

.sheet-print-drawings {
  position: absolute;
  overflow: hidden;
}

.sheet-print-drawing {
  position: absolute;
}
</style>

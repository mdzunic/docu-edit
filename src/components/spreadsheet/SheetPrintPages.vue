<script setup lang="ts">
// The pages of a sheet print job (SPEC-136). `mode="preview"`: in the grid
// area, one `sheet-print-page` per page of the job scaled to fit the area's
// width ("Page N of M"), only the pages in or next to the scrolled window
// rendered — the rest are sized empty boxes. `mode="sheet"`: the pages
// mounted off-screen at 1:1 (fixed at left −20000px, aria-hidden, inert) for
// the browser's print (App.vue's print CSS brings them back, one per sheet
// of paper) and the PDF export's rolling window. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import SheetPrintPage from "./SheetPrintPage.vue";
import { asPlacedPages, asLayoutFor } from "../../views/spreadsheet/useSheetPrint";
import { cssPx } from "../../views/spreadsheet/sheetPrintLayout";

const props = defineProps({
  /** PlacedPage[]: the pages, each with its index in the job. */
  pages: { type: Array, required: true },
  /** The job's page count (the M of "Page N of M"). */
  total: { type: Number, required: true },
  /** (placed, total) → PrintPageLayout. */
  layoutFor: { type: Function, required: true },
  mode: { type: String, default: "preview" },
  /** `page:` names, by position in `pages` (sheet mode, printing). */
  pageNames: { type: Array, default: () => [] },
  fontFamily: { type: String, default: "" },
  theme: { type: Object, default: null },
  imageSrc: { type: Function, default: undefined },
});

const placed = computed(() => asPlacedPages(props.pages));
const layout = computed(() => asLayoutFor(props.layoutFor));
const PAD = 24;
const GAP = 16;

const viewport = useTemplateRef("viewport");
const area = ref({ width: 800, height: 600, scrollTop: 0 });
function measure() {
  const el = viewport.value;
  if (el) {
    area.value = { width: el.clientWidth, height: el.clientHeight, scrollTop: el.scrollTop };
  }
}

const observer = typeof ResizeObserver === "undefined" ? null : new ResizeObserver(() => measure());

onMounted(() => {
  measure();
  if (viewport.value && observer) {
    observer.observe(viewport.value);
  }
});
onBeforeUnmount(() => observer?.disconnect());

const widestPx = computed(() => Math.max(1, ...placed.value.map((each) => (each.page.paperPt.width * 4) / 3)));
const fit = computed(() => (props.mode === "preview" ? Math.min(1, Math.max(0.05, (area.value.width - 2 * PAD) / widestPx.value)) : 1));

/** The top of each page's slot in the preview, and its scaled box. */
const slots = computed(() => {
  let top = PAD;
  return placed.value.map((each) => {
    const width = ((each.page.paperPt.width * 4) / 3) * fit.value;
    const height = ((each.page.paperPt.height * 4) / 3) * fit.value;
    const slot = { top, width, height };
    top += height + GAP;
    return slot;
  });
});

function shown(index = 0) {
  if (props.mode !== "preview") {
    return true;
  }
  const slot = slots.value[index];
  const from = area.value.scrollTop - area.value.height;
  const to = area.value.scrollTop + 2 * area.value.height;
  return slot !== undefined && slot.top + slot.height >= from && slot.top <= to;
}

function label(index = 0) {
  return `Page ${placed.value[index].index + 1} of ${props.total}`;
}
</script>

<template>
  <div
    v-if="mode === 'preview'"
    ref="viewport"
    class="sheet-print-preview"
    data-testid="sheet-print-preview"
    @scroll="measure"
  >
    <div
      v-for="(each, index) in placed"
      :key="each.index"
      class="sheet-print-slot"
      :style="{ width: cssPx(slots[index].width), height: cssPx(slots[index].height) }"
    >
      <div
        v-if="shown(index)"
        class="sheet-print-scaled"
        :style="{ transform: `scale(${fit})` }"
      >
        <SheetPrintPage
          :layout="layout(each, total)"
          :label="label(index)"
          :page-index="each.index"
          :font-family="fontFamily"
          :theme="theme"
          :image-src="imageSrc"
        />
      </div>
      <div
        v-else
        class="sheet-print-placeholder"
        data-testid="sheet-print-page"
        role="group"
        :aria-label="label(index)"
        :data-page-index="each.index"
      />
    </div>
  </div>
  <Teleport
    v-else
    to="body"
  >
    <div
      class="sheet-print-sheet"
      aria-hidden="true"
      :inert="true"
      :style="pageNames[0] ? { page: String(pageNames[0]) } : undefined"
    >
      <SheetPrintPage
        v-for="(each, index) in placed"
        :key="each.index"
        :layout="layout(each, total)"
        :label="label(index)"
        :page-index="each.index"
        :page-name="String(pageNames[index] ?? '')"
        :font-family="fontFamily"
        :theme="theme"
        :image-src="imageSrc"
      />
    </div>
  </Teleport>
</template>

<style scoped>
.sheet-print-preview {
  position: absolute;
  inset: 0;
  overflow: auto;
  padding: var(--space-6);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-5);
  background: var(--surface-sunken);
  z-index: 5;
}

.sheet-print-slot {
  position: relative;
  flex: none;
  box-shadow: var(--shadow-2);
}

.sheet-print-scaled {
  transform-origin: top left;
}

.sheet-print-placeholder {
  width: 100%;
  height: 100%;
  background: var(--surface-raised);
}

.sheet-print-sheet {
  position: fixed;
  left: -20000px;
  top: 0;
  pointer-events: none;
  print-color-adjust: exact;
  -webkit-print-color-adjust: exact;
}
</style>

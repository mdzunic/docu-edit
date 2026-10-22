<script setup lang="ts">
// A chart sheet (SPEC-132 "Chart sheets"): `sheet-chart-sheet` paints the
// sheet's one chart through SheetDrawingView, fitted to the view at its
// anchor's aspect (extEmu cx:cy; the whole area when it has none) —
// letterboxed, centred, at most the visible area below the floating chart
// toolbar — and repainted as the view resizes (ResizeObserver). The chart is
// never moved, resized or deleted as a drawing. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed, onBeforeUnmount, onMounted, ref, useTemplateRef } from "vue";
import { asDrawingContent } from "../../views/spreadsheet/sheetDrawingLayout";
import SheetDrawingView from "./SheetDrawingView.vue";

const props = defineProps({
  /** The chart Sheet on show. */
  sheet: { type: Object, required: true },
  /** The workbook's DocumentTheme. */
  theme: { type: Object, default: null },
});

const area = useTemplateRef("area");
const box = ref({ width: 0, height: 0 });
let stopObserving = () => {};

const drawing = computed(() => (props.sheet.drawings ?? []).find((entry = Object()) => asDrawingContent(entry.content).kind === "chart") ?? null);
const fitted = computed(() => {
  const ext = drawing.value?.anchor?.extEmu;
  const { width, height } = box.value;
  if (!ext || !(ext.cx > 0) || !(ext.cy > 0)) {
    return { width, height };
  }
  const scale = Math.min(width / ext.cx, height / ext.cy);
  return { width: Math.floor(ext.cx * scale), height: Math.floor(ext.cy * scale) };
});

function measure() {
  const element = area.value;
  if (element instanceof HTMLElement) {
    const style = getComputedStyle(element);
    const across = (parseFloat(style.paddingLeft) || 0) + (parseFloat(style.paddingRight) || 0);
    const down = (parseFloat(style.paddingTop) || 0) + (parseFloat(style.paddingBottom) || 0);
    box.value = { width: Math.max(0, element.clientWidth - across), height: Math.max(0, element.clientHeight - down) };
  }
}

onMounted(() => {
  measure();
  if (typeof ResizeObserver !== "undefined" && area.value instanceof HTMLElement) {
    const observer = new ResizeObserver(measure);
    observer.observe(area.value);
    stopObserving = () => observer.disconnect();
  }
});
onBeforeUnmount(() => stopObserving());
</script>

<template>
  <div
    ref="area"
    class="sheet-chart-sheet"
    data-testid="sheet-chart-sheet"
  >
    <div
      v-if="drawing && fitted.width > 0 && fitted.height > 0"
      class="sheet-chart-sheet-chart"
      role="img"
      :aria-label="drawing.altText || drawing.name"
      :style="{ width: `${fitted.width}px`, height: `${fitted.height}px` }"
    >
      <SheetDrawingView
        :content="drawing.content"
        :width="fitted.width"
        :height="fitted.height"
        :theme="props.theme"
      />
    </div>
  </div>
</template>

<style scoped>
.sheet-chart-sheet {
  height: 100%;
  box-sizing: border-box;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
  padding: calc(var(--height-control) + var(--space-4)) var(--space-4) var(--space-4);
  background: var(--surface-sunken);
}

.sheet-chart-sheet-chart {
  position: relative;
  flex: none;
}
</style>

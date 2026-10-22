<script setup lang="ts">
// The Insert menu's surfaces (SPEC-131): Insert ▸ Image…'s hidden file picker
// (`sheet-insert-image-input`, accepting ACCEPTED_IMAGE_ACCEPT_ATTR — bound
// into useSheetDrawings, whose pickImage clicks it; a picked file goes to
// insertImageFile) and Insert ▸ Shapes' popover (`sheet-shape-gallery-popover`,
// under the menu bar's Insert title, focus on its first shape) hosting
// ShapeGallery with `catalogueOnly` — no Connectors or Freeform. A pick arms
// the draw gesture (useSheetDrawings' pickShape); Escape or a press outside
// closes it (UiMenuSurface). Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import { ACCEPTED_IMAGE_ACCEPT_ATTR } from "../../lib/core/media";
import { asShapeGeometry } from "../../lib/core/shapes";
import { asSheetCharts } from "../../views/spreadsheet/useSheetCharts";
import { asSheetDrawings } from "../../views/spreadsheet/useSheetDrawings";
import ChartGallery from "../ChartGallery.vue";
import ShapeGallery from "../ShapeGallery.vue";
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  /** useSheetDrawings' return. */
  drawings: { type: Object, required: true },
  /** SPEC-132: useSheetCharts' return — Insert ▸ Chart's gallery
   * (`sheet-chart-gallery-popover` around ChartGallery; a pick inserts). */
  charts: { type: Object, default: null },
});

const drawings = computed(() => asSheetDrawings(props.drawings));
const charts = computed(() => (props.charts ? asSheetCharts(props.charts) : null));
const picker = useTemplateRef("picker");
/** The popover's viewport position: under the Insert menu's title. */
const at = ref({ left: 16, top: 64 });

onMounted(() => drawings.value.bindImagePicker(picker.value instanceof HTMLInputElement ? picker.value : null));
onBeforeUnmount(() => drawings.value.bindImagePicker(null));

watch(
  () => drawings.value.galleryOpen.value || charts.value?.galleryOpen.value === true,
  async (open) => {
    if (!open) {
      return;
    }
    const title = document.getElementById("menu-title-insert");
    const box = title ? title.getBoundingClientRect() : null;
    if (box) {
      at.value = { left: box.left, top: box.bottom };
    }
    await nextTick();
    const first = document.querySelector('[data-testid="sheet-shape-gallery-popover"] .shape-gallery-item, [data-testid="sheet-chart-gallery-popover"] .chart-gallery-item');
    if (first instanceof HTMLElement) {
      first.focus();
    }
  },
);

/** The picker's change: its file inserted, the picker reset. */
function onPick(event = new Event("change")) {
  const input = event.target;
  if (!(input instanceof HTMLInputElement)) {
    return;
  }
  const file = input.files ? input.files[0] : undefined;
  input.value = "";
  if (file) {
    void drawings.value.insertImageFile(file);
  }
}

function pickShape(geometry = "rect") {
  drawings.value.pickShape(asShapeGeometry(geometry));
}
</script>

<template>
  <input
    ref="picker"
    type="file"
    hidden
    :accept="ACCEPTED_IMAGE_ACCEPT_ATTR"
    data-testid="sheet-insert-image-input"
    aria-hidden="true"
    tabindex="-1"
    @change="onPick"
  >
  <UiMenuSurface
    :open="drawings.galleryOpen.value"
    :anchor="null"
    class="sheet-shape-gallery-popover"
    data-testid="sheet-shape-gallery-popover"
    :style="{ left: `${at.left}px`, top: `${at.top}px` }"
    @dismiss="drawings.closeGallery()"
  >
    <ShapeGallery
      catalogue-only
      @pick="pickShape"
    />
  </UiMenuSurface>
  <UiMenuSurface
    v-if="charts"
    :open="charts.galleryOpen.value"
    :anchor="null"
    class="sheet-shape-gallery-popover"
    data-testid="sheet-chart-gallery-popover"
    :style="{ left: `${at.left}px`, top: `${at.top}px` }"
    @dismiss="charts.closeGallery()"
  >
    <ChartGallery @pick="(chartType) => charts?.pickChartType(chartType)" />
  </UiMenuSurface>
</template>

<style scoped>
.sheet-shape-gallery-popover {
  position: fixed;
  z-index: 40;
  width: 24rem;
}
</style>

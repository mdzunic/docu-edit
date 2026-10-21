<script setup lang="ts">
// One grid pane's drawing layer (SPEC-131 "The drawing layer"):
// `sheet-drawing-layer` (`data-pane` corner | top | left | main), a clipping
// box over the cells and under the selection overlays, in the grid's unscaled
// px (the zoom transform scales it). It paints the drawings
// useDrawingGestures mounts in this pane (sheetDrawingLayout.ts: the pane
// split, the overscan, the primary copy), back to front, each a `sheet-drawing`
// element placed at its pane position — `data-drawing-id`, `data-drawing-name`,
// `data-kind` (image | shape | group | chart | foreign), role img for an image
// and group otherwise, named by its alt text or else its name, aria-selected,
// rotated about its centre — holding its paint (SheetDrawingView) and, when
// selected, its handles (SheetDrawingHandles). Only the primary copy of a
// drawing split across panes is focusable (tabindex 0), named and handled; the
// other copies are aria-hidden paint (`data-primary="false"`). The armed draw
// gesture's box shows as `sheet-drawing-draw-preview`.
// Events go up for the composables: press(id, event) and handle-press(id,
// handle, event) — useDrawingGestures, which keeps them from the grid's own
// pointer handling — key(id, event), menu(id, x, y) (a right-click or the Menu
// key; the default menu suppressed) and focus-drawing(id), useSheetDrawings'.
// A double-click on a drawing never opens the cell editor. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asPaneView, asRectPx, drawingLabel, drawingRole, drawingRotation, inPane } from "../../views/spreadsheet/sheetDrawingLayout";
import SheetDrawingHandles from "./SheetDrawingHandles.vue";
import SheetDrawingView from "./SheetDrawingView.vue";

const props = defineProps({
  /** The pane and what it mounts (sheetDrawingLayout's PaneView). */
  pane: { type: Object, required: true },
  /** The grid's zoom as a fraction. */
  zoom: { type: Number, default: 1 },
  /** The workbook's DocumentTheme. */
  theme: { type: Object, default: null },
  /** mediaId → an image's data URI ("" paints the placeholder). */
  imageSrc: { type: Function, default: undefined },
  /** The armed draw gesture's box (100%-zoom sheet px); null: none. */
  drawBox: { type: Object, default: null },
});

const emit = defineEmits(["press", "handle-press", "key", "menu", "focus-drawing"]);

const view = computed(() => asPaneView(props.pane));
const drawRect = computed(() => {
  const box = asRectPx(props.drawBox);
  return box ? inPane(box, view.value) : null;
});

function px(value = 0) {
  return `${value}px`;
}

function itemStyle(item = view.value.items[0]) {
  if (!item) {
    return {};
  }
  const rotation = drawingRotation(item.drawing);
  return {
    left: px(item.x),
    top: px(item.y),
    width: px(item.width),
    height: px(item.height),
    transform: rotation ? `rotate(${rotation}deg)` : undefined,
  };
}

/** A right-click or the Menu key: the drawing's menu at the pointer (under
 * the element for a key). */
function onMenu(id = "", event = new MouseEvent("contextmenu")) {
  event.preventDefault();
  event.stopPropagation();
  const target = event.currentTarget instanceof Element ? event.currentTarget.getBoundingClientRect() : null;
  const keyed = event.clientX === 0 && event.clientY === 0 && target;
  emit("menu", id, keyed ? target.left : event.clientX, keyed ? target.bottom : event.clientY);
}
</script>

<template>
  <div
    class="sheet-drawing-layer"
    data-testid="sheet-drawing-layer"
    :data-pane="view.id"
    :style="{ left: px(view.left), top: px(view.top), width: px(view.width), height: px(view.height) }"
  >
    <div
      v-for="item in view.items"
      :key="item.drawing.id"
      class="sheet-drawing"
      :class="{ 'is-selected': item.selected }"
      data-testid="sheet-drawing"
      :data-drawing-id="item.drawing.id"
      :data-drawing-name="item.drawing.name"
      :data-kind="item.drawing.content.kind"
      :data-primary="item.primary ? 'true' : 'false'"
      :role="drawingRole(item.drawing)"
      :aria-label="drawingLabel(item.drawing)"
      :aria-selected="item.selected ? 'true' : 'false'"
      :aria-hidden="item.primary ? undefined : 'true'"
      :tabindex="item.primary ? 0 : -1"
      :style="itemStyle(item)"
      @pointerdown="emit('press', item.drawing.id, $event)"
      @pointermove.stop
      @dblclick.stop
      @keydown="emit('key', item.drawing.id, $event)"
      @contextmenu="onMenu(item.drawing.id, $event)"
      @focus="emit('focus-drawing', item.drawing.id)"
    >
      <SheetDrawingView
        :content="item.drawing.content"
        :width="item.width"
        :height="item.height"
        :zoom="props.zoom"
        :theme="props.theme"
        :image-src="props.imageSrc"
        host-rotates
      />
      <SheetDrawingHandles
        v-if="item.selected && item.primary"
        :zoom="props.zoom"
        @press="(handle, event) => emit('handle-press', item.drawing.id, handle, event)"
      />
    </div>
    <div
      v-if="drawRect"
      class="sheet-drawing-draw-preview"
      data-testid="sheet-drawing-draw-preview"
      aria-hidden="true"
      :style="{ left: px(drawRect.left), top: px(drawRect.top), width: px(drawRect.width), height: px(drawRect.height) }"
    />
  </div>
</template>

<style scoped>
/* A pane's clipping box: it takes no pointer itself, so cells under it stay
   the grid's; each drawing takes its own presses. */
.sheet-drawing-layer {
  position: absolute;
  overflow: hidden;
  pointer-events: none;
}

.sheet-drawing {
  position: absolute;
  pointer-events: auto;
  touch-action: none;
  cursor: move;
}

.sheet-drawing.is-selected {
  outline: var(--border-width) solid var(--accent);
}

.sheet-drawing:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
}

.sheet-drawing-draw-preview {
  position: absolute;
  box-sizing: border-box;
  pointer-events: none;
  border: var(--border-width) dashed var(--accent);
}
</style>

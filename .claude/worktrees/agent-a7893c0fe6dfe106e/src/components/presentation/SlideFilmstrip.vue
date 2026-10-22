<script setup lang="ts">
// The slide filmstrip (SPEC-079 listbox + SPEC-080 slide management): a
// listbox of numbered thumbnails, each a scaled-down non-interactive
// SlideStage, plus an "Add slide" button and per-slide management — a context
// menu (right-click, a ⋯ affordance, and Shift+F10 / the Menu key on a focused
// thumbnail) with New slide / Duplicate / Delete / Move up / Move down, and
// pointer drag-reorder with a visible drop indicator (dropping outside
// cancels). Structural intents are emitted up to PresentationView, which owns
// the commands + undo. SPEC-099: the thumbnails are the shared SlideThumb and
// the menu the shared SlideContextMenu, and when the view hands in the shared
// slide `selection` (useSlideSelection) clicks and the selection keys go
// through it — the same multi-selection the slide sorter shows — with
// `selectedIndexes` driving aria-selected. SPEC-104 adds Hide/Show slide to
// the menu; a hidden slide's item carries `filmstrip-hidden` and its thumbnail
// dims (the editor still reaches and edits it like any other slide). Runtime
// prop/emit declarations, not defineProps<{...}>() generics (espree lints .vue
// scripts).
import { computed, nextTick, ref, useTemplateRef } from "vue";
import SlideThumb from "./SlideThumb.vue";
import SlideContextMenu from "./SlideContextMenu.vue";
import {
  anchorPoint,
  asSlides,
  asStageFieldsBySlide,
  filmstripLabel,
  filmstripMenuItems,
  filmstripThumbScale,
  gestureModifiers,
} from "./slideRender";

const props = defineProps({
  slides: { type: Array, default: () => [] },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  /** SPEC-094: the deck's theme, threaded to the stage surfaces. */
  theme: { type: Object, default: null },
  currentIndex: { type: Number, default: 0 },
  /** SPEC-093: each slide's pre-resolved field band (footer/number/date), so a
   * thumbnail's number agrees with the live slide index after a structural
   * change. Indexed to match `slides`. */
  fieldsBySlide: { type: Array, default: () => [] },
  /** SPEC-099: the selected slide indexes (the shared multi-selection). Empty
   * = just the current slide. */
  selectedIndexes: { type: Array, default: () => [] },
  /** SPEC-099: the shared useSlideSelection controller, when the view has one. */
  selection: { type: Object, default: null },
});

const emit = defineEmits(["select", "add", "command", "reorder"]);

/** Target thumbnail width in px; the stage scales down to fit it. */
const THUMB_WIDTH_PX = 168;

const slides = computed(() => asSlides(props.slides));
const fieldsBySlide = computed(() => asStageFieldsBySlide(props.fieldsBySlide));
const scale = computed(() => filmstripThumbScale(props.widthCm, THUMB_WIDTH_PX));
const selected = computed(() => (props.selectedIndexes.length > 0 ? props.selectedIndexes : [props.currentIndex]));

const listboxRef = useTemplateRef("listbox");

// The slide whose context menu is open (-1 = none), and where to place it.
const menuIndex = ref(-1);
const menuPos = ref({ x: 0, y: 0 });
// Drag-reorder state; dropTarget drives the drop indicator.
const dragFrom = ref(-1);
const dropTarget = ref(-1);
const announcement = ref("");

function label(slide = slides.value[0], index = 0) {
  // SPEC-104: the struck-through number is visual only, so say it.
  return `${filmstripLabel(slide, index, slides.value.length)}${slide.hidden ? " (hidden)" : ""}`;
}

// SPEC-104: the slide menu's Hide/Show entry names the state it switches to.
const menuItems = computed(() => filmstripMenuItems(slides.value[menuIndex.value]?.hidden === true));

function optionId(index = 0) {
  return `filmstrip-option-${index}`;
}

function select(index = 0) {
  if (index >= 0 && index < slides.value.length) {
    emit("select", index);
  }
}

function onOptionClick(index = 0, event = new MouseEvent("click")) {
  if (props.selection) {
    props.selection.click(index, gestureModifiers(event));
  } else {
    select(index);
  }
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  // Shift+F10 / the Menu (ContextMenu) key opens the current thumbnail's menu
  // (AC-24). Handled here on the focusable listbox — the options themselves
  // are not focusable, so the aria-activedescendant slide is the target.
  if (event.key === "ContextMenu" || (event.shiftKey && event.key === "F10")) {
    event.preventDefault();
    const index = props.currentIndex;
    const pos = anchorPoint(listboxRef.value, `#${optionId(index)}`);
    openMenuAt(index, pos.x, pos.y);
    return;
  }
  // The shared selection keys (SPEC-099): arrows/Home/End (Shift extends),
  // Cmd/Ctrl+A, Escape — kept from the stage's own window-level shortcuts.
  if (props.selection) {
    if (props.selection.onKeydown(event)) {
      event.preventDefault();
      event.stopPropagation();
    }
    return;
  }
  let next = null;
  switch (event.key) {
    case "ArrowDown":
    case "ArrowRight":
      next = props.currentIndex + 1;
      break;
    case "ArrowUp":
    case "ArrowLeft":
      next = props.currentIndex - 1;
      break;
    case "Home":
      next = 0;
      break;
    case "End":
      next = slides.value.length - 1;
      break;
    default:
      return;
  }
  if (next !== null && next >= 0 && next < slides.value.length) {
    event.preventDefault();
    select(next);
  }
}

// --- context menu ----------------------------------------------------------

function openMenuAt(index = 0, x = 0, y = 0) {
  if (!selected.value.includes(index)) {
    if (props.selection) {
      props.selection.click(index);
    } else {
      select(index);
    }
  }
  menuIndex.value = index;
  menuPos.value = { x, y };
}

function openMenuFromButton(index = 0, event = new MouseEvent("click")) {
  openMenuAt(index, event.clientX, event.clientY);
}

function onContextMenu(index = 0, event = new MouseEvent("contextmenu")) {
  event.preventDefault();
  openMenuAt(index, event.clientX, event.clientY);
}

function closeMenu(refocus = true) {
  const index = menuIndex.value;
  menuIndex.value = -1;
  if (refocus && index >= 0) {
    void nextTick(() => listboxRef.value?.focus());
  }
}

function runMenu(action = "") {
  const index = menuIndex.value;
  menuIndex.value = -1;
  if (index < 0) {
    return;
  }
  if (action === "move-up" || action === "move-down") {
    announceMove(index, action === "move-up" ? index - 1 : index + 1);
  }
  emit("command", { action, index });
}

// --- drag reorder ----------------------------------------------------------
function onDragStart(index = 0, event = new DragEvent("dragstart")) {
  dragFrom.value = index;
  dropTarget.value = index;
  if (event.dataTransfer) {
    event.dataTransfer.effectAllowed = "move";
    event.dataTransfer.setData("text/plain", String(index));
  }
}

function onDragOver(index = 0, event = new DragEvent("dragover")) {
  if (dragFrom.value < 0) {
    return;
  }
  event.preventDefault();
  dropTarget.value = index;
}

function onDrop(index = 0, event = new DragEvent("drop")) {
  event.preventDefault();
  const from = dragFrom.value;
  if (from >= 0 && from !== index) {
    announceMove(from, index);
    emit("reorder", { from, to: index });
  }
  dragFrom.value = -1;
  dropTarget.value = -1;
}

function onDragEnd() {
  // Fires whether or not a drop landed — dropping outside cancels (AC-39).
  dragFrom.value = -1;
  dropTarget.value = -1;
}

function announceMove(from = 0, to = 0) {
  const clamped = Math.min(Math.max(to, 0), slides.value.length - 1);
  announcement.value = `Slide ${from + 1} moved to position ${clamped + 1}`;
}
</script>

<template>
  <div
    class="filmstrip"
    data-testid="slide-filmstrip-panel"
  >
    <div
      ref="listbox"
      class="filmstrip-list"
      role="listbox"
      tabindex="0"
      aria-label="Slides"
      :aria-multiselectable="selection ? 'true' : undefined"
      :aria-activedescendant="optionId(currentIndex)"
      data-testid="slide-filmstrip"
      @keydown="onKeydown"
    >
      <div
        v-for="(slide, index) in slides"
        :id="optionId(index)"
        :key="slide.id"
        class="filmstrip-option"
        :class="{
          selected: index === currentIndex,
          'multi-selected': index !== currentIndex && selected.includes(index),
          'drop-target': index === dropTarget && dragFrom !== index,
          'filmstrip-hidden': slide.hidden === true,
        }"
        role="option"
        :aria-selected="selected.includes(index)"
        :aria-label="label(slide, index)"
        data-testid="filmstrip-slide"
        draggable="true"
        @click="onOptionClick(index, $event)"
        @contextmenu="onContextMenu(index, $event)"
        @dragstart="onDragStart(index, $event)"
        @dragover="onDragOver(index, $event)"
        @drop="onDrop(index, $event)"
        @dragend="onDragEnd"
      >
        <SlideThumb
          :slide="slide"
          :index="index"
          :width-cm="widthCm"
          :height-cm="heightCm"
          :media="media"
          :scale="scale"
          :fields="fieldsBySlide[index] || []"
          :theme="theme"
          :is-hidden="slide.hidden === true"
        />
        <button
          type="button"
          class="slide-menu-button"
          data-testid="presentation-slide-menu"
          aria-haspopup="menu"
          :aria-expanded="menuIndex === index"
          aria-label="Slide options"
          @click.stop="openMenuFromButton(index, $event)"
        >
          ⋯
        </button>
      </div>
    </div>

    <button
      type="button"
      class="add-slide"
      data-testid="presentation-add-slide"
      @click="emit('add')"
    >
      + Add slide
    </button>

    <SlideContextMenu
      v-if="menuIndex >= 0"
      :items="menuItems"
      :x="menuPos.x"
      :y="menuPos.y"
      @run="runMenu"
      @close="closeMenu"
    />

    <div
      class="sr-only"
      aria-live="polite"
      data-testid="filmstrip-live"
    >
      {{ announcement }}
    </div>
  </div>
</template>

<style scoped>
.filmstrip {
  display: flex;
  flex-direction: column;
  min-height: 0;
}

.filmstrip-list {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-3);
  overflow-y: auto;
  outline: none;
  flex: 1 1 auto;
  min-height: 0;
}

.filmstrip-list:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.filmstrip-option {
  position: relative;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  cursor: pointer;
  border-radius: var(--radius-sm);
  padding: var(--space-1);
  border: var(--border-width) solid transparent;
}

.filmstrip-option.selected {
  border-color: var(--accent);
  background: var(--surface-raised);
}

.filmstrip-option.multi-selected {
  border-color: var(--accent);
}

.filmstrip-option.drop-target {
  border-top-color: var(--accent);
  background: var(--surface-raised);
}

.slide-menu-button {
  position: absolute;
  top: var(--space-1);
  right: var(--space-1);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  color: var(--text-secondary);
  cursor: pointer;
  line-height: 1;
  padding: 0 var(--space-1);
}

.slide-menu-button:hover,
.slide-menu-button:focus-visible {
  color: var(--text-primary);
}

.add-slide {
  margin: var(--space-3);
  padding: var(--space-2);
  border: var(--border-width) dashed var(--border-default);
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  font-weight: var(--weight-medium);
}

.add-slide:hover {
  background: var(--surface-raised);
}

.sr-only {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>

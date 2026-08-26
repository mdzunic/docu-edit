<script setup lang="ts">
// The slide filmstrip (SPEC-079 listbox + SPEC-080 slide management): a
// listbox of numbered thumbnails, each a scaled-down non-interactive
// SlideStage, plus an "Add slide" button and per-slide management — a context
// menu (right-click, a ⋯ affordance, and Shift+F10 / the Menu key on a focused
// thumbnail) with New slide / Duplicate / Delete / Move up / Move down, and
// pointer drag-reorder with a visible drop indicator (dropping outside
// cancels). Structural intents are emitted up to PresentationView, which owns
// the commands + undo. Runtime prop/emit declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import { computed, nextTick, ref, useTemplateRef } from "vue";
import SlideStage from "./SlideStage.vue";
import { anchorPoint, asSlides, filmstripLabel, filmstripThumbScale, focusSelector, menuButtons } from "./slideRender";

const props = defineProps({
  slides: { type: Array, default: () => [] },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  currentIndex: { type: Number, default: 0 },
});

const emit = defineEmits(["select", "add", "command", "reorder"]);

/** Target thumbnail width in px; the stage scales down to fit it. */
const THUMB_WIDTH_PX = 168;

const MENU_ITEMS = [
  { action: "new", label: "New slide", testid: "slide-menu-new" },
  { action: "duplicate", label: "Duplicate slide", testid: "slide-menu-duplicate" },
  { action: "delete", label: "Delete slide", testid: "slide-menu-delete" },
  { action: "move-up", label: "Move up", testid: "slide-menu-move-up" },
  { action: "move-down", label: "Move down", testid: "slide-menu-move-down" },
];

const slides = computed(() => asSlides(props.slides));
const scale = computed(() => filmstripThumbScale(props.widthCm, THUMB_WIDTH_PX));

const listboxRef = useTemplateRef("listbox");
const menuRef = useTemplateRef("menu");

// The slide whose context menu is open (-1 = none), and where to place it.
const menuIndex = ref(-1);
const menuPos = ref({ x: 0, y: 0 });
// Drag-reorder state; dropTarget drives the drop indicator.
const dragFrom = ref(-1);
const dropTarget = ref(-1);
const announcement = ref("");

function label(slide = slides.value[0], index = 0) {
  return filmstripLabel(slide, index, slides.value.length);
}

function optionId(index = 0) {
  return `filmstrip-option-${index}`;
}

function select(index = 0) {
  if (index >= 0 && index < slides.value.length) {
    emit("select", index);
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
  select(index);
  menuIndex.value = index;
  menuPos.value = { x, y };
  void nextTick(() => focusSelector(menuRef.value, '[role="menuitem"]'));
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

function onMenuKeydown(event = new KeyboardEvent("keydown")) {
  const items = menuButtons(menuRef.value);
  const active = items.findIndex((item) => item === document.activeElement);
  if (event.key === "ArrowDown") {
    event.preventDefault();
    items[(active + 1) % items.length]?.focus();
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    items[(active - 1 + items.length) % items.length]?.focus();
  } else if (event.key === "Escape") {
    event.preventDefault();
    closeMenu();
  } else if (event.key === "Home") {
    event.preventDefault();
    items[0]?.focus();
  } else if (event.key === "End") {
    event.preventDefault();
    items[items.length - 1]?.focus();
  }
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
      :aria-activedescendant="optionId(currentIndex)"
      data-testid="slide-filmstrip"
      @keydown="onKeydown"
    >
      <div
        v-for="(slide, index) in slides"
        :id="optionId(index)"
        :key="slide.id"
        class="filmstrip-option"
        :class="{ selected: index === currentIndex, 'drop-target': index === dropTarget && dragFrom !== index }"
        role="option"
        :aria-selected="index === currentIndex"
        :aria-label="label(slide, index)"
        data-testid="filmstrip-slide"
        draggable="true"
        @click="select(index)"
        @contextmenu="onContextMenu(index, $event)"
        @dragstart="onDragStart(index, $event)"
        @dragover="onDragOver(index, $event)"
        @drop="onDrop(index, $event)"
        @dragend="onDragEnd"
      >
        <span
          class="thumb-number"
          aria-hidden="true"
        >{{ index + 1 }}</span>
        <div class="thumb-stage">
          <SlideStage
            :slide="slide"
            :width-cm="widthCm"
            :height-cm="heightCm"
            :media="media"
            :scale="scale"
            :interactive="false"
          />
        </div>
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

    <div
      v-if="menuIndex >= 0"
      ref="menu"
      class="slide-context-menu"
      role="menu"
      data-testid="slide-context-menu"
      :style="{ left: `${menuPos.x}px`, top: `${menuPos.y}px` }"
      tabindex="-1"
      @keydown="onMenuKeydown"
    >
      <button
        v-for="item in MENU_ITEMS"
        :key="item.action"
        type="button"
        class="slide-context-menu-item"
        role="menuitem"
        :data-testid="item.testid"
        @click="runMenu(item.action)"
      >
        {{ item.label }}
      </button>
    </div>
    <div
      v-if="menuIndex >= 0"
      class="menu-backdrop"
      data-testid="slide-menu-backdrop"
      @click="closeMenu(false)"
      @contextmenu.prevent="closeMenu(false)"
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

.filmstrip-option.drop-target {
  border-top-color: var(--accent);
  background: var(--surface-raised);
}

.thumb-number {
  flex: none;
  width: 1.5rem;
  text-align: right;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.thumb-stage {
  /* The thumbnail stage is decorative and non-interactive. */
  pointer-events: none;
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

.slide-context-menu {
  position: fixed;
  z-index: 40;
  display: flex;
  flex-direction: column;
  min-width: 10rem;
  padding: var(--space-1);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.slide-context-menu-item {
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.slide-context-menu-item:hover {
  background: var(--surface-sunken);
}

.slide-context-menu-item:focus-visible {
  background: var(--surface-sunken);
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 39;
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

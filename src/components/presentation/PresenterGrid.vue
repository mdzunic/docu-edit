<script setup lang="ts">
// The presenter view's slide grid (SPEC-100): every slide as a numbered
// thumbnail — the shared SlideThumb (a non-interactive SlideStage) at
// filmstripThumbScale, named by filmstripLabel, so there is no second
// thumbnail maths — in a listbox with roving focus: one option is tabbable,
// the arrows move it (Up/Down by a row), Home/End jump, Enter/Space or a click
// chooses, Escape (or g) closes. The view turns `choose` into the owner's
// goTo, so the jump reaches both windows. Runtime declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts).
import { computed, nextTick, onMounted, ref, useTemplateRef } from "vue";
import SlideThumb from "./SlideThumb.vue";
import { asSlides, filmstripLabel, filmstripThumbScale } from "./slideRender";

const props = defineProps({
  slides: { type: Array, default: () => [] },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  theme: { type: Object, default: null },
  currentIndex: { type: Number, default: 0 },
});

const emit = defineEmits(["choose", "close"]);

/** Target thumbnail width in px; each stage scales down to fit it. */
const GRID_THUMB_PX = 192;

const slides = computed(() => asSlides(props.slides));
const scale = computed(() => filmstripThumbScale(props.widthCm, GRID_THUMB_PX));
const focusIndex = ref(props.currentIndex);
const listbox = useTemplateRef("listbox");

function options() {
  return Array.from(listbox.value?.querySelectorAll('[role="option"]') ?? []);
}

function focusOption(index = 0) {
  const count = slides.value.length;
  if (count === 0) return;
  focusIndex.value = Math.min(Math.max(index, 0), count - 1);
  void nextTick(() => {
    const option = options()[focusIndex.value];
    if (option instanceof HTMLElement) option.focus();
  });
}

/** How many thumbnails sit on the first row — the step for Up/Down. */
function columns() {
  const all = options();
  const first = all[0];
  if (!(first instanceof HTMLElement)) return 1;
  return Math.max(1, all.filter((option) => option instanceof HTMLElement && option.offsetTop === first.offsetTop).length);
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  const key = event.key;
  let next = null;
  if (key === "ArrowRight" || key === "ArrowLeft") {
    next = focusIndex.value + (key === "ArrowRight" ? 1 : -1);
  } else if (key === "ArrowDown" || key === "ArrowUp") {
    next = focusIndex.value + (key === "ArrowDown" ? columns() : -columns());
  } else if (key === "Home" || key === "End") {
    next = key === "Home" ? 0 : slides.value.length - 1;
  } else if (key === "Enter" || key === " ") {
    emit("choose", focusIndex.value);
  } else if (key === "Escape" || key === "g" || key === "G") {
    emit("close");
  } else {
    return;
  }
  // Handled here: the presenter's window-level keys must not see it too.
  event.preventDefault();
  event.stopPropagation();
  if (next !== null) focusOption(next);
}

onMounted(() => focusOption(props.currentIndex));
</script>

<template>
  <div
    class="grid-overlay"
    role="dialog"
    aria-label="All slides"
  >
    <div
      ref="listbox"
      class="grid-list"
      role="listbox"
      aria-label="Slides"
      data-testid="presenter-grid"
      @keydown="onKeydown"
    >
      <div
        v-for="(slide, index) in slides"
        :key="slide.id"
        class="grid-option"
        :class="{ current: index === currentIndex }"
        role="option"
        :tabindex="index === focusIndex ? 0 : -1"
        :aria-selected="index === currentIndex"
        :aria-label="filmstripLabel(slide, index, slides.length)"
        data-testid="presenter-grid-slide"
        @click="emit('choose', index)"
        @focus="focusIndex = index"
      >
        <SlideThumb
          :slide="slide"
          :index="index"
          :width-cm="widthCm"
          :height-cm="heightCm"
          :media="media"
          :scale="scale"
          :theme="theme"
          :is-hidden="slide.hidden === true"
          stacked
        />
      </div>
    </div>
  </div>
</template>

<style scoped>
.grid-overlay {
  position: fixed;
  inset: 0;
  z-index: 30;
  overflow-y: auto;
  padding: var(--space-6);
  background: var(--surface-overlay);
}

.grid-list {
  display: flex;
  flex-wrap: wrap;
  align-content: flex-start;
  gap: var(--space-4);
}

.grid-option {
  padding: var(--space-2);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  cursor: pointer;
}

.grid-option.current {
  border-color: var(--accent);
}

.grid-option:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: var(--border-width);
}
</style>

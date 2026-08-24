<script setup lang="ts">
// The slide filmstrip (SPEC-079): a proper listbox of numbered thumbnails —
// each a scaled-down, non-interactive SlideStage — for tests and assistive
// tech. role="listbox"/"option", aria-selected, each option labelled
// "Slide N of M — <title or 'untitled'>", keyboard nav with
// ArrowUp/ArrowDown/Home/End. Dumb: props in, `select` out, no store access.
import { computed } from "vue";
import { cmToPx } from "../../lib/core/units";
import type { MediaPart } from "../../lib/core/media";
import type { Slide } from "../../lib/presentation/model";
import { slideOutline } from "../../lib/presentation/slideModel";
import SlideStage from "./SlideStage.vue";

const props = defineProps<{
  slides: Slide[];
  widthCm: number;
  heightCm: number;
  media?: MediaPart[];
  currentIndex: number;
}>();

const emit = defineEmits<{ (event: "select", index: number): void }>();

/** Target thumbnail width in px; the stage scales down to fit it. */
const THUMB_WIDTH_PX = 168;

const scale = computed(() => THUMB_WIDTH_PX / Math.max(1, cmToPx(props.widthCm)));

function optionLabel(slide: Slide, index: number): string {
  const outline = slideOutline(slide);
  const name = outline.title && outline.title.trim().length > 0 ? outline.title : "untitled";
  return `Slide ${index + 1} of ${props.slides.length} — ${name}`;
}

function optionId(index: number): string {
  return `filmstrip-option-${index}`;
}

function select(index: number): void {
  if (index >= 0 && index < props.slides.length) {
    emit("select", index);
  }
}

function onKeydown(event: KeyboardEvent): void {
  let next: number | null = null;
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
      next = props.slides.length - 1;
      break;
    default:
      return;
  }
  if (next !== null && next >= 0 && next < props.slides.length) {
    event.preventDefault();
    select(next);
  }
}
</script>

<template>
  <div
    class="filmstrip"
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
      :class="{ selected: index === currentIndex }"
      role="option"
      :aria-selected="index === currentIndex"
      :aria-label="optionLabel(slide, index)"
      data-testid="filmstrip-slide"
      @click="select(index)"
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
        />
      </div>
    </div>
  </div>
</template>

<style scoped>
.filmstrip {
  display: flex;
  flex-direction: column;
  gap: var(--space-3, 0.75rem);
  padding: var(--space-3, 0.75rem);
  overflow-y: auto;
  outline: none;
}

.filmstrip:focus-visible {
  box-shadow: inset 0 0 0 2px var(--focus-ring, #4472c4);
}

.filmstrip-option {
  display: flex;
  align-items: center;
  gap: var(--space-2, 0.5rem);
  cursor: pointer;
  border-radius: var(--radius-sm, 4px);
  padding: var(--space-1, 0.25rem);
  border: var(--border-width, 1px) solid transparent;
}

.filmstrip-option.selected {
  border-color: var(--accent, #4472c4);
  background: var(--surface-raised, rgba(68, 114, 196, 0.08));
}

.thumb-number {
  flex: none;
  width: 1.5rem;
  text-align: right;
  color: var(--text-secondary, #666);
  font-size: var(--text-sm, 0.85rem);
  font-weight: var(--weight-bold, 700);
}

.thumb-stage {
  /* The thumbnail stage is decorative and non-interactive. */
  pointer-events: none;
}
</style>

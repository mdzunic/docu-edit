<script setup lang="ts">
// The slide filmstrip (SPEC-079): a proper listbox of numbered thumbnails —
// each a scaled-down, non-interactive SlideStage — for tests and assistive
// tech. role="listbox"/"option", aria-selected, each option labelled
// "Slide N of M — <title or 'untitled'>", keyboard nav with
// ArrowUp/ArrowDown/Home/End. Runtime prop/emit declarations, not
// defineProps<{...}>() generics (espree lints .vue scripts); params take a
// default value so TS infers their type without an annotation.
import { computed } from "vue";
import SlideStage from "./SlideStage.vue";
import { asSlides, filmstripLabel, filmstripThumbScale } from "./slideRender";

const props = defineProps({
  slides: { type: Array, default: () => [] },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  currentIndex: { type: Number, default: 0 },
});

const emit = defineEmits(["select"]);

/** Target thumbnail width in px; the stage scales down to fit it. */
const THUMB_WIDTH_PX = 168;

const slides = computed(() => asSlides(props.slides));
const scale = computed(() => filmstripThumbScale(props.widthCm, THUMB_WIDTH_PX));

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
      :aria-label="label(slide, index)"
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
  gap: var(--space-3);
  padding: var(--space-3);
  overflow-y: auto;
  outline: none;
}

.filmstrip:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.filmstrip-option {
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
</style>

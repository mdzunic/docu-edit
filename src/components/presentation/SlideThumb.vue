<script setup lang="ts">
// One slide thumbnail — its 1-based number beside a scaled-down,
// non-interactive SlideStage — extracted from SlideFilmstrip (SPEC-099) so the
// filmstrip, the slide sorter and the Reuse slides picker render the very
// same thing. The host owns the option semantics (role, aria-selected, click,
// drag); this owns only the picture. `stacked` puts the number under the
// stage (the sorter's grid card) instead of beside it (the filmstrip row).
// Runtime prop declarations, not defineProps<{...}>() generics (espree lints
// .vue scripts).
import SlideStage from "./SlideStage.vue";

defineProps({
  slide: { type: Object, required: true },
  /** 0-based; rendered 1-based. */
  index: { type: Number, default: 0 },
  widthCm: { type: Number, required: true },
  heightCm: { type: Number, required: true },
  media: { type: Array, default: () => [] },
  scale: { type: Number, required: true },
  /** The slide's pre-resolved deck field band (SPEC-093). */
  fields: { type: Array, default: () => [] },
  /** SPEC-094: the deck's theme. */
  theme: { type: Object, default: null },
  stacked: { type: Boolean, default: false },
});
</script>

<template>
  <div
    class="slide-thumb"
    :class="{ stacked }"
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
        :fields="fields"
        :theme="theme"
      />
    </div>
  </div>
</template>

<style scoped>
.slide-thumb {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.slide-thumb.stacked {
  flex-direction: column-reverse;
  align-items: flex-start;
  gap: var(--space-1);
}

.thumb-number {
  flex: none;
  width: 1.5rem;
  text-align: right;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.stacked .thumb-number {
  width: auto;
  text-align: left;
}

.thumb-stage {
  /* The thumbnail stage is decorative and non-interactive. */
  pointer-events: none;
}
</style>

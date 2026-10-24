<script setup lang="ts">
// The slide stage's proofing layer (SPEC-098): find highlights under SPEC-017's
// testids (`find-highlights` / `find-highlight`, `.current` on the current
// hit) and spell-check squiggles under SPEC-035's (`spell-underlines` /
// `spell-underline`). It paints INSIDE the scaled stage surface, so its rects
// are surface px (proofingRects.ts measures them) and the surface's transform
// keeps them registered with the text; `scale` only keeps the squiggle a
// constant on-screen thickness. Every rect is inert — the layer sits above the
// frames, and catching pointer events would swallow the clicks and
// double-clicks that select and edit them — so a right-click is hit-tested
// against the underline rects on the surface instead and emitted as
// `spell-context` (underline's issue index, clientX, clientY). Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed, onBeforeUnmount, onMounted, useTemplateRef } from "vue";
import { asHighlightRects, asSpellUnderlineRects, surfacePoint } from "./proofingRects";

const props = defineProps({
  highlightRects: { type: Array, default: () => [] },
  underlineRects: { type: Array, default: () => [] },
  scale: { type: Number, default: 1 },
});

const emit = defineEmits(["spell-context"]);

const layer = useTemplateRef("layer");
const highlights = computed(() => asHighlightRects(props.highlightRects));
const underlines = computed(() => asSpellUnderlineRects(props.underlineRects));
/** The 6×4 squiggle tile at a constant on-screen size whatever the zoom. */
const squiggleSize = computed(() => `${6 / (props.scale || 1)}px ${4 / (props.scale || 1)}px`);

function box(rect = { left: 0, top: 0, width: 0, height: 0 }) {
  return { left: `${rect.left}px`, top: `${rect.top}px`, width: `${rect.width}px`, height: `${rect.height}px` };
}

function onContextMenu(event = new MouseEvent("contextmenu")) {
  const surface = layer.value?.parentElement;
  if (!surface || underlines.value.length === 0) {
    return;
  }
  const { x, y } = surfacePoint(surface, event.clientX, event.clientY);
  const hit = underlines.value.find((rect) => x >= rect.left && x <= rect.left + rect.width && y >= rect.top && y <= rect.top + rect.height);
  if (hit) {
    event.preventDefault();
    emit("spell-context", hit.issue, event.clientX, event.clientY);
  }
}

onMounted(() => layer.value?.parentElement?.addEventListener("contextmenu", onContextMenu));
onBeforeUnmount(() => layer.value?.parentElement?.removeEventListener("contextmenu", onContextMenu));
</script>

<template>
  <div
    ref="layer"
    class="proofing-layer print-exclude"
    aria-hidden="true"
  >
    <div
      v-if="highlights.length > 0"
      class="find-highlight-layer"
      data-testid="find-highlights"
    >
      <div
        v-for="(rect, index) in highlights"
        :key="index"
        class="find-highlight"
        :class="{ current: rect.current }"
        data-testid="find-highlight"
        :data-current="rect.current ? 'true' : 'false'"
        :style="box(rect)"
      />
    </div>
    <div
      v-if="underlines.length > 0"
      class="spell-underline-layer"
      data-testid="spell-underlines"
    >
      <div
        v-for="(rect, index) in underlines"
        :key="index"
        class="spell-underline"
        data-testid="spell-underline"
        :data-word="rect.word"
        :style="[box(rect), { backgroundSize: squiggleSize }]"
      />
    </div>
  </div>
</template>

<style scoped>
/* Above the frames (their fills are opaque) and under the selection chrome —
   by DOM order in SlideStage's surface, so no z-index; inert throughout (see
   the header). */
.proofing-layer {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

.find-highlight-layer,
.spell-underline-layer {
  position: absolute;
  inset: 0;
}

/* Multiplied over the text rather than painted behind it: a frame's fill
   would hide a highlight underneath. */
.find-highlight {
  position: absolute;
  background: rgb(255 214 0 / 0.45);
  border-radius: 1px;
  mix-blend-mode: multiply;
}

.find-highlight.current {
  background: rgb(255 145 0 / 0.65);
  outline: 1px solid #f57c00;
}

/* SPEC-035's squiggle, at the bottom of the word's glyph box. */
.spell-underline {
  position: absolute;
  background-repeat: repeat-x;
  background-position: left bottom;
  background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='6' height='4' viewBox='0 0 6 4'%3E%3Cpath d='M0 3 Q1.5 0 3 3 T6 3' fill='none' stroke='%23d93025' stroke-width='1'/%3E%3C/svg%3E");
}
</style>

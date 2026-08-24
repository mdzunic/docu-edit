<script setup lang="ts">
// One slide frame rendered absolutely-positioned in slide space (SPEC-079):
// text → nested divs/spans, images → <img> from the media resolver, shapes →
// a devicePixelRatio-crisp <canvas> painted by paintShapeIntoCanvas plus a
// centred text overlay, groups → this component recursively, foreign → a
// labelled dashed placeholder box. Runtime prop declarations, not
// defineProps<{...}>() generics — this repo lints .vue scripts with espree,
// which rejects TS-only syntax; every typed transform lives in slideRender.ts.
import { computed, onMounted, useTemplateRef, watch } from "vue";
import { cmToPx } from "../../lib/core/units";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import {
  asResolver,
  asSlideFrame,
  framePositionStyle,
  groupChildren,
  imageSrc,
  isForeignVisible,
  foreignLabel,
  markedParagraphs,
  paragraphStyle,
  runStyle,
  shapeVisualOf,
  textFrameFill,
} from "./slideRender";

const props = defineProps({
  /** The SlideFrame to render (presentation model). */
  frame: { type: Object, required: true },
  /** Resolves a media id to a `data:` URI (or "" when unresolved). */
  resolveMedia: { type: Function, required: true },
});

const frame = computed(() => asSlideFrame(props.frame));
const kind = computed(() => frame.value.kind);
const position = computed(() => framePositionStyle(frame.value));
const paragraphs = computed(() => markedParagraphs(frame.value));
const textFill = computed(() => textFrameFill(frame.value));
const children = computed(() => groupChildren(frame.value));
const foreignVisible = computed(() => isForeignVisible(frame.value));
const foreignText = computed(() => foreignLabel(frame.value));
const imgSrc = computed(() => imageSrc(frame.value, asResolver(props.resolveMedia)));

const canvasEl = useTemplateRef("canvasEl");

function paint() {
  const el = canvasEl.value;
  const visual = shapeVisualOf(frame.value);
  if (!el || !visual) {
    return;
  }
  try {
    paintShapeIntoCanvas(el, visual, cmToPx(frame.value.widthCm), cmToPx(frame.value.heightCm));
  } catch {
    // A host without a real 2D canvas (jsdom) can't paint; rendering must
    // never break the view.
  }
}

// paintShapeIntoCanvas sizes the backing store at devicePixelRatio, so the
// shape stays crisp at dpr ≥ 2 (AC-43); repaint whenever the frame changes.
onMounted(paint);
watch(() => props.frame, paint, { deep: true });
</script>

<template>
  <!-- Text frame -->
  <div
    v-if="kind === 'text'"
    class="frame text-frame"
    :style="{ ...position, background: textFill }"
    data-testid="slide-text-frame"
  >
    <div
      v-for="(entry, index) in paragraphs"
      :key="index"
      class="paragraph"
      :style="paragraphStyle(entry.para)"
      data-testid="slide-paragraph"
    >
      <span
        v-if="entry.marker"
        class="bullet-marker"
        aria-hidden="true"
      >{{ entry.marker }} </span>
      <span
        v-for="(run, runIndex) in entry.para.runs"
        :key="runIndex"
        :style="runStyle(run)"
      >{{ run.text }}</span>
    </div>
  </div>

  <!-- Image frame -->
  <div
    v-else-if="kind === 'image'"
    class="frame"
    :style="position"
  >
    <img
      class="frame-image"
      :src="imgSrc"
      alt=""
    >
  </div>

  <!-- Shape frame -->
  <div
    v-else-if="kind === 'shape'"
    class="frame shape-frame"
    :style="position"
    data-testid="slide-shape-frame"
  >
    <canvas
      ref="canvasEl"
      class="shape-canvas"
    />
    <div
      v-if="paragraphs.length > 0"
      class="shape-text"
    >
      <div
        v-for="(entry, index) in paragraphs"
        :key="index"
        class="paragraph"
        :style="{ textAlign: 'center' }"
      >
        <span
          v-for="(run, runIndex) in entry.para.runs"
          :key="runIndex"
          :style="runStyle(run)"
        >{{ run.text }}</span>
      </div>
    </div>
  </div>

  <!-- Group frame: children already in slide space, rendered read-only -->
  <div
    v-else-if="kind === 'group'"
    class="frame group-frame"
    :style="position"
    data-testid="slide-group-frame"
  >
    <SlideFrameView
      v-for="child in children"
      :key="child.id"
      :frame="child"
      :resolve-media="resolveMedia"
    />
  </div>

  <!-- Foreign placeholder box -->
  <div
    v-else-if="foreignVisible"
    class="frame foreign-frame"
    :style="position"
    data-testid="slide-foreign-frame"
  >
    <span class="foreign-label">{{ foreignText }}</span>
  </div>
</template>

<style scoped>
.frame {
  box-sizing: border-box;
  overflow: visible;
}

.text-frame {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
}

.paragraph {
  margin: 0;
  line-height: 1.2;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.bullet-marker {
  color: inherit;
}

.frame-image {
  width: 100%;
  height: 100%;
  object-fit: fill;
  display: block;
}

.shape-canvas {
  position: absolute;
  top: 0;
  left: 0;
}

.shape-text {
  position: absolute;
  inset: 0;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  pointer-events: none;
}

.group-frame {
  /* The group box itself paints nothing; children position against the stage
     via their own absolute slide-space coordinates. */
  position: absolute;
  left: 0 !important;
  top: 0 !important;
  width: 0 !important;
  height: 0 !important;
  overflow: visible;
}

.foreign-frame {
  border: 1px dashed var(--border-subtle, #b0b0b0);
  display: flex;
  align-items: center;
  justify-content: center;
}

.foreign-label {
  color: var(--text-secondary, #666);
  font-size: 0.75rem;
  font-family: system-ui, sans-serif;
}
</style>

<script setup lang="ts">
// One slide frame rendered absolutely-positioned in slide space (SPEC-079):
// text → nested divs/spans with inline styles, images → <img> from the media
// resolver, shapes → a devicePixelRatio-crisp <canvas> painted by
// paintShapeIntoCanvas plus a centred text overlay, groups → this component
// recursively, foreign → a labelled dashed placeholder box. Dumb and
// read-only: props in, nothing emitted, no store access.
import { computed, onMounted, onBeforeUnmount, ref, watch } from "vue";
import { cmToPx } from "../../lib/core/units";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import type { ShapeVisual } from "../../lib/core/shapes";
import type { SlideFrame, SlideParagraph } from "../../lib/presentation/model";

const props = defineProps<{
  frame: SlideFrame;
  /** Resolves a media id to a `data:` URI (or "" when unresolved). */
  resolveMedia: (mediaId: string) => string;
}>();

/** pt → px at 96/72, the same ratio the text editor uses. */
const PT_TO_PX = 96 / 72;
const LEVEL_INDENT_CM = 0.95;

function px(cm: number): string {
  return `${cmToPx(cm)}px`;
}

const positionStyle = computed(() => {
  const f = props.frame;
  const style: Record<string, string> = {
    position: "absolute",
    left: px(f.xCm),
    top: px(f.yCm),
    width: px(f.widthCm),
    height: px(f.heightCm),
  };
  if (f.rotationDeg) {
    style.transform = `rotate(${f.rotationDeg}deg)`;
    style.transformOrigin = "center";
  }
  return style;
});

/** A text body's paragraphs annotated with their bullet marker (numbered
 * bullets restart per contiguous same-level run). */
function markedParagraphs(paragraphs: SlideParagraph[]): { para: SlideParagraph; marker: string }[] {
  const counts: Record<number, number> = {};
  let prev: SlideParagraph | undefined;
  return paragraphs.map((para) => {
    let marker = "";
    if (para.bullet === "number") {
      const level = para.level ?? 0;
      if (prev && prev.bullet === "number" && (prev.level ?? 0) === level) {
        counts[level] = (counts[level] ?? 0) + 1;
      } else {
        counts[level] = 1;
      }
      marker = `${counts[level]}.`;
    } else if (para.bullet === "disc") {
      marker = "•";
    }
    prev = para;
    return { para, marker };
  });
}

const textParagraphs = computed(() => {
  if (props.frame.kind === "text") {
    return markedParagraphs(props.frame.paragraphs);
  }
  if (props.frame.kind === "shape" && props.frame.paragraphs) {
    return markedParagraphs(props.frame.paragraphs);
  }
  return [];
});

function paragraphStyle(para: SlideParagraph): Record<string, string> {
  return {
    textAlign: para.align ?? "left",
    marginLeft: para.level ? px(para.level * LEVEL_INDENT_CM) : "0",
  };
}

function runStyle(run: { bold?: true; italic?: true; underline?: true; fontFamily?: string; sizePt?: number; color?: string }): Record<string, string> {
  const style: Record<string, string> = {};
  if (run.bold) style.fontWeight = "700";
  if (run.italic) style.fontStyle = "italic";
  if (run.underline) style.textDecoration = "underline";
  if (run.fontFamily) style.fontFamily = `'${run.fontFamily}', sans-serif`;
  style.fontSize = `${(run.sizePt ?? 18) * PT_TO_PX}px`;
  style.color = run.color ?? "#000000";
  return style;
}

const frameFill = computed(() => {
  const f = props.frame;
  if ((f.kind === "text") && f.fill && f.fill.type === "solid") {
    return f.fill.color;
  }
  return "transparent";
});

// --- shape canvas ----------------------------------------------------------

const canvas = ref<HTMLCanvasElement | null>(null);

const shapeVisual = computed<ShapeVisual | null>(() => {
  const f = props.frame;
  if (f.kind !== "shape") {
    return null;
  }
  const visual: ShapeVisual = { geometry: f.geometry };
  if (f.fill && f.fill.type === "solid") {
    visual.fill = f.fill.color;
  }
  if (f.stroke) {
    visual.outline = { color: f.stroke.color, widthPt: f.stroke.widthPt, dash: "solid" };
  }
  if (f.adjustments && f.adjustments.length > 0) {
    visual.adjustments = f.adjustments;
  }
  // Rotation is applied by the frame's CSS transform (AC-36), so the canvas
  // paints the shape unrotated.
  return visual;
});

function paint(): void {
  const el = canvas.value;
  const visual = shapeVisual.value;
  if (!el || !visual) {
    return;
  }
  paintShapeIntoCanvas(el, visual, cmToPx(props.frame.widthCm), cmToPx(props.frame.heightCm));
}

onMounted(paint);
watch(() => props.frame, paint, { deep: true });

// devicePixelRatio can change (moving a window between displays); repaint so
// the canvas stays crisp (AC-43).
let mql: MediaQueryList | undefined;
onMounted(() => {
  if (typeof window !== "undefined" && typeof window.matchMedia === "function") {
    mql = window.matchMedia(`(resolution: ${window.devicePixelRatio}dppx)`);
    mql.addEventListener?.("change", paint);
  }
});
onBeforeUnmount(() => mql?.removeEventListener?.("change", paint));
</script>

<template>
  <!-- Text frame -->
  <div
    v-if="frame.kind === 'text'"
    class="frame text-frame"
    :style="{ ...positionStyle, background: frameFill }"
    data-testid="slide-text-frame"
  >
    <div
      v-for="(entry, index) in textParagraphs"
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
    v-else-if="frame.kind === 'image'"
    class="frame"
    :style="positionStyle"
  >
    <img
      class="frame-image"
      :src="resolveMedia(frame.mediaId)"
      alt=""
    >
  </div>

  <!-- Shape frame -->
  <div
    v-else-if="frame.kind === 'shape'"
    class="frame shape-frame"
    :style="positionStyle"
    data-testid="slide-shape-frame"
  >
    <canvas
      ref="canvas"
      class="shape-canvas"
    />
    <div
      v-if="textParagraphs.length > 0"
      class="shape-text"
    >
      <div
        v-for="(entry, index) in textParagraphs"
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
    v-else-if="frame.kind === 'group'"
    class="frame group-frame"
    :style="positionStyle"
    data-testid="slide-group-frame"
  >
    <!-- Children carry absolute slide-space coordinates; they render against
         the stage, not this box, so they sit in a zero-size origin layer. -->
    <SlideFrameView
      v-for="child in frame.children"
      :key="child.id"
      :frame="child"
      :resolve-media="resolveMedia"
    />
  </div>

  <!-- Foreign placeholder box -->
  <div
    v-else-if="frame.kind === 'foreign' && !frame.noBounds"
    class="frame foreign-frame"
    :style="positionStyle"
    data-testid="slide-foreign-frame"
  >
    <span class="foreign-label">{{ frame.label }}</span>
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

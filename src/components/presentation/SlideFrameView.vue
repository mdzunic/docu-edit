<script setup lang="ts">
// One slide frame rendered absolutely-positioned in slide space (SPEC-079):
// text → nested divs/spans, images → <img> from the media resolver, shapes →
// a devicePixelRatio-crisp <canvas> painted by paintShapeIntoCanvas plus a
// centred text overlay, groups → this component recursively, foreign → a
// labelled dashed placeholder box. Runtime prop declarations, not
// defineProps<{...}>() generics — this repo lints .vue scripts with espree,
// which rejects TS-only syntax; every typed transform lives in slideRender.ts.
import { computed, nextTick, onMounted, useTemplateRef, watch } from "vue";
import { cmToPx } from "../../lib/core/units";
import { paintShapeIntoCanvas } from "../../lib/core/shapeCanvas";
import {
  asResolver,
  asSlideFrame,
  framePositionStyle,
  groupChildren,
  imageSrc,
  isEmptyTextPlaceholder,
  isForeignVisible,
  foreignLabel,
  markedParagraphs,
  paragraphStyle,
  placeholderPrompt,
  runStyle,
  shapeVisualOf,
  textFrameFill,
} from "./slideRender";

const props = defineProps({
  /** The SlideFrame to render (presentation model). */
  frame: { type: Object, required: true },
  /** Resolves a media id to a `data:` URI (or "" when unresolved). */
  resolveMedia: { type: Function, required: true },
  /** Selection/editing overlays (SPEC-081); only the top-level stage passes
   * these — group children render read-only. */
  selected: { type: Boolean, default: false },
  editing: { type: Boolean, default: false },
});

const emit = defineEmits(["select", "start-edit", "edit-mounted", "edit-unmounted"]);

const frame = computed(() => asSlideFrame(props.frame));
const kind = computed(() => frame.value.kind);
const position = computed(() => framePositionStyle(frame.value));
const paragraphs = computed(() => markedParagraphs(frame.value));
const textFill = computed(() => textFrameFill(frame.value));
const children = computed(() => groupChildren(frame.value));
const foreignVisible = computed(() => isForeignVisible(frame.value));
const foreignText = computed(() => foreignLabel(frame.value));
const imgSrc = computed(() => imageSrc(frame.value, asResolver(props.resolveMedia)));
const showPrompt = computed(() => !props.editing && isEmptyTextPlaceholder(frame.value));
const promptText = computed(() => placeholderPrompt(frame.value));

const canvasEl = useTemplateRef("canvasEl");
const editHost = useTemplateRef("editHost");

// The contenteditable host is created/destroyed with `editing`; hand it to the
// editing composable (which populates it and attaches the caret) once mounted.
watch(
  () => props.editing,
  (isEditing) => {
    if (isEditing) {
      void nextTick(() => {
        if (editHost.value) emit("edit-mounted", editHost.value);
      });
    } else {
      emit("edit-unmounted");
    }
  },
);

function onFrameMousedown(event = new MouseEvent("mousedown")) {
  // Select on pointer-down (before focus moves); editing swallows its own.
  if (props.editing) return;
  event.stopPropagation();
  emit("select", frame.value.id);
}

function onFrameDblclick(event = new MouseEvent("dblclick")) {
  if (frame.value.kind !== "text") return;
  event.stopPropagation();
  // Pass the click's viewport coordinates so the session drops the caret at the
  // click point rather than the end of the frame (AC-1).
  emit("start-edit", frame.value.id, { x: event.clientX, y: event.clientY });
}

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
    :class="{ 'is-selected': selected, 'is-editing': editing }"
    :style="{ ...position, background: textFill }"
    data-testid="slide-text-frame"
    @mousedown="onFrameMousedown"
    @dblclick="onFrameDblclick"
  >
    <!-- Editing session: an empty contenteditable the editing composable owns -->
    <div
      v-if="editing"
      ref="editHost"
      class="ce-host"
      contenteditable="true"
      spellcheck="false"
      data-testid="slide-text-edit"
    />
    <!-- Empty placeholder prompt (never in the model/files/clipboard) -->
    <div
      v-else-if="showPrompt"
      class="ph-prompt"
      data-testid="slide-placeholder-prompt"
    >
      {{ promptText }}
    </div>
    <template v-else>
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
    </template>
    <div
      v-if="selected && !editing"
      class="selection-box"
      aria-hidden="true"
    >
      <span
        v-for="h in 8"
        :key="h"
        :class="`handle handle-${h}`"
      />
    </div>
  </div>

  <!-- Image frame -->
  <div
    v-else-if="kind === 'image'"
    class="frame"
    :class="{ 'is-selected': selected }"
    :style="position"
    @mousedown="onFrameMousedown"
  >
    <img
      class="frame-image"
      :src="imgSrc"
      alt=""
    >
    <div
      v-if="selected"
      class="selection-box"
      aria-hidden="true"
    >
      <span
        v-for="h in 8"
        :key="h"
        :class="`handle handle-${h}`"
      />
    </div>
  </div>

  <!-- Shape frame -->
  <div
    v-else-if="kind === 'shape'"
    class="frame shape-frame"
    :class="{ 'is-selected': selected }"
    :style="position"
    data-testid="slide-shape-frame"
    @mousedown="onFrameMousedown"
  >
    <canvas
      ref="canvasEl"
      class="shape-canvas"
    />
    <div
      v-if="selected"
      class="selection-box"
      aria-hidden="true"
    >
      <span
        v-for="h in 8"
        :key="h"
        :class="`handle handle-${h}`"
      />
    </div>
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
    :class="{ 'is-selected': selected }"
    :style="position"
    data-testid="slide-foreign-frame"
    @mousedown="onFrameMousedown"
  >
    <span class="foreign-label">{{ foreignText }}</span>
    <div
      v-if="selected"
      class="selection-box"
      aria-hidden="true"
    >
      <span
        v-for="h in 8"
        :key="h"
        :class="`handle handle-${h}`"
      />
    </div>
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

.text-frame.is-editing,
.text-frame.is-selected,
.frame.is-selected {
  cursor: text;
}

.text-frame.is-editing {
  outline: 2px solid var(--accent, #3b82f6);
  outline-offset: 1px;
}

.ce-host {
  flex: 1 1 auto;
  min-height: 1em;
  outline: none;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  line-height: 1.2;
  font-size: 24px; /* the renderer's 18pt default (18 * 96/72) */
}

.ce-host :deep(.ce-p) {
  margin: 0;
}

.ph-prompt {
  color: var(--text-secondary, #888);
  font-style: italic;
  font-size: 24px;
  line-height: 1.2;
  pointer-events: none;
  user-select: none;
}

.selection-box {
  position: absolute;
  inset: -2px;
  border: 1px solid var(--accent, #3b82f6);
  pointer-events: none;
}

.handle {
  position: absolute;
  width: 7px;
  height: 7px;
  background: #fff;
  border: 1px solid var(--accent, #3b82f6);
  border-radius: 1px;
}

.handle-1 { top: -4px; left: -4px; }
.handle-2 { top: -4px; left: calc(50% - 3px); }
.handle-3 { top: -4px; right: -4px; }
.handle-4 { top: calc(50% - 3px); right: -4px; }
.handle-5 { bottom: -4px; right: -4px; }
.handle-6 { bottom: -4px; left: calc(50% - 3px); }
.handle-7 { bottom: -4px; left: -4px; }
.handle-8 { top: calc(50% - 3px); left: -4px; }

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

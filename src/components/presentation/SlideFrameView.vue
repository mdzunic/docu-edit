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
  tableColumnWidths,
  tableRenderRows,
  textFrameFill,
  frameAccessibleName,
  frameRole,
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
  /** SPEC-091: the {row,col} of the table cell in a session (null otherwise),
   * so that cell's `<td>` renders the in-cell contenteditable host. */
  editingCell: { type: Object, default: null },
  /** False for a group's children, which render read-only (no selection,
   * focus or test hooks) until ungroup promotes them (SPEC-082). */
  topLevel: { type: Boolean, default: true },
  /** True while an insertion tool is armed: frames ignore pointer-down so the
   * draw gesture on the surface underneath is not stolen as a move. */
  armed: { type: Boolean, default: false },
  /** SPEC-086 playback: an inline style (opacity/transform/clip-path/transition)
   * layered onto the frame's position during a build preview/show. Null in the
   * editor, where frames always render fully visible (AC-25). */
  playStyle: { type: Object, default: null },
});

const emit = defineEmits([
  "select",
  "start-edit",
  "start-edit-cell",
  "cell-select",
  "frame-pointerdown",
  "edit-mounted",
  "edit-unmounted",
]);

const frame = computed(() => asSlideFrame(props.frame));
const kind = computed(() => frame.value.kind);
// AC-51/52/53: each top-level frame is a focusable element with a role, an
// accessible name and the `slide-frame` test id; children carry none.
const testId = computed(() => (props.topLevel ? "slide-frame" : undefined));
const role = computed(() => (props.topLevel ? frameRole(frame.value) : undefined));
const ariaName = computed(() => (props.topLevel ? frameAccessibleName(frame.value) : undefined));
const tabIndex = computed(() => (props.topLevel ? 0 : undefined));
const position = computed(() => framePositionStyle(frame.value));
const paragraphs = computed(() => markedParagraphs(frame.value));
const textFill = computed(() => textFrameFill(frame.value));
const children = computed(() => groupChildren(frame.value));
const foreignVisible = computed(() => isForeignVisible(frame.value));
const tableColumns = computed(() => tableColumnWidths(frame.value));
const tableRows = computed(() => tableRenderRows(frame.value));
const foreignText = computed(() => foreignLabel(frame.value));
const imgSrc = computed(() => imageSrc(frame.value, asResolver(props.resolveMedia)));
const showPrompt = computed(() => !props.editing && isEmptyTextPlaceholder(frame.value));
const promptText = computed(() => placeholderPrompt(frame.value));

const canvasEl = useTemplateRef("canvasEl");
const editHost = useTemplateRef("editHost");
const tableFrameEl = useTemplateRef("tableFrameEl");

// The in-cell contenteditable host is created/destroyed as the editing cell
// changes (double-click, Enter, Tab between cells). Hand the freshly-mounted
// host to the editing composable — the same edit-mounted/edit-unmounted
// contract the text frame uses (SPEC-091 AC-6a).
const editingCellKey = computed(() =>
  kind.value === "table" && props.editing && props.editingCell
    ? `${props.editingCell.row}:${props.editingCell.col}`
    : null,
);
watch(editingCellKey, (key) => {
  if (key) {
    void nextTick(() => {
      const el = tableFrameEl.value?.querySelector(".cell-edit");
      if (el instanceof HTMLElement) emit("edit-mounted", el);
    });
  } else {
    emit("edit-unmounted");
  }
});

/** True for the one owner cell currently in an editing session. */
function isEditingCell(cell = { rowIndex: 0, colIndex: 0 }) {
  return (
    props.editing &&
    props.editingCell != null &&
    props.editingCell.row === cell.rowIndex &&
    props.editingCell.col === cell.colIndex
  );
}

function onCellMousedown(row = 0, col = 0, event = new MouseEvent("mousedown")) {
  // Anchor the panel/session on this cell and select the frame; stop here so a
  // click inside the table never bubbles to the surface deselect (AC-9c).
  event.stopPropagation();
  emit("cell-select", frame.value.id, row, col);
}

function onCellDblclick(row = 0, col = 0, event = new MouseEvent("dblclick")) {
  event.stopPropagation();
  emit("start-edit-cell", frame.value.id, row, col, { x: event.clientX, y: event.clientY });
}

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

function onFramePointerdown(event = new PointerEvent("pointerdown")) {
  // Begin a move drag (SPEC-082); the manipulation composable selects and
  // captures. Editing sessions, group children, and armed insertion (which
  // draws on the surface below) opt out.
  if (props.editing || !props.topLevel || props.armed) return;
  emit("frame-pointerdown", frame.value.id, event);
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
    :style="[{ ...position, background: textFill }, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @dblclick="onFrameDblclick"
    @pointerdown="onFramePointerdown"
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
  </div>

  <!-- Image frame -->
  <div
    v-else-if="kind === 'image'"
    class="frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
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
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
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
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <SlideFrameView
      v-for="child in children"
      :key="child.id"
      :frame="child"
      :resolve-media="resolveMedia"
      :top-level="false"
    />
  </div>

  <!-- Table frame: a real <table> with a colgroup and merged cells (AC-4) -->
  <div
    v-else-if="kind === 'table'"
    ref="tableFrameEl"
    class="frame table-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
  >
    <table
      class="slide-table"
      data-testid="slide-table"
    >
      <colgroup>
        <col
          v-for="(w, ci) in tableColumns"
          :key="ci"
          :style="{ width: w }"
        >
      </colgroup>
      <tbody>
        <tr
          v-for="(row, ri) in tableRows"
          :key="ri"
        >
          <td
            v-for="cell in row"
            :key="cell.colIndex"
            class="slide-table-cell"
            :colspan="cell.colSpan"
            :rowspan="cell.rowSpan"
            :style="cell.style"
            :data-testid="`slide-table-cell-${cell.rowIndex}-${cell.colIndex}`"
            @mousedown="onCellMousedown(cell.rowIndex, cell.colIndex, $event)"
            @dblclick="onCellDblclick(cell.rowIndex, cell.colIndex, $event)"
          >
            <!-- Editing session: the same contenteditable host the text frame
                 uses, owned by the editing composable (SPEC-091 AC-6a) -->
            <div
              v-if="isEditingCell(cell)"
              class="ce-host cell-edit"
              contenteditable="true"
              spellcheck="false"
              data-testid="slide-table-cell-edit"
            />
            <template v-else>
              <div
                v-for="(entry, pi) in cell.paragraphs"
                :key="pi"
                class="paragraph"
                :style="paragraphStyle(entry.para)"
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
          </td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- Foreign placeholder box -->
  <div
    v-else-if="foreignVisible"
    class="frame foreign-frame"
    :class="{ 'is-selected': selected }"
    :style="[position, playStyle]"
    :data-testid="testId"
    :role="role"
    :aria-label="ariaName"
    :tabindex="tabIndex"
    @mousedown="onFrameMousedown"
    @pointerdown="onFramePointerdown"
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

.text-frame.is-editing,
.text-frame.is-selected {
  cursor: text;
}

/* Non-text frames show a move cursor when selected (AC-16); a text box keeps
   its text cursor so a click drops the caret. */
.shape-frame.is-selected,
.frame.is-selected:not(.text-frame) {
  cursor: move;
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

.table-frame {
  overflow: hidden;
}

.slide-table {
  width: 100%;
  height: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}

.slide-table-cell {
  overflow-wrap: anywhere;
  word-break: break-word;
}

/* The in-cell editing host fills the cell and drops the text-frame flex/font
   sizing so it inherits the cell's own type (SPEC-091 AC-6a). */
.slide-table-cell .cell-edit {
  flex: none;
  min-height: 1em;
  font-size: inherit;
  outline: none;
  cursor: text;
}

.table-frame.is-selected {
  outline: 1px solid var(--accent, #3b82f6);
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

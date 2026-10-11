<script setup lang="ts">
// The in-cell editor (SPEC-122 "Editor DOM"): a contenteditable box ALWAYS
// mounted over the active cell — the element that holds focus while the grid
// does, so SPEC-117's navigation keys reach the grid through it unchanged
// and the first printable character (or IME composition) lands natively in
// it. In Ready it is transparent, empty and lets the pointer through to the
// paper; editing, it shows the draft in the cell's resolved font on the
// cell's fill, sized to the cell and growing right (to the viewport's edge)
// and down as lines grow. Named "B3, Revenue" — the address and the shown
// value. On another sheet (a tab clicked while pointing) it keeps focus and
// the draft but shows nothing: the formula bar shows the formula there. A
// formula error shows beneath it in `cell-editor-error`. Input, keys and
// rendering are useDraftSurface's (shared with the formula bar); the state
// is useCellEditing's. Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { computed, useTemplateRef } from "vue";
import { asCellEditing } from "../../views/spreadsheet/useCellEditing";
import { useDraftSurface } from "../../views/spreadsheet/useDraftSurface";

const props = defineProps({
  /** useCellEditing's return. */
  editor: { type: Object, required: true },
  /** The box: position and size over the active cell (useSheetGrid's). */
  boxStyle: { type: Object, default: () => ({}) },
  /** The cell's resolved font, colour and fill while editing. */
  textStyle: { type: Object, default: () => ({}) },
  /** "B3, Revenue". */
  label: { type: String, default: "" },
  /** SPEC-127: the function assists' ARIA (useFunctionAssist's editorAria). */
  aria: { type: Object, default: () => ({}) },
});

const editor = computed(() => asCellEditing(props.editor));
/** The assists' ARIA, a formula error's description taking precedence. */
const hostAria = computed(() => ({ ...props.aria, ...(editor.value.error.value ? { "aria-describedby": "cell-editor-error" } : {}) }));
const host = useTemplateRef("host");
const surface = useDraftSurface(host, () => editor.value, "cell");
</script>

<template>
  <div
    class="cell-editor-box"
    :class="{ 'is-editing': editor.editing.value, 'is-elsewhere': !editor.onShownSheet.value }"
    :style="props.boxStyle"
  >
    <div
      ref="host"
      class="cell-editor"
      data-testid="cell-editor"
      contenteditable="true"
      role="textbox"
      aria-multiline="true"
      :aria-label="props.label"
      :aria-invalid="editor.error.value ? 'true' : undefined"
      v-bind="hostAria"
      :data-mode="editor.mode.value"
      spellcheck="false"
      autocapitalize="off"
      autocorrect="off"
      translate="no"
      :style="editor.editing.value && editor.onShownSheet.value ? props.textStyle : undefined"
      @input="surface.onInput"
      @beforeinput="surface.onBeforeInput"
      @compositionstart="surface.onCompositionStart"
      @compositionend="surface.onCompositionEnd"
      @paste="surface.onPaste"
      @keydown="surface.onKeydown"
      @drop.prevent
      @dragover.prevent
    />
    <div
      v-if="editor.error.value && editor.onShownSheet.value"
      id="cell-editor-error"
      class="cell-editor-error"
      data-testid="cell-editor-error"
      role="alert"
    >
      {{ editor.error.value }}
    </div>
  </div>
</template>

<style scoped>
/* The box sits over the active cell in the paper's unscaled px (the grid's
   zoom scales it with the paper); the editor inside grows from it. */
.cell-editor-box {
  position: absolute;
  pointer-events: none;
}

.cell-editor {
  box-sizing: border-box;
  min-width: 100%;
  min-height: 100%;
  width: max-content;
  max-width: inherit;
  padding: 0 var(--space-1);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  color: transparent;
  background: transparent;
  caret-color: transparent;
  outline: none;
  cursor: cell;
}

/* Editing: the draft on the cell's paper, a text caret, the pointer kept
   (clicks place the caret). */
.is-editing .cell-editor {
  pointer-events: auto;
  caret-color: auto;
  cursor: text;
  box-shadow: var(--shadow-1);
}

.is-editing.is-elsewhere .cell-editor {
  color: transparent;
  background: transparent;
  caret-color: transparent;
  box-shadow: none;
}

.cell-editor-error {
  position: absolute;
  left: 0;
  top: 100%;
  margin-top: var(--space-1);
  padding: var(--space-1) var(--space-2);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
  color: var(--text-danger);
  box-shadow: var(--shadow-2);
  font-family: var(--font-ui);
  font-size: var(--text-xs);
  white-space: nowrap;
}
</style>

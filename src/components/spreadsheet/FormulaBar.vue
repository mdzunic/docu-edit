<script setup lang="ts">
// The formula bar (SPEC-117, editable since SPEC-122): above the grid, the Name Box — a
// read-only input showing the active cell's address (a merge's top-left), or
// "3R × 2C" while a drag is in progress — and the active cell's content,
// SPEC-119's formulaBarText: "=" + its formula, an array's "{=…}", a data
// table's "{=TABLE(…)}", else its edit text (SPEC-118's editText: the value
// as the cell editor would hold it, in the user's locale), line breaks
// preserved. A dynamic array's member shows its anchor's formula muted, with
// data-spill-member="true". SPEC-122: with the view's cell editor
// (`editor`, useCellEditing) and a grid sheet on show (`editable`), the
// content is a contenteditable textbox holding the same draft as the in-cell
// editor (useDraftSurface): typing in either updates both, a click opens
// Edit mode at the clicked character, Enter/Tab/Escape behave as in the
// cell, and a formula's references are coloured as there. Without them it
// is read-only (aria-readonly). The Name Box stays read-only (a Go To field
// is SPEC-134's). Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed, useTemplateRef } from "vue";
import { asCellEditing } from "../../views/spreadsheet/useCellEditing";
import { useDraftSurface } from "../../views/spreadsheet/useDraftSurface";

const props = defineProps({
  /** The Name Box text: "B3", or "3R × 2C" during a drag. */
  address: { type: String, default: "" },
  /** The active cell's content. */
  content: { type: String, default: "" },
  /** The content is a dynamic-array member's anchor formula (muted). */
  spillMember: { type: Boolean, default: false },
  /** useCellEditing's return (SPEC-122); absent = read-only. */
  editor: { type: Object, default: null },
  /** A grid sheet is on show, so the content can be edited. */
  editable: { type: Boolean, default: false },
});

const editor = computed(() => (props.editor && props.editable ? asCellEditing(props.editor) : null));
const editing = computed(() => editor.value?.editing.value === true);
const host = useTemplateRef("host");
const surface = useDraftSurface(host, () => editor.value, "bar", () => props.content);
</script>

<template>
  <div
    class="formula-bar"
    data-testid="formula-bar"
  >
    <input
      class="name-box"
      data-testid="name-box"
      type="text"
      readonly
      aria-label="Name Box"
      :value="props.address"
    >
    <span
      class="formula-bar-fx"
      aria-hidden="true"
    >fx</span>
    <div
      ref="host"
      class="formula-bar-content"
      :class="{ 'is-spill-member': props.spillMember && !editing, 'is-editable': editor }"
      data-testid="formula-bar-content"
      role="textbox"
      :contenteditable="editor ? 'true' : undefined"
      :aria-readonly="editor ? undefined : 'true'"
      aria-label="Cell content"
      aria-multiline="true"
      :data-spill-member="props.spillMember && !editing ? 'true' : undefined"
      spellcheck="false"
      @input="surface.onInput"
      @beforeinput="surface.onBeforeInput"
      @compositionstart="surface.onCompositionStart"
      @compositionend="surface.onCompositionEnd"
      @paste="surface.onPaste"
      @keydown="surface.onKeydown"
      @pointerdown="surface.onPointerDown"
      @focus="surface.onFocus"
      @click="surface.onClick"
      @drop.prevent
      @dragover.prevent
    />
  </div>
</template>

<style scoped>
.formula-bar {
  display: flex;
  flex: none;
  align-items: stretch;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-3);
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
}

.name-box {
  flex: none;
  width: 7rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: inherit;
  font: inherit;
}

.name-box:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: 1px;
}

.formula-bar-fx {
  flex: none;
  align-self: center;
  color: var(--text-secondary);
  font-style: italic;
}

/* Line breaks preserved; a multi-line value is clipped to a few lines. */
.formula-bar-content {
  flex: 1 1 auto;
  min-width: 0;
  max-height: 4.5em;
  padding: var(--space-1) var(--space-2);
  overflow: auto;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.formula-bar-content.is-editable {
  cursor: text;
}

.formula-bar-content.is-editable:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: -2px;
}

/* A spilled cell's formula belongs to its anchor: shown, muted. */
.formula-bar-content.is-spill-member {
  color: var(--text-secondary);
}
</style>

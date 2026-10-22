<script setup lang="ts">
// The grid's note layer (SPEC-135), a dumb view of useCellNotes (`controller`)
// over the cell area, in the paper's unscaled px (the grid's content box
// scales it with the zoom):
// - `sheet-note-box`: each shown note — a `visible` one, or every note while
//   View ▸ "Show all notes" is on — at Excel's default placement beside its
//   cell and its `box` size (sheetNoteLayout's noteBoxPlacement): the author,
//   then the text. Boxes never take the pointer (dragging and resizing them
//   is out of scope), so the cells under them stay reachable.
// - `sheet-note-editor`: the note editor, a plain textarea in a box at the
//   same placement (role="dialog", so the grid's shortcuts leave its keys
//   alone; no key or press reaches the grid's own handling). Escape and
//   Ctrl/Cmd+Enter commit and return to the grid; leaving it commits.
// - the hover popover (SheetNotePopover) beside its cell.
// - the active cell's annotation text in a visually hidden element
//   (NOTE_DESCRIPTION_ID), which the active cell's aria-describedby names.
// Every annotation text is plain text — interpolated, never markup.
// `anchor` is useSheetGrid's noteAnchor (a cell's rect and the height of the
// row above it); `bounds` the cell area's size. Runtime prop declarations
// per the espree constraint (see AppShell.vue).
import { computed, nextTick, useTemplateRef, watch } from "vue";
import { formatA1 } from "../../lib/spreadsheet/address";
import { asNoteAnchorOf, NOTE_POPOVER_WIDTH_PX, noteBoxPlacement, notePopoverPlacement } from "../../views/spreadsheet/sheetNoteLayout";
import { asCellNotes, NOTE_DESCRIPTION_ID } from "../../views/spreadsheet/useCellNotes";
import SheetNotePopover from "./SheetNotePopover.vue";

const props = defineProps({
  /** useCellNotes' return. */
  controller: { type: Object, required: true },
  /** useSheetGrid's noteAnchor: (ref) => { cell, above } | null. */
  anchor: { type: Function, required: true },
  /** The cell area's size (width, height), paper px. */
  bounds: { type: Object, required: true },
});

const notes = computed(() => asCellNotes(props.controller));
const anchorOf = computed(() => asNoteAnchorOf(props.anchor));

function px(value = 0) {
  return `${value}px`;
}

/** The shown notes with their boxes' rects (a cell under a frozen pane has none). */
const boxes = computed(() =>
  notes.value.boxes.value.flatMap((box) => {
    const at = anchorOf.value(box.ref);
    return at ? [{ ...box, address: formatA1(box.ref), rect: noteBoxPlacement(at, box) }] : [];
  }),
);

/** The editor's box, at its note's placement. */
const editorRect = computed(() => {
  const open = notes.value.editor.value;
  const at = open ? anchorOf.value(open.ref) : null;
  return open && at ? noteBoxPlacement(at, open) : null;
});

const popover = computed(() => {
  const view = notes.value.popover.value;
  const at = view ? anchorOf.value(view.ref) : null;
  return view && at ? { view, address: formatA1(view.ref), place: notePopoverPlacement(at.cell, { width: Number(props.bounds.width) || 0 }) } : null;
});

const textarea = useTemplateRef("textarea");
// The editor takes focus as it opens, its caret at the end — once its box
// shows (a cell scrolled out of view is brought back first, so the box can
// appear a frame after the editor opens).
watch(
  () => (editorRect.value ? notes.value.editor.value : null),
  async (open) => {
    if (!open) {
      return;
    }
    await nextTick();
    const el = textarea.value;
    if (el instanceof HTMLTextAreaElement) {
      el.focus({ preventScroll: true });
      el.setSelectionRange(el.value.length, el.value.length);
    }
  },
  { immediate: true },
);

function onInput(event = new Event("input")) {
  const el = event.target;
  if (el instanceof HTMLTextAreaElement) {
    notes.value.input(el.value);
  }
}

/** Escape and Ctrl/Cmd+Enter commit (Excel keeps a note's text on Escape). */
function onEditorKey(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape" || (event.key === "Enter" && (event.ctrlKey || event.metaKey))) {
    event.preventDefault();
    notes.value.commitEditor(true);
  }
}
</script>

<template>
  <div class="sheet-note-layer">
    <div
      :id="NOTE_DESCRIPTION_ID"
      class="sheet-note-description"
      data-testid="sheet-annotation-description"
    >
      {{ notes.activeDescription.value }}
    </div>
    <div
      v-for="box in boxes"
      :key="box.noteId"
      class="sheet-note-box"
      role="note"
      :aria-label="`Note on ${box.address}`"
      data-testid="sheet-note-box"
      :data-note-ref="box.address"
      :style="{ left: px(box.rect.left), top: px(box.rect.top), width: px(box.rect.width), height: px(box.rect.height) }"
    >
      <div
        v-if="box.author"
        class="sheet-note-author"
      >
        {{ box.author }}
      </div>
      <div class="sheet-note-text">
        {{ box.text }}
      </div>
    </div>
    <div
      v-if="editorRect"
      class="sheet-note-box is-editing"
      role="dialog"
      :aria-label="`Note on ${notes.editorAddress.value}`"
      :style="{ left: px(editorRect.left), top: px(editorRect.top), width: px(editorRect.width), height: px(editorRect.height) }"
      @pointerdown.stop
      @mousedown.stop
      @click.stop
      @dblclick.stop
      @contextmenu.stop
      @keydown.stop
    >
      <textarea
        ref="textarea"
        class="sheet-note-editor"
        data-testid="sheet-note-editor"
        :aria-label="`Note on ${notes.editorAddress.value}`"
        :value="notes.draft.value"
        @input="onInput"
        @keydown="onEditorKey"
        @blur="notes.commitEditor(false)"
      />
    </div>
    <SheetNotePopover
      v-if="popover"
      :note="popover.view.note ?? undefined"
      :thread="popover.view.thread ?? undefined"
      :address="popover.address"
      :left="popover.place.left"
      :top="popover.place.top"
      :width="NOTE_POPOVER_WIDTH_PX"
      @enter="notes.hoverPopover(true)"
      @leave="notes.hoverPopover(false)"
      @open-thread="notes.openPopoverThread()"
    />
  </div>
</template>

<style scoped>
/* Over the paper, the overlays and the editor and validation layers, under
   the headers; it only clips — the editor box and the popover take the
   pointer. */
.sheet-note-layer {
  position: absolute;
  z-index: 1;
  overflow: hidden;
  pointer-events: none;
}

/* The active cell's annotation text: for assistive technology only. */
.sheet-note-description {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  clip-path: inset(50%);
  white-space: nowrap;
}

.sheet-note-box {
  position: absolute;
  box-sizing: border-box;
  overflow: hidden;
  padding: var(--space-1) var(--space-2);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  box-shadow: var(--shadow-1);
}

.sheet-note-box.is-editing {
  display: flex;
  padding: 0;
  border-color: var(--accent);
  pointer-events: auto;
}

.sheet-note-author {
  font-weight: var(--weight-bold);
}

.sheet-note-text {
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.sheet-note-editor {
  flex: 1 1 auto;
  box-sizing: border-box;
  min-width: 0;
  padding: var(--space-1) var(--space-2);
  font: inherit;
  color: inherit;
  background: transparent;
  border: 0;
  outline: none;
  resize: none;
}
</style>

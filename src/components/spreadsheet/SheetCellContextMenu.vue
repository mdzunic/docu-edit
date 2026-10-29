<script setup lang="ts">
// The cell context menu (SPEC-123): opened by a right-click in the grid,
// Shift+F10 or the Menu key (useFormatting) at the pointer or under the
// active cell, with "Format cells…" and "Clear contents" — SPEC-124 adds
// "Insert…" and "Delete…" (`structure`: whole rows or columns act at once,
// any other selection opens SheetInsertDeleteDialog); SPEC-125 puts Cut,
// Copy, Paste and Paste special… first (`clipboard`: useSheetClipboard's
// menu — Cut and Copy fire the native events through execCommand, Paste
// reads the async clipboard); SPEC-135 adds "Link…" (`link`: the link
// dialog on the active cell) and the notes and comments items (`annotations`
// — the active cell's state; `annotate` emits "note", "delete-note",
// "toggle-note" or "comment"): New note / Edit note, and on a noted cell
// Delete note and Show note / Hide note; New comment — disabled with its
// reason in an .ods workbook, which holds notes only). Link…, New note and
// New comment are disabled with their reason where what they add could not
// be saved (a CSV/TSV workbook's links; an .xlsx sheet whose source's
// annotations stay as read — AC-11); SPEC-138's `readOnly` disables every
// item but Copy; later specs append theirs. UiMenuSurface owns
// outside-click and Escape dismissal. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  open: { type: Boolean, default: false },
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
  /** SPEC-135: { canAnnotate, hasNote, noteVisible, noteDisabledReason,
   * canComment, commentDisabledReason, canLink, linkDisabledReason } for the
   * active cell. */
  annotations: { type: Object, default: () => ({}) },
  /** SPEC-138: no edit applies — every item but Copy is disabled. */
  readOnly: { type: Boolean, default: false },
});

const emit = defineEmits(["run", "structure", "clipboard", "link", "annotate", "close"]);

const CLIPBOARD = [
  { action: "cut", label: "Cut", testid: "sheet-cell-menu-cut" },
  { action: "copy", label: "Copy", testid: "sheet-cell-menu-copy" },
  { action: "paste", label: "Paste", testid: "sheet-cell-menu-paste" },
  { action: "paste-special", label: "Paste special…", testid: "sheet-cell-menu-paste-special" },
];

function clipboard(action = "copy") {
  emit("close");
  emit("clipboard", action);
}

function choose(command = Object()) {
  emit("close");
  emit("run", command);
}

function structure(mode = "insert") {
  emit("close");
  emit("structure", mode);
}

function link() {
  emit("close");
  emit("link");
}

/** SPEC-135: a field of `annotations`. */
function flag(name = "") {
  return Reflect.get(props.annotations, name) === true;
}
/** SPEC-135: a reason field of `annotations` ("" for none). */
function reasonOf(name = "") {
  const reason = Reflect.get(props.annotations, name);
  return typeof reason === "string" ? reason : "";
}
function commentReason() {
  return reasonOf("commentDisabledReason");
}
/** Link… applies unless `annotations` says otherwise. */
function canLink() {
  return Reflect.get(props.annotations, "canLink") !== false;
}

function annotate(action = "note") {
  emit("close");
  emit("annotate", action);
}
</script>

<template>
  <UiMenuSurface
    :open="props.open"
    :anchor="null"
    class="sheet-cell-menu"
    role="menu"
    aria-label="Cell"
    data-testid="sheet-cell-menu"
    :style="{ left: `${props.x}px`, top: `${props.y}px` }"
    @dismiss="emit('close')"
  >
    <button
      v-for="item in CLIPBOARD"
      :key="item.action"
      type="button"
      role="menuitem"
      :data-testid="item.testid"
      :disabled="props.readOnly && item.action !== 'copy'"
      @mousedown.prevent
      @click="clipboard(item.action)"
    >
      {{ item.label }}
    </button>
    <hr
      class="menu-separator"
      role="separator"
    >
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-format-cells"
      :disabled="props.readOnly"
      @mousedown.prevent
      @click="choose({ kind: 'format-cells' })"
    >
      Format cells…
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-clear-contents"
      :disabled="props.readOnly"
      @mousedown.prevent
      @click="choose({ kind: 'clear', what: 'contents' })"
    >
      Clear contents
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-insert"
      :disabled="props.readOnly"
      @mousedown.prevent
      @click="structure('insert')"
    >
      Insert…
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-delete"
      :disabled="props.readOnly"
      @mousedown.prevent
      @click="structure('delete')"
    >
      Delete…
    </button>
    <hr
      class="menu-separator"
      role="separator"
    >
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-link"
      :disabled="props.readOnly || !canLink()"
      :aria-disabled="canLink() ? undefined : 'true'"
      :title="canLink() ? undefined : reasonOf('linkDisabledReason')"
      @mousedown.prevent
      @click="link"
    >
      Link…
    </button>
    <button
      v-if="flag('canAnnotate') || reasonOf('noteDisabledReason')"
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-note"
      :disabled="props.readOnly || !flag('canAnnotate')"
      :aria-disabled="flag('canAnnotate') ? undefined : 'true'"
      :title="flag('canAnnotate') ? undefined : reasonOf('noteDisabledReason')"
      @mousedown.prevent
      @click="annotate('note')"
    >
      {{ flag("hasNote") ? "Edit note" : "New note" }}
    </button>
    <template v-if="flag('canAnnotate') && flag('hasNote')">
      <button
        type="button"
        role="menuitem"
        data-testid="sheet-cell-menu-delete-note"
        :disabled="props.readOnly"
        @mousedown.prevent
        @click="annotate('delete-note')"
      >
        Delete note
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="sheet-cell-menu-toggle-note"
        :disabled="props.readOnly"
        @mousedown.prevent
        @click="annotate('toggle-note')"
      >
        {{ flag("noteVisible") ? "Hide note" : "Show note" }}
      </button>
    </template>
    <button
      v-if="flag('canComment') || commentReason()"
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-comment"
      :disabled="props.readOnly || !flag('canComment')"
      :aria-disabled="flag('canComment') ? undefined : 'true'"
      :title="flag('canComment') ? undefined : commentReason()"
      @mousedown.prevent
      @click="annotate('comment')"
    >
      New comment
    </button>
  </UiMenuSurface>
</template>

<style scoped>
.sheet-cell-menu {
  position: fixed;
  z-index: 40;
  min-width: 11rem;
}

button {
  display: block;
  width: 100%;
  padding: var(--space-1) var(--space-3);
  font: inherit;
  text-align: left;
  color: var(--text-primary);
  background: transparent;
  border: 0;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

button:hover,
button:focus-visible {
  background: var(--state-hover);
}

button:disabled {
  color: var(--text-secondary);
  cursor: default;
}

button:disabled:hover {
  background: transparent;
}

.menu-separator {
  margin: var(--space-1) 0;
  border: 0;
  border-top: var(--border-width) solid var(--border-subtle);
}
</style>

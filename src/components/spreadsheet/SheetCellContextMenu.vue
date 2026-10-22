<script setup lang="ts">
// The cell context menu (SPEC-123): opened by a right-click in the grid,
// Shift+F10 or the Menu key (useFormatting) at the pointer or under the
// active cell, with "Format cells…" and "Clear contents" — SPEC-124 adds
// "Insert…" and "Delete…" (`structure`: whole rows or columns act at once,
// any other selection opens SheetInsertDeleteDialog); SPEC-125 puts Cut,
// Copy, Paste and Paste special… first (`clipboard`: useSheetClipboard's
// menu — Cut and Copy fire the native events through execCommand, Paste
// reads the async clipboard); SPEC-135 adds "Link…" (`link`: the link
// dialog on the active cell); later specs append theirs. UiMenuSurface owns
// outside-click and Escape dismissal. Runtime prop declarations per the
// espree constraint (see AppShell.vue).
import UiMenuSurface from "../ui/UiMenuSurface.vue";

const props = defineProps({
  open: { type: Boolean, default: false },
  x: { type: Number, default: 0 },
  y: { type: Number, default: 0 },
});

const emit = defineEmits(["run", "structure", "clipboard", "link", "close"]);

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
      @mousedown.prevent
      @click="choose({ kind: 'format-cells' })"
    >
      Format cells…
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-clear-contents"
      @mousedown.prevent
      @click="choose({ kind: 'clear', what: 'contents' })"
    >
      Clear contents
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-insert"
      @mousedown.prevent
      @click="structure('insert')"
    >
      Insert…
    </button>
    <button
      type="button"
      role="menuitem"
      data-testid="sheet-cell-menu-delete"
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
      @mousedown.prevent
      @click="link"
    >
      Link…
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

.menu-separator {
  margin: var(--space-1) 0;
  border: 0;
  border-top: var(--border-width) solid var(--border-subtle);
}
</style>

<script setup lang="ts">
// The editor's context menu (SPEC-016 AC-4..6): Cut, Copy and Paste at the
// right-click point — plus, when opened inside a table cell, the
// row/column section (SPEC-022 AC-17..19) and cell formatting section
// (SPEC-023 AC-6, AC-8, AC-10..23). Emits the chosen action only — the
// editor owns the clipboard and table work. Dismissed by a click anywhere
// outside or Escape.
import { onBeforeUnmount, onMounted } from "vue";
import { modifierKeyLabel } from "../lib/core/platform";

const props = defineProps({
  /** Viewport coordinates of the right-click that opened the menu. */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  /** True when the right-click landed inside a table cell — shows the
   * row/column section (SPEC-022 AC-17..19). */
  tableActions: { type: Boolean, default: false },
  /** True when the active cell-range selection spans more than one cell —
   * shows "Merge cells" (SPEC-023 AC-6). */
  canMergeCells: { type: Boolean, default: false },
  /** True when the right-click landed on a list paragraph — shows the
   * numbering-control section (SPEC-032 AC-14..16). */
  listActions: { type: Boolean, default: false },
});

const emit = defineEmits(["cut", "copy", "paste", "table-action", "list-action", "dismiss"]);

// The numbering-control section's entries (SPEC-032 AC-14..16). The key
// doubles as the emitted action id and the testid suffix, like
// TABLE_ACTIONS.
const LIST_ACTIONS = [
  { key: "restart-numbering", label: "Restart numbering at 1" },
  { key: "continue-numbering", label: "Continue previous numbering" },
  { key: "set-numbering-value", label: "Set numbering value…" },
];

// The table section's entries, in menu order (SPEC-022 AC-17..19). The key
// doubles as the emitted action id and the testid suffix.
const TABLE_ACTIONS = [
  { key: "insert-row-above", label: "Insert row above" },
  { key: "insert-row-below", label: "Insert row below" },
  { key: "insert-col-left", label: "Insert column left" },
  { key: "insert-col-right", label: "Insert column right" },
  { key: "delete-row", label: "Delete row" },
  { key: "delete-col", label: "Delete column" },
  { key: "delete-table", label: "Delete table" },
  // SPEC-046 AC-29, AC-30: "Table -> Repeat row for each..." — an
  // already-repeating row toggles off immediately when this is chosen
  // (see useTables.ts's handleRepeatRowAction), otherwise opens the
  // array-field picker.
  { key: "repeat-row", label: "Repeat row for each…" },
];

// The cell-formatting section's entries (SPEC-023 AC-8, AC-10..23), always
// offered inside a table; "Merge cells" only when a multi-cell range is
// active (AC-6).
const CELL_ACTIONS = [
  { key: "split-cell", label: "Split cell…" },
  { key: "table-properties", label: "Table properties…" },
];

const mod = modifierKeyLabel();

function menuStyle() {
  return { left: `${props.left}px`, top: `${props.top}px` };
}

function handleOutsidePointerDown() {
  emit("dismiss");
}

function handleKeyDown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    emit("dismiss");
  }
}

onMounted(() => {
  document.addEventListener("mousedown", handleOutsidePointerDown);
  document.addEventListener("keydown", handleKeyDown);
});

onBeforeUnmount(() => {
  document.removeEventListener("mousedown", handleOutsidePointerDown);
  document.removeEventListener("keydown", handleKeyDown);
});
</script>

<template>
  <div
    class="context-menu"
    data-testid="editor-context-menu"
    role="menu"
    aria-label="Editing"
    :style="menuStyle()"
    @mousedown.stop
  >
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="context-menu-cut"
      @click="emit('cut')"
    >
      <span>Cut</span>
      <span class="context-menu-shortcut">{{ mod }}+X</span>
    </button>
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="context-menu-copy"
      @click="emit('copy')"
    >
      <span>Copy</span>
      <span class="context-menu-shortcut">{{ mod }}+C</span>
    </button>
    <button
      type="button"
      class="context-menu-item"
      role="menuitem"
      data-testid="context-menu-paste"
      @click="emit('paste')"
    >
      <span>Paste</span>
      <span class="context-menu-shortcut">{{ mod }}+V</span>
    </button>
    <template v-if="listActions">
      <div
        class="context-menu-separator"
        role="separator"
      />
      <button
        v-for="action in LIST_ACTIONS"
        :key="action.key"
        type="button"
        class="context-menu-item"
        role="menuitem"
        :data-testid="`context-menu-${action.key}`"
        @click="emit('list-action', action.key)"
      >
        <span>{{ action.label }}</span>
      </button>
    </template>
    <template v-if="tableActions">
      <div
        class="context-menu-separator"
        role="separator"
      />
      <button
        v-for="action in TABLE_ACTIONS"
        :key="action.key"
        type="button"
        class="context-menu-item"
        role="menuitem"
        :data-testid="`context-menu-${action.key}`"
        @click="emit('table-action', action.key)"
      >
        <span>{{ action.label }}</span>
      </button>
      <div
        class="context-menu-separator"
        role="separator"
      />
      <button
        v-if="canMergeCells"
        type="button"
        class="context-menu-item"
        role="menuitem"
        data-testid="context-menu-merge-cells"
        @click="emit('table-action', 'merge-cells')"
      >
        <span>Merge cells</span>
      </button>
      <button
        v-for="action in CELL_ACTIONS"
        :key="action.key"
        type="button"
        class="context-menu-item"
        role="menuitem"
        :data-testid="`context-menu-${action.key}`"
        @click="emit('table-action', action.key)"
      >
        <span>{{ action.label }}</span>
      </button>
    </template>
  </div>
</template>

<style scoped>
.context-menu {
  position: fixed;
  z-index: 40;
  min-width: 12rem;
  padding: 0.25rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  box-shadow: 0 2px 10px rgb(0 0 0 / 0.25);
}

.context-menu-item {
  font: inherit;
  font-size: 0.875rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 1.5rem;
  width: 100%;
  padding: 0.375rem 0.75rem;
  border: none;
  border-radius: 4px;
  background: transparent;
  color: #202124;
  cursor: pointer;
  text-align: left;
}

.context-menu-item:hover {
  background: #f1f3f4;
}

.context-menu-shortcut {
  color: #5f6368;
  font-size: 0.75rem;
}

.context-menu-separator {
  height: 1px;
  margin: 0.25rem 0;
  background: #e0e0e0;
}
</style>

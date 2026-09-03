<script setup lang="ts">
// The editor's context menu (SPEC-016 AC-4..6): Cut, Copy and Paste at the
// right-click point — plus, when opened inside a table cell, the
// row/column section (SPEC-022 AC-17..19) and cell formatting section
// (SPEC-023 AC-6, AC-8, AC-10..23). Emits the chosen action only — the
// editor owns the clipboard and table work. Dismissed by a click anywhere
// outside or Escape.
import { nextTick, onBeforeUnmount, onMounted, ref, useTemplateRef, watch } from "vue";
import { modifierKeyLabel } from "../lib/core/platform";

const props = defineProps({
  /** Viewport coordinates of the right-click that opened the menu. */
  left: { type: Number, required: true },
  top: { type: Number, required: true },
  /** True when the right-click landed inside a table cell — shows the
   * row/column section (SPEC-022 AC-17..19). */
  tableActions: { type: Boolean, default: false },
  /** Whether the RIGHT-CLICKED row is a header row (SPEC-076 AC-19) — the
   * "Repeat header row" check item's state. */
  rowIsHeader: { type: Boolean, default: false },
  /** True when the active cell-range selection spans more than one cell —
   * shows "Merge cells" (SPEC-023 AC-6). */
  canMergeCells: { type: Boolean, default: false },
  /** True when the right-click landed on a list paragraph — shows the
   * numbering-control section (SPEC-032 AC-14..16). */
  listActions: { type: Boolean, default: false },
  /** True when the right-click landed on a SPEC-053 field — shows the field
   * section (AC-13). */
  fieldActions: { type: Boolean, default: false },
  /** Whether the field under the cursor is locked, so the section offers
   * "Unlock field" instead of "Lock field". */
  fieldLocked: { type: Boolean, default: false },
  /** True when a captionable object (image/chart/table) is selected — shows
   * "Insert caption" (SPEC-055 AC-1). */
  captionAction: { type: Boolean, default: false },
  /** True when a drawing is selected — shows the align/distribute section
   * (SPEC-072 AC-14). */
  drawingActions: { type: Boolean, default: false },
  /** True when three or more drawings are selected, so Distribute applies. */
  canDistributeDrawings: { type: Boolean, default: false },
});

const emit = defineEmits(["cut", "copy", "paste", "table-action", "list-action", "field-action", "caption-action", "drawing-action", "dismiss"]);

// The drawing align/distribute section (SPEC-072 AC-14). The key doubles as the
// emitted action id (`drawing.<key>`) and the testid suffix.
const DRAWING_ALIGN_ACTIONS = [
  { key: "align-left", label: "Align left" },
  { key: "align-center", label: "Align centre" },
  { key: "align-right", label: "Align right" },
  { key: "align-top", label: "Align top" },
  { key: "align-middle", label: "Align middle" },
  { key: "align-bottom", label: "Align bottom" },
];
const DRAWING_DISTRIBUTE_ACTIONS = [
  { key: "distribute-horizontal", label: "Distribute horizontally" },
  { key: "distribute-vertical", label: "Distribute vertically" },
];

// The field section's entries (SPEC-053 AC-13). The key doubles as the
// emitted action id and testid suffix.
const FIELD_ACTIONS = [
  { key: "update", label: "Update field" },
  { key: "edit", label: "Edit field…" },
  { key: "toggle-codes", label: "Toggle field codes" },
  { key: "unlink", label: "Unlink field" },
];

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

// The menu is position:fixed at the click point, so a menu taller than the
// space under the cursor would otherwise hang past the viewport with no way
// to scroll it in — the sections above keep growing (fields, lists, caption),
// and the factory container's taller font metrics overflow before macOS does
// (SPEC-023 AC-14 broke on main this way, 2026-08-18). Clamp the placement to
// the viewport after measuring, the way native context menus shift to fit.
const menuEl = useTemplateRef("menu");
const placed = ref({ left: props.left, top: props.top });

function place() {
  const el = menuEl.value;
  if (!el) {
    return;
  }
  const pad = 8;
  const rect = el.getBoundingClientRect();
  placed.value = {
    left: Math.max(pad, Math.min(props.left, window.innerWidth - rect.width - pad)),
    top: Math.max(pad, Math.min(props.top, window.innerHeight - rect.height - pad)),
  };
}

watch(
  () => [props.left, props.top],
  () => {
    // Re-open at a new point while mounted: measure from the raw spot again.
    placed.value = { left: props.left, top: props.top };
    void nextTick().then(place);
  },
);

function menuStyle() {
  return { left: `${placed.value.left}px`, top: `${placed.value.top}px` };
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
  place();
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
    ref="menu"
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
    <template v-if="fieldActions">
      <div
        class="context-menu-separator"
        role="separator"
      />
      <button
        v-for="action in FIELD_ACTIONS"
        :key="action.key"
        type="button"
        class="context-menu-item"
        role="menuitem"
        :data-testid="`context-menu-field-${action.key}`"
        @click="emit('field-action', action.key)"
      >
        <span>{{ action.label }}</span>
      </button>
      <button
        type="button"
        class="context-menu-item"
        role="menuitem"
        data-testid="context-menu-field-lock"
        @click="emit('field-action', fieldLocked ? 'unlock' : 'lock')"
      >
        <span>{{ fieldLocked ? "Unlock field" : "Lock field" }}</span>
      </button>
    </template>
    <template v-if="captionAction">
      <div
        class="context-menu-separator"
        role="separator"
      />
      <button
        type="button"
        class="context-menu-item"
        role="menuitem"
        data-testid="context-menu-insert-caption"
        @click="emit('caption-action')"
      >
        <span>Insert caption</span>
      </button>
    </template>
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
    <template v-if="drawingActions">
      <div
        class="context-menu-separator"
        role="separator"
      />
      <button
        v-for="action in DRAWING_ALIGN_ACTIONS"
        :key="action.key"
        type="button"
        class="context-menu-item"
        role="menuitem"
        :data-testid="`context-menu-${action.key}`"
        @click="emit('drawing-action', action.key)"
      >
        <span>{{ action.label }}</span>
      </button>
      <button
        v-for="action in DRAWING_DISTRIBUTE_ACTIONS"
        :key="action.key"
        type="button"
        class="context-menu-item"
        role="menuitem"
        :data-testid="`context-menu-${action.key}`"
        :disabled="!canDistributeDrawings"
        @click="emit('drawing-action', action.key)"
      >
        <span>{{ action.label }}</span>
      </button>
    </template>
    <template v-if="tableActions">
      <div
        class="context-menu-separator"
        role="separator"
      />
      <!-- SPEC-076 AC-19: the same check item the Table menu offers, acting
           on the row that was right-clicked rather than the caret's. -->
      <button
        type="button"
        class="context-menu-item"
        role="menuitemcheckbox"
        :aria-checked="rowIsHeader ? 'true' : 'false'"
        data-testid="context-menu-repeat-header-row"
        @click="emit('table-action', 'repeat-header-row')"
      >
        <span class="context-menu-check" aria-hidden="true">{{ rowIsHeader ? "✓" : "" }}</span>
        <span>Repeat header row</span>
      </button>
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
/* The check gutter of the one menuitemcheckbox in this menu (SPEC-076
   AC-19), sized so its label lines up with the plain items' labels. */
.context-menu-check {
  display: inline-block;
  width: 1em;
}

.context-menu {
  position: fixed;
  z-index: 40;
  min-width: 12rem;
  /* Never taller than the screen: the placement clamp above can only SHIFT a
     menu that fits, so a menu with every section present needs a ceiling of
     its own. It is a safety net — the density rule below is what keeps the
     table section inside a short viewport without scrolling at all. */
  max-height: calc(100vh - var(--space-5));
  overflow-y: auto;
  padding: var(--space-1);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}


/* min-height, not vertical padding, sets the 32px hit target (AC-48): it keeps
   each row exactly one --height-control tall, so a menu with every section
   present still clamps inside a short viewport (SPEC-023's Table-properties
   reachability test) instead of growing past it. */
.context-menu-item {
  font: inherit;
  font-size: var(--text-md);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-6);
  width: 100%;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
  text-align: left;
}

.context-menu-item:hover {
  background: var(--state-hover);
}

.context-menu-shortcut {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.context-menu-separator {
  height: 1px;
  margin: var(--space-1) 0;
  background: var(--border-subtle);
}

/* SPEC-076 adds a "Repeat header row" item to the table section, and fifteen
   32px hit targets already fill a 500px-tall viewport. On a short window the
   rows tighten by one spacing step so the whole menu still lands inside it
   without scrolling (SPEC-023's Table-properties reachability test). Stated
   against the token rather than as a literal, so a coarse pointer — where
   --height-control is 44px because the finger needs it — tightens to 38px
   rather than to a mouse-sized target. */
@media (max-height: 640px) {
  .context-menu-item {
    min-height: calc(var(--height-control) - var(--space-2));
  }
}
</style>

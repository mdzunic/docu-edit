<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { PICKER_MAX_COLS, PICKER_MAX_ROWS } from "../lib/tables";

// The "Insert table" toolbar control (SPEC-022 AC-1, AC-2): a dropdown grid
// picker up to 10×10, plus an "Insert table…" entry that opens the dialog
// for larger sizes. Same cluster/menu pattern as ParagraphToolbar's list
// dropdowns; runtime prop/emit declarations per the espree constraint (see
// AppShell.vue).
const props = defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["insert", "open-dialog"]);

const open = ref(false);
const root = useTemplateRef("root");
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

// The grid cell the pointer is over — every cell up-left of it is
// highlighted and a click inserts that many rows × columns.
const hover = ref({ rows: 0, cols: 0 });

const GRID_ROWS = Array.from({ length: PICKER_MAX_ROWS }, (_, index) => index + 1);
const GRID_COLS = Array.from({ length: PICKER_MAX_COLS }, (_, index) => index + 1);

const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value,
});

function toggleMenu() {
  if (!props.disabled) {
    open.value = !open.value;
    hover.value = { rows: 0, cols: 0 };
  }
}

function hoverCell(rows = 0, cols = 0) {
  hover.value = { rows, cols };
}

function chooseCell(rows = 0, cols = 0) {
  open.value = false;
  emit("insert", { rows, cols });
}

function chooseDialog() {
  open.value = false;
  emit("open-dialog");
}

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (open.value && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    open.value = false;
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    open.value = false;
  }
}

document.addEventListener("mousedown", handleDocumentMousedown);
onBeforeUnmount(() => {
  document.removeEventListener("mousedown", handleDocumentMousedown);
});
</script>

<template>
  <div
    ref="root"
    class="table-control"
    role="group"
    aria-label="Insert table"
    @keydown="handleKeydown"
  >
    <button
      ref="button"
      type="button"
      class="control-button"
      data-testid="insert-table-button"
      title="Insert table"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggleMenu"
    >
      Table ▾
    </button>
    <div
      v-if="open"
      ref="menu"
      class="table-menu"
      data-testid="insert-table-menu"
      role="menu"
      aria-label="Insert table size"
      :style="menuStyle"
    >
      <div
        class="grid"
        data-testid="insert-table-grid"
        @mouseleave="hoverCell(0, 0)"
      >
        <div
          v-for="rows in GRID_ROWS"
          :key="rows"
          class="grid-row"
        >
          <button
            v-for="cols in GRID_COLS"
            :key="cols"
            type="button"
            class="grid-cell"
            :class="{ active: rows <= hover.rows && cols <= hover.cols }"
            :data-testid="`table-grid-cell-${rows}-${cols}`"
            :aria-label="`Insert ${rows} by ${cols} table`"
            @mouseenter="hoverCell(rows, cols)"
            @mousedown.prevent
            @click="chooseCell(rows, cols)"
          />
        </div>
      </div>
      <div
        class="grid-label"
        data-testid="insert-table-grid-label"
      >
        {{ hover.rows > 0 ? `${hover.rows} × ${hover.cols}` : "Select table size" }}
      </div>
      <button
        type="button"
        class="menu-item"
        role="menuitem"
        data-testid="insert-table-dialog-item"
        @mousedown.prevent
        @click="chooseDialog"
      >
        Insert table…
      </button>
    </div>
  </div>
</template>

<style scoped>
.table-control {
  position: relative;
  display: flex;
  align-items: center;
}

.control-button {
  font: inherit;
  font-size: 0.875rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.control-button:disabled {
  color: #9e9e9e;
  cursor: default;
}

.table-menu {
  z-index: 30;
  min-width: 12rem;
  padding: 0.5rem;
  border: 1px solid #dadce0;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.2);
}

.grid {
  display: flex;
  flex-direction: column;
  gap: 2px;
}

.grid-row {
  display: flex;
  gap: 2px;
}

.grid-cell {
  width: 1rem;
  height: 1rem;
  padding: 0;
  border: 1px solid #bdbdbd;
  border-radius: 2px;
  background: #fff;
  cursor: pointer;
}

.grid-cell.active {
  background: #d2e3fc;
  border-color: #1a73e8;
}

.grid-label {
  padding: 0.375rem 0 0.25rem;
  font-size: 0.8125rem;
  color: #3c4043;
  text-align: center;
}

.menu-item {
  display: block;
  width: 100%;
  font: inherit;
  font-size: 0.875rem;
  text-align: start;
  padding: 0.375rem 0.5rem;
  border: none;
  border-top: 1px solid #e0e0e0;
  background: none;
  cursor: pointer;
}

.menu-item:hover {
  background: #f1f3f4;
}
</style>

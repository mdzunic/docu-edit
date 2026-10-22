<script setup lang="ts">
// SPEC-133's table design bar (`sheet-table-design`), shown while the active
// cell is in a table and placed BELOW the grid (WorkbookBody), so appearing
// or disappearing never moves a cell under the pointer: the table's name
// (`sheet-table-name` — focus captures the table, Enter or blur renames THAT
// table, Escape restores), Resize table… (`sheet-table-resize`), Convert to
// range (`sheet-table-convert`, confirmed by SheetTableSurfaces), the style
// options (`sheet-table-option-*` checkboxes), the style gallery
// (`sheet-table-styles` → TableStyleGallery) and, in a totals cell, its
// function menu (`sheet-table-totals-menu-button` → `sheet-table-totals-menu`,
// items `sheet-table-totals-<id>`). A locked table (external data) keeps the
// style options only. State and commands are useSheetTables'. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { TOTALS_FUNCTION_LABELS, TOTALS_FUNCTIONS } from "../../lib/spreadsheet/tableCommands";
import { asSheetTables } from "../../views/spreadsheet/useSheetTables";
import UiButton from "../ui/UiButton.vue";
import SheetMenuButton from "./SheetMenuButton.vue";
import TableStyleGallery from "./TableStyleGallery.vue";

const props = defineProps({
  /** useSheetTables' return. */
  tables: { type: Object, required: true },
});

const t = computed(() => asSheetTables(props.tables));
const design = computed(() => t.value.design.value);
/** The gallery trigger's text: the style's title ("Medium 2"), or None. */
const styleTitle = computed(() => (design.value?.styleName ?? "").replace(/^TableStyle(\D+)(\d+)$/, "$1 $2") || "None");

function editName(event = new Event("input")) {
  if (event.target instanceof HTMLInputElement) {
    t.value.editName(event.target.value);
  }
}

/** Enter renames the captured table; Escape restores its name. Both return to the grid. */
function nameKey(event = new KeyboardEvent("keydown")) {
  if (event.key !== "Enter" && event.key !== "Escape") {
    return;
  }
  event.preventDefault();
  event.stopPropagation();
  if (event.key === "Enter") {
    t.value.commitRename();
  } else {
    t.value.cancelRename();
    if (event.target instanceof HTMLInputElement) {
      event.target.value = design.value?.name ?? "";
    }
  }
  t.value.focusGrid();
}

/** A checkbox's change: the command runs, and the box shows the model's
 * state again until the change lands (a refusal leaves it as it was). */
function toggle(row = t.value.optionRows.value[0], event = new Event("change")) {
  if (!row) {
    return;
  }
  t.value.setOption(row.option, !row.checked);
  if (event.target instanceof HTMLInputElement) {
    event.target.checked = row.checked;
  }
}
</script>

<template>
  <section
    v-if="design"
    class="table-design"
    aria-label="Table design"
    data-testid="sheet-table-design"
  >
    <label class="name-field">
      <span class="caption">Table name</span>
      <input
        type="text"
        class="name-input"
        data-testid="sheet-table-name"
        spellcheck="false"
        autocomplete="off"
        :value="t.nameText.value"
        :disabled="design.locked"
        @focus="t.beginRename()"
        @input="editName"
        @keydown="nameKey"
        @blur="t.commitRename()"
      >
    </label>
    <UiButton
      variant="ghost"
      size="sm"
      data-testid="sheet-table-resize"
      :disabled="design.locked"
      @mousedown.prevent
      @click="t.openResize()"
    >
      Resize table…
    </UiButton>
    <UiButton
      variant="ghost"
      size="sm"
      data-testid="sheet-table-convert"
      :disabled="design.locked"
      @mousedown.prevent
      @click="t.askConvert()"
    >
      Convert to range
    </UiButton>
    <span
      class="divider"
      aria-hidden="true"
    />
    <div
      class="options"
      role="group"
      aria-label="Table style options"
    >
      <label
        v-for="row in t.optionRows.value"
        :key="row.option"
        class="option"
        :class="{ disabled: row.disabled }"
        @mousedown.prevent
      >
        <input
          type="checkbox"
          :data-testid="row.testid"
          :checked="row.checked"
          :disabled="row.disabled"
          @change="toggle(row, $event)"
        >
        {{ row.label }}
      </label>
    </div>
    <span
      class="divider"
      aria-hidden="true"
    />
    <SheetMenuButton
      v-slot="{ close }"
      label="Table styles"
      :text="styleTitle"
      testid="sheet-table-styles"
      menu-testid="sheet-table-styles-menu"
      role="dialog"
      placement="above"
    >
      <TableStyleGallery
        :theme="t.theme.value"
        :current="design.styleName"
        @pick="(name) => { t.pickStyle(name); close(); }"
      />
    </SheetMenuButton>
    <SheetMenuButton
      v-if="t.totalsCell.value"
      v-slot="{ close }"
      label="Totals function"
      :text="t.totalsCell.value.label"
      testid="sheet-table-totals-menu-button"
      menu-testid="sheet-table-totals-menu"
      placement="above"
      :disabled="design.locked"
    >
      <button
        v-for="id in TOTALS_FUNCTIONS"
        :key="id"
        type="button"
        role="menuitemradio"
        :aria-checked="t.totalsCell.value?.fn === id ? 'true' : 'false'"
        :data-testid="`sheet-table-totals-${id}`"
        @mousedown.prevent
        @click="t.setTotalsFunction(id); close()"
      >
        {{ TOTALS_FUNCTION_LABELS[id] }}
      </button>
    </SheetMenuButton>
  </section>
</template>

<style scoped>
.table-design {
  flex: none;
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-1) var(--space-2);
  padding: var(--space-1) var(--space-3);
  background: var(--surface-raised);
  border-top: var(--border-width) solid var(--border-subtle);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.name-field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
}

.caption {
  color: var(--text-secondary);
}

.name-input {
  width: 10rem;
  height: var(--height-control);
  padding: 0 var(--space-1);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.name-input:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: 0;
}

.divider {
  align-self: stretch;
  width: var(--border-width);
  background: var(--border-subtle);
}

.options {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-1) var(--space-3);
}

.option {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
  cursor: pointer;
}

.option.disabled {
  color: var(--text-secondary);
  cursor: default;
}

[role="menuitemradio"] {
  display: block;
  width: 100%;
  padding: var(--space-1) var(--space-3);
  font: inherit;
  text-align: left;
  color: var(--text-primary);
  background: transparent;
  border: 0;
  cursor: pointer;
}

[role="menuitemradio"]:hover,
[role="menuitemradio"]:focus-visible {
  background: var(--state-hover);
}

[role="menuitemradio"][aria-checked="true"] {
  font-weight: var(--weight-bold);
}
</style>

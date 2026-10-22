<script setup lang="ts">
// File ▸ "Page setup…" (SPEC-136 "The dialog") on BaseDialog: tabs Page,
// Margins, Header/Footer and Sheet over usePageSetup's dialog state — its
// fields pageSetupForm.ts's (margins in the measurement unit preference),
// the first problem shown as an alert and OK disabled while there is one.
// OK is one undo step ("Page setup"); Cancel, Escape and the backdrop change
// nothing. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import BaseDialog from "../BaseDialog.vue";
import { HEADER_FOOTER_INSERTS, HEADER_FOOTER_PRESETS, HF_PARTS, HF_SECTIONS, MARGIN_LABELS, PAPER_CHOICES } from "../../views/spreadsheet/pageSetupForm";
import { asPageSetupDialog } from "../../views/spreadsheet/usePageSetup";

const props = defineProps({
  /** usePageSetup's dialog. */
  dialog: { type: Object, required: true },
});

const d = computed(() => asPageSetupDialog(props.dialog));
const f = computed(() => d.value.form.value);
const TABS = [
  { id: "page", label: "Page" },
  { id: "margins", label: "Margins" },
  { id: "headerFooter", label: "Header/Footer" },
  { id: "sheet", label: "Sheet" },
];

function applyPreset(part = "header", event = new Event("change")) {
  const target = event.target;
  const preset = HEADER_FOOTER_PRESETS.find((each) => target instanceof HTMLSelectElement && each.id === target.value);
  if (preset && f.value) {
    f.value[part === "footer" ? "footer" : "header"] = { left: "", center: preset.code, right: "" };
  }
}

function focus(part = "header", section = "center") {
  d.value.focused.value = { part: part === "footer" ? "footer" : "header", section: section === "left" ? "left" : section === "right" ? "right" : "center" };
}
</script>

<template>
  <BaseDialog
    v-if="d.open.value && f"
    title="Page setup"
    testid="sheet-page-setup-dialog"
    width="32rem"
    @dismiss="d.cancel()"
  >
    <div
      class="tabs"
      role="tablist"
      aria-label="Page setup"
    >
      <button
        v-for="tab in TABS"
        :key="tab.id"
        type="button"
        role="tab"
        class="tab"
        :aria-selected="d.tab.value === tab.id"
        :data-testid="`sheet-page-setup-tab-${tab.id}`"
        @click="d.selectTab(tab.id)"
      >
        {{ tab.label }}
      </button>
    </div>
    <div
      class="panel"
      role="tabpanel"
    >
      <template v-if="d.tab.value === 'page'">
        <fieldset class="group">
          <legend>Orientation</legend>
          <label><input
            v-model="f.orientation"
            type="radio"
            value="portrait"
            data-testid="sheet-page-portrait"
          > Portrait</label>
          <label><input
            v-model="f.orientation"
            type="radio"
            value="landscape"
            data-testid="sheet-page-landscape"
          > Landscape</label>
        </fieldset>
        <label class="field">Paper size
          <select
            v-model="f.paper"
            data-testid="sheet-page-paper"
          >
            <option
              v-for="name in PAPER_CHOICES"
              :key="name"
              :value="name"
            >{{ name }}</option>
          </select>
        </label>
        <fieldset class="group">
          <legend>Scaling</legend>
          <label><input
            v-model="f.scaleMode"
            type="radio"
            value="percent"
            data-testid="sheet-page-scale-mode-percent"
          > Adjust to</label>
          <input
            v-model="f.scale"
            class="short"
            type="text"
            inputmode="numeric"
            aria-label="Scale (%)"
            data-testid="sheet-page-scale"
            :disabled="f.scaleMode !== 'percent'"
          > %
          <label><input
            v-model="f.scaleMode"
            type="radio"
            value="fit"
            data-testid="sheet-page-scale-mode-fit"
          > Fit to</label>
          <input
            v-model="f.fitWide"
            class="short"
            type="text"
            inputmode="numeric"
            aria-label="Pages wide"
            placeholder="automatic"
            data-testid="sheet-page-fit-wide"
            :disabled="f.scaleMode !== 'fit'"
          > wide by
          <input
            v-model="f.fitTall"
            class="short"
            type="text"
            inputmode="numeric"
            aria-label="Pages tall"
            placeholder="automatic"
            data-testid="sheet-page-fit-tall"
            :disabled="f.scaleMode !== 'fit'"
          > tall
        </fieldset>
        <label class="field">First page number
          <input
            v-model="f.firstNumber"
            type="text"
            inputmode="numeric"
            placeholder="Auto"
            data-testid="sheet-page-first-number"
          >
        </label>
      </template>
      <template v-else-if="d.tab.value === 'margins'">
        <div class="grid">
          <label
            v-for="margin in MARGIN_LABELS"
            :key="margin.field"
            class="field"
          >{{ margin.label }} ({{ d.unit.value }})
            <input
              v-model="f.margins[margin.field]"
              type="text"
              :data-testid="`sheet-margin-${margin.field}`"
            >
          </label>
        </div>
        <fieldset class="group">
          <legend>Centre on page</legend>
          <label><input
            v-model="f.centerHorizontally"
            type="checkbox"
            data-testid="sheet-center-horizontally"
          > Horizontally</label>
          <label><input
            v-model="f.centerVertically"
            type="checkbox"
            data-testid="sheet-center-vertically"
          > Vertically</label>
        </fieldset>
      </template>
      <template v-else-if="d.tab.value === 'headerFooter'">
        <div
          v-for="{ part, label } in HF_PARTS"
          :key="part"
          class="hf"
        >
          <label class="field">{{ label }}
            <select
              :data-testid="`sheet-${part}-preset`"
              @change="applyPreset(part, $event)"
            >
              <option value="">Custom</option>
              <option
                v-for="preset in HEADER_FOOTER_PRESETS"
                :key="preset.id"
                :value="preset.id"
              >{{ preset.label }}</option>
            </select>
          </label>
          <div class="sections">
            <label
              v-for="section in HF_SECTIONS"
              :key="section.id"
              class="field"
            >{{ section.label }}
              <input
                v-model="f[part][section.id]"
                type="text"
                :data-testid="`sheet-${part}-${section.id}`"
                @focus="focus(part, section.id)"
              >
            </label>
          </div>
        </div>
        <div class="inserts">
          <button
            v-for="insert in HEADER_FOOTER_INSERTS"
            :key="insert.id"
            type="button"
            :data-testid="`sheet-hf-insert-${insert.id}`"
            @mousedown.prevent
            @click="d.insertCode(insert.code)"
          >
            {{ insert.label }}
          </button>
        </div>
        <label><input
          v-model="f.differentFirst"
          type="checkbox"
          data-testid="sheet-hf-different-first"
        > Different first page</label>
        <label><input
          v-model="f.differentOddEven"
          type="checkbox"
          data-testid="sheet-hf-different-odd-even"
        > Different odd and even pages</label>
      </template>
      <template v-else>
        <label class="field">Print area
          <input
            v-model="f.printArea"
            type="text"
            placeholder="e.g. A1:H20"
            data-testid="sheet-print-area"
          >
        </label>
        <label class="field">Rows to repeat at top
          <input
            v-model="f.titleRows"
            type="text"
            placeholder="e.g. $1:$1"
            data-testid="sheet-print-title-rows"
          >
        </label>
        <label class="field">Columns to repeat at left
          <input
            v-model="f.titleCols"
            type="text"
            placeholder="e.g. $A:$A"
            data-testid="sheet-print-title-cols"
          >
        </label>
        <label><input
          v-model="f.gridLines"
          type="checkbox"
          data-testid="sheet-print-gridlines"
        > Gridlines</label>
        <label><input
          v-model="f.headings"
          type="checkbox"
          data-testid="sheet-print-headings"
        > Row and column headings</label>
        <fieldset class="group">
          <legend>Page order</legend>
          <label><input
            v-model="f.pageOrder"
            type="radio"
            value="down"
            data-testid="sheet-page-order-down"
          > Down, then over</label>
          <label><input
            v-model="f.pageOrder"
            type="radio"
            value="over"
            data-testid="sheet-page-order-over"
          > Over, then down</label>
        </fieldset>
      </template>
      <p
        v-if="d.problem.value"
        class="problem"
        role="alert"
        data-testid="sheet-page-setup-error"
      >
        {{ d.problem.value.message }}
      </p>
    </div>
    <template #actions>
      <button
        type="button"
        data-testid="sheet-page-setup-cancel"
        @click="d.cancel()"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="sheet-page-setup-ok"
        :disabled="d.problem.value !== null"
        @click="d.ok()"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.tabs {
  display: flex;
  gap: var(--space-2);
  border-bottom: var(--border-width) solid var(--border-default);
  margin-bottom: var(--space-4);
}

.tab {
  font: inherit;
  padding: var(--space-2) var(--space-4);
  border: none;
  background: none;
  color: var(--text-secondary);
  cursor: pointer;
}

.tab[aria-selected="true"] {
  color: var(--text-primary);
  border-bottom: calc(var(--border-width) * 2) solid var(--accent);
}

.panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  min-height: 14rem;
}

.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.group {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-3);
  border: none;
  padding: 0;
  margin: 0;
}

.grid,
.sections {
  display: grid;
  grid-template-columns: repeat(3, 1fr);
  gap: var(--space-3);
}

.short {
  width: 4rem;
}

.inserts {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}

.problem {
  color: var(--danger);
  margin: 0;
}

.primary {
  font-weight: var(--weight-bold);
}
</style>

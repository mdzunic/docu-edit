<script setup lang="ts">
// Format Cells (SPEC-123 "Format Cells"): the tabs Number, Alignment, Font,
// Border, Fill and Protection (`format-cells-tab-<tab>`) over BaseDialog,
// with a live preview of the active cell's value in the pending style
// (`format-cells-preview`). Its state is useFormatCellsDialog's — it opens on
// the active cell's effective style and on OK applies only the fields the
// user changed. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { computed } from "vue";
import { STANDARD_FONT_FAMILIES } from "../../lib/core/fonts";
import { HORIZONTAL_CHOICES, LINE_STYLES, PATTERN_STYLES, VERTICAL_CHOICES } from "../../lib/spreadsheet/formatCells";
import { CURRENCY_SYMBOLS, dateTypeCodes, FRACTION_TYPE_CODES, NUMBER_CATEGORIES, TIME_TYPE_CODES } from "../../lib/spreadsheet/numberCodes";
import { asFormatCellsDialog } from "../../views/spreadsheet/useFormatCellsDialog";
import BaseDialog from "../BaseDialog.vue";
import TextColorPicker from "../TextColorPicker.vue";
import ThemeSwatchRow from "../ThemeSwatchRow.vue";

const props = defineProps({
  /** useFormatCellsDialog's return. */
  dialog: { type: Object, required: true },
});

const d = computed(() => asFormatCellsDialog(props.dialog));
const f = computed(() => d.value.fields.value);
const n = computed(() => d.value.number.value);
const target = computed(() => d.value.target.value);

const TABS = [
  { id: "number", label: "Number" },
  { id: "alignment", label: "Alignment" },
  { id: "font", label: "Font" },
  { id: "border", label: "Border" },
  { id: "fill", label: "Fill" },
  { id: "protection", label: "Protection" },
];
const EDGES = [
  { id: "top", label: "Top" },
  { id: "insideHorizontal", label: "Inside horizontal", testid: "inside-horizontal" },
  { id: "bottom", label: "Bottom" },
  { id: "left", label: "Left" },
  { id: "insideVertical", label: "Inside vertical", testid: "inside-vertical" },
  { id: "right", label: "Right" },
  { id: "diagonalUp", label: "Diagonal up", testid: "diagonal-up" },
  { id: "diagonalDown", label: "Diagonal down", testid: "diagonal-down" },
];
const UNDERLINES = [
  { value: "none", label: "None" },
  { value: "single", label: "Single" },
  { value: "double", label: "Double" },
  { value: "singleAccounting", label: "Single accounting" },
  { value: "doubleAccounting", label: "Double accounting" },
];
const FONT_STYLES = [
  { value: "regular", label: "Regular" },
  { value: "italic", label: "Italic" },
  { value: "bold", label: "Bold" },
  { value: "boldItalic", label: "Bold italic" },
];
const NEGATIVES = [
  { value: "minus", label: "-1,234.10" },
  { value: "red", label: "1,234.10 (red)" },
  { value: "parens", label: "(1,234.10)" },
  { value: "redParens", label: "(1,234.10) (red)" },
];

const usesDecimals = computed(() => ["number", "currency", "accounting", "percentage", "scientific"].includes(n.value.category));
const types = computed(() => {
  const locale = target.value?.locale;
  if (n.value.category === "date" && locale) {
    return dateTypeCodes(locale);
  }
  return n.value.category === "time" ? TIME_TYPE_CODES : n.value.category === "fraction" ? FRACTION_TYPE_CODES : [];
});
const fonts = computed(() => [...new Set([...(target.value?.documentFonts ?? []), ...STANDARD_FONT_FAMILIES])]);

/** An input or select's value, as text. */
function valueOf(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLSelectElement ? el.value : "";
}

function checked(event = new Event("change")) {
  const el = event.target;
  return el instanceof HTMLInputElement && el.checked;
}

function numberOf(event = new Event("change")) {
  const value = Number(valueOf(event));
  return Number.isFinite(value) ? value : 0;
}
</script>

<template>
  <BaseDialog
    title="Format Cells"
    testid="format-cells-dialog"
    width="34rem"
    @dismiss="d.cancel()"
  >
    <div
      class="tabs"
      role="tablist"
    >
      <button
        v-for="tab in TABS"
        :key="tab.id"
        type="button"
        role="tab"
        class="tab"
        :aria-selected="d.tab.value === tab.id ? 'true' : 'false'"
        :data-testid="`format-cells-tab-${tab.id}`"
        @click="d.showTab(tab.id)"
      >
        {{ tab.label }}
      </button>
    </div>
    <div
      class="panel"
      role="tabpanel"
    >
      <template v-if="d.tab.value === 'number'">
        <div class="number">
          <select
            class="category"
            size="11"
            aria-label="Category"
            data-testid="format-cells-category"
            :value="n.category"
            @change="d.setNumber({ category: valueOf($event) })"
          >
            <option
              v-for="category in NUMBER_CATEGORIES"
              :key="category.id"
              :value="category.id"
            >
              {{ category.label }}
            </option>
          </select>
          <div class="options">
            <div class="sample-row">
              Sample <span data-testid="format-cells-sample">{{ d.sample.value }}</span>
            </div>
            <label v-if="usesDecimals">Decimal places
              <input
                type="number"
                min="0"
                max="30"
                data-testid="format-cells-decimals"
                :value="n.decimals ?? 2"
                @input="d.setNumber({ decimals: numberOf($event) })"
              ></label>
            <label v-if="n.category === 'number'"><input
              type="checkbox"
              data-testid="format-cells-thousands"
              :checked="n.thousands === true"
              @change="d.setNumber({ thousands: checked($event) })"
            > Use 1000 separator (,)</label>
            <label v-if="n.category === 'currency' || n.category === 'accounting'">Symbol
              <select
                data-testid="format-cells-symbol"
                :value="n.symbol ?? target?.locale.currencySymbol"
                @change="d.setNumber({ symbol: valueOf($event) })"
              >
                <option value="">None</option>
                <option
                  v-for="symbol in CURRENCY_SYMBOLS"
                  :key="symbol"
                  :value="symbol"
                >{{ symbol }}</option>
              </select></label>
            <label v-if="n.category === 'number' || n.category === 'currency'">Negative numbers
              <select
                data-testid="format-cells-negative"
                :value="n.negative ?? 'minus'"
                @change="d.setNumber({ negative: valueOf($event) })"
              >
                <option
                  v-for="negative in NEGATIVES"
                  :key="negative.value"
                  :value="negative.value"
                >{{ negative.label }}</option>
              </select></label>
            <label v-if="types.length > 0">Type
              <select
                data-testid="format-cells-type"
                :value="n.type ?? types[0]"
                @change="d.setNumber({ type: valueOf($event) })"
              >
                <option
                  v-for="code in types"
                  :key="code"
                  :value="code"
                >{{ code }}</option>
              </select></label>
            <template v-if="n.category === 'custom'">
              <label>Type
                <input
                  type="text"
                  data-testid="format-cells-custom-code"
                  :value="d.customCode.value"
                  :aria-invalid="d.codeError.value ? 'true' : undefined"
                  @input="d.setCustomCode(valueOf($event))"
                ></label>
              <p
                v-if="d.codeError.value"
                class="error"
                role="alert"
                data-testid="format-cells-custom-error"
              >
                {{ d.codeError.value }}
              </p>
              <select
                size="6"
                aria-label="Number formats"
                data-testid="format-cells-custom-list"
                @change="d.setCustomCode(valueOf($event))"
              >
                <option
                  v-for="code in d.customList.value"
                  :key="code"
                  :value="code"
                >
                  {{ code }}
                </option>
              </select>
            </template>
          </div>
        </div>
      </template>
      <template v-else-if="d.tab.value === 'alignment'">
        <div class="grid">
          <label>Horizontal
            <select
              data-testid="format-cells-horizontal"
              :value="f.horizontal"
              @change="d.set('horizontal', valueOf($event))"
            >
              <option
                v-for="choice in HORIZONTAL_CHOICES"
                :key="choice.value"
                :value="choice.value"
              >{{ choice.label }}</option>
            </select></label>
          <label>Indent
            <input
              type="number"
              min="0"
              max="250"
              data-testid="format-cells-indent"
              :value="f.indent"
              @input="d.set('indent', numberOf($event))"
            ></label>
          <label>Vertical
            <select
              data-testid="format-cells-vertical"
              :value="f.vertical"
              @change="d.set('vertical', valueOf($event))"
            >
              <option
                v-for="choice in VERTICAL_CHOICES"
                :key="choice.value"
                :value="choice.value"
              >{{ choice.label }}</option>
            </select></label>
          <label><input
            type="checkbox"
            data-testid="format-cells-wrap"
            :checked="f.wrap"
            @change="d.set('wrap', checked($event))"
          > Wrap text</label>
          <label><input
            type="checkbox"
            data-testid="format-cells-shrink"
            :checked="f.shrink"
            @change="d.set('shrink', checked($event))"
          > Shrink to fit</label>
          <label>Degrees
            <input
              type="number"
              min="-90"
              max="90"
              data-testid="format-cells-rotation"
              :value="f.rotation"
              @input="d.set('rotation', numberOf($event))"
            ></label>
          <label><input
            type="checkbox"
            data-testid="format-cells-vertical-text"
            :checked="f.verticalText"
            @change="d.set('verticalText', checked($event))"
          > Vertical text</label>
        </div>
      </template>
      <template v-else-if="d.tab.value === 'font'">
        <div class="grid">
          <label>Font
            <select
              data-testid="format-cells-font-family"
              :value="f.fontFamily"
              @change="d.set('fontFamily', valueOf($event))"
            >
              <option
                v-for="family in fonts"
                :key="family"
                :value="family"
              >{{ family }}</option>
            </select></label>
          <label>Font style
            <select
              data-testid="format-cells-font-style"
              :value="f.fontStyle"
              @change="d.set('fontStyle', valueOf($event))"
            >
              <option
                v-for="choice in FONT_STYLES"
                :key="choice.value"
                :value="choice.value"
              >{{ choice.label }}</option>
            </select></label>
          <label>Size
            <input
              type="number"
              min="1"
              max="409"
              data-testid="format-cells-font-size"
              :value="f.fontSize"
              @input="d.set('fontSize', numberOf($event))"
            ></label>
          <label>Underline
            <select
              data-testid="format-cells-underline"
              :value="f.underline"
              @change="d.set('underline', valueOf($event))"
            >
              <option
                v-for="choice in UNDERLINES"
                :key="choice.value"
                :value="choice.value"
              >{{ choice.label }}</option>
            </select></label>
          <div
            class="color"
            data-testid="format-cells-font-color"
          >
            Colour
            <TextColorPicker
              :recent="target?.recentColors ?? []"
              :theme="target?.theme"
              @select="d.set('fontColor', $event)"
            />
          </div>
          <label><input
            type="checkbox"
            data-testid="format-cells-strike"
            :checked="f.strike"
            @change="d.set('strike', checked($event))"
          > Strikethrough</label>
          <label><input
            type="checkbox"
            data-testid="format-cells-superscript"
            :checked="f.superscript"
            @change="d.set('superscript', checked($event)); d.set('subscript', false)"
          > Superscript</label>
          <label><input
            type="checkbox"
            data-testid="format-cells-subscript"
            :checked="f.subscript"
            @change="d.set('subscript', checked($event)); d.set('superscript', false)"
          > Subscript</label>
        </div>
      </template>
      <template v-else-if="d.tab.value === 'border'">
        <div class="grid">
          <label>Style
            <select
              data-testid="format-cells-border-style"
              :value="d.lineStyle.value"
              @change="d.setLineStyle(valueOf($event))"
            >
              <option
                v-for="style in LINE_STYLES"
                :key="style"
                :value="style"
              >{{ style }}</option>
            </select></label>
          <div
            class="color"
            data-testid="format-cells-border-color"
          >
            Colour
            <TextColorPicker
              :theme="target?.theme"
              @select="d.lineColor.value = $event"
            />
          </div>
          <div class="presets">
            <button
              v-for="preset in ['none', 'outline', 'inside']"
              :key="preset"
              type="button"
              :data-testid="`format-cells-border-preset-${preset}`"
              :disabled="preset === 'inside' && !d.multiCell.value"
              @click="d.preset(preset)"
            >
              {{ preset === "none" ? "None" : preset === "outline" ? "Outline" : "Inside" }}
            </button>
          </div>
          <div class="edges">
            <button
              v-for="edge in EDGES"
              :key="edge.id"
              type="button"
              :aria-pressed="d.edgeShown(edge.id) ? 'true' : 'false'"
              :data-testid="`format-cells-border-edge-${edge.testid ?? edge.id}`"
              :disabled="(edge.id === 'insideHorizontal' || edge.id === 'insideVertical') && !d.multiCell.value"
              @click="d.toggleEdge(edge.id)"
            >
              {{ edge.label }}
            </button>
          </div>
        </div>
      </template>
      <template v-else-if="d.tab.value === 'fill'">
        <div class="grid">
          <div data-testid="format-cells-fill-color">
            <button
              type="button"
              data-testid="format-cells-no-fill"
              @click="d.set('fillColor', null)"
            >
              No colour
            </button>
            <ThemeSwatchRow
              :theme="target?.theme"
              testid="format-cells-fill-theme"
              @select="d.set('fillColor', $event)"
            />
          </div>
          <label>Pattern style
            <select
              data-testid="format-cells-pattern-style"
              :value="f.patternStyle"
              @change="d.set('patternStyle', valueOf($event))"
            >
              <option
                v-for="pattern in PATTERN_STYLES"
                :key="pattern"
                :value="pattern"
              >{{ pattern }}</option>
            </select></label>
          <div
            class="color"
            data-testid="format-cells-pattern-color"
          >
            Pattern colour
            <TextColorPicker
              :theme="target?.theme"
              @select="d.set('patternColor', $event)"
            />
          </div>
        </div>
      </template>
      <template v-else>
        <div class="grid">
          <label><input
            type="checkbox"
            data-testid="format-cells-locked"
            :checked="f.locked"
            @change="d.set('locked', checked($event))"
          > Locked</label>
          <label><input
            type="checkbox"
            data-testid="format-cells-hidden"
            :checked="f.hidden"
            @change="d.set('hidden', checked($event))"
          > Hidden</label>
          <p class="note">
            Locking cells or hiding formulas has no effect until the sheet is protected.
          </p>
        </div>
      </template>
    </div>
    <div class="preview-row">
      Preview
      <div
        class="preview"
        data-testid="format-cells-preview"
        :style="d.preview.value.box"
      >
        <span :style="d.preview.value.text_">{{ d.preview.value.text }}</span>
      </div>
    </div>
    <template #actions>
      <button
        type="button"
        data-testid="format-cells-cancel"
        @click="d.cancel()"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="format-cells-ok"
        :disabled="d.codeError.value !== ''"
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
  gap: var(--space-1);
  border-bottom: var(--border-width) solid var(--border-default);
  margin-bottom: var(--space-3);
}

.tab {
  font: inherit;
  padding: var(--space-1) var(--space-3);
  background: transparent;
  border: 0;
  border-bottom: calc(2 * var(--border-width)) solid transparent;
  cursor: pointer;
}

.tab[aria-selected="true"] {
  border-bottom-color: var(--accent);
}

.panel {
  min-height: 14rem;
}

.number {
  display: flex;
  gap: var(--space-4);
}

.category {
  min-width: 9rem;
}

.options,
.grid {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.error {
  margin: 0;
  color: var(--text-danger);
}

.color,
.presets,
.edges {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
}

.edges button[aria-pressed="true"] {
  background: var(--surface-selected);
}

.note {
  color: var(--text-secondary);
}

.preview-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-top: var(--space-3);
}

.preview {
  display: flex;
  flex-direction: column;
  width: 10rem;
  height: 2rem;
  overflow: hidden;
  border: var(--border-width) solid var(--border-subtle);
}
</style>

<script setup lang="ts">
import { computed, nextTick, onMounted, reactive, ref, useTemplateRef } from "vue";
import { asStringArray, asStyleOptions, noStyleOptions, styleFormInputFromDraft } from "../lib/styles";
import { asParagraphBorders } from "../lib/documentModel";
import { asTableBandingDraft } from "../lib/tableStyles";
import { detectBorderPreset, noParagraphBorders, noShadingFill } from "../lib/paragraphBorders";
import BaseDialog from "./BaseDialog.vue";
import BordersShadingDialog from "./BordersShadingDialog.vue";

/** How the summary line names each preset (SPEC-074 AC-8's vocabulary). */
const BORDER_PRESET_LABELS = {
  none: "None",
  box: "Box",
  all: "All",
  topAndBottom: "Top and bottom",
  custom: "Custom",
};

// "New style" / "Modify style" (SPEC-042 AC-3, AC-4): one dialog for both —
// `mode` decides the title/submit label and whether `type` is editable (a
// style's family cannot change after creation). BaseDialog owns the shell,
// Escape and initial focus (the [autofocus] name input); selecting the
// prefilled name on top of that is this dialog's own one-shot onMounted.
// Runtime prop/emit declarations, not defineProps<{...}>() generics — see
// AppShell.vue's comment on the same espree constraint.
const props = defineProps({
  mode: { type: String, default: "new" }, // "new" | "modify"
  initialName: { type: String, default: "" },
  initialType: { type: String, default: "paragraph" },
  initialBasedOn: { type: String, default: "" },
  initialNext: { type: String, default: "" },
  initialFontFamily: { type: String, default: "" },
  initialFontSizePt: { type: Number, default: undefined },
  initialBold: { type: Boolean, default: false },
  initialItalic: { type: Boolean, default: false },
  initialUnderline: { type: Boolean, default: false },
  initialColor: { type: String, default: "" },
  initialAlignment: { type: String, default: "" }, // "" | left | center | right | justify
  initialListPresetKey: { type: String, default: "decimal" },
  initialCellShadingFill: { type: String, default: "" },
  initialCellBorderStyle: { type: String, default: "" }, // "" | none | single | double | dashed | dotted
  initialCellBorderWidthPt: { type: Number, default: 1 },
  initialCellBorderColor: { type: String, default: "#000000" },
  /** The table style's banding as an editor draft (SPEC-076 AC-37) —
   * lib/tableStyles.bandingDraftFrom builds it from the style. */
  initialBanding: { type: Object, default: undefined },
  /** The style's own paragraph borders and shading (SPEC-074 AC-18). */
  initialBorders: { type: Object, default: undefined },
  initialShadingFill: { type: String, default: "" },
  /** Every other style's { id, name }, for the "based on"/"next" pickers —
   * the caller filters out the style being modified itself. */
  basedOnOptions: { type: Array, default: () => noStyleOptions() },
  nextOptions: { type: Array, default: () => noStyleOptions() },
  /** Other styles' trimmed, lower-cased names — a collision check input
   * (AC-7), computed by the caller from the live catalogue. */
  existingNames: { type: Array, default: () => [] },
});

const emit = defineEmits(["save", "cancel"]);

const ALIGNMENTS = ["left", "center", "right", "justify"];
const LIST_PRESETS = [
  { key: "decimal", label: "1. / 1.1 / 1.1.1" },
  { key: "outline", label: "I. / A. / 1." },
  { key: "legal", label: "1 / 1.1 / 1.1.1" },
  { key: "bulletMixed", label: "Bulleted (mixed)" },
];
const BORDER_STYLES = ["none", "single", "double", "dashed", "dotted"];

const draft = reactive({
  name: props.initialName,
  type: props.initialType,
  basedOn: props.initialBasedOn,
  next: props.initialNext,
  fontFamily: props.initialFontFamily,
  fontSizePt: props.initialFontSizePt,
  bold: props.initialBold,
  italic: props.initialItalic,
  underline: props.initialUnderline,
  color: props.initialColor,
  alignment: props.initialAlignment,
  listPresetKey: props.initialListPresetKey,
  cellShadingFill: props.initialCellShadingFill,
  cellBorderStyle: props.initialCellBorderStyle,
  cellBorderWidthPt: props.initialCellBorderWidthPt,
  cellBorderColor: props.initialCellBorderColor,
  // SPEC-074 AC-18: edited through the shared Borders and shading dialog,
  // so a style's border controls are literally a paragraph's.
  borders: asParagraphBorders(props.initialBorders),
  shadingFill: props.initialShadingFill,
  // SPEC-076 AC-37: the two band sizes and the six regions.
  banding: asTableBandingDraft(props.initialBanding),
});

/** The nested Borders and shading dialog (AC-18) — the same component the
 * Format menu opens, seeded from and writing back to this draft. */
const bordersOpen = ref(false);

const borderSummary = computed(() => {
  const preset = detectBorderPreset(draft.borders);
  const fill = draft.shadingFill ? `, shaded ${draft.shadingFill}` : "";
  return preset === "none" && !draft.shadingFill ? "None" : `${BORDER_PRESET_LABELS[preset]}${fill}`;
});

function applyStyleBorders(result = { borders: noParagraphBorders(), shadingFill: noShadingFill() }) {
  draft.borders = result.borders ?? undefined;
  draft.shadingFill = result.shadingFill ?? "";
  bordersOpen.value = false;
}

// BaseDialog focuses the [autofocus] name input; selecting the prefilled
// name on top of that (type-to-replace) mirrors RenameDialog.
const nameInput = useTemplateRef("nameInput");
onMounted(async () => {
  await nextTick();
  nameInput.value?.select();
});

const trimmedName = computed(() => draft.name.trim());
const nameError = computed(() => {
  if (trimmedName.value.length === 0) {
    return "Name cannot be empty";
  }
  const names = asStringArray(props.existingNames);
  if (names.includes(trimmedName.value.toLowerCase())) {
    return "A style with this name already exists";
  }
  return "";
});

function save() {
  if (nameError.value) {
    return;
  }
  emit("save", styleFormInputFromDraft({ ...draft, name: trimmedName.value }));
}

const title = computed(() => (props.mode === "new" ? "New style" : "Modify style"));
const submitLabel = computed(() => (props.mode === "new" ? "Create" : "Save"));

// Typed pass-throughs for the untyped Array props (espree bars a generic
// cast inside a .vue script — see documentModel.asCharacterFormatKey).
const basedOnList = computed(() => asStyleOptions(props.basedOnOptions));
const nextList = computed(() => asStyleOptions(props.nextOptions));
</script>

<template>
  <BaseDialog
    :title="title"
    testid="style-editor-dialog"
    @dismiss="emit('cancel')"
  >
    <label class="field">
      <span>Name</span>
      <input
        ref="nameInput"
        v-model="draft.name"
        data-testid="style-name-input"
        autofocus
        @keydown.enter="save"
      >
    </label>
    <p
      v-if="nameError"
      class="dialog-error"
      data-testid="style-name-error"
    >
      {{ nameError }}
    </p>

    <label class="field">
      <span>Type</span>
      <select
        v-model="draft.type"
        data-testid="style-type-select"
        :disabled="mode === 'modify'"
      >
        <option value="paragraph">
          Paragraph
        </option>
        <option value="character">
          Character
        </option>
        <option value="list">
          List
        </option>
        <option value="table">
          Table
        </option>
      </select>
    </label>

    <label class="field">
      <span>Based on</span>
      <select
        v-model="draft.basedOn"
        data-testid="style-based-on-select"
      >
        <option value="">
          (none)
        </option>
        <option
          v-for="option in basedOnList"
          :key="option.id"
          :value="option.id"
        >
          {{ option.name }}
        </option>
      </select>
    </label>

    <label
      v-if="draft.type === 'paragraph' || draft.type === 'list'"
      class="field"
    >
      <span>Style for following paragraph</span>
      <select
        v-model="draft.next"
        data-testid="style-next-select"
      >
        <option value="">
          (same style)
        </option>
        <option
          v-for="option in nextList"
          :key="option.id"
          :value="option.id"
        >
          {{ option.name }}
        </option>
      </select>
    </label>

    <template v-if="draft.type === 'paragraph' || draft.type === 'character' || draft.type === 'list'">
      <fieldset class="field-group">
        <legend>Font</legend>
        <label class="field">
          <span>Font family</span>
          <input
            v-model="draft.fontFamily"
            data-testid="style-font-family-input"
          >
        </label>
        <label class="field">
          <span>Font size (pt)</span>
          <input
            v-model.number="draft.fontSizePt"
            type="number"
            min="1"
            data-testid="style-font-size-input"
          >
        </label>
        <label class="checkbox-field">
          <input
            v-model="draft.bold"
            type="checkbox"
            data-testid="style-bold-checkbox"
          >
          <span>Bold</span>
        </label>
        <label class="checkbox-field">
          <input
            v-model="draft.italic"
            type="checkbox"
            data-testid="style-italic-checkbox"
          >
          <span>Italic</span>
        </label>
        <label class="checkbox-field">
          <input
            v-model="draft.underline"
            type="checkbox"
            data-testid="style-underline-checkbox"
          >
          <span>Underline</span>
        </label>
        <label class="field">
          <span>Colour</span>
          <input
            v-model="draft.color"
            placeholder="#000000"
            data-testid="style-color-input"
          >
        </label>
      </fieldset>
    </template>

    <fieldset
      v-if="draft.type === 'paragraph'"
      class="field-group"
    >
      <legend>Paragraph</legend>
      <label class="field">
        <span>Alignment</span>
        <select
          v-model="draft.alignment"
          data-testid="style-alignment-select"
        >
          <option value="">
            (inherit)
          </option>
          <option
            v-for="alignment in ALIGNMENTS"
            :key="alignment"
            :value="alignment"
          >
            {{ alignment }}
          </option>
        </select>
      </label>
      <!-- SPEC-074 AC-18: a style carries borders and shading, edited by the
           same dialog a paragraph's are. -->
      <div class="field">
        <span>Borders and shading</span>
        <button
          type="button"
          class="borders-button"
          data-testid="style-borders-button"
          @click="bordersOpen = true"
        >
          {{ borderSummary }}…
        </button>
      </div>
    </fieldset>

    <fieldset
      v-if="draft.type === 'list'"
      class="field-group"
    >
      <legend>Numbering</legend>
      <label class="field">
        <span>Preset</span>
        <select
          v-model="draft.listPresetKey"
          data-testid="style-list-preset-select"
        >
          <option
            v-for="preset in LIST_PRESETS"
            :key="preset.key"
            :value="preset.key"
          >
            {{ preset.label }}
          </option>
        </select>
      </label>
    </fieldset>

    <fieldset
      v-if="draft.type === 'table'"
      class="field-group"
    >
      <legend>Borders &amp; shading</legend>
      <label class="field">
        <span>Border style</span>
        <select
          v-model="draft.cellBorderStyle"
          data-testid="style-border-style-select"
        >
          <option value="">
            (none)
          </option>
          <option
            v-for="style in BORDER_STYLES"
            :key="style"
            :value="style"
          >
            {{ style }}
          </option>
        </select>
      </label>
      <label
        v-if="draft.cellBorderStyle"
        class="field"
      >
        <span>Border width (pt)</span>
        <input
          v-model.number="draft.cellBorderWidthPt"
          type="number"
          min="0.25"
          step="0.25"
          data-testid="style-border-width-input"
        >
      </label>
      <label
        v-if="draft.cellBorderStyle"
        class="field"
      >
        <span>Border colour</span>
        <input
          v-model="draft.cellBorderColor"
          data-testid="style-border-color-input"
        >
      </label>
      <label class="field">
        <span>Shading</span>
        <input
          v-model="draft.cellShadingFill"
          placeholder="#eeeeee"
          data-testid="style-shading-input"
        >
      </label>
    </fieldset>

    <!-- SPEC-076 AC-37: the style's conditional formatting. Each region's
         shading, bold and border style; the border width and colour above
         are shared, so a style describes its borders once. -->
    <fieldset
      v-if="draft.type === 'table'"
      class="field-group"
    >
      <legend>Banding</legend>
      <div class="band-sizes">
        <label class="field">
          <span>Row band size</span>
          <input
            v-model.number="draft.banding.rowBandSize"
            type="number"
            min="1"
            step="1"
            data-testid="style-row-band-size"
          >
        </label>
        <label class="field">
          <span>Column band size</span>
          <input
            v-model.number="draft.banding.columnBandSize"
            type="number"
            min="1"
            step="1"
            data-testid="style-column-band-size"
          >
        </label>
      </div>
      <div
        v-for="region in draft.banding.regions"
        :key="region.key"
        class="band-region"
      >
        <span class="band-region-label">{{ region.label }}</span>
        <input
          v-model="region.shading"
          placeholder="#eeeeee"
          :aria-label="`${region.label} shading`"
          :data-testid="`style-banding-shading-${region.key}`"
        >
        <label class="band-region-bold">
          <input
            v-model="region.bold"
            type="checkbox"
            :data-testid="`style-banding-bold-${region.key}`"
          >
          <span>Bold</span>
        </label>
        <select
          v-model="region.borderStyle"
          :aria-label="`${region.label} border style`"
          :data-testid="`style-banding-border-${region.key}`"
        >
          <option value="">
            (no border)
          </option>
          <option
            v-for="style in BORDER_STYLES"
            :key="style"
            :value="style"
          >
            {{ style }}
          </option>
        </select>
      </div>
    </fieldset>

    <template #actions>
      <button
        type="button"
        data-testid="style-editor-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="style-editor-save"
        :disabled="!!nameError"
        @click="save"
      >
        {{ submitLabel }}
      </button>
    </template>
    <BordersShadingDialog
      v-if="bordersOpen"
      title="Style borders and shading"
      :initial="{ borders: draft.borders, shadingFill: draft.shadingFill || undefined }"
      @apply="applyStyleBorders"
      @cancel="bordersOpen = false"
    />
  </BaseDialog>
</template>

<style scoped>
/* The banding grid (SPEC-076 AC-37): one row per region — label, shading,
   bold, border style — under the two band-size fields. */
.band-sizes {
  display: flex;
  gap: var(--space-3);
}

.band-region {
  display: grid;
  grid-template-columns: 8rem 1fr auto 8rem;
  align-items: center;
  gap: var(--space-2);
  margin-top: var(--space-2);
}

.band-region-label {
  font-size: var(--text-md);
  color: var(--text-primary);
}

.band-region-bold {
  display: flex;
  align-items: center;
  gap: var(--space-1);
  font-size: var(--text-md);
  color: var(--text-primary);
}

.borders-button {
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-lg);
  background: var(--surface-base);
  color: inherit;
  cursor: pointer;
  font: inherit;
  text-align: left;
}


/* Fixed 26rem panel with the tall-dialog scroll cap. The actions row now
   sits at the base's 1.25rem gap rather than this dialog's old 0.5rem — a
   deliberate unification with every other dialog's spacing. */
:deep(.dialog) {
  width: 26rem;
  max-height: 85vh;
  overflow-y: auto;
}

.field {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  margin-bottom: var(--space-4);
  font-size: var(--text-md);
}

.field input,
.field select {
  font: inherit;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  margin-bottom: var(--space-3);
  font-size: var(--text-md);
}

.field-group {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  padding: var(--space-3) var(--space-4) var(--space-4);
  margin: 0 0 var(--space-4);
}

.field-group legend {
  font-size: var(--text-sm);
  color: var(--text-secondary);
  padding: 0 var(--space-2);
}

.dialog-error {
  margin: calc(-1 * var(--space-3)) 0 var(--space-4);
  color: var(--text-danger);
  font-size: var(--text-md);
}

.dialog-actions button {
  font: inherit;
  padding: var(--space-3) var(--space-5);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.dialog-actions button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>

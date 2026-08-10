<script setup lang="ts">
import { computed, nextTick, onMounted, reactive, useTemplateRef } from "vue";
import { asStringArray, asStyleOptions, noStyleOptions, styleFormInputFromDraft } from "../lib/styles";

// "New style" / "Modify style" (SPEC-042 AC-3, AC-4): one dialog for both —
// `mode` decides the title/submit label and whether `type` is editable (a
// style's family cannot change after creation). Runtime prop/emit
// declarations, not defineProps<{...}>() generics — see AppShell.vue's
// comment on the same espree constraint.
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
});

const nameInput = useTemplateRef("nameInput");
onMounted(async () => {
  await nextTick();
  nameInput.value?.focus();
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
  <div
    class="dialog-backdrop"
    data-testid="style-editor-dialog"
    @click.self="emit('cancel')"
    @keydown.escape="emit('cancel')"
  >
    <div
      class="dialog"
      role="dialog"
      aria-modal="true"
      :aria-label="title"
    >
      <h2 class="dialog-title">
        {{ title }}
      </h2>

      <label class="field">
        <span>Name</span>
        <input
          ref="nameInput"
          v-model="draft.name"
          data-testid="style-name-input"
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

      <div class="dialog-actions">
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
      </div>
    </div>
  </div>
</template>

<style scoped>
.dialog-backdrop {
  position: fixed;
  inset: 0;
  background: rgb(0 0 0 / 0.4);
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 50;
}

.dialog {
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 4px 16px rgb(0 0 0 / 0.3);
  padding: 1.25rem 1.5rem;
  width: 26rem;
  max-width: 90vw;
  max-height: 85vh;
  overflow-y: auto;
  box-sizing: border-box;
}

.dialog-title {
  margin: 0 0 1rem;
  font-size: 1rem;
  font-weight: 600;
  color: #202124;
}

.field {
  display: flex;
  flex-direction: column;
  gap: 0.25rem;
  margin-bottom: 0.75rem;
  font-size: 0.8125rem;
}

.field input,
.field select {
  font: inherit;
  padding: 0.375rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
}

.checkbox-field {
  display: flex;
  align-items: center;
  gap: 0.375rem;
  margin-bottom: 0.5rem;
  font-size: 0.8125rem;
}

.field-group {
  border: 1px solid #e8eaed;
  border-radius: 4px;
  padding: 0.5rem 0.75rem 0.75rem;
  margin: 0 0 0.75rem;
}

.field-group legend {
  font-size: 0.75rem;
  color: #5f6368;
  padding: 0 0.25rem;
}

.dialog-error {
  margin: -0.5rem 0 0.75rem;
  color: #b3261e;
  font-size: 0.8125rem;
}

.dialog-actions {
  display: flex;
  justify-content: flex-end;
  gap: 0.5rem;
  margin-top: 0.5rem;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.dialog-actions button.primary {
  background: #2c3e50;
  border-color: #2c3e50;
  color: #fff;
}

.dialog-actions button:disabled {
  opacity: 0.5;
  cursor: not-allowed;
}
</style>

<script setup lang="ts">
// The Format → Borders and shading… dialog (SPEC-074 AC-6..AC-10): a preset
// row, then each of the five paragraph edges — top, bottom, left, right and
// between — with its own style, width and colour, its own padding in the
// document's measurement unit (SPEC-064), and a paragraph shading control
// whose "No fill" is a distinct value from white. OK emits the resolved
// borders and fill for the parent to apply through a command; Cancel closes.
// BaseDialog owns the shell, Escape and initial focus.
//
// The same control set is what a *style* edits (AC-18): the style editor
// mounts this dialog with the style's own borders as `initial` and applies
// the result to the style definition instead of the selection.
//
// Every edge/preset decision lives in lib/paragraphBorders.ts's draft
// helpers — this script is parsed by espree (see AppShell.vue) and so cannot
// carry the annotations that state needs, and a pure draft is testable
// without mounting the component.
import { computed, ref, watch } from "vue";
import {
  applyPresetToDraft,
  asBorderDraftEdge,
  asParagraphBorderPreset,
  borderDraftFrom,
  bordersFromDraft,
  asBorderEdge,
  BORDER_DRAFT_EDGES,
  BORDER_PADDING_EDGES,
  detectBorderPreset,
  emptyBorderDraft,
  emptyBorderPaddingText,
  isEventTargetChecked,
  setDraftEdgeStyle,
  setDraftEdgeWidth,
} from "../lib/paragraphBorders";
import { asParagraphBorders } from "../lib/documentModel";
import { BORDER_LINE_STYLES } from "../lib/core/borders";
import {
  asMeasurementUnit,
  formatLength,
  parseLengthToCm,
  CM_PER_INCH,
  POINTS_PER_INCH,
} from "../lib/core/units";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

const props = defineProps({
  /** `{ borders?: ParagraphBorders, shadingFill?: string }` — the formatting
   * of the first paragraph the selection touches, or a style's own. */
  initial: { type: Object, required: true },
  /** The document's measurement-unit preference (SPEC-064): the padding
   * fields display and parse in it; the model stays in points. */
  unit: { type: String, default: "cm" },
  /** Dialog title — the style editor says whose borders these are. */
  title: { type: String, default: "Borders and shading" },
});

const emit = defineEmits(["apply", "cancel"]);

const EDGE_LABELS = {
  top: "Top",
  bottom: "Bottom",
  left: "Left",
  right: "Right",
  between: "Between",
};

const PRESETS = [
  { key: "none", label: "None" },
  { key: "box", label: "Box" },
  { key: "all", label: "All" },
  { key: "topAndBottom", label: "Top and bottom" },
  { key: "custom", label: "Custom" },
];

const STYLE_LABELS = {
  none: "None",
  single: "Single",
  thick: "Thick",
  double: "Double",
  dashed: "Dashed",
  dotted: "Dotted",
  dashDot: "Dash-dot",
  dashDotDot: "Dash-dot-dot",
};

const measurementUnit = computed(() => asMeasurementUnit(props.unit));
const unitLabel = computed(() => props.unit);

const draft = ref(emptyBorderDraft());
draft.value = borderDraftFrom(asParagraphBorders(props.initial.borders));

/** Padding drafts, held as display text in the current unit; points stay
 * canonical (the model's unit, and OOXML's @w:space). */
const paddingDrafts = ref(emptyBorderPaddingText());
for (const edge of BORDER_PADDING_EDGES) {
  paddingDrafts.value[edge] = formatLength(ptToCm(draft.value.paddingPt[edge]), measurementUnit.value);
}

const shadingOn = ref(typeof props.initial.shadingFill === "string");
// "No fill" is distinct from white (AC-10): the checkbox is the distinction,
// so the colour input can hold #ffffff and still mean "white", not "none".
const shadingFill = ref(
  typeof props.initial.shadingFill === "string" ? props.initial.shadingFill : "#fff2cc",
);

/** The preset row reflects what the edges actually are, so hand-editing one
 * edge after picking "Box" moves the selection to "Custom" (AC-8). */
const preset = computed(() => detectBorderPreset(bordersFromDraft(draft.value)));

function ptToCm(pt = 0) {
  return (pt / POINTS_PER_INCH) * CM_PER_INCH;
}

function cmToPt(cm = 0) {
  return (cm / CM_PER_INCH) * POINTS_PER_INCH;
}

/** The padding a field currently holds, in whole points — the grid OOXML's
 * `@w:space` carries. An unparseable entry keeps the draft's value. */
function paddingPtOf(key = "top") {
  const edge = asBorderEdge(key);
  const cm = parseLengthToCm(paddingDrafts.value[edge], measurementUnit.value);
  if (cm === null || cm < 0) {
    return draft.value.paddingPt[edge];
  }
  return Math.min(31, Math.round(cmToPt(cm)));
}

function toggleEdge(key = "top", on = false) {
  const edge = asBorderDraftEdge(key);
  const next = { ...draft.value, paddingPt: { ...draft.value.paddingPt } };
  next[edge] = { ...draft.value[edge], on };
  draft.value = next;
}

function choosePreset(key = "none") {
  draft.value = applyPresetToDraft(asParagraphBorderPreset(key), draft.value);
}

function onStyleChange(key = "top", event = new Event("change")) {
  const target = event.target;
  const value = target instanceof HTMLSelectElement ? target.value : "single";
  draft.value = setDraftEdgeStyle(draft.value, asBorderDraftEdge(key), value);
}

function onWidthChange(key = "top", event = new Event("change")) {
  const target = event.target;
  const value = target instanceof HTMLInputElement ? Number(target.value) : Number.NaN;
  draft.value = setDraftEdgeWidth(draft.value, asBorderDraftEdge(key), value);
}

function onColorChange(key = "top", event = new Event("input")) {
  const target = event.target;
  if (!(target instanceof HTMLInputElement)) {
    return;
  }
  const edge = asBorderDraftEdge(key);
  const next = { ...draft.value, paddingPt: { ...draft.value.paddingPt } };
  next[edge] = { ...draft.value[edge], color: target.value };
  draft.value = next;
}

// Reformat the padding fields when the unit changes, re-deriving each from
// its underlying value so points stay the source of truth (SPEC-064).
watch(
  () => props.unit,
  (next = "cm", previous = "cm") => {
    const previousUnit = asMeasurementUnit(previous);
    const nextUnit = asMeasurementUnit(next);
    for (const edge of BORDER_PADDING_EDGES) {
      const cm = parseLengthToCm(paddingDrafts.value[edge], previousUnit);
      if (cm !== null) {
        paddingDrafts.value[edge] = formatLength(cm, nextUnit);
      }
    }
  },
);

function apply() {
  const resolved = {
    ...draft.value,
    paddingPt: {
      top: paddingPtOf("top"),
      bottom: paddingPtOf("bottom"),
      left: paddingPtOf("left"),
      right: paddingPtOf("right"),
    },
  };
  emit("apply", {
    borders: bordersFromDraft(resolved) ?? null,
    shadingFill: shadingOn.value ? shadingFill.value : null,
  });
}
</script>

<template>
  <BaseDialog
    :title="title"
    testid="borders-shading-dialog"
    @dismiss="emit('cancel')"
  >
    <div class="presets">
      <span class="field-label">Setting</span>
      <div class="preset-buttons">
        <button
          v-for="option in PRESETS"
          :key="option.key"
          type="button"
          class="preset-button"
          :class="{ active: preset === option.key }"
          :data-testid="`borders-preset-${option.key}`"
          :aria-pressed="preset === option.key"
          :disabled="option.key === 'custom'"
          @click="choosePreset(option.key)"
        >
          {{ option.label }}
        </button>
      </div>
    </div>

    <table class="edge-table">
      <thead>
        <tr>
          <th scope="col">
            Edge
          </th>
          <th scope="col">
            Style
          </th>
          <th scope="col">
            Width (pt)
          </th>
          <th scope="col">
            Colour
          </th>
        </tr>
      </thead>
      <tbody>
        <tr
          v-for="edge in BORDER_DRAFT_EDGES"
          :key="edge"
        >
          <th scope="row">
            <label class="edge-toggle">
              <input
                type="checkbox"
                :checked="draft[edge].on"
                :data-testid="`borders-edge-${edge}`"
                :aria-label="`${EDGE_LABELS[edge]} border`"
                @change="toggleEdge(edge, isEventTargetChecked($event.target))"
              >
              {{ EDGE_LABELS[edge] }}
            </label>
          </th>
          <td>
            <select
              class="field-input"
              :value="draft[edge].on ? draft[edge].style : 'none'"
              :data-testid="`borders-style-${edge}`"
              :aria-label="`${EDGE_LABELS[edge]} border style`"
              @change="onStyleChange(edge, $event)"
            >
              <option
                v-for="style in BORDER_LINE_STYLES"
                :key="style"
                :value="style"
              >
                {{ STYLE_LABELS[style] }}
              </option>
            </select>
          </td>
          <td>
            <input
              class="field-input width-input"
              type="number"
              min="0.25"
              max="12"
              step="0.25"
              :value="draft[edge].widthPt"
              :disabled="!draft[edge].on"
              :data-testid="`borders-width-${edge}`"
              :aria-label="`${EDGE_LABELS[edge]} border width in points`"
              @change="onWidthChange(edge, $event)"
            >
          </td>
          <td>
            <input
              class="colour-input"
              type="color"
              :value="draft[edge].color"
              :disabled="!draft[edge].on"
              :data-testid="`borders-color-${edge}`"
              :aria-label="`${EDGE_LABELS[edge]} border colour`"
              @input="onColorChange(edge, $event)"
            >
          </td>
        </tr>
      </tbody>
    </table>

    <fieldset class="padding-group">
      <legend>Padding ({{ unitLabel }})</legend>
      <label
        v-for="edge in BORDER_PADDING_EDGES"
        :key="edge"
        class="field"
      >
        <span class="field-label">{{ edge }}</span>
        <input
          v-model="paddingDrafts[edge]"
          class="field-input"
          type="text"
          inputmode="decimal"
          :data-testid="`borders-padding-${edge}`"
        >
      </label>
    </fieldset>

    <fieldset class="shading-group">
      <legend>Shading</legend>
      <label class="shading-toggle">
        <input
          v-model="shadingOn"
          type="checkbox"
          data-testid="borders-shading-on"
        >
        Fill
      </label>
      <input
        v-model="shadingFill"
        class="colour-input"
        type="color"
        :disabled="!shadingOn"
        data-testid="borders-shading-color"
        aria-label="Shading colour"
      >
      <p class="hint">
        Clear "Fill" for no fill — distinct from a white fill.
      </p>
    </fieldset>

    <template #actions>
      <UiButton
        variant="secondary"
        testid="borders-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        testid="borders-apply"
        @click="apply"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.presets {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 12px;
}

.preset-buttons {
  display: flex;
  gap: 6px;
}

.preset-button {
  padding: 4px 10px;
  border: 1px solid var(--border-subtle, #dadce0);
  border-radius: 4px;
  background: var(--surface-1, #fff);
  cursor: pointer;
  font-size: 13px;
}

.preset-button.active {
  border-color: var(--accent, #1a73e8);
  background: var(--surface-accent-subtle, #e8f0fe);
}

.preset-button:disabled {
  cursor: default;
  opacity: 0.75;
}

.edge-table {
  width: 100%;
  border-collapse: collapse;
  font-size: 13px;
}

.edge-table th,
.edge-table td {
  padding: 4px 6px;
  text-align: left;
  font-weight: normal;
}

.edge-toggle {
  display: flex;
  align-items: center;
  gap: 6px;
  white-space: nowrap;
}

.field-input {
  width: 100%;
  box-sizing: border-box;
  padding: 4px 6px;
  border: 1px solid var(--border-subtle, #dadce0);
  border-radius: 4px;
  background: var(--surface-1, #fff);
  color: inherit;
  font: inherit;
}

.width-input {
  width: 5.5em;
}

.colour-input {
  width: 40px;
  height: 26px;
  padding: 0;
  border: 1px solid var(--border-subtle, #dadce0);
  border-radius: 4px;
  background: var(--surface-1, #fff);
}

.padding-group,
.shading-group {
  margin-top: 12px;
  padding: 8px 10px;
  border: 1px solid var(--border-subtle, #dadce0);
  border-radius: 4px;
}

.padding-group {
  display: flex;
  gap: 10px;
}

.field {
  display: flex;
  flex-direction: column;
  gap: 3px;
  flex: 1;
}

.field-label {
  font-size: 12px;
  text-transform: capitalize;
  color: var(--text-secondary, #5f6368);
}

.shading-group {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
}

.shading-toggle {
  display: flex;
  align-items: center;
  gap: 6px;
  font-size: 13px;
}

.hint {
  margin: 0;
  flex-basis: 100%;
  font-size: 12px;
  color: var(--text-secondary, #5f6368);
}
</style>

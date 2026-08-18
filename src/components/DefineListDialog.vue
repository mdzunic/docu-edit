<script setup lang="ts">
// "Define new multilevel list…" (SPEC-032 AC-5..12): per-level editing of
// the nine levels' number style, prefix/suffix text, levels to include,
// start value, alignment, number position and text position. OK emits the
// full nine-level array; the command normalises and applies it. BaseDialog
// owns the shell and Escape — no field is marked [autofocus] (this dialog
// never focused one), so the panel itself takes focus and Escape works
// immediately. Runtime prop/emit declarations — this script is parsed by
// espree.
import { computed, ref } from "vue";
import { definedMarkerText } from "../lib/core/listDefinitions";
import { asListLevelDefinitions } from "../lib/core/lists";
import BaseDialog from "./BaseDialog.vue";

const props = defineProps({
  /** The nine ListLevelDefinition entries to prefill from (see
   * commands.selectionListDefinitionLevels). */
  initial: { type: Array, required: true },
});

const emit = defineEmits(["apply", "cancel"]);

// The eight marker styles the per-level select offers (the one control
// covers both the number format and bullet glyphs — a bullet level IS the
// "bullet" number format).
const STYLE_OPTIONS = [
  { key: "decimal", label: "1, 2, 3 …" },
  { key: "lowerAlpha", label: "a, b, c …" },
  { key: "upperAlpha", label: "A, B, C …" },
  { key: "lowerRoman", label: "i, ii, iii …" },
  { key: "upperRoman", label: "I, II, III …" },
  { key: "disc", label: "• Bullet" },
  { key: "circle", label: "◦ Bullet" },
  { key: "square", label: "▪ Bullet" },
];

const BULLET_KEYS = ["disc", "circle", "square"];

/** Editable drafts for all nine levels, every field materialised so the
 * form can bind directly; blank/default values are elided again on apply
 * (and normalised once more by the command). */
function draftsFrom(initial = asListLevelDefinitions([])) {
  const drafts = [];
  for (let level = 1; level <= 9; level++) {
    const source = initial[level - 1] ?? { style: "decimal" };
    drafts.push({
      style: source.style ?? "decimal",
      prefix: source.prefix ?? "",
      suffix: source.suffix ?? "",
      includeLevels: String(source.includeLevels ?? 1),
      startAt: String(source.startAt ?? 1),
      alignment: source.alignment ?? "left",
      // Stored positions sit on the twip grid (1.4993… for an entered 1.5);
      // two decimals is the user's own value again, and re-applying it
      // re-snaps to the identical stored value.
      numberPositionCm: String(
        Math.round((source.numberIndentCm ?? level * 1.27 - 0.635) * 100) / 100,
      ),
      textPositionCm: String(Math.round((source.textIndentCm ?? level * 1.27) * 100) / 100),
    });
  }
  return drafts;
}

const levels = ref(draftsFrom(asListLevelDefinitions(props.initial)));
const selectedLevel = ref(1);

const current = computed(() => levels.value[selectedLevel.value - 1]);

const isBullet = computed(() => BULLET_KEYS.includes(current.value.style));

/** 1..selectedLevel — the "levels to include" choices (AC-8). */
const includeOptions = computed(() => {
  const options = [];
  for (let include = 1; include <= selectedLevel.value; include++) {
    options.push(include);
  }
  return options;
});

function levelDefinitionsFromDrafts() {
  return levels.value.map((draft, index) => {
    const level = index + 1;
    const include = Number(draft.includeLevels);
    const startAt = Number(draft.startAt);
    const numberPosition = Number(draft.numberPositionCm);
    const textPosition = Number(draft.textPositionCm);
    return {
      style: draft.style,
      ...(draft.prefix !== "" ? { prefix: draft.prefix } : {}),
      ...(draft.suffix !== "" ? { suffix: draft.suffix } : {}),
      ...(Number.isInteger(include) && include > 1
        ? { includeLevels: Math.min(include, level) }
        : {}),
      ...(Number.isInteger(startAt) && startAt >= 0 && startAt !== 1 ? { startAt } : {}),
      ...(draft.alignment !== "left" ? { alignment: draft.alignment } : {}),
      ...(Number.isFinite(numberPosition) && numberPosition >= 0
        ? { numberIndentCm: numberPosition }
        : {}),
      ...(Number.isFinite(textPosition) && textPosition >= 0
        ? { textIndentCm: textPosition }
        : {}),
    };
  });
}

/** A live sample of the selected level's marker, from each level's own
 * start value. */
const preview = computed(() => {
  const defs = levelDefinitionsFromDrafts();
  const ordinals = defs.map((def) => def.startAt ?? 1);
  return definedMarkerText(
    { defId: "preview", levels: asListLevelDefinitions(defs) },
    selectedLevel.value,
    ordinals,
  );
});

const error = computed(() => {
  const draft = current.value;
  const startAt = Number(draft.startAt);
  if (!Number.isInteger(startAt) || startAt < 0) {
    return "Start value must be a whole number of at least 0";
  }
  for (const raw of [draft.numberPositionCm, draft.textPositionCm]) {
    const parsed = Number(raw);
    if (!Number.isFinite(parsed) || parsed < 0) {
      return "Positions must be non-negative centimetre values";
    }
  }
  return "";
});

function apply() {
  if (error.value) {
    return;
  }
  emit("apply", levelDefinitionsFromDrafts());
}
</script>

<template>
  <BaseDialog
    title="Define new multilevel list"
    testid="define-list-dialog"
    width="26rem"
    @dismiss="emit('cancel')"
  >
    <div class="field-row">
      <label class="field">
        <span class="field-label">Level</span>
        <select
          v-model.number="selectedLevel"
          class="field-input"
          data-testid="define-list-level"
        >
          <option
            v-for="level in 9"
            :key="level"
            :value="level"
          >{{ level }}</option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">Number style</span>
        <select
          v-model="current.style"
          class="field-input"
          data-testid="define-list-style"
        >
          <option
            v-for="option in STYLE_OPTIONS"
            :key="option.key"
            :value="option.key"
          >{{ option.label }}</option>
        </select>
      </label>
    </div>
    <div class="field-row">
      <label class="field">
        <span class="field-label">Prefix</span>
        <input
          v-model="current.prefix"
          class="field-input"
          data-testid="define-list-prefix"
          type="text"
        >
      </label>
      <label class="field">
        <span class="field-label">Suffix</span>
        <input
          v-model="current.suffix"
          class="field-input"
          data-testid="define-list-suffix"
          type="text"
        >
      </label>
    </div>
    <div class="field-row">
      <label class="field">
        <span class="field-label">Levels to include</span>
        <select
          v-model="current.includeLevels"
          class="field-input"
          data-testid="define-list-include"
          :disabled="isBullet"
        >
          <option
            v-for="include in includeOptions"
            :key="include"
            :value="String(include)"
          >{{ include }}</option>
        </select>
      </label>
      <label class="field">
        <span class="field-label">Start at</span>
        <input
          v-model="current.startAt"
          class="field-input"
          data-testid="define-list-start"
          type="number"
          min="0"
          step="1"
          :disabled="isBullet"
          @keydown.enter="apply"
        >
      </label>
      <label class="field">
        <span class="field-label">Alignment</span>
        <select
          v-model="current.alignment"
          class="field-input"
          data-testid="define-list-alignment"
        >
          <option value="left">Left</option>
          <option value="center">Centre</option>
          <option value="right">Right</option>
        </select>
      </label>
    </div>
    <div class="field-row">
      <label class="field">
        <span class="field-label">Number position (cm)</span>
        <input
          v-model="current.numberPositionCm"
          class="field-input"
          data-testid="define-list-number-position"
          type="number"
          min="0"
          step="0.25"
          @keydown.enter="apply"
        >
      </label>
      <label class="field">
        <span class="field-label">Text position (cm)</span>
        <input
          v-model="current.textPositionCm"
          class="field-input"
          data-testid="define-list-text-position"
          type="number"
          min="0"
          step="0.25"
          @keydown.enter="apply"
        >
      </label>
    </div>
    <p class="define-list-preview">
      Preview:
      <span data-testid="define-list-preview">{{ preview }}</span>
    </p>
    <p
      v-if="error"
      class="dialog-error"
      data-testid="define-list-error"
    >
      {{ error }}
    </p>
    <template #actions>
      <button
        type="button"
        data-testid="define-list-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </button>
      <button
        type="button"
        class="primary"
        data-testid="define-list-apply"
        :disabled="!!error"
        @click="apply"
      >
        OK
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.field-row {
  display: flex;
  gap: var(--space-4);
  margin-bottom: var(--space-4);
}

.field {
  flex: 1;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.field-label {
  font-size: var(--text-md);
  color: var(--text-primary);
}

.field-input {
  font: inherit;
  box-sizing: border-box;
  width: 100%;
  padding: var(--space-3) var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
}

.define-list-preview {
  margin: var(--space-3) 0 0;
  font-size: var(--text-md);
  color: var(--text-primary);
}

.dialog-error {
  margin: var(--space-3) 0 0;
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

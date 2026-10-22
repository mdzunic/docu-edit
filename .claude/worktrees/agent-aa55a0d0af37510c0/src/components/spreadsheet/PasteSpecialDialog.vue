<script setup lang="ts">
// SPEC-125's Paste special (`sheet-paste-special-dialog`), opened by
// Ctrl/Cmd+Shift+V, Ctrl+Alt+V (Control+Option+V on macOS), Edit ▸ Paste
// special… or the cell menu once the clipboard was read: what to paste — All,
// Formulas, Values, Formats, Values and number formats
// (`sheet-paste-special-all|formulas|values|formats|values-number-formats`) —
// and Skip blanks and Transpose (`sheet-paste-special-skip-blanks`,
// `-transpose`). Options the clipboard's source cannot supply (paste.ts
// pasteCapabilities) are disabled. OK emits confirm({ what, skipBlanks,
// transpose }); useSheetClipboard pastes. BaseDialog's dismissal is a
// cancel. Runtime prop declarations per the espree constraint (see
// AppShell.vue).
import { ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** PasteCapabilities: { formulas, formats, numberFormats, transpose }. */
  capabilities: { type: Object, default: () => ({ formulas: true, formats: true, numberFormats: true, transpose: true }) },
});

const emit = defineEmits(["confirm", "cancel"]);

const OPTIONS = [
  { what: "all", label: "All", needs: "" },
  { what: "formulas", label: "Formulas", needs: "formulas" },
  { what: "values", label: "Values", needs: "" },
  { what: "formats", label: "Formats", needs: "formats" },
  { what: "values-number-formats", label: "Values and number formats", needs: "numberFormats" },
];

const what = ref("all");
const skipBlanks = ref(false);
const transpose = ref(false);

function enabled(needs = "") {
  return needs === "" || props.capabilities[needs] === true;
}

function confirm() {
  emit("confirm", { what: what.value, skipBlanks: skipBlanks.value, transpose: transpose.value && props.capabilities.transpose === true });
}
</script>

<template>
  <BaseDialog
    title="Paste special"
    testid="sheet-paste-special-dialog"
    width="22rem"
    @dismiss="emit('cancel')"
  >
    <fieldset class="paste-what">
      <legend>Paste</legend>
      <label
        v-for="option in OPTIONS"
        :key="option.what"
        :class="{ 'is-disabled': !enabled(option.needs) }"
      >
        <input
          v-model="what"
          type="radio"
          name="sheet-paste-what"
          :value="option.what"
          :disabled="!enabled(option.needs)"
          :data-testid="`sheet-paste-special-${option.what}`"
        >
        {{ option.label }}
      </label>
    </fieldset>
    <div class="paste-flags">
      <label>
        <input
          v-model="skipBlanks"
          type="checkbox"
          data-testid="sheet-paste-special-skip-blanks"
        >
        Skip blanks
      </label>
      <label :class="{ 'is-disabled': props.capabilities.transpose !== true }">
        <input
          v-model="transpose"
          type="checkbox"
          :disabled="props.capabilities.transpose !== true"
          data-testid="sheet-paste-special-transpose"
        >
        Transpose
      </label>
    </div>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-paste-special-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-paste-special-ok"
        autofocus
        @click="confirm"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.paste-what {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  margin: 0;
  padding: 0;
  border: 0;
}

legend {
  margin-bottom: var(--space-1);
  font-weight: var(--weight-medium);
  color: var(--text-primary);
}

.paste-flags {
  display: flex;
  gap: var(--space-4);
  margin-top: var(--space-3);
}

label {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-primary);
}

label.is-disabled {
  color: var(--text-disabled, var(--text-secondary));
}
</style>

<script setup lang="ts">
// A sheet drawing's "Alt Text…" (SPEC-131): BaseDialog's modal
// (`sheet-drawing-alt-text-dialog`, titled "Alt text", Escape or the backdrop
// cancelling) with the drawing's description in a text area
// (`sheet-drawing-alt-text-input`, focused, opening with the current alt
// text) — what screen readers announce for it and what is written as
// `cNvPr@descr` / `svg:desc`. OK (`sheet-drawing-alt-text-ok`) or Ctrl/Cmd+Enter
// emits confirm(text) — an empty field removes the alt text — and Cancel
// (`sheet-drawing-alt-text-cancel`) emits cancel; useSheetDrawings writes it
// as one "Alt text" step. Runtime prop declarations per the espree constraint
// (see AppShell.vue).
import { ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The drawing's alt text as it stands ("" for none). */
  initial: { type: String, default: "" },
  /** The drawing's name, said in the field's description. */
  name: { type: String, default: "" },
});

const emit = defineEmits(["confirm", "cancel"]);

const text = ref(props.initial);

function confirm() {
  emit("confirm", text.value);
}

/** Enter makes a new line in the description; Ctrl/Cmd+Enter is OK. */
function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    confirm();
  }
}
</script>

<template>
  <BaseDialog
    title="Alt text"
    testid="sheet-drawing-alt-text-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <form
      class="alt-text-form"
      @submit.prevent="confirm"
    >
      <label class="alt-text-field">
        <span>Description</span>
        <textarea
          v-model="text"
          class="alt-text-input"
          rows="4"
          autofocus
          aria-describedby="sheet-drawing-alt-text-hint"
          data-testid="sheet-drawing-alt-text-input"
          @keydown="onKeydown"
        />
      </label>
      <p
        id="sheet-drawing-alt-text-hint"
        class="alt-text-hint"
      >
        {{ props.name ? `Describe ${props.name} for people who cannot see it. Leave it empty to remove the alt text.` : "Leave it empty to remove the alt text." }}
      </p>
    </form>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="sheet-drawing-alt-text-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="sheet-drawing-alt-text-ok"
        @click="confirm"
      >
        OK
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.alt-text-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.alt-text-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-primary);
  font-size: var(--text-md);
}

.alt-text-input {
  padding: var(--space-1) var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  resize: vertical;
}

.alt-text-hint {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}
</style>

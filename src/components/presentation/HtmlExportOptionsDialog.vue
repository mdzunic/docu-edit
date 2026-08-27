<script setup lang="ts">
// The pre-export options dialog for "Export to HTML" (SPEC-088): a small
// BaseDialog with an "Include speaker notes" checkbox (default off); confirm
// emits the chosen options and the parent runs the export. Runtime prop/emit
// declarations per the espree constraint (see AppShell.vue).
import { ref } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";

const emit = defineEmits(["confirm", "cancel"]);

const includeNotes = ref(false);

function confirm() {
  emit("confirm", { includeNotes: includeNotes.value });
}
</script>

<template>
  <BaseDialog
    title="Export to HTML"
    testid="export-html-dialog"
    width="24rem"
    @dismiss="emit('cancel')"
  >
    <p class="intro">
      Save this deck as a single self-contained <code>.html</code> file that
      plays offline in any browser.
    </p>
    <label class="checkbox-row">
      <input
        v-model="includeNotes"
        type="checkbox"
        data-testid="export-html-notes"
      >
      <span>Include speaker notes</span>
    </label>
    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="export-html-cancel"
        @click="emit('cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="export-html-confirm"
        autofocus
        @click="confirm"
      >
        Export
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.intro {
  margin: 0 0 var(--space-3);
  color: var(--text-secondary);
}

.checkbox-row {
  display: flex;
  gap: var(--space-2);
  align-items: center;
  font-size: var(--text-lg);
}
</style>

<script setup lang="ts">
// First-run empty state (SPEC-067 AC-51..53): what the product is, the
// three ways to start, and the local-storage note. Part of the library
// module (src/views/library/) — HomeView wires the actions.
import UiButton from "../../components/ui/UiButton.vue";

const emit = defineEmits(["new-document", "new-from-template", "open-file"]);
</script>

<template>
  <div
    class="empty-state"
    data-testid="empty-state"
  >
    <h2 class="empty-title">
      Welcome to docu-edit
    </h2>
    <p class="empty-lead">
      A private word processor that opens, edits and saves Word, OpenDocument
      and Markdown files — and previews PowerPoint and OpenDocument
      presentations — right here in your browser.
    </p>
    <div class="empty-actions">
      <UiButton
        variant="primary"
        data-testid="new-document"
        @click="emit('new-document')"
      >
        New document
      </UiButton>
      <UiButton
        data-testid="new-from-template"
        @click="emit('new-from-template')"
      >
        New from template…
      </UiButton>
      <UiButton
        data-testid="open-document"
        @click="emit('open-file')"
      >
        Open a file
      </UiButton>
    </div>
    <p class="empty-note">
      Your documents are stored in this browser — nothing is uploaded.
    </p>
  </div>
</template>

<style scoped>
.empty-state {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: var(--space-5);
  color: var(--text-primary);
  text-align: center;
  padding: var(--space-6);
}

.empty-title {
  margin: 0;
  font-size: var(--text-2xl, var(--text-xl));
  font-weight: var(--weight-bold);
}

.empty-lead {
  margin: 0;
  max-width: 32rem;
  color: var(--text-secondary);
  font-size: var(--text-lg);
}

.empty-actions {
  display: flex;
  gap: var(--space-4);
  flex-wrap: wrap;
  justify-content: center;
}

.empty-note {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}
</style>

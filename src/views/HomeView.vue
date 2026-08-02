<script setup lang="ts">
import { useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { getOrCreateDocument } from "../lib/documentStore";

const router = useRouter();

function createNewDocument() {
  const doc = getOrCreateDocument(crypto.randomUUID());
  router.push(`/doc/${doc.id}`);
}
</script>

<template>
  <AppShell document-title="No document">
    <div
      class="empty-state"
      data-testid="empty-state"
    >
      <p class="empty-message">
        No document open
      </p>
      <div class="empty-actions">
        <!-- "Open document" is intentionally inert until SPEC-004/005 wire it
             up (open from disk). -->
        <button
          type="button"
          data-testid="new-document"
          @click="createNewDocument"
        >
          New document
        </button>
        <button
          type="button"
          data-testid="open-document"
        >
          Open document
        </button>
      </div>
    </div>
  </AppShell>
</template>

<style scoped>
.empty-state {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 1rem;
  color: #3c4043;
}

.empty-message {
  margin: 0;
  font-size: 1.125rem;
}

.empty-actions {
  display: flex;
  gap: 0.75rem;
}

.empty-actions button {
  font: inherit;
  padding: 0.5rem 1.25rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
</style>

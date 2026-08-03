<script setup lang="ts">
import { ref, useTemplateRef } from "vue";
import { useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { getOrCreateDocument, storeDocument } from "../lib/documentStore";
import { openDocxFile } from "../lib/openDocument";

const router = useRouter();
const fileInput = useTemplateRef("fileInput");
const openError = ref("");

function createNewDocument() {
  const doc = getOrCreateDocument(crypto.randomUUID());
  router.push(`/doc/${doc.id}`);
}

function pickFile() {
  openError.value = "";
  fileInput.value?.click();
}

async function handleFileChosen() {
  const file = fileInput.value?.files?.[0];
  if (fileInput.value) {
    fileInput.value.value = "";
  }
  if (!file) {
    return;
  }
  try {
    const doc = await openDocxFile(file);
    openError.value = "";
    storeDocument(doc);
    router.push(`/doc/${doc.id}`);
  } catch (error) {
    openError.value = error instanceof Error ? error.message : String(error);
  }
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
          @click="pickFile"
        >
          Open document
        </button>
        <input
          ref="fileInput"
          type="file"
          accept=".docx"
          data-testid="open-document-input"
          class="visually-hidden"
          @change="handleFileChosen"
        >
      </div>
      <p
        v-if="openError"
        class="open-error"
        role="alert"
        data-testid="open-document-error"
      >
        {{ openError }}
      </p>
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

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: -1px;
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}

.open-error {
  margin: 0;
  color: #b3261e;
  max-width: 32rem;
  text-align: center;
}
</style>

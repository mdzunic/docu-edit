<script setup lang="ts">
import { onMounted, ref, useTemplateRef } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { openDocumentFile } from "../lib/openDocument";
import { openDocumentFromUrl } from "../lib/openDocumentFromUrl";
import { storeDocument } from "../lib/documentStore";

const route = useRoute();
const router = useRouter();
const fileInput = useTemplateRef("fileInput");

// "loading" while the fetch is in flight, "error" once it's failed —
// success replaces the route with /doc/:id, so there's no "done" state to
// render here (SPEC-009 AC-9, AC-10, AC-16).
const status = ref("loading");
const errorReason = ref("");

function urlParam() {
  const value = route.query.url;
  return typeof value === "string" ? value : "";
}

async function run() {
  const url = urlParam();
  if (!url) {
    status.value = "error";
    errorReason.value = "No URL was given to open.";
    return;
  }
  status.value = "loading";
  errorReason.value = "";
  try {
    const doc = await openDocumentFromUrl(url);
    storeDocument(doc);
    // Replace, not push: reloading /doc/:id afterwards must not re-fetch
    // the URL (AC-16), and the failed /open?url=... entry shouldn't sit in
    // history either.
    router.replace(`/doc/${doc.id}`);
  } catch (error) {
    status.value = "error";
    errorReason.value = error instanceof Error ? error.message : String(error);
  }
}

onMounted(run);

function pickLocalFile() {
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
    const doc = await openDocumentFile(file);
    storeDocument(doc);
    router.replace(`/doc/${doc.id}`);
  } catch (error) {
    status.value = "error";
    errorReason.value = error instanceof Error ? error.message : String(error);
  }
}
</script>

<template>
  <AppShell document-title="No document">
    <div
      v-if="status === 'loading'"
      class="status-panel"
      data-testid="open-from-url-loading"
      role="status"
    >
      <p>Opening document…</p>
    </div>
    <div
      v-else
      class="status-panel"
      data-testid="open-from-url-error"
      role="alert"
    >
      <p class="error-title">
        Couldn't open that link
      </p>
      <p
        class="error-reason"
        data-testid="open-from-url-reason"
      >
        {{ errorReason }}
      </p>
      <button
        type="button"
        data-testid="open-from-url-pick-file"
        @click="pickLocalFile"
      >
        Choose a file instead
      </button>
    </div>
    <input
      ref="fileInput"
      type="file"
      accept=".docx,.odt,.md,.markdown"
      data-testid="open-from-url-input"
      class="visually-hidden"
      @change="handleFileChosen"
    >
  </AppShell>
</template>

<style scoped>
.status-panel {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.75rem;
  color: #3c4043;
  text-align: center;
  padding: 1rem;
}

.error-title {
  margin: 0;
  font-size: 1.125rem;
  color: #202124;
}

.error-reason {
  margin: 0;
  color: #b3261e;
  max-width: 32rem;
}

.status-panel button {
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
</style>

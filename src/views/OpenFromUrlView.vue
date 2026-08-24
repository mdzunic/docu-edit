<script setup lang="ts">
import { onMounted, ref, useTemplateRef } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { useTheme } from "../composables/useTheme";
import { openAnyDocumentFile } from "../lib/openDocument";
import { openDocumentFromUrl } from "../lib/openDocumentFromUrl";
import { storeDocument, storePresentation } from "../lib/documentStore";
import { editorPathFor } from "../lib/documentKinds";

const route = useRoute();
const router = useRouter();

// No menu bar here either — read the shared composable so this view is themed
// by the same tokens and follows the OS live (SPEC-059 AC-53).
useTheme();
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
    const opened = await openDocumentFromUrl(url);
    if (opened.kind === "presentation") {
      storePresentation(opened.model);
    } else {
      storeDocument(opened.model);
    }
    // Replace, not push: reloading /doc/:id afterwards must not re-fetch
    // the URL (AC-16), and the failed /open?url=... entry shouldn't sit in
    // history either.
    router.replace(editorPathFor(opened.kind, opened.model.id));
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
    const opened = await openAnyDocumentFile(file);
    if (opened.kind === "presentation") {
      storePresentation(opened.model);
    } else {
      storeDocument(opened.model);
    }
    router.replace(editorPathFor(opened.kind, opened.model.id));
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
  gap: var(--space-4);
  color: var(--text-primary);
  text-align: center;
  padding: var(--space-5);
}

.error-title {
  margin: 0;
  font-size: var(--text-xl);
  color: var(--text-primary);
}

.error-reason {
  margin: 0;
  color: var(--text-danger);
  max-width: 32rem;
}

.status-panel button {
  font: inherit;
  padding: var(--space-3) var(--space-6);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>

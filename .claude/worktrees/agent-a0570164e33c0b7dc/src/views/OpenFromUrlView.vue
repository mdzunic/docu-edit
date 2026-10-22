<script setup lang="ts">
import { computed, ref, useTemplateRef } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { useTheme } from "../composables/useTheme";
import { openAnyDocumentFile } from "../lib/openDocument";
import { openDocumentFromUrl } from "../lib/openDocumentFromUrl";
import { storeOpenedFile } from "../lib/documentStore";
import { useToasts } from "../composables/useToasts";
import { editorPathFor, OPEN_FILE_ACCEPT } from "../lib/documentKinds";

const route = useRoute();
const router = useRouter();
const { addToast } = useToasts();

// No menu bar here either — read the shared composable so this view is themed
// by the same tokens and follows the OS live (SPEC-059 AC-53).
useTheme();
const fileInput = useTemplateRef("fileInput");

function urlParam() {
  const value = route.query.url;
  return typeof value === "string" ? value : "";
}

// "confirm" until the reader asks for the document, "loading" while the
// fetch is in flight, "error" once it's failed — success replaces the route
// with /doc/:id, so there's no "done" state to render here (SPEC-009 AC-9,
// AC-10, AC-16).
//
// The fetch waits for a click rather than firing from onMounted: `url` is
// whatever a link put in the query string, so an auto-fetch made any link a
// way to send a request from the reader's browser to a host of the sender's
// choosing and file the response in the reader's library, with nothing shown
// and nothing to decline. Showing the address first is what makes it theirs
// to refuse (AC-9).
const status = ref(urlParam() === "" ? "error" : "confirm");
const errorReason = ref(urlParam() === "" ? "No URL was given to open." : "");

const targetUrl = computed(() => urlParam());

/** The host the link points at, shown on its own line so it survives a long
 * or padded path — the part of the address that says who is being asked. An
 * unparseable URL has no host to show; `run` rejects it on click. */
const targetHost = computed(() => {
  try {
    return new URL(targetUrl.value).host;
  } catch {
    return "";
  }
});

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
    storeOpenedFile(opened, addToast);
    // Replace, not push: reloading /doc/:id afterwards must not re-fetch
    // the URL (AC-16), and the failed /open?url=... entry shouldn't sit in
    // history either.
    router.replace(editorPathFor(opened.kind, opened.model.id));
  } catch (error) {
    status.value = "error";
    errorReason.value = error instanceof Error ? error.message : String(error);
  }
}

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
    storeOpenedFile(opened, addToast);
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
      v-if="status === 'confirm'"
      class="status-panel"
      data-testid="open-from-url-confirm"
    >
      <p class="confirm-title">
        Open this document?
      </p>
      <p class="confirm-host">
        <span
          v-if="targetHost"
          data-testid="open-from-url-host"
        >{{ targetHost }}</span>
      </p>
      <p
        class="confirm-url"
        data-testid="open-from-url-target"
      >
        {{ targetUrl }}
      </p>
      <div class="confirm-actions">
        <button
          type="button"
          class="primary"
          data-testid="open-from-url-confirm-button"
          @click="run"
        >
          Open document
        </button>
        <button
          type="button"
          data-testid="open-from-url-pick-file-instead"
          @click="pickLocalFile"
        >
          Choose a file instead
        </button>
      </div>
    </div>
    <div
      v-else-if="status === 'loading'"
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
      tabindex="-1"
      aria-hidden="true"
      :accept="OPEN_FILE_ACCEPT"
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

.confirm-title {
  margin: 0;
  font-size: var(--text-xl);
  color: var(--text-primary);
}

.confirm-host {
  margin: 0;
  font-size: var(--text-lg);
  color: var(--text-primary);
}

/* The full address, wrapped rather than truncated: a link that pads the
   path to push the real host out of view is exactly what this panel is
   here to expose. */
.confirm-url {
  margin: 0;
  max-width: 32rem;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  overflow-wrap: anywhere;
}

.confirm-actions {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
  justify-content: center;
}

.status-panel button {
  font: inherit;
  padding: var(--space-3) var(--space-6);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

/* The repo's primary-button treatment (BookmarkDialog, ChartOptionsDialog,
   …) so the confirming action reads the same here as everywhere else. */
.status-panel button.primary {
  background: var(--surface-brand);
  border-color: var(--surface-brand);
  color: var(--text-on-brand);
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

<script setup lang="ts">
import { computed, nextTick, onMounted, onUnmounted, ref, useTemplateRef } from "vue";
import { useRouter } from "vue-router";
import ShortcutsDialog from "./ShortcutsDialog.vue";
import StorageWarningBanner from "./StorageWarningBanner.vue";
import ToastHost from "./ToastHost.vue";
import { useToasts } from "../composables/useToasts";
import { storeDocument } from "../lib/documentStore";
import {
  hasSupportedExtension,
  openDocumentFile,
  UNSUPPORTED_FILE_TYPE_MESSAGE,
} from "../lib/openDocument";
import { findShortcut, isOutsideDocumentInput } from "../lib/shortcuts";

// Runtime prop/emit declarations (not the defineProps<{...}>()/defineEmits<{...}>()
// generics) because ESLint parses .vue scripts with espree, which rejects
// TS-only syntax (generics, type annotations, casts); vue-tsc still infers
// the types from them. useTemplateRef needs no generic for the same reason —
// it infers the element type from the matching template `ref` attribute.
const props = defineProps({
  /** Shown in the app bar next to the product name. */
  documentTitle: { type: String, required: true },
  /** Whether clicking the title enters an editable text field. */
  titleEditable: { type: Boolean, default: false },
  /** Shown as a badge next to the title, e.g. "DOCX". Omitted when unset. */
  formatBadge: { type: String, default: "" },
});

const emit = defineEmits(["update:documentTitle"]);

const isEditingTitle = ref(false);
const draftTitle = ref("");
const titleInput = useTemplateRef("titleInput");

async function startEditingTitle() {
  if (!props.titleEditable) {
    return;
  }
  draftTitle.value = props.documentTitle;
  isEditingTitle.value = true;
  await nextTick();
  titleInput.value?.focus();
  titleInput.value?.select();
}

function commitTitleEdit() {
  if (!isEditingTitle.value) {
    return;
  }
  isEditingTitle.value = false;
  emit("update:documentTitle", draftTitle.value);
}

function cancelTitleEdit() {
  isEditingTitle.value = false;
}

// Window-level drag-and-drop (SPEC-009): the shell is the app's whole
// window, so it owns the drop target regardless of which view (home or an
// open document) is mounted. dragenter/dragleave fire once per element
// boundary crossed, including into and out of every child, so a depth
// counter — not a boolean — is what makes the overlay track "is a file
// being dragged anywhere over the window" correctly.
const router = useRouter();
const { addToast } = useToasts();
const dragDepth = ref(0);
const showDropOverlay = computed(() => dragDepth.value > 0);
const dropError = ref("");

// Default values (rather than `event: DragEvent` annotations) are how this
// file infers parameter types elsewhere too — ESLint parses .vue scripts
// with plain espree, which rejects TS-only syntax; the default expression
// is only ever evaluated if a caller omits the argument, which none here
// do, so it never actually runs.
function isFileDrag(event = new DragEvent("dragenter")) {
  const types = event.dataTransfer?.types;
  return !!types && types.includes("Files");
}

function handleDragEnter(event = new DragEvent("dragenter")) {
  if (!isFileDrag(event)) {
    return;
  }
  event.preventDefault();
  dragDepth.value += 1;
}

function handleDragOver(event = new DragEvent("dragover")) {
  if (!isFileDrag(event)) {
    return;
  }
  // Required so the browser fires `drop` instead of rejecting it / opening
  // the file itself.
  event.preventDefault();
}

function handleDragLeave(event = new DragEvent("dragleave")) {
  if (!isFileDrag(event)) {
    return;
  }
  dragDepth.value = Math.max(0, dragDepth.value - 1);
}

// Until this resolves successfully, nothing here touches the currently
// open document (AC-8) — only a rejected/errored drop sets dropError, and
// only a parsed drop calls storeDocument/navigates.
async function handleDrop(event = new DragEvent("drop")) {
  if (!isFileDrag(event)) {
    return;
  }
  event.preventDefault();
  dragDepth.value = 0;
  const files = event.dataTransfer?.files;
  if (!files || files.length === 0) {
    return;
  }
  dropError.value = "";
  const file = files[0];
  if (!hasSupportedExtension(file.name)) {
    dropError.value = UNSUPPORTED_FILE_TYPE_MESSAGE;
    return;
  }
  try {
    const doc = await openDocumentFile(file);
    storeDocument(doc);
    if (files.length > 1) {
      addToast(`Opened ${doc.title} — only one file at a time`);
    }
    router?.push(`/doc/${doc.id}`);
  } catch (error) {
    dropError.value = error instanceof Error ? error.message : String(error);
  }
}

onMounted(() => {
  window.addEventListener("dragenter", handleDragEnter);
  window.addEventListener("dragover", handleDragOver);
  window.addEventListener("dragleave", handleDragLeave);
  window.addEventListener("drop", handleDrop);
});
onUnmounted(() => {
  window.removeEventListener("dragenter", handleDragEnter);
  window.removeEventListener("dragover", handleDragOver);
  window.removeEventListener("dragleave", handleDragLeave);
  window.removeEventListener("drop", handleDrop);
});

// Ctrl/Cmd+/ opens the shortcut reference (SPEC-021 AC-6) from anywhere in
// the shell, not just the document canvas — so it lives here rather than in
// DocumentView, alongside every other view AppShell wraps. Guarded the same
// way every registry-driven shortcut is (AC-12..14): it does nothing while
// focus is inside a dialog (including its own, so Ctrl+/ can't reopen itself
// mid-search), the find bar, or another text input outside the document.
const shortcutsOpen = ref(false);

function handleGlobalKeydown(event = new KeyboardEvent("keydown")) {
  if (event.altKey || isOutsideDocumentInput(event.target)) {
    return;
  }
  if (findShortcut(event)?.id === "view.shortcuts") {
    event.preventDefault();
    shortcutsOpen.value = true;
  }
}

onMounted(() => {
  document.addEventListener("keydown", handleGlobalKeydown);
});
onUnmounted(() => {
  document.removeEventListener("keydown", handleGlobalKeydown);
});
</script>

<template>
  <div class="shell">
    <header
      class="app-bar"
      data-testid="app-bar"
    >
      <span
        class="product-name"
        data-testid="product-name"
      >docu-edit</span>
      <h1
        v-if="!isEditingTitle"
        class="document-title"
        data-testid="document-title"
        :tabindex="titleEditable ? 0 : undefined"
        :role="titleEditable ? 'button' : undefined"
        @click="startEditingTitle"
        @keydown.enter="startEditingTitle"
      >
        {{ documentTitle }}
      </h1>
      <input
        v-else
        ref="titleInput"
        v-model="draftTitle"
        class="document-title document-title-input"
        data-testid="document-title-input"
        @keydown.enter="commitTitleEdit"
        @keydown.escape="cancelTitleEdit"
        @blur="commitTitleEdit"
      >
      <span
        v-if="formatBadge"
        class="format-badge"
        data-testid="format-badge"
      >{{ formatBadge }}</span>
      <div class="app-bar-actions">
        <slot name="app-bar-actions" />
      </div>
    </header>
    <StorageWarningBanner />
    <div
      class="toolbar"
      role="toolbar"
      aria-label="Document toolbar"
      data-testid="toolbar"
    >
      <slot name="toolbar" />
    </div>
    <main
      class="canvas"
      data-testid="canvas"
    >
      <slot />
    </main>
    <ToastHost />
    <ShortcutsDialog
      v-if="shortcutsOpen"
      @close="shortcutsOpen = false"
    />
    <div
      v-if="showDropOverlay"
      class="drop-overlay"
      data-testid="drop-overlay"
    >
      Drop to open
    </div>
    <p
      v-if="dropError"
      class="drop-error"
      role="alert"
      data-testid="drop-error"
    >
      {{ dropError }}
    </p>
  </div>
</template>

<style scoped>
/* The shell owns the viewport: it fills #app exactly and hides its own
   overflow, so the page never grows a scrollbar — only .canvas scrolls. */
.shell {
  height: 100%;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.app-bar {
  flex: none;
  display: flex;
  align-items: baseline;
  gap: 0.75rem;
  padding: 0.5rem 1rem;
  background: #2c3e50;
  color: #fff;
}

.product-name {
  font-weight: 600;
}

.document-title {
  margin: 0;
  font-size: 1rem;
  font-weight: 400;
}

h1.document-title[role="button"] {
  cursor: pointer;
}

.document-title-input {
  font: inherit;
  color: inherit;
  background: rgb(255 255 255 / 0.15);
  border: 1px solid rgb(255 255 255 / 0.5);
  border-radius: 2px;
  padding: 0 0.25rem;
}

/* Pushed to the far edge of the app bar; align-self keeps the actions
   vertically centred even though the bar's items align to baseline. */
.app-bar-actions {
  margin-left: auto;
  align-self: center;
}

.format-badge {
  font-size: 0.6875rem;
  font-weight: 600;
  letter-spacing: 0.03em;
  padding: 0.125rem 0.375rem;
  border-radius: 3px;
  background: rgb(255 255 255 / 0.15);
  color: inherit;
}

/* A single non-wrapping row: contents that overflow the viewport width are
   reached by scrolling the strip, never by wrapping or clipping. */
.toolbar {
  flex: none;
  display: flex;
  flex-wrap: nowrap;
  align-items: center;
  gap: 0.25rem;
  min-height: 2.5rem;
  padding: 0.25rem 1rem;
  overflow-x: auto;
  overflow-y: hidden;
  white-space: nowrap;
  background: #f5f5f5;
  border-bottom: 1px solid #d0d0d0;
  box-sizing: border-box;
}

/* min-height: 0 lets the canvas shrink below its content height, so tall
   content scrolls here instead of stretching the shell past the viewport. */
.canvas {
  flex: 1 1 0;
  min-height: 0;
  overflow-y: auto;
  background: #e8eaed;
}

.drop-overlay {
  position: fixed;
  inset: 0;
  z-index: 200;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 1.5rem;
  font-weight: 600;
  color: #fff;
  background: rgb(26 82 118 / 0.85);
  border: 4px dashed rgb(255 255 255 / 0.8);
  pointer-events: none;
}

.drop-error {
  position: fixed;
  bottom: 1.5rem;
  left: 50%;
  transform: translateX(-50%);
  z-index: 150;
  margin: 0;
  max-width: 32rem;
  padding: 0.625rem 1.25rem;
  border-radius: 4px;
  background: #fdecea;
  color: #b3261e;
  box-shadow: 0 2px 6px rgb(0 0 0 / 0.2);
}
</style>

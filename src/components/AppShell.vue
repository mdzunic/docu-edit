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
import { installTierObserver, useTier } from "../composables/useResponsiveTier";

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

// SPEC-068 — the shell root is the width the three layout tiers key off. A
// ResizeObserver here (not device sniffing) means a narrow desktop window
// behaves identically to a tablet of the same width (AC-1), and the tier
// switches live as the window resizes with no reload (AC-2). The tier is
// module-level shared state (useResponsiveTier) so it also reaches chrome that
// DocumentView slots into this shell, which provide/inject could not.
const shellRoot = useTemplateRef("shellRoot");
const tier = useTier();
let stopTierObserver = () => {};
onMounted(() => {
  stopTierObserver = installTierObserver(() => shellRoot.value);
});
onUnmounted(() => {
  stopTierObserver();
});

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

// The product-name home link (SPEC-060). Client-side navigation so the
// editor's route guard runs its autosave flush before leaving; `router?`
// because AppShell mounts without a router in some unit tests.
function goHome() {
  router?.push("/");
}

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
  <div
    ref="shellRoot"
    class="shell"
    :data-tier="tier"
    data-testid="app-shell"
  >
    <header
      class="app-bar"
      data-testid="app-bar"
    >
      <!-- The leading control slot (SPEC-060): the editor fills it with the
           Close control, left of the product name. Empty in views with no
           open document, so their bar is unchanged. -->
      <slot name="app-bar-leading" />
      <!-- The product name is the app's home affordance (SPEC-060): a real
           link to the library in every view, so it reads as a link to
           assistive technology. In the editor the route guard flushes any
           pending autosave before this navigation lands, exactly as the Close
           control does — @click.prevent keeps it a client-side navigation so
           that guard runs (a full document load would bypass it). -->
      <a
        class="product-name"
        data-testid="product-name"
        href="/"
        @click.prevent="goHome"
      >docu-edit</a>
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
    <!-- The menu bar (SPEC-049) sits between the app bar and the toolbar
         strip. Only rendered when a view supplies one, so the home and
         open-from-URL shells keep their previous layout. -->
    <div
      v-if="$slots.menubar"
      class="menubar-region"
    >
      <slot name="menubar" />
    </div>
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
    <div
      class="status-bar"
      data-testid="status-bar"
    >
      <slot name="status-bar" />
    </div>
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
  gap: var(--space-4);
  /* One stated height (AC-15): the app bar reads the brand surface token. */
  height: var(--height-app-bar);
  padding: var(--space-3) var(--space-5);
  background: var(--surface-brand);
  color: var(--text-on-brand);
  box-sizing: border-box;
}

.product-name {
  font-weight: var(--weight-bold);
  color: inherit;
  text-decoration: none;
  border-radius: var(--radius-sm);
}

.product-name:hover {
  text-decoration: underline;
}

.product-name:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.document-title {
  margin: 0;
  font-size: var(--text-lg);
  font-weight: var(--weight-regular);
}

h1.document-title[role="button"] {
  cursor: pointer;
}

.document-title-input {
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  padding: 0 var(--space-2);
}

/* Pushed to the far edge of the app bar; align-self keeps the actions
   vertically centred even though the bar's items align to baseline. */
.app-bar-actions {
  margin-left: auto;
  align-self: center;
}

.format-badge {
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  letter-spacing: 0.03em;
  padding: var(--space-1) var(--space-2);
  border-radius: var(--radius-sm);
  background: transparent;
  border: var(--border-width) solid var(--text-on-brand);
  color: inherit;
}

/* The menu bar's own row. It is flush against the app bar above and the
   toolbar strip below; the MenuBar component paints its own background and
   border, so this wrapper only reserves the row and never clips the popups
   that hang out of it. */
.menubar-region {
  flex: none;
}

/* A single non-wrapping row: contents that overflow the viewport width are
   reached by scrolling the strip, never by wrapping or clipping. */
.toolbar {
  flex: none;
  display: flex;
  flex-wrap: nowrap;
  align-items: center;
  gap: var(--space-2);
  /* One stated height (AC-49): the strip never sizes to its content. */
  min-height: var(--height-toolbar);
  /* --space-2 inline padding: the full control set plus platform font-width
     variance must fit the 1280px no-overflow budget (SPEC-049 AC-4). */
  padding: var(--space-2);
  overflow-x: auto;
  overflow-y: hidden;
  white-space: nowrap;
  background: var(--surface-raised);
  border-bottom: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
}

/* min-height: 0 lets the canvas shrink below its content height, so tall
   content scrolls here instead of stretching the shell past the viewport. */
.canvas {
  flex: 1 1 0;
  min-height: 0;
  overflow-y: auto;
  background: var(--surface-canvas);
}

/* The status bar (SPEC-028 AC-3): a slim strip under the canvas the view
   fills with the caret's page indicator. Present even when the slot is
   empty so the layout never jumps. */
.status-bar {
  flex: none;
  display: flex;
  align-items: center;
  gap: var(--space-5);
  /* One stated height (AC-51): the status bar never sizes to its content. */
  min-height: var(--height-status-bar);
  padding: var(--space-1) var(--space-5);
  border-top: var(--border-width) solid var(--border-default);
  background: var(--surface-raised);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.drop-overlay {
  position: fixed;
  inset: 0;
  z-index: 200;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-on-brand);
  background: var(--surface-overlay);
  border: calc(4 * var(--border-width)) dashed var(--text-on-brand);
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
  padding: var(--space-3) var(--space-6);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-danger);
  border: var(--border-width) solid var(--danger);
  box-shadow: var(--shadow-2);
}
</style>

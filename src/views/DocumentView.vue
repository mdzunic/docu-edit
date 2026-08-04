<script setup lang="ts">
import { computed, nextTick, onBeforeUnmount, onMounted, ref, shallowRef, triggerRef, useTemplateRef, watch } from "vue";
import { useRoute, useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import DownloadMenu from "../components/DownloadMenu.vue";
import { useAutosave } from "../composables/useAutosave";
import { useToasts } from "../composables/useToasts";
import { documentExists } from "../lib/documents";
import { DOCX_MIME_TYPE, docxFileName, writeDocx } from "../lib/docx/writer";
import { ODT_MIMETYPE, odtFileName, writeOdt } from "../lib/odf/writer";
import {
  deleteDocument,
  getOrCreateDocument,
  touchDocument,
} from "../lib/documentStore";
import { DEFAULT_TITLE } from "../lib/documentModel";
import { acquireEditLock, noEditLock } from "../lib/editLock";
import { modelRangeFromSelection, noElement, setDomCaret } from "../lib/domCaret";
import {
  deleteBackward,
  deleteForward,
  deleteRange,
  insertText,
  noCaret,
  splitParagraph,
} from "../lib/textEditing";
import { headingTag, paragraphStyle, runKey, runStyle } from "../lib/renderFormatting";

const route = useRoute();
const router = useRouter();
const { addToast } = useToasts();
const documentId = computed(() => String(route.params.id ?? ""));

// A SHALLOW ref, deliberately: the model must stay a plain object because
// IndexedDB's structured clone rejects Vue's reactive proxies outright
// (DataCloneError), and autosave persists this exact object. Mutations
// therefore happen on the raw model and re-render via triggerRef — which is
// also what keeps 200-page documents cheap: no proxy per block/run. The
// initial value is computed inline (rather than starting from `null`) so TS
// infers the `DocumentModel | null` type from this expression alone —
// a generic would need TS syntax, which ESLint's espree parser for .vue
// scripts can't parse.
const doc = shallowRef(
  documentExists(documentId.value)
    ? getOrCreateDocument(documentId.value)
    : null,
);

// --- edit lock (SPEC-010: two tabs on the same document) -------------------
// "pending" until the Web Lock resolves; "busy" renders the conflict bar and
// a read-only canvas until the other tab closes and the queued lock grants.
const lockState = ref("pending");
const editable = computed(() => lockState.value === "acquired" && !!doc.value);

// --- caret and scroll (SPEC-010: restored on reopen) -----------------------
// Plain variables, not refs: nothing renders from them, they are only read
// when a draft or commit is written. noCaret() seeds the CaretPosition|null
// type, which espree bars this file from annotating directly.
let caret = noCaret();
let scrollTop = 0;
let scrollContainer = noElement();
let editLock = noEditLock();

const contentRoot = useTemplateRef("contentRoot");

const autosave = useAutosave({
  getCaret: () => caret,
  getScroll: () => scrollTop,
  isActive: () => editable.value,
  applyRestored: (restored) => {
    doc.value = restored.model;
    caret = restored.caret;
    scrollTop = restored.scroll;
    void focusAndRestore();
  },
});
const { saveState, saveStateText, recovery, keepRecovered, discardRecovered } = autosave;

const recoveryTime = computed(() =>
  recovery.value ? new Date(recovery.value.savedAt).toLocaleTimeString() : "",
);

async function focusAndRestore() {
  await nextTick();
  attachScrollListener();
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  root.focus();
  caret = caret ?? { section: 0, block: 0, offset: 0 };
  setDomCaret(root, caret);
  // After focus/caret placement, which may themselves scroll.
  if (scrollContainer) {
    scrollContainer.scrollTop = scrollTop;
  }
}

// --- session lifecycle -----------------------------------------------------

function openSession(id = "") {
  lockState.value = "pending";
  editLock = acquireEditLock(id, (state) => {
    if (documentId.value !== id || !doc.value) {
      return;
    }
    lockState.value = state === "acquired" ? "acquired" : "busy";
    if (state === "acquired") {
      void autosave.beginSession(id);
    }
  });
}

function closeSession() {
  // An in-app close (navigation, unmount) is a clean close: commit, unless
  // this tab never owned the document.
  autosave.endSession(lockState.value === "acquired");
  if (editLock) {
    editLock.release();
    editLock = null;
  }
  lockState.value = "pending";
  caret = noCaret();
  scrollTop = 0;
}

watch(documentId, (id) => {
  closeSession();
  doc.value = documentExists(id) ? getOrCreateDocument(id) : null;
  if (doc.value) {
    openSession(id);
  }
});

function attachScrollListener() {
  const root = contentRoot.value;
  const container = root ? root.closest("main") : null;
  if (container && container !== scrollContainer) {
    if (scrollContainer) {
      scrollContainer.removeEventListener("scroll", handleScroll);
    }
    scrollContainer = container;
    scrollContainer.addEventListener("scroll", handleScroll, { passive: true });
  }
}

function handleScroll() {
  if (scrollContainer) {
    scrollTop = scrollContainer.scrollTop;
  }
}

function handleSelectionChange() {
  const root = contentRoot.value;
  if (!root || !editable.value) {
    return;
  }
  const range = modelRangeFromSelection(root, document.getSelection());
  if (range) {
    caret = range.end;
  }
}

onMounted(() => {
  document.addEventListener("selectionchange", handleSelectionChange);
  attachScrollListener();
  if (doc.value) {
    openSession(documentId.value);
  }
});

onBeforeUnmount(() => {
  closeSession();
  document.removeEventListener("selectionchange", handleSelectionChange);
  if (scrollContainer) {
    scrollContainer.removeEventListener("scroll", handleScroll);
    scrollContainer = null;
  }
});

// --- typing (SPEC-010's minimal editing surface; SPEC-011 will extend it) --
// Fully controlled: every beforeinput is cancelled and the supported ones
// re-applied to the model, which re-renders and re-places the caret. The
// model therefore never drifts from the DOM, and unsupported input types
// are simply inert rather than corrupting.

function handleBeforeInput(event = new InputEvent("beforeinput")) {
  event.preventDefault();
  const root = contentRoot.value;
  const model = doc.value;
  if (!root || !model || !editable.value) {
    return;
  }
  const selection = modelRangeFromSelection(root, document.getSelection());
  const start = selection ? selection.start : caret;
  const end = selection ? selection.end : caret;
  if (!start || !end) {
    return;
  }
  const collapsed =
    start.section === end.section &&
    start.block === end.block &&
    start.offset === end.offset;
  const type = event.inputType;
  let next = null;
  if (type === "insertText" || type === "insertFromPaste") {
    const text =
      event.data ??
      (event.dataTransfer ? event.dataTransfer.getData("text/plain") : "");
    if (text) {
      const base = collapsed ? start : deleteRange(model, start, end);
      next = insertText(model, base, text);
    }
  } else if (type === "insertParagraph") {
    const base = collapsed ? start : deleteRange(model, start, end);
    next = splitParagraph(model, base);
  } else if (type === "insertLineBreak") {
    // Shift+Enter: a forced line break, stored as a literal "\n" that
    // pre-wrap rendering displays without a new paragraph.
    const base = collapsed ? start : deleteRange(model, start, end);
    next = insertText(model, base, "\n");
  } else if (type === "deleteContentBackward") {
    next = collapsed ? deleteBackward(model, start) : deleteRange(model, start, end);
  } else if (type === "deleteContentForward") {
    next = collapsed ? deleteForward(model, start) : deleteRange(model, start, end);
  } else if (type === "deleteByCut" || type === "deleteContent") {
    // The clipboard copy half of a cut happened in the earlier `cut` event;
    // cancelling beforeinput only leaves the deletion to us.
    next = collapsed ? null : deleteRange(model, start, end);
  }
  if (!next) {
    return;
  }
  caret = next;
  triggerRef(doc);
  autosave.markDirty();
  void restoreCaretAfterRender();
}

async function restoreCaretAfterRender() {
  await nextTick();
  const root = contentRoot.value;
  if (root && caret) {
    setDomCaret(root, caret);
  }
}

// ---------------------------------------------------------------------------

const pageStyle = computed(() => {
  if (!doc.value) {
    return {};
  }
  const { widthCm, heightCm } = doc.value.defaultStyles.pageSize;
  return { width: `${widthCm}cm`, height: `${heightCm}cm` };
});

const marginStyle = computed(() => {
  if (!doc.value) {
    return {};
  }
  const { topCm, rightCm, bottomCm, leftCm } = doc.value.defaultStyles.margins;
  return {
    top: `${topCm}cm`,
    right: `${rightCm}cm`,
    bottom: `${bottomCm}cm`,
    left: `${leftCm}cm`,
  };
});

const formatBadge = computed(() => doc.value?.format?.toUpperCase() ?? "");

// Sections flattened to one list so a single v-for can carry v-memo (Vue
// bars v-memo on elements nested inside another v-for). The memo skips
// re-rendering paragraphs whose block object is unchanged — textEditing
// replaces edited blocks — which is what keeps typing responsive on very
// large documents.
const flatBlocks = computed(() => {
  if (!doc.value) {
    return [];
  }
  return doc.value.sections.flatMap((section, sectionIndex) =>
    section.blocks.map((block, blockIndex) => ({
      key: `${sectionIndex}:${blockIndex}`,
      block,
      section: sectionIndex,
      blockIndex,
    })),
  );
});
const downloadFormat = computed(() =>
  doc.value && doc.value.format === "odt" ? "odt" : "docx",
);

function handleTitleUpdate(newTitle = "") {
  if (!doc.value) {
    return;
  }
  const trimmed = newTitle.trim();
  doc.value.title = trimmed.length > 0 ? trimmed : DEFAULT_TITLE;
  triggerRef(doc);
  // The assignment above already mutated the stored model (doc.value IS the
  // raw stored model); touching records the modification time and persists.
  touchDocument(doc.value.id);
}

async function handleDownload(format = "") {
  if (!doc.value || (format !== "docx" && format !== "odt")) {
    return;
  }
  const odt = format === "odt";
  const bytes = odt ? await writeOdt(doc.value) : await writeDocx(doc.value);
  const fileName = odt ? odtFileName(doc.value.title) : docxFileName(doc.value.title);
  // A transient anchor pointed at an object URL is the download mechanism
  // that needs no new browser permissions; revocation is deferred so the
  // browser has started streaming the blob before the URL disappears.
  const url = URL.createObjectURL(new Blob([bytes], { type: odt ? ODT_MIMETYPE : DOCX_MIME_TYPE }));
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  setTimeout(() => URL.revokeObjectURL(url), 10_000);
  // A download is SPEC-003's explicit save: promote the state to committed.
  void autosave.commitNow();
}

const confirmingDelete = ref(false);

function deleteOpenDocument() {
  if (!doc.value) {
    return;
  }
  const { id, title } = doc.value;
  confirmingDelete.value = false;
  deleteDocument(id);
  addToast(`Deleted ${title}`);
  router.push("/");
}
</script>

<template>
  <AppShell
    v-if="doc"
    :document-title="doc.title"
    :title-editable="editable"
    :format-badge="formatBadge"
    @update:document-title="handleTitleUpdate"
  >
    <template #app-bar-actions>
      <div class="app-bar-cluster">
        <span
          class="save-state"
          data-testid="save-state"
          :data-save-state="saveState"
        >{{ saveStateText }}</span>
        <a
          v-if="saveState === 'error'"
          class="save-state-download"
          data-testid="save-state-download"
          href="#"
          @click.prevent="handleDownload(downloadFormat)"
        >Download a copy</a>
        <DownloadMenu @download="handleDownload" />
      </div>
    </template>
    <template #toolbar>
      <button
        v-if="editable"
        type="button"
        class="toolbar-button"
        data-testid="delete-document"
        @click="confirmingDelete = true"
      >
        Delete
      </button>
    </template>
    <div
      v-if="lockState === 'busy'"
      class="notice-bar"
      role="status"
      data-testid="conflict-bar"
    >
      This document is open in another tab
    </div>
    <div
      v-if="recovery"
      class="notice-bar"
      role="status"
      data-testid="recovery-bar"
    >
      <span>Recovered unsaved changes from {{ recoveryTime }}</span>
      <button
        type="button"
        class="notice-button"
        data-testid="recovery-keep"
        @click="keepRecovered"
      >
        Keep
      </button>
      <button
        type="button"
        class="notice-button"
        data-testid="recovery-discard"
        @click="discardRecovered"
      >
        Discard
      </button>
    </div>
    <div class="page-wrap">
      <div
        class="page"
        data-testid="document-page"
        :style="pageStyle"
      >
        <div
          class="page-margin"
          data-testid="page-margin"
          :style="marginStyle"
        >
          <div
            ref="contentRoot"
            class="document-content"
            data-testid="document-content"
            role="textbox"
            aria-multiline="true"
            aria-label="Document content"
            :aria-readonly="editable ? undefined : 'true'"
            :contenteditable="editable"
            spellcheck="false"
            @beforeinput="handleBeforeInput"
          >
            <component
              :is="headingTag(entry.block.headingLevel)"
              v-for="entry in flatBlocks"
              :key="entry.key"
              v-memo="[entry.block, entry.section, entry.blockIndex]"
              class="doc-paragraph"
              :style="paragraphStyle(entry.block)"
              :data-section="entry.section"
              :data-block="entry.blockIndex"
            >
              <template v-if="entry.block.runs.length > 0">
                <span
                  v-for="(run, runIndex) in entry.block.runs"
                  :key="runKey(run, runIndex)"
                  :style="runStyle(run.formatting)"
                >{{ run.text }}</span>
              </template>
              <br v-else>
            </component>
          </div>
        </div>
      </div>
    </div>
    <ConfirmDeleteDialog
      v-if="confirmingDelete"
      :document-title="doc.title"
      @confirm="deleteOpenDocument"
      @cancel="confirmingDelete = false"
    />
  </AppShell>
  <AppShell
    v-else
    document-title="No document"
  >
    <div
      class="not-found"
      data-testid="not-found"
    >
      <p class="not-found-message">
        Document not found
      </p>
      <RouterLink to="/">
        Back to home
      </RouterLink>
    </div>
  </AppShell>
</template>

<style scoped>
.toolbar-button {
  font: inherit;
  font-size: 0.875rem;
  padding: 0.25rem 0.75rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.app-bar-cluster {
  display: flex;
  align-items: center;
  gap: 0.75rem;
}

.save-state {
  font-size: 0.8125rem;
  color: rgb(255 255 255 / 0.85);
  white-space: nowrap;
}

.save-state-download {
  font-size: 0.8125rem;
  color: #fff;
}

/* Sticky so the conflict/recovery notice stays visible while the canvas
   scrolls; they are mutually exclusive, so a single top slot suffices. */
.notice-bar {
  position: sticky;
  top: 0;
  z-index: 10;
  display: flex;
  align-items: center;
  gap: 0.75rem;
  padding: 0.5rem 1rem;
  background: #e8f0fe;
  border-bottom: 1px solid #aecbfa;
  color: #174ea6;
  font-size: 0.875rem;
}

.notice-button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.125rem 0.75rem;
  border: 1px solid #aecbfa;
  border-radius: 4px;
  background: #fff;
  color: #174ea6;
  cursor: pointer;
}

.page-wrap {
  padding: 2rem 1rem;
}

/* A page-shaped surface sized to the document's own defaultStyles (in real
   centimetres via CSS's absolute cm unit) — later specs render content onto
   it. .page-margin marks where the page margins end and content begins. */
.page {
  position: relative;
  margin: 0 auto;
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3);
  box-sizing: border-box;
}

.page-margin {
  position: absolute;
  border: 1px dashed #c8c8c8;
}

.document-content {
  height: 100%;
  overflow: hidden;
  color: #202124;
  outline: none;
}

.doc-paragraph {
  margin: 0;
  font-weight: normal;
  font-size: inherit;
}

.not-found {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.5rem;
  color: #3c4043;
}

.not-found-message {
  margin: 0;
  font-size: 1.125rem;
}
</style>

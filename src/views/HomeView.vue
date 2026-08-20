<script setup lang="ts">
import { useRouter } from "vue-router";
import { useTheme } from "../composables/useTheme";
import AppShell from "../components/AppShell.vue";
import { openDocumentFile } from "../lib/openDocument";
import { computed, nextTick, onMounted, onUnmounted, ref, useTemplateRef } from "vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import RenameDialog from "../components/RenameDialog.vue";
import TemplateGalleryDialog from "../components/TemplateGalleryDialog.vue";
import {
  deleteDocument,
  duplicateDocument,
  getOrCreateDocument,
  initLibrary,
  listDocuments,
  renameDocument,
  storeDocument
} from "../lib/documentStore";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { instantiateTemplate } from "../lib/templateStore";
import { createDocument } from "../lib/documentModel";
import { formatBadgeLabel, formatRelativeTime } from "../lib/format";

const router = useRouter();

// The library has no menu bar, so it reads the shared theme composable to keep
// the OS-appearance listener alive and be themed by the same tokens (SPEC-059
// AC-52). The preference is set from the editor's View menu.
useTheme();
const fileInput = useTemplateRef("fileInput");
const openError = ref("");

// main.ts awaits initLibrary() before mounting, so listDocuments() is already
// hydrated here in the running app; the refresh-on-init matters only when the
// view is mounted directly (tests).
const docs = ref(listDocuments());
const now = ref(Date.now());
function refresh() {
  docs.value = listDocuments();
  now.value = Date.now();
}
initLibrary().then(refresh);

// Relative times go stale while the page sits open; re-render them each minute.
const clockTimer = setInterval(() => {
  now.value = Date.now();
}, 60_000);

// Ids ("" = none) rather than object refs, so state stays valid across
// refreshes and the espree-parsed script needs no type annotations.
const menuFor = ref("");
const renameFor = ref("");
const deleteFor = ref("");

const renameTarget = computed(() =>
  docs.value.find((entry) => entry.id === renameFor.value),
);
const deleteTarget = computed(() =>
  docs.value.find((entry) => entry.id === deleteFor.value),
);

function closeMenu() {
  menuFor.value = "";
}

// Any click that bubbles to the document (i.e. outside the actions cell,
// which stops propagation) closes the open row menu.
onMounted(() => document.addEventListener("click", closeMenu));

// Preload the editor route's chunk while the library sits idle (SPEC-047
// AC-3): every action on this view ends in router.push("/doc/…"), and
// without this the fetch + evaluation of the editor's whole module graph
// lands inside the open path, between choosing a file and its first page.
// Same specifier as router.ts's lazy route, so the router reuses the
// already-loaded module.
onMounted(() => {
  void import("./DocumentView.vue");
});
onUnmounted(() => {
  document.removeEventListener("click", closeMenu);
  clearInterval(clockTimer);
});

// Returning from a closed document (SPEC-060): closeOpenDocument navigates
// here with the closed id in router state (kept out of the URL). Scroll its
// row into view and flash it for one --motion-base interval so a user sees
// where they came from — the flash is suppressed under prefers-reduced-motion.
const highlightId = ref("");

function prefersReducedMotion() {
  return (
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-reduced-motion: reduce)").matches
  );
}

/** The --motion-base design token in milliseconds — the highlight's lifetime,
 * read from the CSS so it tracks the one source of truth (and collapses to 0
 * under reduced motion, where the token itself is 0ms). */
function motionBaseMs() {
  const raw = getComputedStyle(document.documentElement)
    .getPropertyValue("--motion-base")
    .trim();
  const value = parseFloat(raw);
  if (!Number.isFinite(value)) {
    return 200;
  }
  return raw.endsWith("ms") ? value : value * 1000;
}

onMounted(async () => {
  const closedId = window.history.state?.closedDocId;
  if (typeof closedId !== "string" || closedId.length === 0) {
    return;
  }
  await nextTick();
  const row = Array.from(document.querySelectorAll("[data-doc-id]")).find(
    (el) => el.getAttribute("data-doc-id") === closedId,
  );
  if (row && typeof row.scrollIntoView === "function") {
    row.scrollIntoView({ block: "center" });
  }
  if (prefersReducedMotion()) {
    return;
  }
  highlightId.value = closedId;
  setTimeout(() => {
    if (highlightId.value === closedId) {
      highlightId.value = "";
    }
  }, motionBaseMs());
});

const showGallery = ref(false);

function createNewDocument() {
  const doc = getOrCreateDocument(crypto.randomUUID());
  router.push(`/doc/${doc.id}`);
}

function openTemplateGallery() {
  showGallery.value = true;
}

function chooseBlankDocument() {
  showGallery.value = false;
  const doc = getOrCreateDocument(crypto.randomUUID());
  router.push(`/doc/${doc.id}`);
}

function chooseBuiltinTemplate(id = "") {
  showGallery.value = false;
  const template = BUILTIN_TEMPLATES.find((entry) => entry.id === id);
  if (!template) {
    return;
  }
  const doc = template.build();
  storeDocument(doc);
  router.push(`/doc/${doc.id}`);
}

function chooseUserTemplate(id = "") {
  showGallery.value = false;
  const doc = instantiateTemplate(id);
  storeDocument(doc);
  router.push(`/doc/${doc.id}`);
}

function chooseTemplateFile(doc = createDocument("")) {
  showGallery.value = false;
  storeDocument(doc);
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
    const doc = await openDocumentFile(file);
    openError.value = "";
    storeDocument(doc);
    router.push(`/doc/${doc.id}`);
  } catch (error) {
    openError.value = error instanceof Error ? error.message : String(error);
  }
}

function openDocument(id = "") {
  router.push(`/doc/${id}`);
}

function toggleMenu(id = "") {
  menuFor.value = menuFor.value === id ? "" : id;
}

function startRename(id = "") {
  renameFor.value = id;
  closeMenu();
}

function applyRename(newTitle = "") {
  renameDocument(renameFor.value, newTitle);
  renameFor.value = "";
  refresh();
}

function duplicateEntry(id = "") {
  duplicateDocument(id);
  closeMenu();
  refresh();
}

function startDelete(id = "") {
  deleteFor.value = id;
  closeMenu();
}

function confirmDelete() {
  deleteDocument(deleteFor.value);
  deleteFor.value = "";
  refresh();
}
</script>

<template>
  <AppShell document-title="No document">
    <div
      v-if="docs.length === 0"
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
          data-testid="new-from-template"
          @click="openTemplateGallery"
        >
          New from template…
        </button>
        <button
          type="button"
          data-testid="open-document"
          @click="pickFile"
        >
          Open document
        </button>
      </div>
    </div>
    <div
      v-else
      class="library"
      data-testid="library"
    >
      <div class="library-header">
        <h2 class="library-heading">
          Documents
        </h2>
        <div class="library-actions">
          <button
            type="button"
            class="new-button"
            data-testid="new-document"
            @click="createNewDocument"
          >
            New document
          </button>
          <button
            type="button"
            class="new-button"
            data-testid="new-from-template"
            @click="openTemplateGallery"
          >
            New from template…
          </button>
          <button
            type="button"
            class="new-button"
            data-testid="open-document"
            @click="pickFile"
          >
            Open document
          </button>
        </div>
      </div>
      <table
        class="library-table"
        data-testid="library-list"
      >
        <thead>
          <tr>
            <th scope="col">
              Title
            </th>
            <th scope="col">
              Last modified
            </th>
            <th scope="col">
              Format
            </th>
            <th scope="col">
              <span class="visually-hidden">Actions</span>
            </th>
          </tr>
        </thead>
        <tbody>
          <tr
            v-for="entry in docs"
            :key="entry.id"
            class="library-row"
            :class="{ 'row-highlight': entry.id === highlightId }"
            data-testid="library-row"
            :data-doc-id="entry.id"
            @click="openDocument(entry.id)"
          >
            <td
              class="cell-title"
              data-testid="row-title"
            >
              {{ entry.title }}
            </td>
            <td
              class="cell-modified"
              data-testid="row-modified"
            >
              {{ formatRelativeTime(entry.lastModified, now) }}
            </td>
            <td class="cell-format">
              <span
                class="badge"
                data-testid="row-format"
              >{{ formatBadgeLabel(entry.format) }}</span>
            </td>
            <td
              class="cell-actions"
              @click.stop
            >
              <button
                type="button"
                class="menu-button"
                data-testid="row-menu"
                :aria-label="`Actions for ${entry.title}`"
                aria-haspopup="menu"
                :aria-expanded="menuFor === entry.id"
                @click="toggleMenu(entry.id)"
              >
                ⋯
              </button>
              <div
                v-if="menuFor === entry.id"
                class="row-menu"
                role="menu"
                data-testid="row-menu-list"
              >
                <button
                  type="button"
                  role="menuitem"
                  data-testid="menu-rename"
                  @click="startRename(entry.id)"
                >
                  Rename
                </button>
                <button
                  type="button"
                  role="menuitem"
                  data-testid="menu-duplicate"
                  @click="duplicateEntry(entry.id)"
                >
                  Duplicate
                </button>
                <button
                  type="button"
                  role="menuitem"
                  data-testid="menu-delete"
                  @click="startDelete(entry.id)"
                >
                  Delete
                </button>
              </div>
            </td>
          </tr>
        </tbody>
      </table>
    </div>
    <!-- Shared by both branches: the picker must stay reachable once the
         library is non-empty, not only from the empty state. -->
    <input
      ref="fileInput"
      type="file"
      accept=".docx,.odt,.md,.markdown"
      data-testid="open-document-input"
      class="visually-hidden"
      @change="handleFileChosen"
    >
    <p
      v-if="openError"
      class="open-error"
      role="alert"
      data-testid="open-document-error"
    >
      {{ openError }}
    </p>
    <RenameDialog
      v-if="renameTarget"
      :current-title="renameTarget.title"
      @save="applyRename"
      @cancel="renameFor = ''"
    />
    <ConfirmDeleteDialog
      v-if="deleteTarget"
      :document-title="deleteTarget.title"
      @confirm="confirmDelete"
      @cancel="deleteFor = ''"
    />
    <TemplateGalleryDialog
      v-if="showGallery"
      @choose-blank="chooseBlankDocument"
      @choose-builtin="chooseBuiltinTemplate"
      @choose-user="chooseUserTemplate"
      @choose-file="chooseTemplateFile"
      @close="showGallery = false"
    />
  </AppShell>
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
}

.empty-message {
  margin: 0;
  font-size: var(--text-xl);
}

.empty-actions {
  display: flex;
  gap: var(--space-4);
}

.empty-actions button,
.new-button {
  font: inherit;
  padding: var(--space-3) var(--space-6);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.library {
  max-width: 56rem;
  margin: 0 auto;
  padding: var(--space-7) var(--space-5);
}

.library-actions {
  display: flex;
  gap: var(--space-3);
}

.library-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: var(--space-5);
}

.library-heading {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.library-table {
  width: 100%;
  border-collapse: collapse;
  background: var(--surface-base);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-1);
}

.library-table th {
  text-align: left;
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  padding: var(--space-3) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-default);
}

.library-table td {
  padding: var(--space-3) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-subtle);
  color: var(--text-primary);
}

.library-row {
  cursor: pointer;
}

.library-row:hover {
  background: var(--surface-raised);
}

/* The just-closed row flashes once on return (SPEC-060 AC-22): a single
   fade from the selected surface back to transparent over one --motion-base
   interval. Purely an animation, so it leaves no lingering background; under
   prefers-reduced-motion the class is never applied (and the animation is
   stilled here too, defensively). */
.row-highlight {
  animation: row-flash var(--motion-base) var(--easing);
}

@keyframes row-flash {
  from {
    background: var(--surface-selected);
  }
  to {
    background: transparent;
  }
}

@media (prefers-reduced-motion: reduce) {
  .row-highlight {
    animation: none;
  }
}

.cell-title {
  font-weight: var(--weight-medium);
  overflow-wrap: anywhere;
}

.cell-modified {
  color: var(--text-secondary);
  white-space: nowrap;
}

.badge {
  display: inline-block;
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.cell-actions {
  position: relative;
  width: 2.5rem;
  text-align: right;
}

.menu-button {
  font: inherit;
  font-weight: var(--weight-bold);
  line-height: 1;
  padding: var(--space-2) var(--space-3);
  border: none;
  border-radius: var(--radius-sm);
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.menu-button:hover {
  background: var(--surface-sunken);
}

.row-menu {
  position: absolute;
  right: 0.5rem;
  top: 2rem;
  display: flex;
  flex-direction: column;
  min-width: 8rem;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-2);
  z-index: 10;
}

.row-menu button {
  font: inherit;
  text-align: left;
  padding: var(--space-3) var(--space-5);
  border: none;
  background: var(--surface-base);
  cursor: pointer;
  color: var(--text-primary);
}

.row-menu button:hover {
  background: var(--surface-sunken);
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

.open-error {
  margin: 0;
  color: var(--text-danger);
  max-width: 32rem;
  text-align: center;
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}
</style>

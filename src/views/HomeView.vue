<script setup lang="ts">
import { useRouter } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { openDocxFile } from "../lib/openDocument";
import { computed, onMounted, onUnmounted, ref, useTemplateRef } from "vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import RenameDialog from "../components/RenameDialog.vue";
import {
  deleteDocument,
  duplicateDocument,
  getOrCreateDocument,
  initLibrary,
  listDocuments,
  renameDocument,
  storeDocument
} from "../lib/documentStore";
import { formatBadgeLabel, formatRelativeTime } from "../lib/format";

const router = useRouter();
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
onUnmounted(() => {
  document.removeEventListener("click", closeMenu);
  clearInterval(clockTimer);
});

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
    <div
      v-else
      class="library"
      data-testid="library"
    >
      <div class="library-header">
        <h2 class="library-heading">
          Documents
        </h2>
        <button
          type="button"
          class="new-button"
          data-testid="new-document"
          @click="createNewDocument"
        >
          New document
        </button>
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
            data-testid="library-row"
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

.empty-actions button,
.new-button {
  font: inherit;
  padding: 0.5rem 1.25rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.library {
  max-width: 56rem;
  margin: 0 auto;
  padding: 2rem 1rem;
}

.library-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin-bottom: 1rem;
}

.library-heading {
  margin: 0;
  font-size: 1.125rem;
  font-weight: 600;
  color: #202124;
}

.library-table {
  width: 100%;
  border-collapse: collapse;
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.2);
}

.library-table th {
  text-align: left;
  font-size: 0.8125rem;
  font-weight: 600;
  color: #5f6368;
  padding: 0.625rem 1rem;
  border-bottom: 1px solid #e0e0e0;
}

.library-table td {
  padding: 0.625rem 1rem;
  border-bottom: 1px solid #f0f0f0;
  color: #202124;
}

.library-row {
  cursor: pointer;
}

.library-row:hover {
  background: #f8f9fa;
}

.cell-title {
  font-weight: 500;
  overflow-wrap: anywhere;
}

.cell-modified {
  color: #5f6368;
  white-space: nowrap;
}

.badge {
  display: inline-block;
  padding: 0.125rem 0.5rem;
  border-radius: 999px;
  background: #e8f0fe;
  color: #1a5276;
  font-size: 0.75rem;
  font-weight: 600;
}

.cell-actions {
  position: relative;
  width: 2.5rem;
  text-align: right;
}

.menu-button {
  font: inherit;
  font-weight: 700;
  line-height: 1;
  padding: 0.25rem 0.5rem;
  border: none;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  color: #5f6368;
}

.menu-button:hover {
  background: #e8eaed;
}

.row-menu {
  position: absolute;
  right: 0.5rem;
  top: 2rem;
  display: flex;
  flex-direction: column;
  min-width: 8rem;
  background: #fff;
  border: 1px solid #e0e0e0;
  border-radius: 4px;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.2);
  z-index: 10;
}

.row-menu button {
  font: inherit;
  text-align: left;
  padding: 0.5rem 1rem;
  border: none;
  background: #fff;
  cursor: pointer;
  color: #202124;
}

.row-menu button:hover {
  background: #f1f3f4;
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
  overflow: hidden;
  clip: rect(0 0 0 0);
  white-space: nowrap;
}
</style>

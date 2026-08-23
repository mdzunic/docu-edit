<script setup lang="ts">
import { useRouter } from "vue-router";
import { computed, nextTick, onMounted, onUnmounted, ref, useTemplateRef, watch } from "vue";
import { useTheme } from "../composables/useTheme";
import { useToasts } from "../composables/useToasts";
import { useLibrary } from "../composables/useLibrary";
import AppShell from "../components/AppShell.vue";
import UiButton from "../components/ui/UiButton.vue";
import ConfirmDeleteDialog from "../components/ConfirmDeleteDialog.vue";
import RenameDialog from "../components/RenameDialog.vue";
import TemplateGalleryDialog from "../components/TemplateGalleryDialog.vue";
import CommandPalette from "../components/CommandPalette.vue";
import RowMenu from "../components/library/RowMenu.vue";
import LibraryCard from "../components/library/LibraryCard.vue";
import HighlightedTitle from "../components/library/HighlightedTitle.vue";
import { openDocumentFile } from "../lib/openDocument";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";
import { COMMANDS, loadRecentCommands, pushRecentCommand } from "../lib/commandCatalogue";
import {
  getOrCreateDocument,
  initLibrary,
  renameDocument,
  storeDocument,
} from "../lib/documentStore";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { instantiateTemplate } from "../lib/templateStore";
import { createDocument } from "../lib/documentModel";
import { formatBadgeLabel, formatFileSize, formatRelativeTime } from "../lib/format";
import { FILTERABLE_FORMATS } from "../lib/libraryView";

const router = useRouter();
const { addToast } = useToasts();

// The library has no menu bar, so it reads the shared theme composable to keep
// the OS-appearance listener alive and be themed by the same tokens (SPEC-059).
useTheme();

const lib = useLibrary();

const fileInput = useTemplateRef("fileInput");
const openError = ref("");

// main.ts awaits initLibrary() before mounting, so the store is hydrated here in
// the running app; the refresh-on-init matters only when the view is mounted
// directly (tests).
lib.refresh();
initLibrary().then(() => lib.refresh());

// Relative times go stale while the page sits open; re-render them each minute.
const clockTimer = setInterval(() => {
  lib.now.value = Date.now();
}, 60_000);

// Storage quota (SPEC-067 AC-47): the browser's estimate, read once. Absent in
// engines without the Storage API, in which case the panel shows the used bytes
// without a percentage.
const quotaBytes = ref(0);
onMounted(async () => {
  try {
    if (navigator.storage?.estimate) {
      const estimate = await navigator.storage.estimate();
      quotaBytes.value = estimate.quota ?? 0;
    }
  } catch {
    // No estimate available — fall back to bytes-only display.
  }
});
const usagePercent = computed(() =>
  quotaBytes.value > 0
    ? Math.min(100, Math.round((lib.usage.value.totalBytes / quotaBytes.value) * 100))
    : null,
);
const usageHigh = computed(() => usagePercent.value !== null && usagePercent.value >= 80);

// --- overflow menu -------------------------------------------------------

const menuFor = ref("");
function toggleMenu(id = "") {
  menuFor.value = menuFor.value === id ? "" : id;
}
function closeMenu() {
  menuFor.value = "";
}
onMounted(() => document.addEventListener("click", closeMenu));

// Preload the editor route's chunk while the library sits idle (SPEC-047 AC-3).
onMounted(() => {
  void import("./DocumentView.vue");
});
onUnmounted(() => {
  document.removeEventListener("click", closeMenu);
  document.removeEventListener("keydown", handleLibraryPaletteKey);
  clearInterval(clockTimer);
});

function libraryMenuItems(starred = false) {
  return [
    { id: "star", label: starred ? "Unstar" : "Star", testid: "menu-star" },
    { id: "rename", label: "Rename", testid: "menu-rename" },
    { id: "duplicate", label: "Duplicate", testid: "menu-duplicate" },
    { id: "delete", label: "Delete", testid: "menu-delete", danger: true },
  ];
}
const trashMenuItems = [
  { id: "restore", label: "Restore", testid: "menu-restore" },
  { id: "delete-permanently", label: "Delete permanently", testid: "menu-delete-permanently", danger: true },
];

function onMenuSelect(payload = { id: "", action: "" }) {
  closeMenu();
  const { id, action } = payload;
  if (action === "star") {
    lib.toggleStar(id);
  } else if (action === "rename") {
    renameFor.value = id;
  } else if (action === "duplicate") {
    lib.duplicateDocuments([id]);
  } else if (action === "delete") {
    trashBatch([id]);
  } else if (action === "restore") {
    lib.restoreDocuments([id]);
  } else if (action === "delete-permanently") {
    permanentDeleteFor.value = id;
  }
}

// --- selection + open ----------------------------------------------------

function openDocument(id = "") {
  router.push(`/doc/${id}`);
}

function onSelectModifiers(id = "", modifiers = { shift: false, ctrlOrMeta: false }) {
  lib.selectWithModifiers(id, modifiers);
  lib.focusedId.value = id;
}

function onRowClick(id = "", event = new MouseEvent("click")) {
  if (event.shiftKey || event.ctrlKey || event.metaKey) {
    event.preventDefault();
    onSelectModifiers(id, {
      shift: event.shiftKey,
      ctrlOrMeta: event.ctrlKey || event.metaKey,
    });
    return;
  }
  openDocument(id);
}

function rowTabindex(id = "", index = 0) {
  const focused = lib.focusedId.value;
  if (focused === id) {
    return 0;
  }
  return focused === "" && index === 0 ? 0 : -1;
}

// --- keyboard navigation (AC-57..62) -------------------------------------

async function focusDom() {
  await nextTick();
  const el = document.querySelector<HTMLElement>(`[data-doc-id="${lib.focusedId.value}"]`);
  el?.focus();
}

function handleContentKeydown(event = new KeyboardEvent("keydown")) {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "a") {
    event.preventDefault();
    lib.selectAll();
    return;
  }
  const key = event.key;
  if (key === "ArrowDown" || key === "ArrowRight") {
    event.preventDefault();
    lib.focusNext(1);
    void focusDom();
  } else if (key === "ArrowUp" || key === "ArrowLeft") {
    event.preventDefault();
    lib.focusNext(-1);
    void focusDom();
  } else if (key === "Enter") {
    if (lib.focusedId.value) {
      event.preventDefault();
      openDocument(lib.focusedId.value);
    }
  } else if (key === " ") {
    if (lib.focusedId.value) {
      event.preventDefault();
      lib.toggleSelection(lib.focusedId.value);
    }
  } else if (key === "Delete" || key === "Backspace") {
    if (!lib.showTrash.value && lib.focusedId.value) {
      event.preventDefault();
      const ids = lib.hasSelection.value ? lib.selectedIds.value : [lib.focusedId.value];
      trashBatch(ids);
    }
  } else if (key === "F2") {
    if (!lib.showTrash.value && lib.focusedId.value) {
      event.preventDefault();
      renameFor.value = lib.focusedId.value;
    }
  }
}

// --- trash + undo --------------------------------------------------------

function trashBatch(ids = [""]) {
  const titles = ids.map((id) => lib.titleOf(id));
  const batch = lib.trashDocuments(ids);
  if (batch.length === 0) {
    return;
  }
  const message =
    batch.length === 1
      ? `Moved “${titles[0]}” to Trash`
      : `Moved ${batch.length} documents to Trash`;
  addToast(message, { label: "Undo", run: () => lib.restoreDocuments(batch) });
}

// --- bulk actions --------------------------------------------------------

function bulkTrash() {
  trashBatch(lib.selectedIds.value);
}
function bulkDuplicate() {
  lib.duplicateDocuments(lib.selectedIds.value);
}
function bulkStar() {
  lib.starDocuments(lib.selectedIds.value);
}
function bulkDownload() {
  void lib.downloadDocuments(lib.selectedIds.value);
}

// --- rename + permanent delete dialogs -----------------------------------

const renameFor = ref("");
const renameTarget = computed(() => lib.titleOf(renameFor.value));
const permanentDeleteFor = ref("");
const permanentDeleteTitle = computed(() => lib.titleOf(permanentDeleteFor.value));
const confirmingEmptyTrash = ref(false);

function applyRename(newTitle = "") {
  renameDocument(renameFor.value, newTitle);
  renameFor.value = "";
  lib.refresh();
}

function confirmPermanentDelete() {
  lib.deletePermanently([permanentDeleteFor.value]);
  permanentDeleteFor.value = "";
}

function confirmEmptyTrash() {
  lib.emptyTrash();
  confirmingEmptyTrash.value = false;
}

// --- creating / opening documents (SPEC-003, unchanged behaviour) --------

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

// --- command palette (SPEC-062) ------------------------------------------

// Preferences (SPEC-066) are reachable from the library too (AC-39): the
// app-bar button and the library command palette both open the dialog AppShell
// hosts.
const preferencesDialog = usePreferencesDialog();

const libraryCommands = COMMANDS.filter((command) => command.scope === "library");
const libraryCommandIds = libraryCommands.map((command) => command.id);
const commandPaletteOpen = ref(false);
const paletteRecentIds = ref(loadRecentCommands());

function handleLibraryPaletteKey(event = new KeyboardEvent("keydown")) {
  if ((event.metaKey || event.ctrlKey) && !event.altKey && event.key.toLowerCase() === "k") {
    event.preventDefault();
    commandPaletteOpen.value = true;
  }
}
onMounted(() => document.addEventListener("keydown", handleLibraryPaletteKey));

function runLibraryCommand(id = "") {
  paletteRecentIds.value = pushRecentCommand(id, paletteRecentIds.value);
  commandPaletteOpen.value = false;
  if (id === "library.new-document") {
    createNewDocument();
  } else if (id === "library.new-from-template") {
    openTemplateGallery();
  } else if (id === "library.open") {
    pickFile();
  } else if (id === "library.preferences") {
    preferencesDialog.open();
  }
}

// --- SPEC-060 returning highlight ----------------------------------------

const highlightId = ref("");

function prefersReducedMotion() {
  return (
    typeof window.matchMedia === "function" &&
    window.matchMedia("(prefers-reduced-motion: reduce)").matches
  );
}
function motionBaseMs() {
  const raw = getComputedStyle(document.documentElement).getPropertyValue("--motion-base").trim();
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
  const row = document.querySelector(`[data-doc-id="${closedId}"]`);
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

// A plain string ref drives the sort <select> (v-model), avoiding a TS `as`
// cast on the change event (barred in .vue). String(...) forces Ref<string> so
// v-model can assign to it; setSortKey validates the value.
const sortKeyChoice = ref(String(lib.sort.value.key));
watch(sortKeyChoice, (value) => lib.setSortKey(value));
</script>

<template>
  <AppShell document-title="No document">
    <template #app-bar-actions>
      <UiButton
        variant="ghost"
        data-testid="library-preferences"
        @click="preferencesDialog.open()"
      >
        Preferences
      </UiButton>
    </template>
    <!-- First-run: no documents anywhere (AC-51..53). -->
    <div
      v-if="lib.libraryEmpty.value && lib.trashCount.value === 0 && !lib.showTrash.value"
      class="empty-state"
      data-testid="empty-state"
    >
      <h2 class="empty-title">
        Welcome to docu-edit
      </h2>
      <p class="empty-lead">
        A private word processor that opens, edits and saves Word, OpenDocument
        and Markdown files — right here in your browser.
      </p>
      <div class="empty-actions">
        <UiButton
          variant="primary"
          data-testid="new-document"
          @click="createNewDocument"
        >
          New document
        </UiButton>
        <UiButton
          data-testid="new-from-template"
          @click="openTemplateGallery"
        >
          New from template…
        </UiButton>
        <UiButton
          data-testid="open-document"
          @click="pickFile"
        >
          Open a file
        </UiButton>
      </div>
      <p class="empty-note">
        Your documents are stored in this browser — nothing is uploaded.
      </p>
    </div>

    <!-- Trash view (AC-36..39, AC-45). -->
    <div
      v-else-if="lib.showTrash.value"
      class="library"
      data-testid="trash-view"
    >
      <div class="library-header">
        <div class="header-line">
          <UiButton
            variant="ghost"
            data-testid="back-to-library"
            @click="lib.showTrash.value = false"
          >
            ← Library
          </UiButton>
          <h2 class="library-heading">
            Trash
          </h2>
          <div class="header-spacer" />
          <UiButton
            variant="danger"
            data-testid="empty-trash"
            :disabled="lib.trashCount.value === 0"
            @click="confirmingEmptyTrash = true"
          >
            Empty trash
          </UiButton>
        </div>
        <p class="trash-note">
          Documents here keep their id and version history. Restore returns them
          to your library intact.
        </p>
      </div>
      <p
        v-if="lib.trashCount.value === 0"
        class="state-message"
        data-testid="trash-empty"
      >
        The trash is empty.
      </p>
      <table
        v-else
        class="library-table"
        role="grid"
        data-testid="trash-list"
        @keydown="handleContentKeydown"
      >
        <thead>
          <tr>
            <th scope="col">
              Title
            </th>
            <th scope="col">
              Trashed
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
            v-for="(entry, index) in lib.trashVisible.value"
            :key="entry.id"
            class="library-row"
            role="row"
            data-testid="library-row"
            :data-doc-id="entry.id"
            :tabindex="rowTabindex(entry.id, index)"
            @focus="lib.focusedId.value = entry.id"
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
              {{ formatRelativeTime(entry.trashedAt ?? entry.lastModified, lib.now.value) }}
            </td>
            <td class="cell-format">
              <span
                class="badge"
                data-testid="row-format"
              >{{ formatBadgeLabel(entry.format) }}</span>
            </td>
            <td class="cell-actions">
              <UiButton
                size="sm"
                data-testid="row-restore"
                @click="lib.restoreDocuments([entry.id])"
              >
                Restore
              </UiButton>
              <RowMenu
                :open="menuFor === entry.id"
                :label="`Actions for ${entry.title}`"
                :items="trashMenuItems"
                @toggle="toggleMenu(entry.id)"
                @close="closeMenu"
                @select="(action) => onMenuSelect({ id: entry.id, action })"
              />
            </td>
          </tr>
        </tbody>
      </table>
    </div>

    <!-- The library shell (AC-1, AC-8..50). -->
    <div
      v-else
      class="library"
      data-testid="library"
    >
      <div class="library-header">
        <div class="header-line">
          <h2 class="library-heading">
            Documents
          </h2>
          <input
            v-model="lib.search.value"
            type="search"
            class="library-search"
            data-testid="library-search"
            placeholder="Search documents…"
            aria-label="Search documents"
          >
          <div class="sort-control">
            <label
              class="visually-hidden"
              for="library-sort"
            >Sort by</label>
            <select
              id="library-sort"
              v-model="sortKeyChoice"
              class="sort-select"
              data-testid="sort-key"
            >
              <option value="modified">
                Last modified
              </option>
              <option value="title">
                Title
              </option>
              <option value="size">
                Size
              </option>
            </select>
            <button
              type="button"
              class="sort-direction"
              data-testid="sort-direction"
              :aria-label="lib.sort.value.direction === 'asc' ? 'Ascending' : 'Descending'"
              @click="lib.toggleSortDirection()"
            >
              {{ lib.sort.value.direction === "asc" ? "↑" : "↓" }}
            </button>
          </div>
          <div
            class="view-toggle"
            role="group"
            aria-label="View"
          >
            <button
              type="button"
              class="toggle-button"
              data-testid="view-list"
              :class="{ 'is-active': lib.view.value === 'list' }"
              :aria-pressed="lib.view.value === 'list'"
              @click="lib.setView('list')"
            >
              List
            </button>
            <button
              type="button"
              class="toggle-button"
              data-testid="view-grid"
              :class="{ 'is-active': lib.view.value === 'grid' }"
              :aria-pressed="lib.view.value === 'grid'"
              @click="lib.setView('grid')"
            >
              Grid
            </button>
          </div>
          <div class="header-spacer" />
          <button
            type="button"
            class="trash-link"
            data-testid="trash-link"
            @click="lib.showTrash.value = true"
          >
            Trash <span
              class="trash-count"
              data-testid="trash-count"
            >({{ lib.trashCount.value }})</span>
          </button>
          <div class="library-actions">
            <UiButton
              variant="primary"
              data-testid="new-document"
              @click="createNewDocument"
            >
              New document
            </UiButton>
            <UiButton
              data-testid="new-from-template"
              @click="openTemplateGallery"
            >
              New from template…
            </UiButton>
            <UiButton
              data-testid="open-document"
              @click="pickFile"
            >
              Open document
            </UiButton>
          </div>
        </div>
        <div
          class="filter-chips"
          role="group"
          aria-label="Filters"
        >
          <button
            v-for="format in FILTERABLE_FORMATS"
            :key="format"
            type="button"
            class="chip"
            :data-testid="`filter-${format}`"
            :class="{ 'is-active': lib.activeFormats.value.has(format) }"
            :aria-pressed="lib.activeFormats.value.has(format)"
            @click="lib.toggleFormat(format)"
          >
            {{ formatBadgeLabel(format) }}
            <span class="chip-count">{{ lib.counts.value.formats[format] }}</span>
          </button>
          <button
            type="button"
            class="chip"
            data-testid="filter-starred"
            :class="{ 'is-active': lib.starredOnly.value }"
            :aria-pressed="lib.starredOnly.value"
            @click="lib.toggleStarredFilter()"
          >
            ★ Starred
            <span class="chip-count">{{ lib.counts.value.starred }}</span>
          </button>
        </div>
      </div>

      <!-- Bulk action bar (AC-27..32). -->
      <div
        v-if="lib.hasSelection.value"
        class="selection-bar"
        data-testid="selection-bar"
        role="toolbar"
        aria-label="Bulk actions"
      >
        <span
          class="selection-count"
          data-testid="selection-count"
        >
          {{ lib.selectionCount.value }} selected
        </span>
        <UiButton
          size="sm"
          data-testid="bulk-download"
          @click="bulkDownload"
        >
          Download
        </UiButton>
        <UiButton
          size="sm"
          data-testid="bulk-duplicate"
          @click="bulkDuplicate"
        >
          Duplicate
        </UiButton>
        <UiButton
          size="sm"
          data-testid="bulk-star"
          @click="bulkStar"
        >
          Star
        </UiButton>
        <UiButton
          size="sm"
          variant="danger"
          data-testid="bulk-trash"
          @click="bulkTrash"
        >
          Move to trash
        </UiButton>
        <UiButton
          size="sm"
          variant="ghost"
          data-testid="clear-selection"
          @click="lib.clearSelection()"
        >
          Clear selection
        </UiButton>
      </div>

      <!-- Storage usage (AC-47..50). -->
      <div
        class="storage-panel"
        data-testid="storage-panel"
      >
        <span
          class="storage-usage"
          data-testid="storage-usage"
        >
          {{ formatFileSize(lib.usage.value.totalBytes) }}<template v-if="usagePercent !== null"> of your browser's storage ({{ usagePercent }}%)</template> in use
        </span>
        <span
          v-if="lib.largestDocuments.value.length > 0"
          class="storage-largest"
        >
          Largest:
          <span
            v-for="doc in lib.largestDocuments.value.slice(0, 3)"
            :key="doc.id"
            class="largest-doc"
            data-testid="largest-doc"
          >{{ doc.title }} ({{ formatFileSize(doc.bytes) }})</span>
        </span>
        <button
          v-if="usageHigh || lib.trashCount.value > 0"
          type="button"
          class="storage-trash-link"
          data-testid="storage-trash-link"
          @click="lib.showTrash.value = true"
        >
          {{ usageHigh ? "Storage is running low — empty the trash" : "Review the trash" }}
        </button>
      </div>

      <!-- No results (AC-54..56). -->
      <div
        v-if="lib.noResults.value"
        class="state-message no-results"
        data-testid="no-results"
      >
        <p>No documents match your search or filters.</p>
        <UiButton
          data-testid="clear-filters"
          @click="lib.clearFilters()"
        >
          Clear search and filters
        </UiButton>
      </div>

      <!-- Library empty but trash has documents. -->
      <div
        v-else-if="lib.libraryEmpty.value"
        class="state-message"
        data-testid="library-empty"
      >
        <p>Your library is empty. {{ lib.trashCount.value }} document(s) in the trash.</p>
      </div>

      <!-- List view (AC-1). -->
      <table
        v-else-if="lib.view.value === 'list'"
        class="library-table"
        role="grid"
        data-testid="library-list"
        @keydown="handleContentKeydown"
      >
        <thead>
          <tr role="row">
            <th
              scope="col"
              role="columnheader"
              class="col-select"
            >
              <span class="visually-hidden">Select</span>
            </th>
            <th
              scope="col"
              role="columnheader"
              class="col-star"
            >
              <span class="visually-hidden">Starred</span>
            </th>
            <th
              scope="col"
              role="columnheader"
            >
              Title
            </th>
            <th
              scope="col"
              role="columnheader"
            >
              Last modified
            </th>
            <th
              scope="col"
              role="columnheader"
            >
              Format
            </th>
            <th
              scope="col"
              role="columnheader"
            >
              <span class="visually-hidden">Actions</span>
            </th>
          </tr>
        </thead>
        <tbody>
          <tr
            v-for="(entry, index) in lib.visible.value"
            :key="entry.id"
            class="library-row"
            role="row"
            data-testid="library-row"
            :data-doc-id="entry.id"
            :class="{ 'row-highlight': entry.id === highlightId, 'is-selected': lib.isSelected(entry.id) }"
            :tabindex="rowTabindex(entry.id, index)"
            :aria-selected="lib.isSelected(entry.id)"
            @click="onRowClick(entry.id, $event)"
            @focus="lib.focusedId.value = entry.id"
          >
            <td
              class="cell-select"
              @click.stop
            >
              <input
                type="checkbox"
                data-testid="row-select"
                :checked="lib.isSelected(entry.id)"
                :aria-label="`Select ${entry.title}`"
                @change="lib.toggleSelection(entry.id)"
              >
            </td>
            <td
              class="cell-star"
              @click.stop
            >
              <button
                type="button"
                class="star-button"
                data-testid="row-star"
                :class="{ 'is-starred': entry.starred }"
                :aria-pressed="entry.starred"
                :aria-label="entry.starred ? `Unstar ${entry.title}` : `Star ${entry.title}`"
                @click="lib.toggleStar(entry.id)"
              >
                {{ entry.starred ? "★" : "☆" }}
              </button>
            </td>
            <td
              class="cell-title"
              data-testid="row-title"
            >
              <HighlightedTitle
                :title="entry.title"
                :search="lib.search.value"
              />
            </td>
            <td
              class="cell-modified"
              data-testid="row-modified"
            >
              {{ formatRelativeTime(entry.lastModified, lib.now.value) }}
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
              <RowMenu
                :open="menuFor === entry.id"
                :label="`Actions for ${entry.title}`"
                :items="libraryMenuItems(entry.starred)"
                @toggle="toggleMenu(entry.id)"
                @close="closeMenu"
                @select="(action) => onMenuSelect({ id: entry.id, action })"
              />
            </td>
          </tr>
        </tbody>
      </table>

      <!-- Grid view (AC-1, AC-3..7). -->
      <div
        v-else
        class="library-grid"
        role="grid"
        data-testid="library-grid"
        @keydown="handleContentKeydown"
      >
        <LibraryCard
          v-for="(entry, index) in lib.visible.value"
          :key="entry.id"
          :entry="entry"
          :search="lib.search.value"
          :now="lib.now.value"
          :selected="lib.isSelected(entry.id)"
          :focused="rowTabindex(entry.id, index) === 0"
          :menu-open="menuFor === entry.id"
          :highlight="entry.id === highlightId"
          :menu-items="libraryMenuItems(entry.starred)"
          @open="openDocument"
          @select="onSelectModifiers($event.id, { shift: $event.shift, ctrlOrMeta: $event.ctrlOrMeta })"
          @toggle-check="lib.toggleSelection"
          @toggle-star="lib.toggleStar"
          @menu-toggle="toggleMenu"
          @menu-close="closeMenu"
          @menu-select="onMenuSelect"
          @focus="lib.focusedId.value = $event"
        />
      </div>
    </div>

    <!-- Shared across every branch. -->
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
      v-if="renameFor"
      :current-title="renameTarget"
      @save="applyRename"
      @cancel="renameFor = ''"
    />
    <ConfirmDeleteDialog
      v-if="permanentDeleteFor"
      :document-title="permanentDeleteTitle"
      kind="document permanently"
      @confirm="confirmPermanentDelete"
      @cancel="permanentDeleteFor = ''"
    />
    <ConfirmDeleteDialog
      v-else-if="confirmingEmptyTrash"
      :document-title="`${lib.trashCount.value} document(s) in the trash`"
      kind="everything in the trash"
      @confirm="confirmEmptyTrash"
      @cancel="confirmingEmptyTrash = false"
    />
    <TemplateGalleryDialog
      v-if="showGallery"
      @choose-blank="chooseBlankDocument"
      @choose-builtin="chooseBuiltinTemplate"
      @choose-user="chooseUserTemplate"
      @choose-file="chooseTemplateFile"
      @close="showGallery = false"
    />
    <CommandPalette
      v-if="commandPaletteOpen"
      :commands="libraryCommands"
      :common-ids="libraryCommandIds"
      :recent-ids="paletteRecentIds"
      @run="runLibraryCommand"
      @dismiss="commandPaletteOpen = false"
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
  text-align: center;
  padding: var(--space-6);
}

.empty-title {
  margin: 0;
  font-size: var(--text-2xl, var(--text-xl));
  font-weight: var(--weight-bold);
}

.empty-lead {
  margin: 0;
  max-width: 32rem;
  color: var(--text-secondary);
  font-size: var(--text-lg);
}

.empty-actions {
  display: flex;
  gap: var(--space-4);
  flex-wrap: wrap;
  justify-content: center;
}

.empty-note {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.library {
  max-width: 64rem;
  margin: 0 auto;
  padding: var(--space-7) var(--space-5);
}

.library-header {
  margin-bottom: var(--space-5);
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
}

.header-line {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.header-spacer {
  flex: 1 1 auto;
}

.library-heading {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.library-search {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  min-width: 12rem;
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.sort-control {
  display: flex;
  align-items: center;
  gap: var(--space-1);
}

.sort-select {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
}

.sort-direction,
.toggle-button,
.trash-link {
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.view-toggle {
  display: inline-flex;
}

.view-toggle .toggle-button:first-child {
  border-top-right-radius: 0;
  border-bottom-right-radius: 0;
}

.view-toggle .toggle-button:last-child {
  border-top-left-radius: 0;
  border-bottom-left-radius: 0;
  border-left: none;
}

.toggle-button.is-active {
  background: var(--surface-selected);
  color: var(--text-accent);
  font-weight: var(--weight-bold);
}

.trash-count {
  color: var(--text-secondary);
}

.library-actions {
  display: flex;
  gap: var(--space-3);
}

.filter-chips {
  display: flex;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.chip {
  font: inherit;
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-pill);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.chip.is-active {
  background: var(--surface-selected);
  color: var(--text-accent);
  border-color: var(--text-accent);
  font-weight: var(--weight-bold);
}

.chip-count {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.chip.is-active .chip-count {
  color: var(--text-accent);
}

.selection-bar {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
  margin-bottom: var(--space-4);
  border-radius: var(--radius-md);
  background: var(--surface-selected);
}

.selection-count {
  font-weight: var(--weight-bold);
  color: var(--text-accent);
  margin-right: auto;
}

.storage-panel {
  display: flex;
  align-items: center;
  gap: var(--space-4);
  flex-wrap: wrap;
  padding: var(--space-3) var(--space-4);
  margin-bottom: var(--space-5);
  border-radius: var(--radius-md);
  background: var(--surface-raised);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.storage-largest {
  display: flex;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.largest-doc:not(:last-child)::after {
  content: ",";
}

.storage-trash-link {
  font: inherit;
  margin-left: auto;
  border: none;
  background: transparent;
  color: var(--text-accent);
  cursor: pointer;
  text-decoration: underline;
}

.state-message {
  padding: var(--space-7) var(--space-5);
  text-align: center;
  color: var(--text-secondary);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-4);
}

.no-results {
  border: var(--border-width) dashed var(--border-default);
  border-radius: var(--radius-md);
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

.col-select,
.col-star,
.cell-select,
.cell-star {
  width: 2.5rem;
  padding-right: 0;
}

.library-row {
  cursor: pointer;
}

.library-row:hover {
  background: var(--surface-raised);
}

.library-row:focus-visible {
  outline: var(--focus-ring-width, 2px) solid var(--border-focus, var(--text-accent));
  outline-offset: -2px;
}

.library-row.is-selected {
  background: var(--surface-selected);
}

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

.star-button {
  font-size: var(--text-lg);
  line-height: 1;
  border: none;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.star-button.is-starred {
  color: var(--text-accent);
}

.cell-title {
  font-weight: var(--weight-medium);
  overflow-wrap: anywhere;
}

.cell-modified {
  color: var(--text-secondary);
  white-space: nowrap;
}

.cell-actions {
  position: relative;
  text-align: right;
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

.library-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(14rem, 1fr));
  gap: var(--space-5);
}

.trash-note,
.trash-link ~ .trash-note {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
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
  margin: var(--space-4) auto 0;
  color: var(--text-danger);
  max-width: 32rem;
  text-align: center;
}
</style>

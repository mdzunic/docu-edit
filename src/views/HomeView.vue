<script setup lang="ts">
// The landing page: a thin wiring hub over the library module
// (src/views/library/* + useLibrary + lib/libraryView), the same
// hub-and-composables shape DocumentView has. Product knowledge stays here
// — what can be created (text documents, templates), which file kinds open,
// and where each kind routes (documentKinds.ts) — while the library
// components stay product-neutral. HomeView.size.spec.ts holds the line
// budget that keeps it this way.
import { useRouter } from "vue-router";
import { computed, nextTick, onMounted, onUnmounted, ref, useTemplateRef } from "vue";
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
import LibraryControls from "./library/LibraryControls.vue";
import LibraryEmptyState from "./library/LibraryEmptyState.vue";
import LibraryList from "./library/LibraryList.vue";
import LibrarySelectionBar from "./library/LibrarySelectionBar.vue";
import LibraryStoragePanel from "./library/LibraryStoragePanel.vue";
import LibraryTrash from "./library/LibraryTrash.vue";
import { useTier } from "../composables/useResponsiveTier";
import { openAnyDocumentFile } from "../lib/openDocument";
import { editorPathFor, OPEN_FILE_ACCEPT } from "../lib/documentKinds";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";
import { COMMANDS, loadRecentCommands, pushRecentCommand } from "../lib/commandCatalogue";
import { libraryRowMenuItems } from "../lib/libraryMenu";
import {
  getOrCreateDocument,
  initLibrary,
  renameDocument,
  storeDocument,
  storePresentation,
} from "../lib/documentStore";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { instantiateTemplate } from "../lib/templateStore";
import { createDocument } from "../lib/documentModel";

const router = useRouter();
const { addToast } = useToasts();

// The library has no menu bar, so it reads the shared theme composable to keep
// the OS-appearance listener alive and be themed by the same tokens (SPEC-059).
useTheme();

const lib = useLibrary();

// SPEC-068 AC-38 — at the compact tier the library's three header actions
// collapse to one primary action (New document) plus an overflow menu holding
// the rest; the cards reflow to one column (CSS) and search stays visible.
const tier = useTier();
const compact = computed(() => tier.value === "compact");
const headerActionsMenuOpen = ref(false);
const HEADER_OVERFLOW_ACTIONS = [
  { id: "new-from-template", label: "New from template…", testid: "new-from-template" },
  { id: "open-document", label: "Open document", testid: "open-document" },
];
function runHeaderAction(id = "") {
  headerActionsMenuOpen.value = false;
  if (id === "new-from-template") {
    openTemplateGallery();
  } else if (id === "open-document") {
    pickFile();
  }
}

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

// The accept list spans every kind the app opens (see documentKinds.ts).
const openFileAccept = OPEN_FILE_ACCEPT;

function openDocument(id = "") {
  // Kind-aware routing: text documents open the editor, presentations the
  // presentation view (see documentKinds.editorPathFor).
  router.push(editorPathFor(lib.kindOf(id), id));
}

function onSelectModifiers(id = "", modifiers = { shift: false, ctrlOrMeta: false }) {
  lib.selectWithModifiers(id, modifiers);
  lib.focusedId.value = id;
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

function bulkTrash() {
  trashBatch(lib.selectedIds.value);
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
    const opened = await openAnyDocumentFile(file);
    openError.value = "";
    if (opened.kind === "presentation") {
      storePresentation(opened.model);
    } else {
      storeDocument(opened.model);
    }
    router.push(editorPathFor(opened.kind, opened.model.id));
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
</script>

<template>
  <AppShell document-title="No document">
    <template #app-bar-actions>
      <UiButton
        variant="ghost"
        class="app-bar-preferences"
        data-testid="library-preferences"
        @click="preferencesDialog.open()"
      >
        Preferences
      </UiButton>
    </template>
    <!-- First-run: no documents anywhere (AC-51..53). -->
    <LibraryEmptyState
      v-if="lib.libraryEmpty.value && lib.trashCount.value === 0 && !lib.showTrash.value"
      @new-document="createNewDocument"
      @new-from-template="openTemplateGallery"
      @open-file="pickFile"
    />

    <!-- Trash view (AC-36..39, AC-45). -->
    <LibraryTrash
      v-else-if="lib.showTrash.value"
      :lib="lib"
      :menu-for="menuFor"
      @empty-trash="confirmingEmptyTrash = true"
      @keydown="handleContentKeydown"
      @menu-toggle="toggleMenu"
      @menu-close="closeMenu"
      @menu-select="onMenuSelect"
    />

    <!-- The library shell (AC-1, AC-8..50). -->
    <div
      v-else
      class="library"
      data-testid="library"
    >
      <LibraryControls :lib="lib">
        <template #actions>
          <UiButton
            variant="primary"
            data-testid="new-document"
            @click="createNewDocument"
          >
            New document
          </UiButton>
          <template v-if="!compact">
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
          </template>
          <RowMenu
            v-else
            :open="headerActionsMenuOpen"
            label="More actions"
            :items="HEADER_OVERFLOW_ACTIONS"
            button-testid="library-actions-overflow"
            menu-testid="library-actions-overflow-menu"
            @toggle="headerActionsMenuOpen = !headerActionsMenuOpen"
            @close="headerActionsMenuOpen = false"
            @select="runHeaderAction"
          />
        </template>
      </LibraryControls>

      <!-- Bulk action bar (AC-27..32). -->
      <LibrarySelectionBar
        v-if="lib.hasSelection.value"
        :lib="lib"
        @trash="bulkTrash"
      />

      <!-- Storage usage (AC-47..50). -->
      <LibraryStoragePanel :lib="lib" />

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
      <LibraryList
        v-else-if="lib.view.value === 'list'"
        :lib="lib"
        :menu-for="menuFor"
        :highlight-id="highlightId"
        @open="openDocument"
        @select-modifiers="onSelectModifiers($event.id, { shift: $event.shift, ctrlOrMeta: $event.ctrlOrMeta })"
        @keydown="handleContentKeydown"
        @menu-toggle="toggleMenu"
        @menu-close="closeMenu"
        @menu-select="onMenuSelect"
      />

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
          :focused="lib.rowTabindex(entry.id, index) === 0"
          :menu-open="menuFor === entry.id"
          :highlight="entry.id === highlightId"
          :menu-items="libraryRowMenuItems(entry.starred)"
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
      :accept="openFileAccept"
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
/* The ghost button's greys are tuned for light surfaces; on the brand bar it
   takes the on-brand foreground and a translucent wash (the same treatment
   the editor's Close control gets) so it reads in both themes. */
.app-bar-preferences {
  color: var(--text-on-brand);
}

.app-bar-preferences:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.app-bar-preferences:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}

.library {
  max-width: 64rem;
  margin: 0 auto;
  padding: var(--space-7) var(--space-5);
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

.library-grid {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(14rem, 1fr));
  gap: var(--space-5);
}

/* SPEC-068 AC-37 — the compact tier (<640px, matching useResponsiveTier):
   cards reflow to a single column. The search box's own full-width row
   lives with it in LibraryControls. */
@media (max-width: 639.98px) {
  .library-grid {
    grid-template-columns: 1fr;
  }
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

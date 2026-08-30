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
import LibraryControls from "./library/LibraryControls.vue";
import LibraryEmptyState from "./library/LibraryEmptyState.vue";
import LibraryDialogs from "./library/LibraryDialogs.vue";
import LibraryGrid from "./library/LibraryGrid.vue";
import LibraryHeaderActions from "./library/LibraryHeaderActions.vue";
import LibraryList from "./library/LibraryList.vue";
import LibraryQuickLook from "./library/LibraryQuickLook.vue";
import LibraryRail from "./library/LibraryRail.vue";
import LibrarySearchBox from "./library/LibrarySearchBox.vue";
import LibrarySelectionBar from "./library/LibrarySelectionBar.vue";
import LibraryStateMessage from "./library/LibraryStateMessage.vue";
import LibraryStoragePanel from "./library/LibraryStoragePanel.vue";
import LibraryTrash from "./library/LibraryTrash.vue";
import { openAnyDocumentFile } from "../lib/openDocument";
import { editorPathFor, OPEN_FILE_ACCEPT } from "../lib/documentKinds";
import { usePreferencesDialog } from "../composables/usePreferencesDialog";
import { COMMANDS, loadRecentCommands, pushRecentCommand } from "../lib/commandCatalogue";
import {
  getOrCreateDocument,
  initLibrary,
  renameDocument,
  storeDocument,
  storeMedia,
  storePresentation,
} from "../lib/documentStore";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { instantiateTemplate } from "../lib/templateStore";
import { createDocument } from "../lib/documentModel";
import { createBlankPresentation } from "../lib/presentation/slides";
import { useTier } from "../composables/useResponsiveTier";
import { motionBaseMs, prefersReducedMotion } from "../lib/core/motion";

const router = useRouter();
const { addToast } = useToasts();

// The library has no menu bar, so it reads the shared theme composable to keep
// the OS-appearance listener alive and be themed by the same tokens (SPEC-059).
useTheme();

const lib = useLibrary();

// SPEC-068 AC-38 — at the compact tier the library's header actions collapse
// to one primary action (New document) plus an overflow menu; that component
// owns the menu state, HomeView only supplies the tier flag.
const tier = useTier();
const compact = computed(() => tier.value === "compact");
// The handoff's full-tier layout (rail + quick look, search in the app bar).
const fullTier = computed(() => tier.value === "full");

const openError = ref("");
const dialogs = useTemplateRef("dialogs");

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

// Grid plain click: sole-select into quick look (open = dblclick/Enter/pane).
function onCardPreview(id = "") {
  lib.selectOnly(id);
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
// A blank deck (SPEC-080): stored via storePresentation, carrying "pptx" from
// birth, then straight to the presentation view.
function createNewPresentation() {
  const model = createBlankPresentation();
  storePresentation(model);
  router.push(`/slides/${model.id}`);
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
  dialogs.value?.pick();
}
async function handleFileChosen(file = null) {
  if (!file) {
    return;
  }
  try {
    const opened = await openAnyDocumentFile(file);
    openError.value = "";
    if (opened.kind === "presentation") {
      storePresentation(opened.model);
    } else if (opened.kind === "media") {
      storeMedia(opened.model);
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
  <AppShell>
    <template
      v-if="fullTier"
      #app-bar-center
    >
      <LibrarySearchBox :lib="lib" />
    </template>
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
    >
      <template #extra-actions>
        <UiButton
          data-testid="new-presentation"
          @click="createNewPresentation"
        >
          New presentation
        </UiButton>
      </template>
    </LibraryEmptyState>

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

    <!-- The library shell (AC-1, AC-8..50); full tier = rail | main | quick look. -->
    <div
      v-else
      class="library"
      :class="{ 'library--full': fullTier }"
      data-testid="library"
    >
      <LibraryRail
        v-if="fullTier"
        :lib="lib"
      >
        <template #actions>
          <LibraryHeaderActions
            :compact="false"
            @new-document="createNewDocument"
            @new-presentation="createNewPresentation"
            @new-from-template="openTemplateGallery"
            @open-document="pickFile"
          />
        </template>
        <template #storage>
          <LibraryStoragePanel :lib="lib" />
        </template>
      </LibraryRail>

      <div class="library-main">
        <LibraryControls
          :lib="lib"
          :rail="fullTier"
        >
          <template
            v-if="!fullTier"
            #actions
          >
            <LibraryHeaderActions
              :compact="compact"
              @new-document="createNewDocument"
              @new-presentation="createNewPresentation"
              @new-from-template="openTemplateGallery"
              @open-document="pickFile"
            />
          </template>
        </LibraryControls>

        <!-- Bulk action bar (AC-27..32). -->
        <LibrarySelectionBar
          v-if="lib.hasSelection.value"
          :lib="lib"
          @trash="bulkTrash"
        />

        <!-- Storage usage (AC-47..50); the rail hosts it at the full tier. -->
        <LibraryStoragePanel
          v-if="!fullTier"
          :lib="lib"
        />

        <!-- No results (AC-54..56) / empty-with-trash note. -->
        <LibraryStateMessage
          v-if="lib.noResults.value"
          no-results
          @clear-filters="lib.clearFilters()"
        />
        <LibraryStateMessage
          v-else-if="lib.libraryEmpty.value"
          :trash-count="lib.trashCount.value"
        />

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
        <LibraryGrid
          v-else
          :lib="lib"
          :menu-for="menuFor"
          :highlight-id="highlightId"
          @open="openDocument"
          @preview="onCardPreview"
          @select-modifiers="onSelectModifiers($event.id, { shift: $event.shift, ctrlOrMeta: $event.ctrlOrMeta })"
          @keydown="handleContentKeydown"
          @menu-toggle="toggleMenu"
          @menu-close="closeMenu"
          @menu-select="onMenuSelect"
          @focus="lib.focusedId.value = $event"
        />
      </div>

      <LibraryQuickLook
        v-if="fullTier"
        :lib="lib"
        @open="openDocument"
        @action="onMenuSelect"
        @download="lib.downloadDocuments([$event])"
      />
    </div>

    <!-- Shared across every branch. -->
    <LibraryDialogs
      ref="dialogs"
      :open-file-accept="openFileAccept"
      :open-error="openError"
      :rename-for="renameFor"
      :rename-target="renameTarget"
      :permanent-delete-for="permanentDeleteFor"
      :permanent-delete-title="permanentDeleteTitle"
      :confirming-empty-trash="confirmingEmptyTrash"
      :trash-count="lib.trashCount.value"
      :show-gallery="showGallery"
      :command-palette-open="commandPaletteOpen"
      :library-commands="libraryCommands"
      :library-command-ids="libraryCommandIds"
      :palette-recent-ids="paletteRecentIds"
      @rename-save="applyRename"
      @rename-cancel="renameFor = ''"
      @permanent-delete-confirm="confirmPermanentDelete"
      @permanent-delete-cancel="permanentDeleteFor = ''"
      @empty-trash-confirm="confirmEmptyTrash"
      @empty-trash-cancel="confirmingEmptyTrash = false"
      @template-choose-blank="chooseBlankDocument"
      @template-choose-builtin="chooseBuiltinTemplate"
      @template-choose-user="chooseUserTemplate"
      @template-choose-file="chooseTemplateFile"
      @template-close="showGallery = false"
      @command-run="runLibraryCommand"
      @command-dismiss="commandPaletteOpen = false"
      @file-chosen="handleFileChosen"
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

/* The modernist handoff's full-tier frame: rail | main | quick look. The two
   side columns are sticky so the file list scrolls between them. */
.library--full {
  max-width: none;
  display: grid;
  grid-template-columns: 14rem minmax(0, 1fr) 19rem;
  gap: var(--space-6);
  align-items: stretch;
  padding: var(--space-6) var(--space-6) var(--space-7);
}

/* The side panes pin to the viewport while the file list scrolls between
   them: sticky against the canvas scrollport, sized to the visible height so
   the rail bottom-pins its storage meter and the column rules span the view.
   align-self overrides the grid stretch — a full-row pane would leave sticky
   no travel. */
.library--full > :deep(.library-rail),
.library--full > :deep(.quick-look) {
  position: sticky;
  top: var(--space-6);
  align-self: start;
  height: calc(100vh - var(--height-app-bar) - 2 * var(--space-6));
  overflow-y: auto;
}

/* The main column's header — title, sort/view controls and the filter chips —
   pins too, so only the rows scroll. The negative top margin + matching
   padding extend the header's painted box up over the frame's top padding,
   so rows sliding under it never peek through the gap above. */
.library--full .library-main :deep(.library-header) {
  position: sticky;
  top: 0;
  z-index: 2;
  background: var(--surface-canvas);
  margin-top: calc(-1 * var(--space-6));
  padding-top: var(--space-6);
  padding-bottom: var(--space-3);
  margin-bottom: var(--space-3);
}
</style>

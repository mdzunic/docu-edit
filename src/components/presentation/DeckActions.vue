<script setup lang="ts">
// The deck's app-bar action cluster (Present, download, save state, Close) —
// extracted from PresentationView so the view stays a thin wiring hub (its size
// budget only ratchets down). Pure presentation + event forwarding; the view
// owns the commands these fire.
//
// It is deliberately the word processor's cluster plus Present: save state,
// the download control, Close. Four controls left when the deck line folded
// into the app bar, because eight of them overflowed the bar at 1024px and put
// Close off the right edge where nothing could reach it — and each had a truer
// home anyway:
//
//   · Notes was never an action. It opens the notes panel exactly as Design,
//     Table and Media open theirs, so it joined SlideRibbon's panel toggles
//     (its slides-notes-toggle testid unchanged).
//   · Undo/redo went to the ribbon row, which is the deck's always-visible
//     command line — where DocumentToolbar's band 1 keeps the word processor's
//     pair, and where the handoff's own 2c strip starts.
//   · Deck setup… is a rare command the File menu already carries
//     (slides-menu-deck-setup); the handoff's 2c bar holds only Download,
//     Present and Close.
//
// Close is now the same icon button the word processor's cluster ends with,
// rather than a wider text button saying the same thing.
//
// SPEC-100: Present gains a ▾ menu beside it — PresentationDownloadButton's
// split-button pattern, keyboard-operable (Enter/Space or ArrowDown opens it
// on its first item, arrows/Home/End move, Escape closes back to the caret) —
// offering the show and "Presenter view"; `present` carries which.
//
// Runtime prop/emit declarations, not the generic forms (espree lints .vue
// scripts).
import { nextTick, ref, useTemplateRef } from "vue";
import SaveStateBadge from "../SaveStateBadge.vue";
import UiButton from "../ui/UiButton.vue";
import UiIconButton from "../ui/UiIconButton.vue";
import PresentationDownloadButton from "./PresentationDownloadButton.vue";

const props = defineProps({
  /** SPEC-010's readout, for the deck. The word processor has shown this all
   * along and the slides side showed nothing, which left the product's own
   * "stored in this browser" promise uncheckable on half of it. Since
   * SPEC-103 it is the deck's autosave state ("saved" | "saving" | "error" |
   * "recovered"), and "error" adds the download link beside it. */
  saveState: { type: String, default: "saved" },
  saveStateText: { type: String, default: "" },
  canPresent: { type: Boolean, default: false },
  otherFormatLabel: { type: String, default: "" },
});

const emit = defineEmits([
  "present",
  "download",
  "download-copy",
  "export-html",
  "export-pdf",
  "export-images",
  "close",
]);

const menuOpen = ref(false);
const presentMenu = useTemplateRef("presentMenu");
const presentCaret = useTemplateRef("presentCaret");

function menuItems() {
  return Array.from(presentMenu.value?.querySelectorAll('[role="menuitem"]') ?? []);
}

function focusItem(index = 0) {
  const items = menuItems();
  const item = items[(index + items.length) % Math.max(1, items.length)];
  if (item instanceof HTMLElement) item.focus();
}

function openMenu() {
  menuOpen.value = true;
  void nextTick(() => focusItem(0));
}

function closeMenu(refocus = false) {
  menuOpen.value = false;
  if (refocus) presentCaret.value?.$el?.focus();
}

function toggleMenu() {
  if (menuOpen.value) closeMenu();
  else openMenu();
}

function onCaretKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "ArrowDown") {
    event.preventDefault();
    openMenu();
  }
}

function onMenuKeydown(event = new KeyboardEvent("keydown")) {
  const items = menuItems();
  const at = items.indexOf(document.activeElement ?? document.body);
  if (event.key === "ArrowDown" || event.key === "ArrowUp") {
    focusItem(event.key === "ArrowDown" ? at + 1 : at - 1);
  } else if (event.key === "Home" || event.key === "End") {
    focusItem(event.key === "Home" ? 0 : items.length - 1);
  } else if (event.key === "Escape") {
    closeMenu(true);
  } else if (event.key === "Tab") {
    closeMenu();
    return;
  } else {
    return;
  }
  // Handled here: the editor's window-level shortcuts must not see it too.
  event.preventDefault();
  event.stopPropagation();
}

function present(mode = "show") {
  closeMenu();
  emit("present", mode);
}
</script>

<template>
  <div class="deck-actions">
    <div class="present-split">
      <UiButton
        variant="primary"
        data-testid="slides-present"
        :disabled="!canPresent"
        :title="canPresent ? 'Start slideshow' : 'Add a slide to present'"
        @click="present('show')"
      >
        Present
      </UiButton>
      <UiButton
        ref="presentCaret"
        variant="primary"
        class="present-caret"
        data-testid="presentation-present-menu"
        aria-haspopup="menu"
        :aria-expanded="menuOpen"
        aria-label="More ways to present"
        :disabled="!canPresent"
        @click="toggleMenu"
        @keydown="onCaretKeydown"
      >
        ▾
      </UiButton>
      <div
        v-if="menuOpen"
        ref="presentMenu"
        class="present-menu"
        role="menu"
        aria-label="Present"
        data-testid="presentation-present-menu-popup"
        @keydown="onMenuKeydown"
      >
        <button
          type="button"
          role="menuitem"
          class="present-menu-item"
          data-testid="presentation-present-show"
          @click="present('show')"
        >
          From current slide
        </button>
        <!-- The work order names this item presentation-presenter; the spec
             presentation-present-presenter — both resolve to it. -->
        <button
          type="button"
          role="menuitem"
          class="present-menu-item"
          data-testid="presentation-present-presenter"
          @click="present('presenter')"
        >
          <span data-testid="presentation-presenter">Presenter view</span>
        </button>
      </div>
      <div
        v-if="menuOpen"
        class="menu-backdrop"
        @click="closeMenu()"
      />
    </div>
    <PresentationDownloadButton
      :other-format-label="otherFormatLabel"
      @download="emit('download')"
      @download-copy="emit('download-copy')"
      @export-html="emit('export-html')"
      @export-pdf="emit('export-pdf')"
      @export-images="emit('export-images')"
    />
    <SaveStateBadge
      v-if="saveStateText"
      class="deck-save-state"
      testid="deck-save-state"
      :save-state="saveState"
      :text="saveStateText"
    />
    <!-- SPEC-103: the escape hatch when storage refuses the deck — the
         word processor's save-state-download, for slides. -->
    <a
      v-if="props.saveState === 'error'"
      class="deck-save-state-download"
      data-testid="deck-save-state-download"
      href="#"
      @click.prevent="emit('download')"
    >Download a copy</a>
    <UiIconButton
      class="close-button"
      icon="close"
      label="Close presentation"
      data-testid="presentation-close"
      @click="emit('close')"
    />
  </div>
</template>

<style scoped>
.deck-actions {
  display: flex;
  flex: none;
  gap: var(--space-3);
  align-items: center;
}

/* SPEC-100: the Present split button — PresentationDownloadButton's menu. */
.present-split {
  position: relative;
  display: flex;
  gap: var(--space-1);
  align-items: center;
}

.present-caret {
  padding-inline: var(--space-2);
}

.present-menu {
  position: absolute;
  right: 0;
  top: calc(100% + var(--space-1));
  z-index: 40;
  min-width: 14rem;
  padding: var(--space-1);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
}

.present-menu-item {
  display: block;
  width: 100%;
  text-align: left;
  padding: var(--space-2) var(--space-3);
  border: 0;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  cursor: pointer;
}

.present-menu-item:hover {
  background: var(--surface-sunken);
}

.present-menu-item:focus-visible {
  background: var(--surface-sunken);
  outline: var(--border-width) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--border-width));
}

.menu-backdrop {
  position: fixed;
  inset: 0;
  z-index: 39;
}

/* Quieter than the buttons around it: a state to glance at, not an action. */
.deck-save-state {
  color: var(--text-secondary);
}

.deck-save-state-download {
  font-size: var(--text-md);
  color: var(--text-on-brand);
  white-space: nowrap;
}

/* The icon button's default greys are tuned for light surfaces; on the brand
   bar it takes the on-brand foreground and a translucent wash so it reads in
   both themes — the same treatment EditorSaveStatus gives the word processor's
   Close, because it is now the same control. */
.deck-actions .close-button {
  color: var(--text-on-brand);
}

.deck-actions .close-button:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.deck-actions .close-button:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}
</style>

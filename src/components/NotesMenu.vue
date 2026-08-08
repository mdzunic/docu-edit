<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { shortcutLabel } from "../lib/shortcuts";

// The Insert → Footnote/Endnote menu (SPEC-031 AC-1, AC-3, AC-17..20):
// inserting either note kind and the numbering settings dialog start here —
// the same floating-menu shape HeaderFooterMenu uses.
const props = defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["insert-footnote", "insert-endnote", "settings"]);

const open = ref(false);
const root = useTemplateRef("root");
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value,
});

const footnoteShortcut = shortcutLabel("insert.footnote");
const endnoteShortcut = shortcutLabel("insert.endnote");

function toggleMenu() {
  if (!props.disabled) {
    open.value = !open.value;
  }
}

function chooseInsertFootnote() {
  open.value = false;
  emit("insert-footnote");
}

function chooseInsertEndnote() {
  open.value = false;
  emit("insert-endnote");
}

function chooseSettings() {
  open.value = false;
  emit("settings");
}

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  if (open.value && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    open.value = false;
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    open.value = false;
  }
}

document.addEventListener("mousedown", handleDocumentMousedown);
onBeforeUnmount(() => {
  document.removeEventListener("mousedown", handleDocumentMousedown);
});
</script>

<template>
  <div
    ref="root"
    class="notes-menu"
    @keydown="handleKeydown"
  >
    <button
      ref="button"
      type="button"
      class="toolbar-button"
      data-testid="notes-menu-button"
      title="Insert footnote/endnote"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggleMenu"
    >
      Notes ▾
    </button>
    <div
      v-if="open"
      ref="menu"
      class="notes-menu-popup"
      data-testid="notes-menu"
      role="menu"
      aria-label="Footnotes and endnotes"
      :style="menuStyle"
    >
      <button
        type="button"
        role="menuitem"
        data-testid="notes-menu-footnote"
        :title="`Insert footnote (${footnoteShortcut})`"
        @mousedown.prevent
        @click="chooseInsertFootnote"
      >
        Footnote
        <span class="menu-shortcut">{{ footnoteShortcut }}</span>
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="notes-menu-endnote"
        :title="`Insert endnote (${endnoteShortcut})`"
        @mousedown.prevent
        @click="chooseInsertEndnote"
      >
        Endnote
        <span class="menu-shortcut">{{ endnoteShortcut }}</span>
      </button>
      <hr class="notes-menu-separator">
      <button
        type="button"
        role="menuitem"
        data-testid="notes-menu-settings"
        @mousedown.prevent
        @click="chooseSettings"
      >
        Footnote settings…
      </button>
    </div>
  </div>
</template>

<style scoped>
.notes-menu {
  display: inline-flex;
}

.toolbar-button {
  border: 1px solid transparent;
  background: transparent;
  border-radius: 4px;
  padding: 0.25rem 0.5rem;
  font-size: 0.85rem;
  color: #3c4043;
  cursor: pointer;
  white-space: nowrap;
}

.toolbar-button:hover:not(:disabled) {
  background: #f1f3f4;
}

.toolbar-button:disabled {
  color: #9aa0a6;
  cursor: default;
}

.notes-menu-popup {
  z-index: 30;
  display: flex;
  flex-direction: column;
  min-width: 14rem;
  padding: 0.25rem 0;
  border: 1px solid #dadce0;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.2);
}

.notes-menu-popup [role="menuitem"] {
  display: flex;
  justify-content: space-between;
  gap: 1rem;
  border: 0;
  background: transparent;
  padding: 0.4rem 1rem;
  text-align: left;
  font-size: 0.85rem;
  color: #202124;
  cursor: pointer;
}

.notes-menu-popup [role="menuitem"]:hover:not(:disabled) {
  background: #f1f3f4;
}

.menu-shortcut {
  color: #80868b;
}

.notes-menu-separator {
  margin: 0.25rem 0;
  border: 0;
  border-top: 1px solid #e0e0e0;
}
</style>

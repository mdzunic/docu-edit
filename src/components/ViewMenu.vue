<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// The View menu (SPEC-037 AC-13): toggles the document outline pane. Same
// floating-menu shape as NotesMenu/HeaderFooterMenu. The "Document outline"
// item is a checkbox-style toggle whose checked state mirrors the pane's
// current open/closed state.
const props = defineProps({
  disabled: { type: Boolean, default: false },
  /** Whether the outline pane is currently open (drives the checkmark). */
  outlineOpen: { type: Boolean, default: false },
  /** Whether the comments pane is currently open (SPEC-038). */
  commentsPaneOpen: { type: Boolean, default: false },
  /** Whether the comments-only view is active (SPEC-038). */
  commentsOnlyView: { type: Boolean, default: false },
});

const emit = defineEmits(["toggle-outline", "toggle-comments-pane", "toggle-comments-only"]);

const open = ref(false);
const root = useTemplateRef("root");
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value,
});

function toggleMenu() {
  if (!props.disabled) {
    open.value = !open.value;
  }
}

function chooseToggleOutline() {
  open.value = false;
  emit("toggle-outline");
}

function chooseToggleCommentsPane() {
  open.value = false;
  emit("toggle-comments-pane");
}

function chooseToggleCommentsOnly() {
  open.value = false;
  emit("toggle-comments-only");
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
    class="view-menu"
    @keydown="handleKeydown"
  >
    <button
      ref="button"
      type="button"
      class="toolbar-button"
      data-testid="view-menu-button"
      title="View options"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggleMenu"
    >
      View ▾
    </button>
    <div
      v-if="open"
      ref="menu"
      class="view-menu-popup"
      data-testid="view-menu"
      role="menu"
      aria-label="View"
      :style="menuStyle"
    >
      <button
        type="button"
        role="menuitemcheckbox"
        data-testid="view-menu-outline"
        :aria-checked="outlineOpen ? 'true' : 'false'"
        @mousedown.prevent
        @click="chooseToggleOutline"
      >
        <span class="menu-check">{{ outlineOpen ? "✓" : "" }}</span>
        Document outline
      </button>
      <button
        type="button"
        role="menuitemcheckbox"
        data-testid="view-menu-comments-pane"
        :aria-checked="commentsPaneOpen ? 'true' : 'false'"
        @mousedown.prevent
        @click="chooseToggleCommentsPane"
      >
        <span class="menu-check">{{ commentsPaneOpen ? "✓" : "" }}</span>
        Comments pane
      </button>
      <button
        type="button"
        role="menuitemcheckbox"
        data-testid="view-menu-comments-only"
        :aria-checked="commentsOnlyView ? 'true' : 'false'"
        @mousedown.prevent
        @click="chooseToggleCommentsOnly"
      >
        <span class="menu-check">{{ commentsOnlyView ? "✓" : "" }}</span>
        Comments-only view
      </button>
    </div>
  </div>
</template>

<style scoped>
.view-menu {
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

.view-menu-popup {
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

.view-menu-popup [role="menuitemcheckbox"] {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  border: 0;
  background: transparent;
  padding: 0.4rem 1rem;
  text-align: left;
  font-size: 0.85rem;
  color: #202124;
  cursor: pointer;
}

.view-menu-popup [role="menuitemcheckbox"]:hover:not(:disabled) {
  background: #f1f3f4;
}

.menu-check {
  display: inline-block;
  width: 1rem;
  color: #1a73e8;
}
</style>

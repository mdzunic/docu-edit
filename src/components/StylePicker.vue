<script setup lang="ts">
import { computed, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { asStyleEntries, noStyleEntries } from "../lib/styles";
import UiMenuSurface from "./ui/UiMenuSurface.vue";

// The paragraph style picker (SPEC-019): a toolbar dropdown showing the
// current paragraph's style name (AC-1) whose menu lists the document's
// styles, each entry previewed in its own style (AC-2, AC-8). Right-clicking
// an entry opens its context menu with "Update <style> to match" (AC-12)
// and "Reset style" (AC-14). Runtime prop/emit declarations, not the
// defineProps<{...}>() generics — see AppShell.vue's comment on the same
// espree constraint.
const props = defineProps({
  /** The picker list (lib/styles.stylePickerEntries): id, name and the
   * preview's inline style per entry. */
  entries: { type: Array, default: () => noStyleEntries() },
  /** The current paragraph's resolved named style id ("" when none). */
  currentId: { type: String, default: "" },
  /** The current paragraph's style display name (AC-1). */
  currentName: { type: String, default: "Normal" },
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["apply", "update-style", "reset-style"]);

const styleList = computed(() => asStyleEntries(props.entries));

const open = ref(false);
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

// Fixed-position menu escaping the toolbar strip's clipping — the same
// pattern as ParagraphToolbar's spacing menu. The measured element is
// UiMenuSurface's rendered box ($el).
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value?.$el ?? null,
});

// The per-entry context menu (AC-12, AC-14), placed at the pointer.
const entryMenu = ref({ visible: false, id: "", name: "", left: 0, top: 0 });

function toggleMenu() {
  if (!props.disabled) {
    entryMenu.value = { ...entryMenu.value, visible: false };
    open.value = !open.value;
  }
}

function closeAll() {
  open.value = false;
  entryMenu.value = { ...entryMenu.value, visible: false };
}

function chooseStyle(styleId = "") {
  closeAll();
  emit("apply", styleId);
}

function openEntryMenu(styleId = "", name = "", event = new MouseEvent("contextmenu")) {
  entryMenu.value = {
    visible: true,
    id: styleId,
    name,
    left: event.clientX,
    top: event.clientY,
  };
}

function chooseUpdate() {
  const styleId = entryMenu.value.id;
  closeAll();
  emit("update-style", styleId);
}

function chooseReset() {
  closeAll();
  emit("reset-style");
}

function closeEntryMenu() {
  entryMenu.value = { ...entryMenu.value, visible: false };
}

// Outside-click and Escape dismissal (and focus return) are owned by
// UiMenuSurface — one instance per popup — so this component no longer
// hand-rolls a document listener or a root Escape handler.
</script>

<template>
  <div class="style-picker">
    <button
      ref="button"
      type="button"
      class="picker-button"
      data-testid="style-picker-button"
      :disabled="disabled"
      title="Paragraph styles"
      aria-label="Paragraph styles"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      @mousedown.prevent
      @click="toggleMenu"
    >
      <span
        class="current-name"
        data-testid="style-picker-current"
      >{{ currentName }}</span>
      <span aria-hidden="true">▾</span>
    </button>
    <UiMenuSurface
      ref="menu"
      :open="open"
      :anchor="button"
      class="style-menu"
      data-testid="style-picker-menu"
      role="menu"
      aria-label="Paragraph styles"
      :style="menuStyle"
      @dismiss="closeAll"
    >
      <button
        v-for="entry in styleList"
        :key="entry.id"
        type="button"
        role="menuitemradio"
        class="style-option"
        :aria-checked="entry.id === currentId || (currentId === '' && entry.id === 'Normal') ? 'true' : 'false'"
        :class="{ checked: entry.id === currentId || (currentId === '' && entry.id === 'Normal') }"
        :data-testid="`style-option-${entry.id}`"
        @mousedown.prevent
        @click="chooseStyle(entry.id)"
        @contextmenu.prevent="openEntryMenu(entry.id, entry.name, $event)"
      >
        <span :style="entry.previewStyle">{{ entry.name }}</span>
      </button>
    </UiMenuSurface>
    <UiMenuSurface
      :open="entryMenu.visible"
      class="entry-menu"
      data-testid="style-context-menu"
      role="menu"
      :aria-label="`Style actions for ${entryMenu.name}`"
      :style="{ left: `${entryMenu.left}px`, top: `${entryMenu.top}px` }"
      @dismiss="closeEntryMenu"
    >
      <button
        type="button"
        role="menuitem"
        data-testid="style-update"
        @mousedown.prevent
        @click="chooseUpdate"
      >
        Update {{ entryMenu.name }} to match
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="style-reset"
        @mousedown.prevent
        @click="chooseReset"
      >
        Reset style
      </button>
    </UiMenuSurface>
  </div>
</template>

<style scoped>
.style-picker {
  display: inline-flex;
  position: relative;
}

.picker-button {
  font: inherit;
  font-size: var(--text-md);
  height: var(--height-control);
  min-width: 6.5rem;
  padding: 0 var(--space-3);
  display: inline-flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.picker-button:disabled {
  opacity: 0.5;
  cursor: default;
}

.current-name {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 9rem;
}

/* position/left/top come inline from useFloatingMenu; the box is UiMenuSurface's. */
.style-menu {
  z-index: 30;
  min-width: 12rem;
  max-height: 60vh;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
}

.style-option {
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: left;
  background: none;
  border: none;
  padding: var(--space-2) var(--space-4);
  border-radius: var(--radius-sm);
  cursor: pointer;
  white-space: nowrap;
}

.style-option:hover {
  background: var(--state-hover);
}

.style-option.checked {
  background: var(--surface-selected);
}

/* position comes inline (pointer coords); the box is UiMenuSurface's. */
.entry-menu {
  position: fixed;
  z-index: 40;
  display: flex;
  flex-direction: column;
  min-width: 11rem;
}

.entry-menu button {
  font: inherit;
  font-size: var(--text-md);
  color: var(--text-primary);
  text-align: left;
  background: none;
  border: none;
  padding: var(--space-2) var(--space-4);
  border-radius: var(--radius-sm);
  cursor: pointer;
  white-space: nowrap;
}

.entry-menu button:hover {
  background: var(--state-hover);
}
</style>

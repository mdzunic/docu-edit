<script setup lang="ts">
import { computed, onBeforeUnmount, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";
import { asStyleEntries, noStyleEntries } from "../lib/styles";

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
const root = useTemplateRef("root");
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

// Fixed-position menu escaping the toolbar strip's clipping — the same
// pattern as ParagraphToolbar's spacing menu.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value,
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

function handleDocumentMousedown(event = new MouseEvent("mousedown")) {
  const anyOpen = open.value || entryMenu.value.visible;
  if (anyOpen && root.value && event.target instanceof Node && !root.value.contains(event.target)) {
    closeAll();
  }
}

function handleKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key === "Escape") {
    closeAll();
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
    class="style-picker"
    @keydown="handleKeydown"
  >
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
    <div
      v-if="open"
      ref="menu"
      class="style-menu"
      data-testid="style-picker-menu"
      role="menu"
      aria-label="Paragraph styles"
      :style="menuStyle"
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
    </div>
    <div
      v-if="entryMenu.visible"
      class="entry-menu"
      data-testid="style-context-menu"
      role="menu"
      :aria-label="`Style actions for ${entryMenu.name}`"
      :style="{ left: `${entryMenu.left}px`, top: `${entryMenu.top}px` }"
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
    </div>
  </div>
</template>

<style scoped>
.style-picker {
  display: inline-flex;
  position: relative;
}

.picker-button {
  font: inherit;
  font-size: 0.875rem;
  height: 1.75rem;
  min-width: 6.5rem;
  padding: 0 0.5rem;
  display: inline-flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.375rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
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

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   toolbar's vertical clipping), like ParagraphToolbar's menus. */
.style-menu {
  z-index: 30;
  min-width: 12rem;
  max-height: 60vh;
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  padding: 0.25rem 0;
  border: 1px solid #bdbdbd;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.style-option {
  font: inherit;
  font-size: 0.875rem;
  color: #202124;
  text-align: left;
  background: none;
  border: none;
  padding: 0.375rem 1rem;
  cursor: pointer;
  white-space: nowrap;
}

.style-option:hover {
  background: #f1f3f4;
}

.style-option.checked {
  background: #d2e3fc;
}

.entry-menu {
  position: fixed;
  z-index: 40;
  display: flex;
  flex-direction: column;
  min-width: 11rem;
  padding: 0.25rem 0;
  border: 1px solid #bdbdbd;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.entry-menu button {
  font: inherit;
  font-size: 0.875rem;
  color: #202124;
  text-align: left;
  background: none;
  border: none;
  padding: 0.375rem 1rem;
  cursor: pointer;
  white-space: nowrap;
}

.entry-menu button:hover {
  background: #f1f3f4;
}
</style>

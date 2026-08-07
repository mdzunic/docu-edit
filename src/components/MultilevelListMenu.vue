<script setup lang="ts">
// The "Multilevel list" toolbar control (SPEC-032 AC-1..4): a gallery menu
// of the four presets, plus "Number headings 1–3" (AC-13) and "Define new
// multilevel list…" (AC-5..12). Emits the chosen action only — the editor
// owns the commands. Runtime prop/emit declarations, not the generic form —
// this script is parsed by espree (see AppShell.vue).
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { GALLERY_PRESETS } from "../lib/listDefinitions";
import { useFloatingMenu } from "../composables/useFloatingMenu";

const props = defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["preset", "link-headings", "define"]);

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

function choosePreset(key = "") {
  open.value = false;
  emit("preset", key);
}

function chooseLinkHeadings() {
  open.value = false;
  emit("link-headings");
}

function chooseDefine() {
  open.value = false;
  emit("define");
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
    class="multilevel-cluster"
    @keydown="handleKeydown"
  >
    <button
      ref="button"
      type="button"
      class="toolbar-button"
      data-testid="multilevel-list-button"
      :disabled="disabled"
      title="Multilevel list"
      aria-label="Multilevel list"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      @mousedown.prevent
      @click="toggleMenu"
    >
      1.1≡▾
    </button>
    <div
      v-if="open"
      ref="menu"
      class="multilevel-menu"
      data-testid="multilevel-list-menu"
      role="menu"
      aria-label="Multilevel list"
      :style="menuStyle"
    >
      <button
        v-for="preset in GALLERY_PRESETS"
        :key="preset.key"
        type="button"
        role="menuitem"
        :data-testid="`multilevel-preset-${preset.key}`"
        @mousedown.prevent
        @click="choosePreset(preset.key)"
      >
        <span class="preset-label">{{ preset.label }}</span>
        <span class="preset-preview">{{ preset.preview }}</span>
      </button>
      <hr class="menu-separator">
      <button
        type="button"
        role="menuitem"
        data-testid="multilevel-link-headings"
        @mousedown.prevent
        @click="chooseLinkHeadings"
      >
        Number headings 1–3
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="multilevel-define-list"
        @mousedown.prevent
        @click="chooseDefine"
      >
        Define new multilevel list…
      </button>
    </div>
  </div>
</template>

<style scoped>
.multilevel-cluster {
  display: inline-flex;
}

.toolbar-button {
  font: inherit;
  font-size: 0.8125rem;
  height: 1.75rem;
  padding: 0 0.375rem;
  line-height: 1;
  display: inline-flex;
  align-items: center;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.toolbar-button:disabled {
  opacity: 0.5;
  cursor: default;
}

/* position/left/top come inline from useFloatingMenu (fixed, escapes the
   toolbar's vertical clipping). */
.multilevel-menu {
  z-index: 30;
  display: flex;
  flex-direction: column;
  min-width: 14rem;
  padding: 0.25rem 0;
  border: 1px solid #bdbdbd;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 2px 8px rgb(0 0 0 / 0.25);
}

.multilevel-menu button {
  font: inherit;
  font-size: 0.875rem;
  color: #202124;
  text-align: left;
  background: none;
  border: none;
  padding: 0.375rem 1rem;
  cursor: pointer;
  white-space: nowrap;
  display: flex;
  justify-content: space-between;
  gap: 1.5rem;
}

.multilevel-menu button:hover {
  background: #f1f3f4;
}

.preset-preview {
  color: #5f6368;
  font-size: 0.75rem;
}

.menu-separator {
  margin: 0.25rem 0;
  border: none;
  border-top: 1px solid #e0e0e0;
}
</style>

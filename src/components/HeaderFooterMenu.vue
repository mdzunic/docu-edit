<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// The Insert → Header/Footer menu (SPEC-029 AC-2, AC-14..19, AC-11, AC-12,
// AC-22, AC-23): entering header/footer editing, inserting a page number
// and the per-section header/footer options all start here — the same
// floating-menu shape InsertBreakMenu uses.
const props = defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["insert-header", "insert-footer", "insert-page-number", "options"]);

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

function chooseInsertHeader() {
  open.value = false;
  emit("insert-header");
}

function chooseInsertFooter() {
  open.value = false;
  emit("insert-footer");
}

function chooseInsertPageNumber() {
  open.value = false;
  emit("insert-page-number");
}

function chooseOptions() {
  open.value = false;
  emit("options");
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
    class="header-footer-menu"
    @keydown="handleKeydown"
  >
    <button
      ref="button"
      type="button"
      class="toolbar-button"
      data-testid="header-footer-menu-button"
      title="Insert header/footer"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggleMenu"
    >
      Header/Footer ▾
    </button>
    <div
      v-if="open"
      ref="menu"
      class="hf-menu"
      data-testid="header-footer-menu"
      role="menu"
      aria-label="Header/footer"
      :style="menuStyle"
    >
      <button
        type="button"
        role="menuitem"
        data-testid="hf-menu-header"
        @mousedown.prevent
        @click="chooseInsertHeader"
      >
        Header
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="hf-menu-footer"
        @mousedown.prevent
        @click="chooseInsertFooter"
      >
        Footer
      </button>
      <hr class="hf-menu-separator">
      <button
        type="button"
        role="menuitem"
        data-testid="hf-menu-page-number"
        @mousedown.prevent
        @click="chooseInsertPageNumber"
      >
        Page number…
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="hf-menu-options"
        @mousedown.prevent
        @click="chooseOptions"
      >
        Header/footer options…
      </button>
    </div>
  </div>
</template>

<style scoped>
.header-footer-menu {
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

.hf-menu {
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

.hf-menu [role="menuitem"] {
  border: 0;
  background: transparent;
  padding: 0.4rem 1rem;
  text-align: left;
  font-size: 0.85rem;
  color: #202124;
  cursor: pointer;
}

.hf-menu [role="menuitem"]:hover:not(:disabled) {
  background: #f1f3f4;
}

.hf-menu-separator {
  margin: 0.25rem 0;
  border: 0;
  border-top: 1px solid #e0e0e0;
}
</style>

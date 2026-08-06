<script setup lang="ts">
import { onBeforeUnmount, ref, useTemplateRef } from "vue";
import { shortcutLabel } from "../lib/shortcuts";
import { useFloatingMenu } from "../composables/useFloatingMenu";

// The Insert → Break menu (SPEC-028 AC-6..9): page break, column break
// (disabled until SPEC-030 gives it meaning) and the two section breaks.
// Runtime prop/emit declarations — see AppShell.vue's espree comment.
const props = defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["insert-break", "insert-section-break"]);

const open = ref(false);
const root = useTemplateRef("root");
const button = useTemplateRef("button");
const menu = useTemplateRef("menu");

// Fixed-position menu escaping the toolbar strip's overflow clipping — the
// same pattern as ParagraphToolbar/StylePicker. A plain absolute-positioned
// dropdown here would be clipped invisible by .toolbar's overflow-y.
const { menuStyle } = useFloatingMenu({
  open: () => open.value,
  anchor: () => button.value,
  menu: () => menu.value,
});

const pageBreakTitle = `Page break (${shortcutLabel("insert.pageBreak")})`;

function toggleMenu() {
  if (!props.disabled) {
    open.value = !open.value;
  }
}

function chooseBreak(kind = "") {
  open.value = false;
  emit("insert-break", kind);
}

function chooseSectionBreak(start = "") {
  open.value = false;
  emit("insert-section-break", start);
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
    class="insert-break"
    @keydown="handleKeydown"
  >
    <button
      ref="button"
      type="button"
      class="toolbar-button"
      data-testid="insert-break-button"
      title="Insert break"
      aria-haspopup="menu"
      :aria-expanded="open ? 'true' : 'false'"
      :disabled="disabled"
      @mousedown.prevent
      @click="toggleMenu"
    >
      Insert break ▾
    </button>
    <div
      v-if="open"
      ref="menu"
      class="break-menu"
      data-testid="insert-break-menu"
      role="menu"
      aria-label="Insert break"
      :style="menuStyle"
    >
      <button
        type="button"
        role="menuitem"
        data-testid="break-page"
        :title="pageBreakTitle"
        @mousedown.prevent
        @click="chooseBreak('page')"
      >
        Page break
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="break-column"
        disabled
        title="Requires columns (coming soon)"
      >
        Column break
      </button>
      <hr class="break-separator">
      <button
        type="button"
        role="menuitem"
        data-testid="break-section-next"
        @mousedown.prevent
        @click="chooseSectionBreak('nextPage')"
      >
        Section break (next page)
      </button>
      <button
        type="button"
        role="menuitem"
        data-testid="break-section-continuous"
        @mousedown.prevent
        @click="chooseSectionBreak('continuous')"
      >
        Section break (continuous)
      </button>
    </div>
  </div>
</template>

<style scoped>
.insert-break {
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

/* position/left/top come inline from useFloatingMenu's menuStyle. */
.break-menu {
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

.break-menu [role="menuitem"] {
  border: 0;
  background: transparent;
  padding: 0.4rem 1rem;
  text-align: left;
  font-size: 0.85rem;
  color: #202124;
  cursor: pointer;
}

.break-menu [role="menuitem"]:hover:not(:disabled) {
  background: #f1f3f4;
}

.break-menu [role="menuitem"]:disabled {
  color: #9aa0a6;
  cursor: default;
}

.break-separator {
  margin: 0.25rem 0;
  border: 0;
  border-top: 1px solid #e0e0e0;
}
</style>

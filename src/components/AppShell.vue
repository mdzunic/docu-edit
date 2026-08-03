<script setup lang="ts">
import { nextTick, ref, useTemplateRef } from "vue";
import StorageWarningBanner from "./StorageWarningBanner.vue";
import ToastHost from "./ToastHost.vue";

// Runtime prop/emit declarations (not the defineProps<{...}>()/defineEmits<{...}>()
// generics) because ESLint parses .vue scripts with espree, which rejects
// TS-only syntax (generics, type annotations, casts); vue-tsc still infers
// the types from them. useTemplateRef needs no generic for the same reason —
// it infers the element type from the matching template `ref` attribute.
const props = defineProps({
  /** Shown in the app bar next to the product name. */
  documentTitle: { type: String, required: true },
  /** Whether clicking the title enters an editable text field. */
  titleEditable: { type: Boolean, default: false },
  /** Shown as a badge next to the title, e.g. "DOCX". Omitted when unset. */
  formatBadge: { type: String, default: "" },
});

const emit = defineEmits(["update:documentTitle"]);

const isEditingTitle = ref(false);
const draftTitle = ref("");
const titleInput = useTemplateRef("titleInput");

async function startEditingTitle() {
  if (!props.titleEditable) {
    return;
  }
  draftTitle.value = props.documentTitle;
  isEditingTitle.value = true;
  await nextTick();
  titleInput.value?.focus();
  titleInput.value?.select();
}

function commitTitleEdit() {
  if (!isEditingTitle.value) {
    return;
  }
  isEditingTitle.value = false;
  emit("update:documentTitle", draftTitle.value);
}

function cancelTitleEdit() {
  isEditingTitle.value = false;
}
</script>

<template>
  <div class="shell">
    <header
      class="app-bar"
      data-testid="app-bar"
    >
      <span
        class="product-name"
        data-testid="product-name"
      >docu-edit</span>
      <h1
        v-if="!isEditingTitle"
        class="document-title"
        data-testid="document-title"
        :tabindex="titleEditable ? 0 : undefined"
        :role="titleEditable ? 'button' : undefined"
        @click="startEditingTitle"
        @keydown.enter="startEditingTitle"
      >
        {{ documentTitle }}
      </h1>
      <input
        v-else
        ref="titleInput"
        v-model="draftTitle"
        class="document-title document-title-input"
        data-testid="document-title-input"
        @keydown.enter="commitTitleEdit"
        @keydown.escape="cancelTitleEdit"
        @blur="commitTitleEdit"
      >
      <span
        v-if="formatBadge"
        class="format-badge"
        data-testid="format-badge"
      >{{ formatBadge }}</span>
      <div class="app-bar-actions">
        <slot name="app-bar-actions" />
      </div>
    </header>
    <StorageWarningBanner />
    <div
      class="toolbar"
      role="toolbar"
      aria-label="Document toolbar"
      data-testid="toolbar"
    >
      <slot name="toolbar" />
    </div>
    <main
      class="canvas"
      data-testid="canvas"
    >
      <slot />
    </main>
    <ToastHost />
  </div>
</template>

<style scoped>
/* The shell owns the viewport: it fills #app exactly and hides its own
   overflow, so the page never grows a scrollbar — only .canvas scrolls. */
.shell {
  height: 100%;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.app-bar {
  flex: none;
  display: flex;
  align-items: baseline;
  gap: 0.75rem;
  padding: 0.5rem 1rem;
  background: #2c3e50;
  color: #fff;
}

.product-name {
  font-weight: 600;
}

.document-title {
  margin: 0;
  font-size: 1rem;
  font-weight: 400;
}

h1.document-title[role="button"] {
  cursor: pointer;
}

.document-title-input {
  font: inherit;
  color: inherit;
  background: rgb(255 255 255 / 0.15);
  border: 1px solid rgb(255 255 255 / 0.5);
  border-radius: 2px;
  padding: 0 0.25rem;
}

/* Pushed to the far edge of the app bar; align-self keeps the actions
   vertically centred even though the bar's items align to baseline. */
.app-bar-actions {
  margin-left: auto;
  align-self: center;
}

.format-badge {
  font-size: 0.6875rem;
  font-weight: 600;
  letter-spacing: 0.03em;
  padding: 0.125rem 0.375rem;
  border-radius: 3px;
  background: rgb(255 255 255 / 0.15);
  color: inherit;
}

/* A single non-wrapping row: contents that overflow the viewport width are
   reached by scrolling the strip, never by wrapping or clipping. */
.toolbar {
  flex: none;
  display: flex;
  flex-wrap: nowrap;
  align-items: center;
  gap: 0.25rem;
  min-height: 2.5rem;
  padding: 0.25rem 1rem;
  overflow-x: auto;
  overflow-y: hidden;
  white-space: nowrap;
  background: #f5f5f5;
  border-bottom: 1px solid #d0d0d0;
  box-sizing: border-box;
}

/* min-height: 0 lets the canvas shrink below its content height, so tall
   content scrolls here instead of stretching the shell past the viewport. */
.canvas {
  flex: 1 1 0;
  min-height: 0;
  overflow-y: auto;
  background: #e8eaed;
}
</style>

<script setup lang="ts">
import { useTemplateRef } from "vue";
import { ACCEPTED_IMAGE_ACCEPT_ATTR } from "../lib/images";

// The "Insert image" toolbar control (SPEC-024 AC-1): a plain button that
// drives a hidden native file picker — no dropdown, unlike TableInsertControl,
// since a file picker needs no in-app grid. Runtime prop/emit declarations
// per the espree constraint (see AppShell.vue).
const props = defineProps({
  disabled: { type: Boolean, default: false },
});

const emit = defineEmits(["insert-file"]);

const input = useTemplateRef("input");

function openPicker() {
  if (!props.disabled) {
    input.value?.click();
  }
}

function handleChange(event = new Event("change")) {
  const target = event.target;
  const file = target instanceof HTMLInputElement ? target.files?.[0] : undefined;
  if (file) {
    emit("insert-file", file);
  }
  if (target instanceof HTMLInputElement) {
    // Clears the value so picking the same file twice in a row still fires
    // a change event.
    target.value = "";
  }
}
</script>

<template>
  <div class="image-control">
    <button
      type="button"
      class="control-button"
      data-testid="insert-image-button"
      title="Insert image"
      :disabled="disabled"
      @mousedown.prevent
      @click="openPicker"
    >
      Image
    </button>
    <input
      ref="input"
      type="file"
      class="file-input"
      data-testid="insert-image-file-input"
      :accept="ACCEPTED_IMAGE_ACCEPT_ATTR"
      @change="handleChange"
    >
  </div>
</template>

<style scoped>
.image-control {
  display: flex;
  align-items: center;
}

.control-button {
  font: inherit;
  font-size: 0.875rem;
  padding: 0.25rem 0.5rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
  white-space: nowrap;
}

.control-button:disabled {
  color: #9e9e9e;
  cursor: default;
}

.file-input {
  display: none;
}
</style>

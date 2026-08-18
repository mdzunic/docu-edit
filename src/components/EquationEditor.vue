<script setup lang="ts">
// SPEC-056 equation editor. A focused panel over an equation that turns
// keystrokes and palette clicks into operations on the pure edit engine
// (src/lib/math/edit.ts) and renders the result live (AC-4). It hosts the
// template palette (AC-5), captures Tab/Shift+Tab (AC-6/7), the arrow keys
// (AC-12) and Backspace (AC-13), and edits the alt text (AC-20). Every editing
// rule lives in the framework-free engine; this component is the thin driver.
// Runtime prop/emit declarations (not generics) — espree cannot parse TS in a
// .vue script (see MathBox.vue's note).
import { computed, onMounted, useTemplateRef } from "vue";
import { layoutRow } from "../lib/math/render";
import { CONSTRUCT_TEMPLATES } from "../lib/math/edit";
import { asMathNodes } from "../lib/documentModel";
import MathBox from "./MathBox.vue";

const props = defineProps({
  /** The current MathNode[] body of the equation being edited. */
  body: { type: Array, required: true },
  /** The regenerated linear form, shown as feedback. */
  linear: { type: String, default: "" },
  /** The equation's alt text (AC-20). */
  altText: { type: String, default: "" },
});

const emit = defineEmits(["op", "altText", "commit", "close"]);

const templates = CONSTRUCT_TEMPLATES;

const box = computed(() => layoutRow(asMathNodes(props.body)));

const surface = useTemplateRef("surface");
const altInput = useTemplateRef("altInput");
onMounted(() => surface.value?.focus());

function op(kind = "", extra = {}) {
  emit("op", { kind, ...extra });
}

function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.isComposing) return;
  const key = event.key;
  if (key === "Escape" || key === "Enter") {
    event.preventDefault();
    emit("commit");
    return;
  }
  if (key === "Tab") {
    event.preventDefault();
    op(event.shiftKey ? "shiftTab" : "tab");
    return;
  }
  if (key === "ArrowLeft") return finish(event, "left");
  if (key === "ArrowRight") return finish(event, "right");
  if (key === "ArrowUp") return finish(event, "up");
  if (key === "ArrowDown") return finish(event, "down");
  if (key === "Backspace") return finish(event, "backspace");
  // A single printable character (not a shortcut) is linear input (AC-4).
  if (key.length === 1 && !event.ctrlKey && !event.metaKey && !event.altKey) {
    event.preventDefault();
    op("type", { text: key });
  }
}

function finish(event = new KeyboardEvent("keydown"), kind = "") {
  event.preventDefault();
  op(kind);
}

function onAltInput() {
  emit("altText", altInput.value?.value ?? "");
}
</script>

<template>
  <div
    class="equation-editor"
    role="dialog"
    aria-label="Equation editor"
    data-testid="equation-editor"
  >
    <div
      class="equation-editor-palette"
      role="toolbar"
      aria-label="Equation templates"
    >
      <button
        v-for="t in templates"
        :key="t.id"
        type="button"
        class="equation-template"
        :data-testid="`equation-template-${t.id}`"
        :title="t.label"
        @mousedown.prevent
        @click="op('template', { id: t.id })"
      >
        {{ t.label }}
      </button>
    </div>

    <div
      ref="surface"
      class="equation-editor-surface"
      tabindex="0"
      data-testid="equation-editor-surface"
      @keydown="onKeydown"
    >
      <MathBox :box="box" />
      <span
        v-if="asMathNodes(body).length === 0"
        class="equation-editor-placeholder"
      >Type an equation…</span>
    </div>

    <div
      class="equation-editor-linear"
      data-testid="equation-editor-linear"
    >
      {{ linear }}
    </div>

    <label class="equation-editor-alt">
      <span>Alt text</span>
      <input
        ref="altInput"
        type="text"
        data-testid="equation-alt-text"
        :value="altText"
        placeholder="Describe this equation"
        @input="onAltInput"
      >
    </label>

    <div class="equation-editor-actions">
      <button
        type="button"
        data-testid="equation-editor-done"
        @click="emit('commit')"
      >
        Done
      </button>
    </div>
  </div>
</template>

<style scoped>
.equation-editor {
  display: inline-flex;
  flex-direction: column;
  gap: 0.5rem;
  padding: 0.6rem;
  border: 1px solid #b0b0b0;
  border-radius: 6px;
  background: #fff;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.15);
  font-family: system-ui, sans-serif;
}
.equation-editor-palette {
  display: flex;
  flex-wrap: wrap;
  gap: 0.25rem;
  max-width: 22rem;
}
.equation-template {
  font-size: 0.75rem;
  padding: 0.15rem 0.4rem;
  border: 1px solid #ccc;
  border-radius: 3px;
  background: #f7f7f7;
  cursor: pointer;
}
.equation-template:hover {
  background: #ececec;
}
.equation-editor-surface {
  min-width: 8rem;
  min-height: 2.2rem;
  padding: 0.4rem 0.6rem;
  border: 1px solid #d0d0d0;
  border-radius: 4px;
  outline: none;
  font-size: 1.4rem;
}
.equation-editor-surface:focus {
  border-color: #4a80f0;
  box-shadow: 0 0 0 2px rgba(74, 128, 240, 0.25);
}
.equation-editor-placeholder {
  color: #999;
  font-size: 0.9rem;
}
.equation-editor-linear {
  font-family: "Liberation Mono", monospace;
  font-size: 0.75rem;
  color: #555;
  word-break: break-all;
}
.equation-editor-alt {
  display: flex;
  flex-direction: column;
  gap: 0.15rem;
  font-size: 0.75rem;
  color: #444;
}
.equation-editor-alt input {
  padding: 0.2rem 0.35rem;
  border: 1px solid #ccc;
  border-radius: 3px;
}
.equation-editor-actions {
  display: flex;
  justify-content: flex-end;
}
.equation-editor-actions button {
  padding: 0.25rem 0.8rem;
  border: 1px solid #4a80f0;
  background: #4a80f0;
  color: #fff;
  border-radius: 4px;
  cursor: pointer;
}
</style>

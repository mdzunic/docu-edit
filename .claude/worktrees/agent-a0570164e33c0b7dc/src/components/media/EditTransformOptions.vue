<script setup lang="ts">
// The rotate tool's options bar (SPEC-112): rotate left/right (also
// Cmd/Ctrl+[ and Cmd/Ctrl+]), flip either way, and resize — target pixels
// or a percentage, aspect-locked by default, bilinear or nearest.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { MAX_EDIT_EDGE_PX } from "../../lib/core/imageEdits";
import { asImageEditor } from "../../views/media/useImageEditor";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const size = computed(() => e.value.commands.outputSize.value ?? { widthPx: 1, heightPx: 1 });
const percent = computed(() => {
  const base = e.value.commands.resizeBase.value;
  return base ? Math.round((size.value.widthPx / base.widthPx) * 100) : 100;
});

function numberOf(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement ? Number(target.value) : Number.NaN;
}

function checked(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement && target.checked;
}

function chooseResample(event = new Event("change")) {
  const target = event.target;
  e.value.commands.setResample(target instanceof HTMLSelectElement && target.value === "nearest" ? "nearest" : "bilinear");
}
</script>

<template>
  <div
    class="options-bar"
    role="group"
    aria-label="Rotate, flip and resize options"
  >
    <UiButton
      size="sm"
      data-testid="edit-rotate-left"
      title="Rotate left (Ctrl/Cmd+[)"
      @click="e.commands.rotate(-1)"
    >
      Rotate left
    </UiButton>
    <UiButton
      size="sm"
      data-testid="edit-rotate-right"
      title="Rotate right (Ctrl/Cmd+])"
      @click="e.commands.rotate(1)"
    >
      Rotate right
    </UiButton>
    <UiButton
      size="sm"
      data-testid="edit-flip-horizontal"
      @click="e.commands.flip('horizontal')"
    >
      Flip horizontal
    </UiButton>
    <UiButton
      size="sm"
      data-testid="edit-flip-vertical"
      @click="e.commands.flip('vertical')"
    >
      Flip vertical
    </UiButton>
    <span
      class="divider"
      aria-hidden="true"
    />
    <label class="field">
      <span>Width</span>
      <input
        type="number"
        class="number"
        data-testid="edit-resize-width"
        min="1"
        :max="MAX_EDIT_EDGE_PX"
        :value="size.widthPx"
        @change="e.commands.setResizeWidth(numberOf($event))"
      >
    </label>
    <label class="field">
      <span>Height</span>
      <input
        type="number"
        class="number"
        data-testid="edit-resize-height"
        min="1"
        :max="MAX_EDIT_EDGE_PX"
        :value="size.heightPx"
        @change="e.commands.setResizeHeight(numberOf($event))"
      >
    </label>
    <label class="field">
      <span>Scale %</span>
      <input
        type="number"
        class="number"
        data-testid="edit-resize-percent"
        min="1"
        max="1000"
        :value="percent"
        @change="e.commands.setResizePercent(numberOf($event))"
      >
    </label>
    <label class="field">
      <input
        type="checkbox"
        data-testid="edit-resize-lock"
        :checked="e.commands.resizeLocked.value"
        @change="e.commands.resizeLocked.value = checked($event)"
      >
      <span>Keep aspect ratio</span>
    </label>
    <label class="field">
      <span>Resampling</span>
      <select
        data-testid="edit-resize-resample"
        :value="e.commands.resample.value"
        @change="chooseResample"
      >
        <option value="bilinear">Bilinear</option>
        <option value="nearest">Nearest neighbour</option>
      </select>
    </label>
    <UiButton
      size="sm"
      data-testid="edit-resize-reset"
      @click="e.commands.clearResize()"
    >
      Reset size
    </UiButton>
  </div>
</template>

<style scoped>
.options-bar {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.field {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
}

.number {
  width: 5.5rem;
  font: inherit;
}

.divider {
  width: var(--border-width);
  align-self: stretch;
  background: var(--border-subtle);
}
</style>

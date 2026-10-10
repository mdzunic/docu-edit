<script setup lang="ts">
// The redact and annotate tools' options bar (SPEC-112): the redaction
// colour, or the annotation's shape, colour, stroke width, opacity and
// (for text) its words and size — then "Add" for a keyboard user, and the
// focusable list of what is on the image: arrow keys nudge the focused
// object 1 px (10 with Shift), Alt+arrow keys resize it, Delete removes it.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { ANNOTATION_KINDS, annotationNoun } from "../../lib/core/imageEdits";
import { asImageEditor } from "../../views/media/useImageEditor";
import UiButton from "../ui/UiButton.vue";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
  /** "redact" or "annotate". */
  mode: { type: String, required: true },
});

const e = computed(() => asImageEditor(props.editor));
const tool = computed(() => (props.mode === "redact" ? "redact" : "annotate"));
const o = computed(() => e.value.objects);
const shown = computed(() => o.value.objects.value.filter((entry) => entry.kind === tool.value));

function text(event = new Event("input")) {
  const target = event.target;
  return target instanceof HTMLInputElement || target instanceof HTMLSelectElement ? target.value : "";
}

function chooseShape(event = new Event("change")) {
  const shape = ANNOTATION_KINDS.find((kind) => kind === text(event));
  if (shape) {
    o.value.shape.value = shape;
  }
}

function clampNumber(event = new Event("change"), min = 0, max = 0) {
  const number = Number(text(event));
  return Number.isFinite(number) ? Math.min(max, Math.max(min, number)) : min;
}

function onObjectKey(index = 0, event = new KeyboardEvent("keydown")) {
  if (o.value.handleKey(index, event)) {
    event.preventDefault();
  }
}
</script>

<template>
  <div
    class="options-bar"
    role="group"
    :aria-label="tool === 'redact' ? 'Redact options' : 'Annotate options'"
  >
    <template v-if="tool === 'redact'">
      <label class="field">
        <span>Colour</span>
        <input
          type="color"
          data-testid="redact-color"
          :value="o.redactColor.value"
          @change="o.redactColor.value = text($event)"
        >
      </label>
      <UiButton
        size="sm"
        data-testid="redact-add"
        :disabled="o.full.value"
        @click="o.add('redact')"
      >
        Add redaction
      </UiButton>
      <span class="hint">Drag over the image to cover an area.</span>
    </template>
    <template v-else>
      <label class="field">
        <span>Shape</span>
        <select
          data-testid="annotate-shape"
          :value="o.shape.value"
          @change="chooseShape"
        >
          <option
            v-for="kind in ANNOTATION_KINDS"
            :key="kind"
            :value="kind"
          >
            {{ annotationNoun(kind) }}
          </option>
        </select>
      </label>
      <label class="field">
        <span>Colour</span>
        <input
          type="color"
          data-testid="annotate-color"
          :value="o.color.value"
          @change="o.color.value = text($event)"
        >
      </label>
      <label class="field">
        <span>Stroke</span>
        <input
          type="range"
          data-testid="annotate-stroke"
          min="1"
          max="40"
          :value="o.strokeWidth.value"
          @input="o.strokeWidth.value = clampNumber($event, 1, 40)"
        >
      </label>
      <input
        type="number"
        class="number"
        data-testid="annotate-stroke-number"
        aria-label="Stroke width in pixels"
        min="1"
        max="40"
        :value="o.strokeWidth.value"
        @change="o.strokeWidth.value = clampNumber($event, 1, 40)"
      >
      <label class="field">
        <span>Opacity</span>
        <input
          type="range"
          data-testid="annotate-opacity"
          min="5"
          max="100"
          :value="Math.round(o.opacity.value * 100)"
          @input="o.opacity.value = clampNumber($event, 5, 100) / 100"
        >
      </label>
      <input
        type="number"
        class="number"
        data-testid="annotate-opacity-number"
        aria-label="Opacity in percent"
        min="5"
        max="100"
        :value="Math.round(o.opacity.value * 100)"
        @change="o.opacity.value = clampNumber($event, 5, 100) / 100"
      >
      <template v-if="o.shape.value === 'text'">
        <label class="field">
          <span>Text</span>
          <input
            type="text"
            data-testid="annotate-text"
            :value="o.text.value"
            @input="o.text.value = text($event)"
          >
        </label>
        <label class="field">
          <span>Size</span>
          <input
            type="number"
            class="number"
            data-testid="annotate-font-size"
            min="4"
            max="400"
            :value="o.fontSize.value"
            @change="o.fontSize.value = clampNumber($event, 4, 400)"
          >
        </label>
      </template>
      <UiButton
        size="sm"
        data-testid="annotate-add"
        :disabled="o.full.value"
        @click="o.add('annotate')"
      >
        Add {{ annotationNoun(o.shape.value) }}
      </UiButton>
    </template>
    <ul
      class="object-list"
      data-testid="edit-objects"
      :aria-label="`${tool === 'redact' ? 'Redactions' : 'Annotations'} — arrow keys move, Alt+arrow keys resize, Delete removes`"
    >
      <li
        v-for="entry in shown"
        :key="entry.index"
        class="object"
        :class="{ 'is-selected': o.selected.value === entry.index }"
        tabindex="0"
        data-testid="edit-object"
        :aria-label="`${entry.label}${entry.detail ? ` ${entry.detail}` : ''}. Arrow keys move it; Alt+arrow keys resize it; Delete removes it.`"
        @focus="o.selected.value = entry.index"
        @keydown="onObjectKey(entry.index, $event)"
      >
        <span>{{ entry.label }}</span>
        <span class="detail">{{ entry.detail }}</span>
        <button
          type="button"
          class="remove"
          data-testid="edit-object-remove"
          :aria-label="`Delete ${entry.label}`"
          @click="o.remove(entry.index)"
        >
          ×
        </button>
      </li>
    </ul>
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
  width: 4.5rem;
  font: inherit;
}

.hint {
  color: var(--text-secondary);
}

.object-list {
  flex-basis: 100%;
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
  margin: 0;
  padding: 0;
  list-style: none;
}

.object {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-raised);
}

.object.is-selected {
  border-color: var(--accent);
  background: var(--surface-selected);
}

.object:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: var(--space-1);
}

.detail {
  color: var(--text-secondary);
  font-variant-numeric: tabular-nums;
}

.remove {
  font: inherit;
  border: none;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
}
</style>

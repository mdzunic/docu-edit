<script setup lang="ts">
// The data-validation input message (SPEC-130): `validation-input-message`,
// a role="tooltip" (id VALIDATION_INPUT_MESSAGE_ID — the active sheet-cell's
// aria-describedby while it shows) holding the rule's title in bold above its
// text, line breaks kept. SheetGrid places it in its validation layer from
// the active cell's rect (unscaled paper px, so it follows scroll and zoom):
// below-right of the cell, flipped left or up to stay inside the cell area —
// validationChrome's placeInputMessage, given the size measured here. It
// never takes the pointer. Runtime prop declarations per the espree
// constraint (see AppShell.vue).
import { computed, onMounted, ref, useTemplateRef, watch } from "vue";
import { placeInputMessage, VALIDATION_INPUT_MESSAGE_ID } from "../../views/spreadsheet/validationChrome";

const props = defineProps({
  /** The rule's input title (shown in bold). */
  title: { type: String, default: "" },
  /** The rule's input message ("\n" are line breaks). */
  text: { type: String, default: "" },
  /** The active cell's rect (left, top, width, height), paper px. */
  cellRect: { type: Object, required: true },
  /** The cell area's size (width, height), paper px. */
  bounds: { type: Object, required: true },
});

const box = useTemplateRef("box");
const size = ref({ width: 0, height: 0 });

/** The box's own size, unscaled (offset sizes ignore the zoom transform). */
function measure() {
  const el = box.value;
  if (el && (el.offsetWidth !== size.value.width || el.offsetHeight !== size.value.height)) {
    size.value = { width: el.offsetWidth, height: el.offsetHeight };
  }
}

onMounted(measure);
watch(() => [props.title, props.text], measure, { flush: "post" });

const position = computed(() => {
  const cell = props.cellRect;
  const area = props.bounds;
  return placeInputMessage(
    { left: Number(cell.left) || 0, top: Number(cell.top) || 0, width: Number(cell.width) || 0, height: Number(cell.height) || 0 },
    size.value,
    { width: Number(area.width) || 0, height: Number(area.height) || 0 },
  );
});
</script>

<template>
  <div
    :id="VALIDATION_INPUT_MESSAGE_ID"
    ref="box"
    class="validation-input-message"
    role="tooltip"
    data-testid="validation-input-message"
    :style="{ left: `${position.left}px`, top: `${position.top}px` }"
  >
    <div
      v-if="props.title"
      class="validation-input-title"
      data-testid="validation-input-title"
    >
      {{ props.title }}
    </div>
    <div
      v-if="props.text"
      class="validation-input-text"
      data-testid="validation-input-text"
    >
      {{ props.text }}
    </div>
  </div>
</template>

<style scoped>
.validation-input-message {
  position: absolute;
  box-sizing: border-box;
  max-width: 16rem;
  padding: var(--space-1) var(--space-2);
  pointer-events: none;
  color: var(--text-primary);
  background: var(--surface-raised);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  box-shadow: var(--shadow-2);
  font-family: var(--font-ui);
  font-size: var(--text-sm);
  white-space: pre-line;
  overflow-wrap: anywhere;
}

.validation-input-title {
  font-weight: var(--weight-bold);
}
</style>

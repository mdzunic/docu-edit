<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";
import ShapeStylePanel from "./ShapeStylePanel.vue";

// The Format → Drawing → Shape formatting… dialog (SPEC-073 AC-17): the
// second host of ShapeStylePanel, so every full-model picker is reachable
// without the contextual toolbar. Changes apply LIVE, exactly like
// ShapePositionDialog — each panel emission is one labelled undo entry in
// the owner; Close only dismisses. Runtime prop declarations per the espree
// constraint this file's siblings share.
defineProps({
  /** The selection's styling snapshot — ShapeStylePanel's `styling` prop. */
  styling: { type: Object, required: true },
  /** The document's theme (SPEC-077), forwarded to ShapeStylePanel. */
  theme: { type: Object, default: null },
});

const emit = defineEmits([
  "fill",
  "outline",
  "shadow",
  "style-preset",
  "preview-fill",
  "preview-outline",
  "pick-picture",
  "dismiss",
]);
</script>

<template>
  <BaseDialog
    title="Shape formatting"
    testid="shape-format-dialog"
    @dismiss="emit('dismiss')"
  >
    <ShapeStylePanel
      :styling="styling"
      :theme="theme"
      @fill="(value) => emit('fill', value)"
      @outline="(value) => emit('outline', value)"
      @shadow="(value) => emit('shadow', value)"
      @preset="(id) => emit('style-preset', id)"
      @preview-fill="(value) => emit('preview-fill', value)"
      @preview-outline="(value) => emit('preview-outline', value)"
      @pick-picture="emit('pick-picture')"
    />
    <template #actions>
      <UiButton
        data-testid="shape-format-close"
        @click="emit('dismiss')"
      >
        Close
      </UiButton>
    </template>
  </BaseDialog>
</template>

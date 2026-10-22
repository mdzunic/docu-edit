<script setup lang="ts">
import ContextualToolbar from "./ContextualToolbar.vue";
import ShapeToolbar from "./ShapeToolbar.vue";
import ShapeFormatDialog from "./ShapeFormatDialog.vue";

// The selected shape's floating toolbar plus the SPEC-073 Shape formatting…
// dialog, as one island so DocumentView keeps a single wiring call within
// its line budget (see DocumentView.size.spec.ts). `state` carries every
// display prop ShapeToolbar reads (useShapes.shapeToolbarState builds it);
// all events pass through $attrs to the toolbar AND the dialog, which emit
// the same event names for the same styling actions — one handler set in
// the owner, two hosts (AC-17). Runtime prop declarations per the espree
// constraint every .vue script here shares.
defineOptions({ inheritAttrs: false });

defineProps({
  /** shapeToolbarStyle(): { left, top, placement }. */
  box: { type: Object, required: true },
  /** shapeToolbarState(): ShapeToolbar's display props + `styling`. */
  state: { type: Object, required: true },
  /** Whether the Format → Drawing → Shape formatting… dialog is open. */
  formatDialogOpen: { type: Boolean, default: false },
});

const emit = defineEmits(["format-dismiss"]);
</script>

<template>
  <ContextualToolbar
    class="print-exclude"
    testid="shape-toolbar"
    label="Shape formatting"
    context="Shape"
    :left="box.left"
    :top="box.top"
    :lift="box.placement === 'above'"
  >
    <ShapeToolbar v-bind="{ ...state, ...$attrs }" />
  </ContextualToolbar>
  <ShapeFormatDialog
    v-if="formatDialogOpen"
    :styling="state.styling || {}"
    v-bind="$attrs"
    @dismiss="emit('format-dismiss')"
  />
</template>

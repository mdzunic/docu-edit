<script setup lang="ts">
// The Toolbar Redesign's control caption — the tiny uppercase word printed
// under a control's icon in the banded strip ("BOLD", "SPACING", "MORE").
//
// One component rather than a repeated rule because five different controls
// wear it: UiToggleButton, UiIconButton, DocumentToolbar's overflow button and
// ParagraphToolbar's own bare buttons. It is always rendered and hidden
// through --toolbar-caption-display, which the strip sets: the planner has to
// measure the captioned and un-captioned geometries against each other, so
// both must really lay out, and a caption must never disappear from the DOM
// between measure and paint.
//
// Decorative by construction: the control it sits in already carries the
// accessible name, so this is aria-hidden and the caption may be a shorter
// word than the label ("Strike" under "Strikethrough (⌘⇧X)").
defineProps({
  text: { type: String, required: true },
});
</script>

<template>
  <span
    class="ui-control-caption"
    aria-hidden="true"
  >{{ text }}</span>
</template>

<style scoped>
.ui-control-caption {
  display: var(--toolbar-caption-display, block);
  font-family: var(--font-caption);
  font-size: var(--text-caption);
  font-weight: var(--weight-caption);
  letter-spacing: var(--tracking-caption);
  line-height: var(--line-tight);
  text-transform: uppercase;
  white-space: nowrap;
}
</style>

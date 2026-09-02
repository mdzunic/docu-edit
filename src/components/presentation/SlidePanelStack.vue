<script setup lang="ts">
// The stack of below-stage editing panels (SPEC-091 AC-18): Notes (SPEC-083),
// Transitions (SPEC-085), Animations (SPEC-086) and Table properties
// (SPEC-091), each shown when its controller is open and a slide is present.
// Extracted from PresentationView so a new panel costs the view one component,
// not another v-if block. Runtime prop declarations per the espree constraint.
import NotesPanel from "./NotesPanel.vue";
import TransitionsPanel from "./TransitionsPanel.vue";
import AnimationsPanel from "./AnimationsPanel.vue";
import TablePropertiesPanel from "./TablePropertiesPanel.vue";
import MediaPanel from "./MediaPanel.vue";

defineProps({
  notes: { type: Object, required: true },
  transitions: { type: Object, required: true },
  animations: { type: Object, required: true },
  tables: { type: Object, required: true },
  media: { type: Object, required: true },
  slideNumber: { type: Number, default: 1 },
  hasSlide: { type: Boolean, default: false },
});
</script>

<template>
  <NotesPanel
    v-if="notes.open.value && hasSlide"
    :value="notes.notesText.value"
    :slide-number="slideNumber"
    @input="notes.onInput"
  />
  <TransitionsPanel
    v-if="transitions.open.value && hasSlide"
    :controller="transitions"
    :slide-number="slideNumber"
  />
  <AnimationsPanel
    v-if="animations.open.value && hasSlide"
    :controller="animations"
    :slide-number="slideNumber"
  />
  <TablePropertiesPanel
    v-if="tables.open.value && tables.hasTable.value && hasSlide"
    :controller="tables"
    :slide-number="slideNumber"
  />
  <MediaPanel
    v-if="media.open.value && hasSlide"
    :controller="media"
    :slide-number="slideNumber"
  />
</template>

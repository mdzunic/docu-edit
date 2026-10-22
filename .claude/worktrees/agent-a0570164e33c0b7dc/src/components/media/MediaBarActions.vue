<script setup lang="ts">
// The media viewer's app-bar actions (SPEC-107's Close, SPEC-112's Edit):
// lifted out of MediaView so the view stays a wiring hub within its line
// budget. "Edit" opens the image editor and is offered only for an image
// that is not already being edited; Close returns to the library.
// Runtime prop declarations — see AppShell.vue.
import UiButton from "../ui/UiButton.vue";
import UiIconButton from "../ui/UiIconButton.vue";

const props = defineProps({
  /** Whether the document is an image the editor can open. */
  canEdit: { type: Boolean, default: false },
  /** Whether the editor is open. */
  editing: { type: Boolean, default: false },
});

const emit = defineEmits(["edit", "close"]);
</script>

<template>
  <UiButton
    v-if="props.canEdit && !props.editing"
    class="media-edit-open"
    variant="ghost"
    data-testid="media-edit-open"
    @click="emit('edit')"
  >
    Edit
  </UiButton>
  <UiIconButton
    class="media-close"
    icon="close"
    label="Close document"
    data-testid="close-document"
    @click="emit('close')"
  />
</template>

<style scoped>
/* The controls' default greys are tuned for light surfaces; on the brand
   bar they take the on-brand foreground and a translucent wash (the same
   treatment the editor's Close control gets). */
.media-close,
.media-edit-open {
  color: var(--text-on-brand);
}

.media-close:not(:disabled):hover,
.media-edit-open:not(:disabled):hover {
  color: var(--text-on-brand);
  background: color-mix(in srgb, var(--text-on-brand) 12%, transparent);
}

.media-close:not(:disabled):active,
.media-edit-open:not(:disabled):active {
  background: color-mix(in srgb, var(--text-on-brand) 20%, transparent);
}

.media-edit-open {
  margin-right: var(--space-2);
}
</style>

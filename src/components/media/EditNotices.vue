<script setup lang="ts">
// The image editor's notices (SPEC-112): the persistent redaction warning
// (shown whenever the stack holds a redaction — a cover is not an erasure
// while the original is stored), the SVG and animated-GIF notes, the
// refusal (an image over the canvas ceiling, or one that cannot be read),
// the running action's progress (role="status") and the last outcome.
// Runtime prop declarations — see AppShell.vue.
import { computed } from "vue";
import { RouterLink } from "vue-router";
import { asImageEditor, GIF_NOTICE, REDACTION_NOTICE, SVG_NOTICE } from "../../views/media/useImageEditor";

const props = defineProps({
  /** The useImageEditor controller. */
  editor: { type: Object, required: true },
});

const e = computed(() => asImageEditor(props.editor));
</script>

<template>
  <div class="edit-notices">
    <p
      v-if="e.phase.value === 'refused'"
      class="notice is-error"
      data-testid="edit-refusal"
      role="alert"
    >
      {{ e.refusal.value }}
    </p>
    <p
      v-if="e.isAnimatedGif.value"
      class="notice"
      data-testid="edit-gif-notice"
      role="note"
    >
      {{ GIF_NOTICE }}
    </p>
    <p
      v-if="e.isSvg.value"
      class="notice"
      data-testid="edit-svg-notice"
      role="note"
    >
      {{ SVG_NOTICE }}
    </p>
    <p
      v-if="e.hasRedaction.value"
      class="notice is-warning"
      data-testid="edit-redaction-notice"
      role="note"
    >
      {{ REDACTION_NOTICE }}
    </p>
    <p
      class="notice is-status"
      data-testid="edit-progress"
      role="status"
      :hidden="!e.actions.progress.value"
    >
      {{ e.actions.progress.value }}
    </p>
    <p
      v-if="e.actions.failure.value"
      class="notice is-error"
      data-testid="edit-failure"
      role="alert"
    >
      {{ e.actions.failure.value }}
    </p>
    <p
      v-if="e.actions.notice.value"
      class="notice is-status"
      data-testid="edit-notice"
      role="status"
    >
      {{ e.actions.notice.value }}
      <RouterLink
        v-if="e.actions.copyId.value && e.actions.notice.value.startsWith('Saved a copy')"
        :to="`/media/${e.actions.copyId.value}`"
        data-testid="edit-copy-open"
      >
        Open the copy
      </RouterLink>
    </p>
  </div>
</template>

<style scoped>
.edit-notices {
  display: flex;
  flex-direction: column;
}

.notice {
  margin: 0;
  padding: var(--space-2) var(--space-4);
  border-bottom: var(--border-width) solid var(--border-subtle);
  background: var(--surface-base);
  color: var(--text-primary);
  font-size: var(--text-sm);
}

.notice[hidden] {
  display: none;
}

.notice.is-warning {
  background: color-mix(in srgb, var(--warning) 14%, var(--surface-base));
  font-weight: var(--weight-bold);
}

.notice.is-error {
  color: var(--text-danger);
}

.notice.is-status {
  color: var(--text-secondary);
}
</style>

<script setup lang="ts">
// SPEC-115: the plan step before there is a plan — the reply streaming in as
// plain text (so a slow local model shows progress; a reply is never rendered
// as HTML or Markdown), the one repair in progress, and the ends: the provider
// off, unreachable or declined (nothing was sent, and the model-free designer
// is on offer), or a reply that failed its one repair (the error naming the
// path, the raw reply in the designer-raw disclosure, and "Try again").
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import UiButton from "../../components/ui/UiButton.vue";
import { asDeckDesigner } from "./designerFields";

const props = defineProps({
  /** The useDeckDesigner() session. */
  controller: { type: Object, required: true },
});

const c = asDeckDesigner(props.controller);
</script>

<template>
  <div
    v-if="c.phase.value === 'streaming' || c.phase.value === 'repairing'"
    class="reply"
  >
    <p
      class="reply-status"
      data-testid="designer-progress"
      aria-live="polite"
    >
      {{ c.phase.value === "repairing" ? "The reply did not check out, so it was sent back once to be corrected…" : "Waiting for the model…" }}
    </p>
    <pre
      class="reply-stream"
      data-testid="designer-stream"
      tabindex="0"
      aria-label="The reply so far"
    >{{ c.streamText.value }}</pre>
  </div>
  <div
    v-else-if="c.phase.value === 'failed'"
    class="reply"
  >
    <p
      v-if="c.offline.value"
      class="reply-offline"
      data-testid="designer-offline"
      role="status"
    >
      {{ c.offline.value }}
    </p>
    <template v-else>
      <p
        class="reply-error"
        data-testid="designer-error"
        role="alert"
      >
        {{ c.error.value }}
      </p>
      <details
        v-if="c.raw.value"
        class="reply-raw"
        data-testid="designer-raw"
      >
        <summary>The model's raw reply</summary>
        <pre>{{ c.raw.value }}</pre>
      </details>
    </template>
    <div class="reply-actions">
      <UiButton
        v-if="!c.offline.value"
        data-testid="designer-try-again"
        @click="c.tryAgain()"
      >
        Try again
      </UiButton>
      <UiButton
        v-if="c.mode === 'new'"
        data-testid="designer-outline-offer"
        :disabled="!c.canOutline.value"
        :title="c.canOutline.value ? undefined : 'Go back and choose a source to outline it without a model'"
        @click="c.outline()"
      >
        Outline this document
      </UiButton>
      <UiButton
        v-else
        data-testid="designer-by-hand-offer"
        @click="c.outline()"
      >
        Rearrange by hand
      </UiButton>
    </div>
  </div>
</template>

<style scoped>
.reply {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.reply-status,
.reply-offline {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.reply-error {
  margin: 0;
  color: var(--text-danger);
  font-size: var(--text-sm);
  overflow-wrap: anywhere;
}

.reply-stream,
.reply-raw pre {
  max-height: 16rem;
  margin: 0;
  overflow: auto;
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--text-primary);
  font-family: var(--font-mono, monospace);
  font-size: var(--text-xs);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.reply-actions {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
}
</style>

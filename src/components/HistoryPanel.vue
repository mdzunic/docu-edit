<script setup lang="ts">
// The version history panel both editing products share (SPEC-138, extracted
// from SPEC-103's DeckHistoryPanel): newest first, each version with its time,
// author (the preferences' authorName), the product's own counts and change
// (the controller's `statLabel`) and what took it; the actions Preview,
// Restore, Download (the controller's `downloadLabels`), Name/Remove name and
// Delete; a named-only filter; the restore note; the paused notice; and the
// preview banner, whose body a product fills through the `preview` slot (the
// deck its own read-only stage; a workbook previews in its grid instead). The
// root's testid is a prop — `slides-history-panel` (DeckHistoryPanel),
// `sheet-history-panel` (WorkbookHistoryPanel) — and every `history-*` testid
// is SPEC-103's, unchanged. While the controller is `readOnly` (another tab
// holds the document) the actions that write are disabled.
//
// SPEC-085/086's panel contract: all logic lives in the controller passed in,
// this component emits nothing, and its typed transforms live in
// historyPanelRender.ts. Runtime prop declarations, not defineProps<{...}>()
// generics (espree lints .vue scripts).
import { computed } from "vue";
import UiButton from "./ui/UiButton.vue";
import { asEntries, asHistoryController, triggerLabel, versionTime } from "./historyPanelRender";

const props = defineProps({
  /** The history controller (useDeckHistory, useWorkbookHistory). */
  controller: { type: Object, required: true },
  /** The root's data-testid. */
  testid: { type: String, required: true },
  /** The region's accessible name. */
  label: { type: String, default: "Version history" },
  /** What a restore's undo brings back, said under the title. */
  restoreNote: { type: String, default: "" },
  /** The preview banner's sentence. */
  previewLabel: { type: String, default: "Previewing an earlier version, read-only" },
  /** The banner's way out. */
  exitLabel: { type: String, default: "Back to current" },
});

const c = asHistoryController(props.controller);
const entries = computed(() => asEntries(c.visibleVersions.value));
const readOnly = computed(() => c.readOnly?.value === true);
const notice = computed(() => c.notice?.value ?? "");

function onNamedOnly(event = new Event("change")) {
  const el = event.target;
  if (el instanceof HTMLInputElement) c.namedOnly.value = el.checked;
}

function onName(versionId = "", event = new Event("submit")) {
  const form = event.currentTarget;
  const input = form instanceof HTMLFormElement ? form.elements.namedItem("name") : null;
  if (input instanceof HTMLInputElement && !readOnly.value) void c.nameVersion(versionId, input.value);
}
</script>

<template>
  <section
    class="history-panel"
    :data-testid="props.testid"
    :aria-label="props.label"
  >
    <header class="history-head">
      <h2 class="history-title">
        Version history
      </h2>
      <label class="history-filter">
        <input
          type="checkbox"
          data-testid="history-named-only"
          :checked="c.namedOnly.value"
          @change="onNamedOnly"
        >
        Named versions only
      </label>
    </header>
    <p
      v-if="props.restoreNote"
      class="history-note"
    >
      {{ props.restoreNote }}
    </p>
    <p
      v-if="c.pausedMessage.value"
      class="history-paused"
      role="status"
      data-testid="history-paused-notice"
    >
      {{ c.pausedMessage.value }}
    </p>
    <p
      v-if="notice"
      class="history-paused"
      role="alert"
      data-testid="history-notice"
    >
      {{ notice }}
    </p>

    <div
      v-if="c.previewModel.value"
      class="history-preview"
      data-testid="history-preview-banner"
    >
      <div class="history-preview-head">
        <span>{{ props.previewLabel }}</span>
        <UiButton
          size="sm"
          data-testid="history-exit-preview"
          @click="c.exitPreview()"
        >
          {{ props.exitLabel }}
        </UiButton>
      </div>
      <slot name="preview" />
    </div>

    <p
      v-if="entries.length === 0"
      class="history-empty"
      data-testid="history-empty"
    >
      No versions yet
    </p>
    <ol
      v-else
      class="history-list"
      data-testid="history-list"
    >
      <li
        v-for="entry in entries"
        :key="entry.id"
        class="history-version"
        :class="{ 'is-previewing': entry.id === c.previewVersionId.value }"
        data-testid="history-version"
      >
        <div class="history-meta">
          <span>{{ versionTime(entry) }}</span>
          <span class="history-author">{{ entry.author }}</span>
        </div>
        <div
          class="history-name"
          data-testid="history-name"
        >
          {{ entry.name || "Unnamed version" }}
        </div>
        <div class="history-facts">
          {{ c.statLabel(entry) }}<template v-if="triggerLabel(entry)">
            · <span data-testid="history-trigger">{{ triggerLabel(entry) }}</span>
          </template>
        </div>
        <div class="history-actions">
          <UiButton
            size="sm"
            data-testid="history-preview"
            @click="c.preview(entry.id)"
          >
            Preview
          </UiButton>
          <UiButton
            size="sm"
            data-testid="history-restore"
            :disabled="readOnly"
            @click="c.restore(entry.id)"
          >
            Restore
          </UiButton>
          <UiButton
            size="sm"
            data-testid="history-download"
            :title="c.downloadLabels.value.title"
            @click="c.downloadVersion(entry.id)"
          >
            {{ c.downloadLabels.value.button }}
          </UiButton>
          <form
            class="history-name-form"
            @submit.prevent="(event) => onName(entry.id, event)"
          >
            <input
              name="name"
              type="text"
              placeholder="Name this version"
              aria-label="Version name"
              data-testid="history-name-input"
              :value="entry.name || ''"
              :disabled="readOnly"
            >
            <UiButton
              size="sm"
              type="submit"
              data-testid="history-name-submit"
              :disabled="readOnly"
            >
              Name
            </UiButton>
          </form>
          <UiButton
            v-if="entry.name"
            size="sm"
            data-testid="history-remove-name"
            :disabled="readOnly"
            @click="c.removeName(entry.id)"
          >
            Remove name
          </UiButton>
          <UiButton
            size="sm"
            variant="ghost"
            data-testid="history-delete"
            :disabled="readOnly"
            @click="c.deleteVersion(entry.id)"
          >
            Delete
          </UiButton>
        </div>
      </li>
    </ol>
  </section>
</template>

<style scoped>
.history-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  background: var(--surface-default);
  color: var(--text-primary);
  font-size: var(--text-sm);
  overflow-y: auto;
}

.history-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.history-title {
  margin: 0;
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.history-filter {
  display: inline-flex;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-secondary);
}

.history-note {
  margin: 0;
  color: var(--text-secondary);
}

.history-paused {
  margin: 0;
  padding: var(--space-2) var(--space-3);
  border: var(--border-width) solid var(--warning);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--warning);
}

.history-preview {
  display: flex;
  flex-direction: column;
  align-items: flex-start;
  gap: var(--space-2);
  padding: var(--space-3);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-sm);
  background: var(--surface-selected);
}

.history-preview-head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  align-self: stretch;
  color: var(--text-accent);
}

.history-empty {
  margin: 0;
  font-style: italic;
  color: var(--text-secondary);
}

.history-list {
  list-style: none;
  margin: 0;
  padding: 0;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.history-version {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  background: var(--surface-base);
}

.history-version.is-previewing {
  border-color: var(--accent);
  background: var(--surface-selected);
}

.history-meta {
  display: flex;
  justify-content: space-between;
  gap: var(--space-3);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.history-author,
.history-name {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.history-facts {
  color: var(--text-secondary);
}

.history-actions {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  gap: var(--space-2);
  margin-top: var(--space-1);
}

.history-name-form {
  display: inline-flex;
  gap: var(--space-1);
}

.history-name-form input {
  width: 9rem;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-size: var(--text-xs);
}
</style>

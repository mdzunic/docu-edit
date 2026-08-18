<script setup lang="ts">
import { computed } from "vue";
import { formatTimestamp } from "../lib/format";
import { asVersionListEntries } from "../views/document/useVersionHistory";

// The version history pane (SPEC-040): a newest-first list of versions with
// timestamp, author and word-count delta (AC-7..10), a named-only filter
// (AC-18), and per-version actions — select for read-only diff preview
// (AC-11..13), restore (AC-14), name/remove name (AC-17) and download
// (AC-19, AC-20). A simpler docked list than CommentsPane's margin-anchored
// cards, the same shape as TrackChangesPane.
//
// Runtime prop/emit declarations, not the generic form — espree parses .vue
// scripts and rejects TS-only syntax (see TrackChangesPane.vue).
const props = defineProps({
  versions: { type: Array, default: () => [] },
  pausedMessage: { type: String, default: "" },
  namedOnlyFilter: { type: Boolean, default: false },
  previewVersionId: { type: String, default: "" },
  editable: { type: Boolean, default: true },
});

const emit = defineEmits([
  "close",
  "update:named-only-filter",
  "select",
  "exit-preview",
  "restore",
  "name",
  "remove-name",
  "delete",
  "download",
]);

const versionsTyped = computed(() => asVersionListEntries(props.versions));

function onFilterChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("update:named-only-filter", target.checked);
  }
}

function onNameSubmit(versionId = "", event = new Event("submit")) {
  event.preventDefault();
  const form = event.target;
  if (!(form instanceof HTMLFormElement)) {
    return;
  }
  const input = form.elements.namedItem("name");
  if (input instanceof HTMLInputElement) {
    emit("name", versionId, input.value);
  }
}

function wordDeltaLabel(delta = 0) {
  if (delta > 0) {
    return `+${delta} words`;
  }
  if (delta < 0) {
    return `${delta} words`;
  }
  return "no word change";
}
</script>

<template>
  <aside
    class="version-history-pane"
    aria-label="Version history"
    data-testid="version-history-pane"
  >
    <div class="version-history-pane-header">
      <span class="version-history-pane-title">Version history</span>
      <button
        type="button"
        class="version-history-close"
        data-testid="version-history-pane-close"
        aria-label="Close version history pane"
        @click="emit('close')"
      >
        ×
      </button>
    </div>

    <div
      v-if="pausedMessage"
      class="version-history-paused"
      data-testid="version-history-paused-notice"
    >
      {{ pausedMessage }}
    </div>

    <div
      v-if="previewVersionId"
      class="version-history-preview-banner"
      data-testid="version-history-preview-banner"
    >
      <span>Previewing an earlier version, read-only</span>
      <button
        type="button"
        data-testid="version-history-exit-preview"
        @click="emit('exit-preview')"
      >
        Back to current
      </button>
    </div>

    <label class="version-history-filter">
      <input
        type="checkbox"
        data-testid="version-history-named-only"
        :checked="namedOnlyFilter"
        @change="onFilterChange"
      >
      Named versions only
    </label>

    <p
      v-if="versionsTyped.length === 0"
      class="version-history-empty"
      data-testid="version-history-empty"
    >
      No versions yet
    </p>
    <ul
      v-else
      class="version-history-list"
    >
      <li
        v-for="entry in versionsTyped"
        :key="entry.id"
        class="version-history-item"
        :class="{ 'version-history-item-active': entry.id === previewVersionId }"
        :data-testid="`version-history-item-${entry.id}`"
        @click="emit('select', entry.id)"
      >
        <div class="version-history-item-header">
          <span class="version-history-timestamp">{{ formatTimestamp(new Date(entry.timestamp)) }}</span>
          <span class="version-history-author">{{ entry.author }}</span>
        </div>
        <div
          class="version-history-name"
          data-testid="version-history-name"
        >
          {{ entry.name || "Unnamed version" }}
        </div>
        <div class="version-history-delta">
          {{ wordDeltaLabel(entry.wordDelta) }}
        </div>
        <div
          v-if="editable"
          class="version-history-actions"
          @click.stop
        >
          <button
            type="button"
            data-testid="version-history-restore"
            @click="emit('restore', entry.id)"
          >
            Restore this version
          </button>
          <button
            type="button"
            data-testid="version-history-download-docx"
            @click="emit('download', entry.id, 'docx')"
          >
            Download .docx
          </button>
          <button
            type="button"
            data-testid="version-history-download-odt"
            @click="emit('download', entry.id, 'odt')"
          >
            Download .odt
          </button>
          <form
            class="version-history-name-form"
            @submit="(event) => onNameSubmit(entry.id, event)"
          >
            <input
              name="name"
              type="text"
              placeholder="Name this version"
              data-testid="version-history-name-input"
              :value="entry.name || ''"
            >
            <button
              type="submit"
              data-testid="version-history-name-submit"
            >
              Name
            </button>
          </form>
          <button
            v-if="entry.name"
            type="button"
            data-testid="version-history-remove-name"
            @click="emit('remove-name', entry.id)"
          >
            Remove name
          </button>
          <button
            type="button"
            data-testid="version-history-delete"
            @click="emit('delete', entry.id)"
          >
            Delete version
          </button>
        </div>
      </li>
    </ul>
  </aside>
</template>

<style scoped>
.version-history-pane {
  flex: none;
  align-self: flex-start;
  position: sticky;
  top: 0;
  width: 20rem;
  max-height: 100vh;
  display: flex;
  flex-direction: column;
  min-height: 0;
  overflow-y: auto;
  background: var(--surface-raised);
  border-left: var(--border-width) solid var(--border-default);
  box-sizing: border-box;
  font-size: var(--text-md);
}

.version-history-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  padding: var(--space-3) var(--space-4);
}

.version-history-pane-title {
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.version-history-close {
  border: none;
  background: transparent;
  font-size: var(--text-lg);
  line-height: 1;
  cursor: pointer;
  color: var(--text-secondary);
}

.version-history-paused {
  margin: 0 var(--space-4) var(--space-3);
  padding: var(--space-3) var(--space-3);
  background: var(--surface-sunken);
  border: var(--border-width) solid var(--warning);
  border-radius: var(--radius-sm);
  color: var(--warning);
  font-size: var(--text-sm);
}

.version-history-preview-banner {
  margin: 0 var(--space-4) var(--space-3);
  padding: var(--space-3) var(--space-3);
  background: var(--surface-selected);
  border: var(--border-width) solid var(--accent);
  border-radius: var(--radius-sm);
  color: var(--text-accent);
  font-size: var(--text-sm);
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
}

.version-history-filter {
  flex: none;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  padding: 0 var(--space-4) var(--space-3);
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.version-history-empty {
  margin: 0;
  padding: var(--space-2) var(--space-4) var(--space-4);
  font-style: italic;
  color: var(--text-secondary);
}

.version-history-list {
  list-style: none;
  margin: 0;
  padding: 0 var(--space-3) var(--space-3);
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.version-history-item {
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  padding: var(--space-3);
  cursor: pointer;
}

.version-history-item-active {
  border-color: var(--accent);
  background: var(--surface-selected);
}

.version-history-item-header {
  display: flex;
  justify-content: space-between;
  gap: var(--space-3);
  font-size: var(--text-xs);
  color: var(--text-secondary);
  margin-bottom: var(--space-2);
}

.version-history-author {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.version-history-name {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.version-history-delta {
  color: var(--text-secondary);
  font-style: italic;
}

.version-history-actions {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
  margin-top: var(--space-3);
}

.version-history-actions button {
  border: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  font-size: var(--text-xs);
  cursor: pointer;
}

.version-history-name-form {
  display: flex;
  gap: var(--space-2);
}

.version-history-name-form input {
  font-size: var(--text-xs);
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  width: 8rem;
}
</style>

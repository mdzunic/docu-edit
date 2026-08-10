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
  background: #f8f9fa;
  border-left: 1px solid #dadce0;
  box-sizing: border-box;
  font-size: 0.8125rem;
}

.version-history-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
  padding: 0.5rem 0.75rem;
}

.version-history-pane-title {
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}

.version-history-close {
  border: none;
  background: transparent;
  font-size: 1rem;
  line-height: 1;
  cursor: pointer;
  color: #5f6368;
}

.version-history-paused {
  margin: 0 0.75rem 0.5rem;
  padding: 0.4rem 0.5rem;
  background: #fef7e0;
  border: 1px solid #f9ab00;
  border-radius: 4px;
  color: #7c4b00;
  font-size: 0.75rem;
}

.version-history-preview-banner {
  margin: 0 0.75rem 0.5rem;
  padding: 0.4rem 0.5rem;
  background: #e8f0fe;
  border: 1px solid #1a73e8;
  border-radius: 4px;
  color: #174ea6;
  font-size: 0.75rem;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
}

.version-history-filter {
  flex: none;
  display: flex;
  align-items: center;
  gap: 0.35rem;
  padding: 0 0.75rem 0.5rem;
  font-size: 0.75rem;
  color: #5f6368;
}

.version-history-empty {
  margin: 0;
  padding: 0.25rem 0.75rem 0.75rem;
  font-style: italic;
  color: #5f6368;
}

.version-history-list {
  list-style: none;
  margin: 0;
  padding: 0 0.5rem 0.5rem;
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
}

.version-history-item {
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 0.5rem;
  cursor: pointer;
}

.version-history-item-active {
  border-color: #1a73e8;
  background: #e8f0fe;
}

.version-history-item-header {
  display: flex;
  justify-content: space-between;
  gap: 0.5rem;
  font-size: 0.7rem;
  color: #5f6368;
  margin-bottom: 0.2rem;
}

.version-history-author {
  font-weight: 600;
  color: #202124;
}

.version-history-name {
  font-weight: 600;
  color: #202124;
}

.version-history-delta {
  color: #5f6368;
  font-style: italic;
}

.version-history-actions {
  display: flex;
  flex-wrap: wrap;
  gap: 0.35rem;
  margin-top: 0.4rem;
}

.version-history-actions button {
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.15rem 0.4rem;
  font-size: 0.68rem;
  cursor: pointer;
}

.version-history-name-form {
  display: flex;
  gap: 0.25rem;
}

.version-history-name-form input {
  font-size: 0.68rem;
  padding: 0.1rem 0.3rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
  width: 8rem;
}
</style>

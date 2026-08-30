<script setup lang="ts">
// The library's quick-look panel (the modernist handoff's right column):
// a live preview of the focused or single-selected entry, its key facts,
// and the primary open action with the row actions beside it. Rendered
// only at the full tier. Product-neutral: it emits intents (`open`,
// `action` with the same payloads the row overflow menu produces, plus
// `download`) and never knows what opening a document means.
// `lib` is the LibraryController from useLibrary (runtime-typed Object —
// the .vue ESLint parser bars the PropType syntax; see useLibrary.ts).
import { computed } from "vue";
import LibraryThumbnail from "../../components/library/LibraryThumbnail.vue";
import { formatBadgeLabel, formatFileSize, formatRelativeTime } from "../../lib/format";

const props = defineProps({
  lib: { type: Object, required: true },
});

const emit = defineEmits(["open", "action", "download"]);

// More than one selected entry previews nothing — the pane says so instead.
const selectionCount = computed(() => props.lib.selectedIds.value.length);

// A single selected entry wins (an explicit choice); otherwise the entry the
// keyboard focus sits on. Entries hidden by the current filters preview
// nothing — the panel never shows a row the list does not.
const entry = computed(() => {
  const ids = props.lib.selectedIds.value;
  const id = ids.length === 1 ? ids[0] : String(props.lib.focusedId.value || "");
  if (!id) {
    return null;
  }
  for (const candidate of props.lib.visible.value) {
    if (candidate.id === id) {
      return candidate;
    }
  }
  return null;
});

const facts = computed(() => {
  if (!entry.value) {
    return [];
  }
  return [
    { label: "Modified", value: formatRelativeTime(entry.value.lastModified, props.lib.now.value) },
    { label: "Size", value: formatFileSize(props.lib.sizeOf(entry.value.id)) },
    { label: "Format", value: formatBadgeLabel(entry.value.format) },
  ];
});
</script>

<template>
  <aside
    class="quick-look"
    data-testid="quick-look"
    aria-label="Quick look"
  >
    <h3 class="quick-look-heading">
      Quick look
    </h3>

    <div
      v-if="selectionCount > 1"
      class="quick-look-empty"
      data-testid="quick-look-multiple"
    >
      {{ selectionCount }} documents selected. A preview shows for a single
      document — the bar above the list holds the actions that apply to all
      of them.
    </div>

    <div
      v-else-if="!entry"
      class="quick-look-empty"
      data-testid="quick-look-empty"
    >
      Click or focus a document to preview it here. Nothing opens until you
      ask it to.
    </div>

    <template v-else>
      <div class="quick-look-preview">
        <LibraryThumbnail
          :id="entry.id"
          :key="entry.id"
          :last-modified="entry.lastModified"
          :format="entry.format"
        />
      </div>

      <div class="quick-look-title-row">
        <span
          class="quick-look-title"
          data-testid="quick-look-title"
        >{{ entry.title }}</span>
        <span class="quick-look-badge">{{ formatBadgeLabel(entry.format) }}</span>
      </div>

      <dl class="quick-look-facts">
        <div
          v-for="fact in facts"
          :key="fact.label"
          class="fact-row"
        >
          <dt>{{ fact.label }}</dt>
          <dd>{{ fact.value }}</dd>
        </div>
      </dl>

      <div class="quick-look-actions">
        <button
          type="button"
          class="quick-look-open"
          data-testid="quick-look-open"
          @click="emit('open', entry.id)"
        >
          Open
        </button>
        <div class="action-grid">
          <button
            type="button"
            class="quick-look-action"
            data-testid="quick-look-download"
            @click="emit('download', entry.id)"
          >
            Download
          </button>
          <button
            type="button"
            class="quick-look-action"
            data-testid="quick-look-duplicate"
            @click="emit('action', { id: entry.id, action: 'duplicate' })"
          >
            Duplicate
          </button>
          <button
            type="button"
            class="quick-look-action"
            data-testid="quick-look-rename"
            @click="emit('action', { id: entry.id, action: 'rename' })"
          >
            Rename
          </button>
          <button
            type="button"
            class="quick-look-action"
            data-testid="quick-look-trash"
            @click="emit('action', { id: entry.id, action: 'delete' })"
          >
            Move to trash
          </button>
        </div>
      </div>
    </template>
  </aside>
</template>

<style scoped>
.quick-look {
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
  padding-inline-start: var(--space-5);
  /* The handoff's heavier column rule. Height comes from HomeView's sticky
     layout, not the grid row (see LibraryRail's note). */
  border-inline-start: calc(2 * var(--border-width)) solid var(--border-default);
}

.quick-look-heading {
  margin: 0;
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
  letter-spacing: 0.08em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.quick-look-empty {
  font-size: var(--text-md);
  color: var(--text-secondary);
  line-height: var(--line-normal);
  text-wrap: pretty;
}

/* The preview well: the handoff's inset ground with the thumbnail centred. */
.quick-look-preview {
  display: grid;
  place-items: center;
  padding: var(--space-5);
  min-height: 12rem;
  background: var(--surface-sunken);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
}

.quick-look-preview > :deep(*) {
  width: 9rem;
}

.quick-look-title-row {
  display: flex;
  align-items: baseline;
  gap: var(--space-3);
  min-width: 0;
}

.quick-look-title {
  font-size: var(--text-lg);
  font-weight: var(--weight-bold);
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.quick-look-badge {
  flex: none;
  font-size: var(--text-xs);
  letter-spacing: 0.06em;
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--accent-subtle);
  color: var(--text-accent);
}

.quick-look-facts {
  margin: 0;
  display: flex;
  flex-direction: column;
}

.fact-row {
  display: flex;
  justify-content: space-between;
  gap: var(--space-4);
  padding: var(--space-2) 0;
  border-bottom: var(--border-width) solid var(--border-subtle);
  font-size: var(--text-md);
}

.fact-row dt {
  color: var(--text-secondary);
}

.fact-row dd {
  margin: 0;
  text-align: end;
}

.quick-look-actions {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.quick-look-open {
  font: inherit;
  font-size: var(--text-button);
  font-weight: var(--weight-button);
  min-height: var(--height-control);
  border: none;
  border-radius: var(--radius-sm);
  background: var(--accent);
  color: var(--text-on-brand);
  cursor: pointer;
}

.quick-look-open:hover {
  background: var(--accent-hover);
}

.action-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: var(--space-3);
}

.quick-look-action {
  font: inherit;
  font-size: var(--text-button);
  font-weight: var(--weight-button);
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}

.quick-look-action:hover {
  background: var(--state-hover);
}
</style>

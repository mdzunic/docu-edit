<script setup lang="ts">
import { computed } from "vue";
import { asStylesPaneEntries, noStylesPaneEntries } from "../lib/styles";

// The "Styles" side pane (SPEC-042): every style in the document grouped by
// type (AC-1), each row showing whether it is in use (AC-2). The caret's
// active paragraph/character style is read out at the top (AC-10). Runtime
// prop/emit declarations, not defineProps<{...}>() generics — see
// AppShell.vue's comment on the same espree constraint OutlinePane.vue
// already follows.
const props = defineProps({
  /** lib/styles.stylesPaneEntries(model) — every named style, all families. */
  entries: { type: Array, default: () => noStylesPaneEntries() },
  /** The caret's resolved paragraph style id ("" for Normal). */
  caretParagraphStyleId: { type: String, default: "" },
  caretParagraphStyleName: { type: String, default: "Normal" },
  /** The caret's active character style id ("" for none). */
  caretCharacterStyleId: { type: String, default: "" },
  caretCharacterStyleName: { type: String, default: "" },
});

const emit = defineEmits(["close", "new-style", "modify-style", "delete-style", "apply-style"]);

const GROUPS = [
  { type: "paragraph", label: "Paragraph" },
  { type: "character", label: "Character" },
  { type: "list", label: "List" },
  { type: "table", label: "Table" },
];

const grouped = computed(() => {
  const list = asStylesPaneEntries(props.entries);
  return GROUPS.map((group) => ({
    ...group,
    entries: list.filter((entry) => entry.type === group.type),
  }));
});
</script>

<template>
  <nav
    class="styles-pane"
    aria-label="Styles"
    data-testid="styles-pane"
  >
    <div class="styles-pane-header">
      <h2 class="styles-pane-title">
        Styles
      </h2>
      <button
        type="button"
        class="styles-pane-close"
        aria-label="Close Styles pane"
        data-testid="styles-pane-close"
        @click="emit('close')"
      >
        ×
      </button>
    </div>

    <div
      class="caret-readout"
      data-testid="styles-caret-readout"
    >
      <p>
        <span class="caret-readout-label">Paragraph:</span>
        <span data-testid="caret-paragraph-style">{{ caretParagraphStyleName || "Normal" }}</span>
      </p>
      <p>
        <span class="caret-readout-label">Character:</span>
        <span data-testid="caret-character-style">{{ caretCharacterStyleName || "None" }}</span>
      </p>
    </div>

    <div class="styles-pane-groups">
      <section
        v-for="group in grouped"
        :key="group.type"
        class="styles-group"
        :data-testid="`styles-group-${group.type}`"
      >
        <div class="styles-group-header">
          <h3>{{ group.label }}</h3>
          <button
            type="button"
            class="new-style-button"
            :data-testid="`new-style-${group.type}`"
            @click="emit('new-style', group.type)"
          >
            New style…
          </button>
        </div>
        <ul
          v-if="group.entries.length > 0"
          class="styles-list"
        >
          <li
            v-for="entry in group.entries"
            :key="entry.id"
            class="style-row"
            :data-testid="`style-row-${entry.id}`"
          >
            <button
              type="button"
              class="style-apply"
              :data-testid="`style-apply-${entry.id}`"
              @click="emit('apply-style', { id: entry.id, type: entry.type })"
            >
              {{ entry.name }}
            </button>
            <span
              v-if="entry.inUse"
              class="in-use-badge"
              :data-testid="`style-in-use-${entry.id}`"
            >In use</span>
            <span
              v-if="entry.unmodeled"
              class="unmodeled-badge"
              :data-testid="`style-unmodeled-${entry.id}`"
              title="This style was read from the source document and cannot be fully edited here"
            >Unmodeled</span>
            <button
              type="button"
              class="style-action"
              :disabled="entry.unmodeled"
              :data-testid="`style-modify-${entry.id}`"
              @click="emit('modify-style', entry.id)"
            >
              Modify…
            </button>
            <button
              type="button"
              class="style-action"
              :data-testid="`style-delete-${entry.id}`"
              @click="emit('delete-style', entry.id)"
            >
              Delete
            </button>
          </li>
        </ul>
        <p
          v-else
          class="styles-group-empty"
        >
          No {{ group.label.toLowerCase() }} styles yet.
        </p>
      </section>
    </div>
  </nav>
</template>

<style scoped>
.styles-pane {
  width: 18rem;
  height: 100%;
  overflow-y: auto;
  border-left: var(--border-width) solid var(--border-default);
  background: var(--surface-base);
  box-sizing: border-box;
  font-size: var(--text-md);
}

.styles-pane-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: var(--space-4) var(--space-5) var(--space-3);
}

.styles-pane-title {
  margin: 0;
  font-size: var(--text-lg);
  font-weight: var(--weight-bold);
}

.styles-pane-close {
  border: none;
  background: transparent;
  font-size: var(--text-xl);
  cursor: pointer;
  line-height: 1;
}

.caret-readout {
  padding: 0 var(--space-5) var(--space-3);
  border-bottom: var(--border-width) solid var(--border-subtle);
}

.caret-readout p {
  margin: 0 0 var(--space-2);
}

.caret-readout-label {
  font-weight: var(--weight-bold);
  margin-right: var(--space-3);
}

.styles-pane-groups {
  padding: var(--space-2) 0 var(--space-5);
}

.styles-group {
  padding: var(--space-3) var(--space-5);
}

.styles-group-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.styles-group-header h3 {
  margin: 0;
  font-size: var(--text-md);
  text-transform: uppercase;
  letter-spacing: 0.02em;
  color: var(--text-secondary);
}

.new-style-button {
  font: inherit;
  font-size: var(--text-sm);
  border: none;
  background: transparent;
  color: var(--text-accent);
  cursor: pointer;
  padding: var(--space-1) var(--space-2);
}

.styles-group-empty {
  color: var(--text-secondary);
  margin: var(--space-3) 0 0;
}

.styles-list {
  list-style: none;
  margin: var(--space-3) 0 0;
  padding: 0;
}

.style-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  padding: var(--space-2) 0;
}

.style-apply {
  flex: 1;
  text-align: left;
  border: none;
  background: transparent;
  font: inherit;
  cursor: pointer;
  padding: var(--space-1) 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.in-use-badge {
  font-size: var(--text-xs);
  color: var(--text-accent);
  background: var(--surface-selected);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  white-space: nowrap;
}

.unmodeled-badge {
  font-size: var(--text-xs);
  color: var(--text-secondary);
  background: var(--surface-sunken);
  border-radius: var(--radius-sm);
  padding: var(--space-1) var(--space-3);
  white-space: nowrap;
}

.style-action {
  font-size: var(--text-sm);
  border: none;
  background: transparent;
  color: var(--text-secondary);
  cursor: pointer;
  padding: var(--space-1) var(--space-2);
}

.style-action:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
</style>

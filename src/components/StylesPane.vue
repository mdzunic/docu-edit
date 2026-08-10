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
  border-left: 1px solid #dadce0;
  background: #fff;
  box-sizing: border-box;
  font-size: 0.8125rem;
}

.styles-pane-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0.75rem 1rem 0.5rem;
}

.styles-pane-title {
  margin: 0;
  font-size: 0.9375rem;
  font-weight: 600;
}

.styles-pane-close {
  border: none;
  background: transparent;
  font-size: 1.125rem;
  cursor: pointer;
  line-height: 1;
}

.caret-readout {
  padding: 0 1rem 0.5rem;
  border-bottom: 1px solid #e8eaed;
}

.caret-readout p {
  margin: 0 0 0.25rem;
}

.caret-readout-label {
  font-weight: 600;
  margin-right: 0.375rem;
}

.styles-pane-groups {
  padding: 0.25rem 0 1rem;
}

.styles-group {
  padding: 0.5rem 1rem;
}

.styles-group-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
}

.styles-group-header h3 {
  margin: 0;
  font-size: 0.8125rem;
  text-transform: uppercase;
  letter-spacing: 0.02em;
  color: #5f6368;
}

.new-style-button {
  font: inherit;
  font-size: 0.75rem;
  border: none;
  background: transparent;
  color: #1a73e8;
  cursor: pointer;
  padding: 0.125rem 0.25rem;
}

.styles-group-empty {
  color: #80868b;
  margin: 0.375rem 0 0;
}

.styles-list {
  list-style: none;
  margin: 0.375rem 0 0;
  padding: 0;
}

.style-row {
  display: flex;
  align-items: center;
  gap: 0.375rem;
  padding: 0.25rem 0;
}

.style-apply {
  flex: 1;
  text-align: left;
  border: none;
  background: transparent;
  font: inherit;
  cursor: pointer;
  padding: 0.125rem 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.in-use-badge {
  font-size: 0.6875rem;
  color: #1a73e8;
  background: #e8f0fe;
  border-radius: 3px;
  padding: 0.0625rem 0.375rem;
  white-space: nowrap;
}

.unmodeled-badge {
  font-size: 0.6875rem;
  color: #5f6368;
  background: #f1f3f4;
  border-radius: 3px;
  padding: 0.0625rem 0.375rem;
  white-space: nowrap;
}

.style-action {
  font-size: 0.75rem;
  border: none;
  background: transparent;
  color: #5f6368;
  cursor: pointer;
  padding: 0.125rem 0.25rem;
}

.style-action:disabled {
  opacity: 0.4;
  cursor: not-allowed;
}
</style>

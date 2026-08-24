<script setup lang="ts">
// The trash view (SPEC-067 AC-36..39, AC-45): header with the back link and
// "Empty trash", the note, and the trashed-documents table with Restore +
// overflow menu per row. Emits the destructive intents (empty-trash, the
// menu's delete-permanently) up to HomeView, which owns the confirm
// dialogs.
import UiButton from "../../components/ui/UiButton.vue";
import RowMenu from "../../components/library/RowMenu.vue";
import { formatBadgeLabel, formatRelativeTime } from "../../lib/format";
import { TRASH_MENU_ITEMS } from "../../lib/libraryMenu";

const props = defineProps({
  lib: { type: Object, required: true },
  menuFor: { type: String, default: "" },
});

const emit = defineEmits(["empty-trash", "keydown", "menu-toggle", "menu-close", "menu-select"]);
</script>

<template>
  <div
    class="library"
    data-testid="trash-view"
  >
    <div class="library-header">
      <div class="header-line">
        <UiButton
          variant="ghost"
          data-testid="back-to-library"
          @click="props.lib.showTrash.value = false"
        >
          ← Library
        </UiButton>
        <h2 class="library-heading">
          Trash
        </h2>
        <div class="header-spacer" />
        <UiButton
          variant="danger"
          data-testid="empty-trash"
          :disabled="props.lib.trashCount.value === 0"
          @click="emit('empty-trash')"
        >
          Empty trash
        </UiButton>
      </div>
      <p class="trash-note">
        Documents here keep their id and version history. Restore returns them
        to your library intact.
      </p>
    </div>
    <p
      v-if="props.lib.trashCount.value === 0"
      class="state-message"
      data-testid="trash-empty"
    >
      The trash is empty.
    </p>
    <table
      v-else
      class="library-table"
      role="grid"
      data-testid="trash-list"
      @keydown="emit('keydown', $event)"
    >
      <thead>
        <tr>
          <th scope="col">
            Title
          </th>
          <th scope="col">
            Trashed
          </th>
          <th scope="col">
            Format
          </th>
          <th scope="col">
            <span class="visually-hidden">Actions</span>
          </th>
        </tr>
      </thead>
      <tbody>
        <tr
          v-for="(entry, index) in props.lib.trashVisible.value"
          :key="entry.id"
          class="library-row"
          role="row"
          data-testid="library-row"
          :data-doc-id="entry.id"
          :tabindex="props.lib.rowTabindex(entry.id, index)"
          @focus="props.lib.focusedId.value = entry.id"
        >
          <td
            class="cell-title"
            data-testid="row-title"
          >
            {{ entry.title }}
          </td>
          <td
            class="cell-modified"
            data-testid="row-modified"
          >
            {{ formatRelativeTime(entry.trashedAt ?? entry.lastModified, props.lib.now.value) }}
          </td>
          <td class="cell-format">
            <span
              class="badge"
              data-testid="row-format"
            >{{ formatBadgeLabel(entry.format) }}</span>
          </td>
          <td class="cell-actions">
            <UiButton
              size="sm"
              data-testid="row-restore"
              @click="props.lib.restoreDocuments([entry.id])"
            >
              Restore
            </UiButton>
            <RowMenu
              :open="props.menuFor === entry.id"
              :label="`Actions for ${entry.title}`"
              :items="TRASH_MENU_ITEMS"
              @toggle="emit('menu-toggle', entry.id)"
              @close="emit('menu-close')"
              @select="(action) => emit('menu-select', { id: entry.id, action })"
            />
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<style scoped>
.library {
  max-width: 64rem;
  margin: 0 auto;
  padding: var(--space-7) var(--space-5);
}

.library-header {
  margin-bottom: var(--space-5);
  display: flex;
  flex-direction: column;
  gap: var(--space-4);
}

.header-line {
  display: flex;
  align-items: center;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.header-spacer {
  flex: 1 1 auto;
}

.library-heading {
  margin: 0;
  font-size: var(--text-xl);
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.trash-note {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.state-message {
  padding: var(--space-7) var(--space-5);
  text-align: center;
  color: var(--text-secondary);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-4);
}

.library-table {
  width: 100%;
  border-collapse: collapse;
  background: var(--surface-base);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-1);
}

.library-table th {
  text-align: left;
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  padding: var(--space-3) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-default);
}

.library-table td {
  padding: var(--space-3) var(--space-5);
  border-bottom: var(--border-width) solid var(--border-subtle);
  color: var(--text-primary);
}

.library-row {
  cursor: pointer;
}

.library-row:hover {
  background: var(--surface-raised);
}

.library-row:focus-visible {
  outline: var(--focus-ring-width, 2px) solid var(--border-focus, var(--text-accent));
  outline-offset: -2px;
}

.cell-title {
  font-weight: var(--weight-medium);
  overflow-wrap: anywhere;
}

.cell-modified {
  color: var(--text-secondary);
  white-space: nowrap;
}

.cell-actions {
  position: relative;
  text-align: right;
  white-space: nowrap;
}

.badge {
  display: inline-block;
  padding: var(--space-1) var(--space-3);
  border-radius: var(--radius-pill);
  background: var(--surface-selected);
  color: var(--text-accent);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.visually-hidden {
  position: absolute;
  width: 1px;
  height: 1px;
  padding: 0;
  margin: calc(-1 * var(--border-width));
  overflow: hidden;
  clip: rect(0, 0, 0, 0);
  white-space: nowrap;
  border: 0;
}
</style>

<script setup lang="ts">
import BaseDialog from "./BaseDialog.vue";

// The bookmark manager (SPEC-025 AC-23..25): lists the document's bookmarks
// in document order, jumps to one, deletes one. Pure presentation — the
// parent owns the jump (caret + scroll) and the delete command. BaseDialog
// owns the shell, Escape and the focus trap; @dismiss maps onto the close
// emit callers already handle.
// Runtime prop declarations — see AppShell.vue's espree comment.
defineProps({
  /** Bookmark names in document order. */
  names: { type: Array, default: () => [] },
});

const emit = defineEmits(["jump", "remove", "close"]);
</script>

<template>
  <BaseDialog
    title="Bookmarks"
    testid="bookmark-manager"
    width="22rem"
    @dismiss="emit('close')"
  >
    <p
      v-if="names.length === 0"
      class="empty"
      data-testid="bookmark-manager-empty"
    >
      No bookmarks in this document
    </p>
    <ul
      v-else
      class="bookmark-list"
    >
      <li
        v-for="name in names"
        :key="String(name)"
        class="bookmark-row"
        data-testid="bookmark-row"
      >
        <span class="bookmark-name">{{ name }}</span>
        <button
          type="button"
          :data-testid="`bookmark-jump-${name}`"
          :aria-label="`Go to bookmark ${name}`"
          @click="emit('jump', name)"
        >
          Go to
        </button>
        <button
          type="button"
          :data-testid="`bookmark-delete-${name}`"
          :aria-label="`Delete bookmark ${name}`"
          @click="emit('remove', name)"
        >
          Delete
        </button>
      </li>
    </ul>
    <template #actions>
      <button
        type="button"
        data-testid="bookmark-manager-close"
        @click="emit('close')"
      >
        Close
      </button>
    </template>
  </BaseDialog>
</template>

<style scoped>
.empty {
  margin: 0.5rem 0;
  font-size: 0.875rem;
  color: #5f6368;
}

.bookmark-list {
  list-style: none;
  margin: 0;
  padding: 0;
  max-height: 14rem;
  overflow-y: auto;
}

.bookmark-row {
  display: flex;
  align-items: center;
  gap: 0.5rem;
  padding: 0.25rem 0;
}

.bookmark-name {
  flex: 1;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 0.875rem;
}

.bookmark-row button {
  font: inherit;
  font-size: 0.8125rem;
  padding: 0.25rem 0.625rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}

.bookmark-row button:hover {
  background: #f1f3f4;
}

.dialog-actions button {
  font: inherit;
  padding: 0.375rem 1rem;
  border: 1px solid #bdbdbd;
  border-radius: 4px;
  background: #fff;
  cursor: pointer;
}
</style>

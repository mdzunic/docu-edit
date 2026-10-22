<script setup lang="ts">
// The comments pane (SPEC-135 "Threads"): the `comments` pane of SPEC-129's
// sheet side panel, a dumb view of useSheetComments (`controller`).
// `sheet-comments-panel` (aria-label "Comments") holds the identity field
// (`sheet-comment-author-name`), the filter (`sheet-comment-filter`: all |
// unresolved), the composer New comment opens on the active cell
// (`sheet-comment-composer`), then one `sheet-comment-group` per sheet in tab
// order (its heading the sheet's name) with its threads by (row, col): each
// `sheet-comment-thread-${threadId}` — its cell's address (activating the
// thread selects its sheet and cell), its `sheet-comment-message-${messageId}`
// rows (author, the `commentInitials` chip, formatTimestamp, text; Edit and
// Delete — `sheet-comment-edit` / `-delete` — on the identity's own messages
// only), Resolve or Reopen (`sheet-comment-resolve` / `-reopen`) and, while
// open, a reply (`sheet-comment-reply-input` / `-submit`). An .ods workbook
// holds notes only: the pane lists its notes (`sheet-comment-note-${noteId}`)
// instead. `sheet-comment-empty` when there is nothing to list. Every text
// is plain text — interpolated, never markup. Runtime prop declarations per
// the espree constraint (see AppShell.vue).
import { computed, nextTick, useTemplateRef, watch } from "vue";
import { asSheetComments } from "../../views/spreadsheet/useSheetComments";

const props = defineProps({
  /** useSheetComments' return. */
  controller: { type: Object, required: true },
});

const c = computed(() => asSheetComments(props.controller));

/** A field's value, as text. */
function valueOf(event = new Event("input")) {
  const el = event.target;
  return el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement || el instanceof HTMLSelectElement ? el.value : "";
}

/** Ctrl/Cmd+Enter in a textarea runs `submit`. */
function submitOnModEnter(event = new KeyboardEvent("keydown"), submit = () => {}) {
  if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    submit();
  }
}

const panel = useTemplateRef("panel");
// The composer takes focus as it opens; a thread the panel is opened on
// scrolls into view.
watch(
  () => c.value.composer.value,
  async (open) => {
    if (open) {
      await nextTick();
      panel.value?.querySelector<HTMLTextAreaElement>('[data-testid="sheet-comment-composer-input"]')?.focus();
    }
  },
  { immediate: true },
);
watch(
  () => c.value.focusedId.value,
  async (id) => {
    if (id) {
      await nextTick();
      const cards = panel.value ? [...panel.value.querySelectorAll<HTMLElement>("[data-item-id]")] : [];
      cards.find((card) => card.dataset.itemId === id)?.scrollIntoView?.({ block: "nearest" });
    }
  },
  { immediate: true },
);
</script>

<template>
  <section
    ref="panel"
    class="sheet-comments"
    data-testid="sheet-comments-panel"
    aria-label="Comments"
  >
    <div class="sheet-comments-controls">
      <label>
        Your name
        <input
          class="sheet-comments-field"
          data-testid="sheet-comment-author-name"
          :value="c.author.value"
          @change="c.setAuthor(valueOf($event))"
        >
      </label>
      <label v-if="!c.notesOnly.value">
        Show
        <select
          class="sheet-comments-field"
          data-testid="sheet-comment-filter"
          :value="c.filter.value"
          @change="c.setFilter(valueOf($event))"
        >
          <option value="all">All comments</option>
          <option value="unresolved">Unresolved</option>
        </select>
      </label>
    </div>
    <p
      v-if="c.notesOnly.value"
      class="sheet-comments-note"
    >
      OpenDocument spreadsheets support notes only.
    </p>
    <div
      v-if="c.composer.value"
      class="sheet-comment-card"
      data-testid="sheet-comment-composer"
    >
      <div class="sheet-comment-where">
        New comment on {{ c.composer.value.address }}
      </div>
      <textarea
        class="sheet-comment-input"
        data-testid="sheet-comment-composer-input"
        aria-label="New comment"
        :value="c.composerDraft.value"
        @input="c.composerDraft.value = valueOf($event)"
        @keydown="submitOnModEnter($event, c.submitComposer)"
        @keydown.esc.prevent="c.cancelComposer()"
      />
      <div class="sheet-comment-actions">
        <button
          type="button"
          data-testid="sheet-comment-composer-submit"
          :disabled="c.composerDraft.value.trim() === ''"
          @click="c.submitComposer()"
        >
          Comment
        </button>
        <button
          type="button"
          data-testid="sheet-comment-composer-cancel"
          @click="c.cancelComposer()"
        >
          Cancel
        </button>
      </div>
    </div>
    <p
      v-if="c.groups.value.length === 0"
      class="sheet-comments-empty"
      data-testid="sheet-comment-empty"
    >
      {{ c.notesOnly.value ? "No notes in this workbook." : c.filter.value === "unresolved" ? "No unresolved comments." : "No comments yet. Select a cell and choose New comment." }}
    </p>
    <section
      v-for="group in c.groups.value"
      :key="group.sheetId"
      class="sheet-comment-group"
      data-testid="sheet-comment-group"
      :aria-label="group.sheetName"
    >
      <h3 class="sheet-comment-group-heading">
        {{ group.sheetName }}
      </h3>
      <template
        v-for="item in group.items"
        :key="item.kind === 'thread' ? item.threadId : item.noteId"
      >
        <article
          v-if="item.kind === 'thread'"
          class="sheet-comment-card"
          :class="{ 'is-resolved': item.resolved, 'is-focused': c.focusedId.value === item.threadId }"
          :data-testid="`sheet-comment-thread-${item.threadId}`"
          :data-item-id="item.threadId"
          :data-resolved="item.resolved ? 'true' : undefined"
        >
          <button
            type="button"
            class="sheet-comment-where"
            data-testid="sheet-comment-activate"
            :aria-label="`${group.sheetName}!${item.address}`"
            @click="c.activate(group.sheetId, item.threadId, item.ref)"
          >
            {{ item.address }}<span v-if="item.resolved"> · Resolved</span>
          </button>
          <div
            v-for="message in item.messages"
            :key="message.messageId"
            class="sheet-comment-message"
            :data-testid="`sheet-comment-message-${message.messageId}`"
          >
            <div class="sheet-comment-header">
              <span
                class="sheet-comment-initials"
                data-testid="sheet-comment-initials"
                aria-hidden="true"
              >{{ message.initials }}</span>
              <span
                class="sheet-comment-author"
                data-testid="sheet-comment-author"
              >{{ message.author }}</span>
              <time
                v-if="message.time"
                class="sheet-comment-time"
              >{{ message.time }}</time>
            </div>
            <template v-if="c.editing.value?.threadId === item.threadId && c.editing.value?.messageId === message.messageId">
              <textarea
                class="sheet-comment-input"
                data-testid="sheet-comment-edit-input"
                aria-label="Edit comment"
                :value="c.editDraft.value"
                @input="c.editDraft.value = valueOf($event)"
                @keydown="submitOnModEnter($event, c.saveEdit)"
                @keydown.esc.prevent="c.cancelEdit()"
              />
              <div class="sheet-comment-actions">
                <button
                  type="button"
                  data-testid="sheet-comment-edit-save"
                  @click="c.saveEdit()"
                >
                  Save
                </button>
                <button
                  type="button"
                  data-testid="sheet-comment-edit-cancel"
                  @click="c.cancelEdit()"
                >
                  Cancel
                </button>
              </div>
            </template>
            <template v-else>
              <p
                class="sheet-comment-text"
                data-testid="sheet-comment-text"
              >
                {{ message.text }}
              </p>
              <div
                v-if="message.own"
                class="sheet-comment-actions"
              >
                <button
                  type="button"
                  data-testid="sheet-comment-edit"
                  @click="c.startEdit(item.threadId, message.messageId)"
                >
                  Edit
                </button>
                <button
                  type="button"
                  data-testid="sheet-comment-delete"
                  @click="c.deleteMessage(item.threadId, message.messageId)"
                >
                  Delete
                </button>
              </div>
            </template>
          </div>
          <div
            v-if="!item.resolved"
            class="sheet-comment-reply"
          >
            <textarea
              class="sheet-comment-input"
              data-testid="sheet-comment-reply-input"
              aria-label="Reply"
              placeholder="Reply…"
              :value="c.replyDrafts.value[item.threadId] ?? ''"
              @input="c.setReplyDraft(item.threadId, valueOf($event))"
              @keydown="submitOnModEnter($event, () => c.reply(item.threadId))"
            />
          </div>
          <div class="sheet-comment-actions">
            <button
              v-if="!item.resolved"
              type="button"
              data-testid="sheet-comment-reply-submit"
              :disabled="(c.replyDrafts.value[item.threadId] ?? '').trim() === ''"
              @click="c.reply(item.threadId)"
            >
              Reply
            </button>
            <button
              v-if="!item.resolved"
              type="button"
              data-testid="sheet-comment-resolve"
              @click="c.resolve(item.threadId)"
            >
              Resolve
            </button>
            <button
              v-else
              type="button"
              data-testid="sheet-comment-reopen"
              @click="c.reopen(item.threadId)"
            >
              Reopen
            </button>
          </div>
        </article>
        <article
          v-else
          class="sheet-comment-card"
          :class="{ 'is-focused': c.focusedId.value === item.noteId }"
          :data-testid="`sheet-comment-note-${item.noteId}`"
          :data-item-id="item.noteId"
        >
          <button
            type="button"
            class="sheet-comment-where"
            data-testid="sheet-comment-activate"
            :aria-label="`${group.sheetName}!${item.address}`"
            @click="c.activate(group.sheetId, item.noteId, item.ref)"
          >
            {{ item.address }} · Note
          </button>
          <div class="sheet-comment-message">
            <div class="sheet-comment-header">
              <span
                v-if="item.initials"
                class="sheet-comment-initials"
                aria-hidden="true"
              >{{ item.initials }}</span>
              <span class="sheet-comment-author">{{ item.author }}</span>
              <time
                v-if="item.time"
                class="sheet-comment-time"
              >{{ item.time }}</time>
            </div>
            <p
              class="sheet-comment-text"
              data-testid="sheet-comment-text"
            >
              {{ item.text }}
            </p>
          </div>
        </article>
      </template>
    </section>
  </section>
</template>

<style scoped>
.sheet-comments {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-3);
  font-size: var(--text-sm);
  color: var(--text-primary);
}

.sheet-comments-controls {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
}

.sheet-comments-controls label {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.sheet-comments-field {
  font: inherit;
  font-size: var(--text-sm);
  padding: var(--space-1) var(--space-2);
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.sheet-comments-note,
.sheet-comments-empty {
  margin: 0;
  font-style: italic;
  color: var(--text-secondary);
}

.sheet-comment-group {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.sheet-comment-group-heading {
  margin: 0;
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: var(--text-secondary);
}

.sheet-comment-card {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-3);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
}

.sheet-comment-card.is-focused {
  border-color: var(--accent);
  box-shadow: var(--shadow-1);
}

.sheet-comment-card.is-resolved {
  opacity: 0.75;
}

.sheet-comment-where {
  align-self: flex-start;
  padding: 0;
  font: inherit;
  font-size: var(--text-xs);
  font-weight: var(--weight-bold);
  color: var(--text-accent);
  background: none;
  border: 0;
  cursor: pointer;
}

.sheet-comment-message + .sheet-comment-message {
  padding-top: var(--space-2);
  border-top: var(--border-width) solid var(--border-subtle);
}

.sheet-comment-header {
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-xs);
  color: var(--text-secondary);
}

.sheet-comment-initials {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 1.5rem;
  height: 1.5rem;
  padding: 0 var(--space-1);
  box-sizing: border-box;
  font-weight: var(--weight-bold);
  color: var(--text-accent);
  background: var(--surface-sunken);
  border: var(--border-width) solid var(--border-default);
  border-radius: 999px;
}

.sheet-comment-author {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.sheet-comment-time {
  margin-left: auto;
}

.sheet-comment-text {
  margin: var(--space-1) 0 0;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}

.sheet-comment-input {
  width: 100%;
  min-height: 2.5rem;
  box-sizing: border-box;
  padding: var(--space-2);
  font: inherit;
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  resize: vertical;
}

.sheet-comment-actions {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
}

.sheet-comment-actions button {
  padding: var(--space-1) var(--space-3);
  font: inherit;
  font-size: var(--text-xs);
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.sheet-comment-actions button:disabled {
  cursor: default;
  opacity: 0.6;
}
</style>

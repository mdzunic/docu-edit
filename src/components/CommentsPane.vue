<script setup lang="ts">
import { computed, reactive, ref } from "vue";
import { formatTimestamp } from "../lib/format";
import {
  asCommentAuthorList,
  asCommentMessage,
  asCommentThreadTops,
  asVisibleCommentThreads,
  noEditingMessageId,
  noReplyDrafts,
} from "../views/document/useComments";

// The comment threads pane (SPEC-038): a right-hand margin dock (open
// threads, each pinned beside its anchored range, plus a "Resolved"
// section) that toggles into a flat "comments-only view" with open/resolved
// and author filters (AC "A comments-only view lists all threads in
// document order with filters for open/resolved and by author"). Runtime
// prop/emit declarations, not the generic form — espree parses .vue scripts
// and rejects TS-only syntax (see OutlinePane.vue for the same constraint).
const props = defineProps({
  openThreads: { type: Array, default: () => [] },
  resolvedThreads: { type: Array, default: () => [] },
  filteredThreads: { type: Array, default: () => [] },
  commentsOnlyView: { type: Boolean, default: false },
  /** threadId -> px offset (a Map), for margin-view card positioning. */
  threadTops: { type: Object, default: () => new Map() },
  activeThreadId: { type: String, default: null },
  authorName: { type: String, default: "You" },
  commentFilter: { type: String, default: "all" },
  commentAuthorFilter: { type: String, default: "" },
  authors: { type: Array, default: () => [] },
  composerOpen: { type: Boolean, default: false },
  editable: { type: Boolean, default: true },
});

const emit = defineEmits([
  "close",
  "toggle-comments-only",
  "update:comment-filter",
  "update:author-filter",
  "update:author-name",
  "submit-comment",
  "cancel-composer",
  "focus-thread",
  "submit-reply",
  "save-edit",
  "delete-message",
  "resolve",
  "reopen",
  "delete-thread",
]);

// The runtime Array/Object props above carry no element type (the espree
// constraint bars writing PropType<...> here) — these computed casts give
// the template typed access, the same seam asOutlineNodes gives OutlinePane.
const openThreadsTyped = computed(() => asVisibleCommentThreads(props.openThreads));
const resolvedThreadsTyped = computed(() => asVisibleCommentThreads(props.resolvedThreads));
const filteredThreadsTyped = computed(() => asVisibleCommentThreads(props.filteredThreads));
const threadTopsTyped = computed(() => asCommentThreadTops(props.threadTops));
const authorsTyped = computed(() => asCommentAuthorList(props.authors));

const composerText = ref("");
function submitComposer() {
  if (composerText.value.trim().length === 0) {
    return;
  }
  emit("submit-comment", composerText.value);
  composerText.value = "";
}
function cancelComposer() {
  composerText.value = "";
  emit("cancel-composer");
}

// One reply draft per thread, so switching between open threads keeps each
// one's unsent text.
const replyDrafts = reactive(noReplyDrafts());
function replyDraft(threadId = "") {
  return replyDrafts[threadId] ?? "";
}
function updateReplyDraft(threadId = "", event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLTextAreaElement) {
    replyDrafts[threadId] = target.value;
  }
}
function submitReply(threadId = "") {
  const text = replyDraft(threadId);
  if (text.trim().length === 0) {
    return;
  }
  emit("submit-reply", threadId, text);
  replyDrafts[threadId] = "";
}

// Only one message edits at a time across the whole pane.
const editingMessageId = ref(noEditingMessageId());
const editDraftText = ref("");
function startEdit(message = asCommentMessage({ messageId: "", text: "" })) {
  editingMessageId.value = message.messageId;
  editDraftText.value = message.text;
}
function cancelEdit() {
  editingMessageId.value = null;
  editDraftText.value = "";
}
function saveEdit(threadId = "") {
  if (editDraftText.value.trim().length === 0) {
    return;
  }
  emit("save-edit", threadId, editingMessageId.value, editDraftText.value);
  editingMessageId.value = null;
  editDraftText.value = "";
}

function canModify(message = asCommentMessage({ author: "" })) {
  return props.editable && message.author === props.authorName;
}

function onFilterChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("update:comment-filter", target.value);
  }
}
function onAuthorFilterChange(event = new Event("change")) {
  const target = event.target;
  if (target instanceof HTMLSelectElement) {
    emit("update:author-filter", target.value);
  }
}
function onAuthorNameInput(event = new Event("input")) {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    emit("update:author-name", target.value);
  }
}
</script>

<template>
  <aside
    class="comments-pane"
    aria-label="Comments"
    data-testid="comments-pane"
  >
    <div class="comments-pane-header">
      <span class="comments-pane-title">Comments</span>
      <button
        type="button"
        class="comments-only-toggle"
        :aria-pressed="commentsOnlyView"
        data-testid="comments-only-toggle"
        @click="emit('toggle-comments-only')"
      >
        {{ commentsOnlyView ? "Margin view" : "Comments-only view" }}
      </button>
    </div>

    <div class="comments-identity">
      <label>
        Comment as
        <input
          type="text"
          class="comments-author-input"
          data-testid="comment-author-name-input"
          :value="authorName"
          @change="onAuthorNameInput"
        >
      </label>
    </div>

    <div
      v-if="commentsOnlyView"
      class="comments-filters"
    >
      <label>
        Show
        <select
          class="comments-filter-select"
          data-testid="comment-filter-select"
          :value="commentFilter"
          @change="onFilterChange"
        >
          <option value="all">
            All
          </option>
          <option value="open">
            Open
          </option>
          <option value="resolved">
            Resolved
          </option>
        </select>
      </label>
      <label>
        Author
        <select
          class="comments-filter-select"
          data-testid="comment-author-filter-select"
          :value="commentAuthorFilter"
          @change="onAuthorFilterChange"
        >
          <option value="">
            All authors
          </option>
          <option
            v-for="author in authorsTyped"
            :key="author"
            :value="author"
          >
            {{ author }}
          </option>
        </select>
      </label>
    </div>

    <div
      v-if="composerOpen"
      class="comment-composer"
      data-testid="comment-composer"
    >
      <textarea
        v-model="composerText"
        class="comment-composer-input"
        placeholder="Add a comment…"
        data-testid="comment-composer-input"
      />
      <div class="comment-actions">
        <button
          type="button"
          data-testid="comment-composer-submit"
          @click="submitComposer"
        >
          Comment
        </button>
        <button
          type="button"
          @click="cancelComposer"
        >
          Cancel
        </button>
      </div>
    </div>

    <template v-if="commentsOnlyView">
      <p
        v-if="filteredThreadsTyped.length === 0"
        class="comments-empty"
        data-testid="comments-empty"
      >
        No comments
      </p>
      <div class="comments-only-list">
        <article
          v-for="entry in filteredThreadsTyped"
          :key="entry.thread.threadId"
          class="comment-thread"
          :class="{ 'is-active': entry.thread.threadId === activeThreadId, 'is-resolved': entry.thread.resolved }"
          :data-testid="`comment-thread-${entry.thread.threadId}`"
          @click="emit('focus-thread', entry.thread.threadId)"
        >
          <div
            v-for="message in entry.thread.messages"
            :key="message.messageId"
            class="comment-message"
            :data-testid="`comment-message-${message.messageId}`"
          >
            <div class="comment-message-header">
              <span class="comment-author">{{ message.author }}</span>
              <span class="comment-timestamp">{{ formatTimestamp(new Date(message.createdAt)) }}</span>
            </div>
            <template v-if="editingMessageId === message.messageId">
              <textarea
                v-model="editDraftText"
                class="comment-edit-input"
                @click.stop
              />
              <div class="comment-actions">
                <button
                  type="button"
                  @click.stop="saveEdit(entry.thread.threadId)"
                >
                  Save
                </button>
                <button
                  type="button"
                  @click.stop="cancelEdit"
                >
                  Cancel
                </button>
              </div>
            </template>
            <template v-else>
              <p class="comment-text">
                {{ message.text }}
              </p>
              <div
                v-if="canModify(message)"
                class="comment-actions"
              >
                <button
                  type="button"
                  @click.stop="startEdit(message)"
                >
                  Edit
                </button>
                <button
                  type="button"
                  @click.stop="emit('delete-message', entry.thread.threadId, message.messageId)"
                >
                  Delete
                </button>
              </div>
            </template>
          </div>
          <div
            v-if="editable"
            class="comment-thread-actions"
          >
            <button
              v-if="!entry.thread.resolved"
              type="button"
              @click.stop="emit('resolve', entry.thread.threadId)"
            >
              Resolve
            </button>
            <button
              v-else
              type="button"
              @click.stop="emit('reopen', entry.thread.threadId)"
            >
              Reopen
            </button>
            <button
              type="button"
              @click.stop="emit('delete-thread', entry.thread.threadId)"
            >
              Delete
            </button>
          </div>
        </article>
      </div>
    </template>

    <template v-else>
      <p
        v-if="openThreadsTyped.length === 0 && resolvedThreadsTyped.length === 0"
        class="comments-empty"
        data-testid="comments-empty"
      >
        No comments yet
      </p>
      <div class="comments-open-list">
        <article
          v-for="entry in openThreadsTyped"
          :key="entry.thread.threadId"
          class="comment-thread"
          :class="{ 'is-active': entry.thread.threadId === activeThreadId }"
          :style="threadTopsTyped.has(entry.thread.threadId) ? { top: `${threadTopsTyped.get(entry.thread.threadId)}px` } : {}"
          :data-testid="`comment-thread-${entry.thread.threadId}`"
          @click="emit('focus-thread', entry.thread.threadId)"
        >
          <div
            v-for="message in entry.thread.messages"
            :key="message.messageId"
            class="comment-message"
            :data-testid="`comment-message-${message.messageId}`"
          >
            <div class="comment-message-header">
              <span class="comment-author">{{ message.author }}</span>
              <span class="comment-timestamp">{{ formatTimestamp(new Date(message.createdAt)) }}</span>
            </div>
            <template v-if="editingMessageId === message.messageId">
              <textarea
                v-model="editDraftText"
                class="comment-edit-input"
                @click.stop
              />
              <div class="comment-actions">
                <button
                  type="button"
                  @click.stop="saveEdit(entry.thread.threadId)"
                >
                  Save
                </button>
                <button
                  type="button"
                  @click.stop="cancelEdit"
                >
                  Cancel
                </button>
              </div>
            </template>
            <template v-else>
              <p class="comment-text">
                {{ message.text }}
              </p>
              <div
                v-if="canModify(message)"
                class="comment-actions"
              >
                <button
                  type="button"
                  @click.stop="startEdit(message)"
                >
                  Edit
                </button>
                <button
                  type="button"
                  @click.stop="emit('delete-message', entry.thread.threadId, message.messageId)"
                >
                  Delete
                </button>
              </div>
            </template>
          </div>
          <div
            v-if="editable"
            class="comment-reply"
          >
            <textarea
              class="comment-reply-input"
              placeholder="Reply…"
              data-testid="comment-reply-input"
              :value="replyDraft(entry.thread.threadId)"
              @click.stop
              @input="updateReplyDraft(entry.thread.threadId, $event)"
            />
            <div class="comment-actions">
              <button
                type="button"
                @click.stop="submitReply(entry.thread.threadId)"
              >
                Reply
              </button>
              <button
                type="button"
                @click.stop="emit('resolve', entry.thread.threadId)"
              >
                Resolve
              </button>
              <button
                type="button"
                @click.stop="emit('delete-thread', entry.thread.threadId)"
              >
                Delete
              </button>
            </div>
          </div>
        </article>
      </div>

      <div
        v-if="resolvedThreadsTyped.length > 0"
        class="comments-resolved-section"
      >
        <div class="comments-resolved-heading">
          Resolved
        </div>
        <article
          v-for="entry in resolvedThreadsTyped"
          :key="entry.thread.threadId"
          class="comment-thread is-resolved"
          :class="{ 'is-active': entry.thread.threadId === activeThreadId }"
          :data-testid="`comment-thread-${entry.thread.threadId}`"
          @click="emit('focus-thread', entry.thread.threadId)"
        >
          <div
            v-for="message in entry.thread.messages"
            :key="message.messageId"
            class="comment-message"
            :data-testid="`comment-message-${message.messageId}`"
          >
            <div class="comment-message-header">
              <span class="comment-author">{{ message.author }}</span>
              <span class="comment-timestamp">{{ formatTimestamp(new Date(message.createdAt)) }}</span>
            </div>
            <p class="comment-text">
              {{ message.text }}
            </p>
          </div>
          <div
            v-if="editable"
            class="comment-thread-actions"
          >
            <button
              type="button"
              @click.stop="emit('reopen', entry.thread.threadId)"
            >
              Reopen
            </button>
            <button
              type="button"
              @click.stop="emit('delete-thread', entry.thread.threadId)"
            >
              Delete
            </button>
          </div>
        </article>
      </div>
    </template>
  </aside>
</template>

<style scoped>
.comments-pane {
  flex: none;
  align-self: flex-start;
  position: sticky;
  top: 0;
  width: 18rem;
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

.comments-pane-header {
  flex: none;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 0.5rem;
  padding: 0.5rem 0.75rem;
}

.comments-pane-title {
  font-size: 0.75rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}

.comments-only-toggle {
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.15rem 0.5rem;
  font-size: 0.7rem;
  cursor: pointer;
}

.comments-only-toggle[aria-pressed="true"] {
  background: #e8f0fe;
  color: #1a73e8;
  border-color: #1a73e8;
}

.comments-identity,
.comments-filters {
  flex: none;
  display: flex;
  gap: 0.5rem;
  padding: 0 0.75rem 0.5rem;
  flex-wrap: wrap;
}

.comments-identity label,
.comments-filters label {
  display: flex;
  flex-direction: column;
  gap: 0.15rem;
  font-size: 0.7rem;
  color: #5f6368;
}

.comments-author-input,
.comments-filter-select {
  font-size: 0.75rem;
  padding: 0.15rem 0.3rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
}

.comments-empty {
  margin: 0;
  padding: 0.25rem 0.75rem 0.75rem;
  font-style: italic;
  color: #5f6368;
}

.comments-open-list {
  position: relative;
  flex: 1 1 auto;
}

.comments-only-list {
  display: flex;
  flex-direction: column;
  gap: 0.5rem;
  padding: 0 0.5rem 0.5rem;
}

.comment-composer {
  margin: 0 0.75rem 0.5rem;
  padding: 0.5rem;
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
}

.comment-thread {
  background: #fff;
  border: 1px solid #dadce0;
  border-radius: 6px;
  padding: 0.5rem;
  margin: 0 0.5rem 0.5rem;
  cursor: pointer;
}

.comments-open-list .comment-thread {
  position: absolute;
  left: 0.5rem;
  right: 0.5rem;
  margin: 0 0 0.5rem;
}

.comment-thread.is-active {
  border-color: #1a73e8;
  box-shadow: 0 0 0 1px #1a73e8;
}

.comment-thread.is-resolved {
  opacity: 0.75;
}

.comment-message + .comment-message {
  margin-top: 0.4rem;
  padding-top: 0.4rem;
  border-top: 1px solid #f1f3f4;
}

.comment-message-header {
  display: flex;
  justify-content: space-between;
  gap: 0.5rem;
  font-size: 0.7rem;
  color: #5f6368;
  margin-bottom: 0.2rem;
}

.comment-author {
  font-weight: 600;
  color: #202124;
}

.comment-text {
  margin: 0;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  color: #202124;
}

.comment-actions {
  display: flex;
  gap: 0.4rem;
  margin-top: 0.3rem;
}

.comment-actions button,
.comment-thread-actions button {
  border: 1px solid #dadce0;
  background: #fff;
  border-radius: 4px;
  padding: 0.1rem 0.45rem;
  font-size: 0.7rem;
  cursor: pointer;
}

.comment-thread-actions {
  display: flex;
  gap: 0.4rem;
  margin-top: 0.5rem;
}

.comment-reply {
  margin-top: 0.5rem;
  padding-top: 0.4rem;
  border-top: 1px solid #f1f3f4;
}

.comment-composer-input,
.comment-edit-input,
.comment-reply-input {
  width: 100%;
  box-sizing: border-box;
  resize: vertical;
  min-height: 2.5rem;
  font: inherit;
  padding: 0.3rem;
  border: 1px solid #dadce0;
  border-radius: 4px;
}

.comments-resolved-section {
  flex: none;
  margin-top: 0.5rem;
  border-top: 1px solid #dadce0;
  padding-top: 0.5rem;
}

.comments-resolved-heading {
  padding: 0 0.75rem 0.4rem;
  font-size: 0.7rem;
  font-weight: 600;
  letter-spacing: 0.04em;
  text-transform: uppercase;
  color: #5f6368;
}
</style>

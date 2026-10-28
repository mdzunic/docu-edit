<script setup lang="ts">
// SPEC-114 (experimental): the "Everything" scope's surface in the library's
// main column — the indexing progress line, the "Not searched" list (every
// document the index could not read, with its reason; never silently
// absent), and, while the field holds a query, the grouped results that
// replace the rows: one per document, with its title, kind badge, match
// count and up to three snippets, matched words marked the way SPEC-067
// marks a title (HighlightedText). Loaded only by HomeView's ExperimentGate
// while the library-search experiment is on.
//
// Product-neutral like the rest of src/views/library: it emits intents —
// `showing` (the results replace the rows) and `open` ({ id, find, at }: the
// document, what its find surface should look for, and the body offset of
// the first match) — and HomeView routes them. Keyboard: the rows'
// conventions (arrows move, Enter opens) with a roving tabindex; from the
// search field, ArrowDown enters the list and Enter opens the first result.
// `lib` is the LibraryController (runtime-typed Object — see useLibrary.ts).
import { computed, nextTick, onBeforeUnmount, onMounted, ref, watch } from "vue";
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import HighlightedText from "../../components/library/HighlightedText.vue";
import { entryIds, noSearchHit, useLibrarySearch } from "../../composables/useLibrarySearch";
import { formatBadgeLabel } from "../../lib/format";

const props = defineProps({
  lib: { type: Object, required: true },
});

const emit = defineEmits(["open", "showing"]);

const search = useLibrarySearch();
const everything = computed(() => search.scope.value === "everything");
const text = computed(() => String(props.lib.search.value));
/** Results replace the rows while there is a query to answer. */
const showing = computed(() => everything.value && text.value.trim() !== "");
watch(showing, (value) => emit("showing", value), { immediate: true });
onBeforeUnmount(() => emit("showing", false));

/** Active filter chips narrow the results too. */
function onlyIds() {
  const lib = props.lib;
  return lib.activeFormats.value.size > 0 || lib.starredOnly.value ? entryIds(lib.visible.value) : undefined;
}

watch([text, everything, () => props.lib.visible.value], () => search.query(text.value, onlyIds()), { immediate: true });
// A document added, renamed or trashed changes the library's counts: bring
// the index up to date (unchanged documents cost nothing — their stamps hold).
watch(() => props.lib.counts.value, () => search.refreshIndex());
onMounted(() => search.refreshIndex());

const progress = computed(() => search.progress.value);
const progressLabel = computed(() => `Indexed ${progress.value.done} of ${progress.value.total}`);
const hits = computed(() => search.hits.value);
const unsearched = computed(() =>
  search.unsearched.value
    .filter((status) => props.lib.formatOf(status.docId) !== undefined)
    .map((status) => ({ ...status, title: props.lib.titleOf(status.docId) })),
);
const noneCount = computed(() => unsearched.value.filter((status) => status.coverage === "none").length);
const errorText = computed(() => {
  const error = search.error.value;
  return error ? `${error.error} (at character ${error.position + 1})` : "";
});
const answered = computed(() => search.answered.value === text.value && !search.searching.value);

function badge(id = "") {
  const format = props.lib.formatOf(id);
  return format ? formatBadgeLabel(format) : "";
}

function matchLabel(count = 0) {
  return count === 1 ? "1 match" : `${count} matches`;
}

// --- open and keyboard ---------------------------------------------------------

function open(hit = noSearchHit()) {
  search.opened(hit.id);
  emit("open", { id: hit.id, find: hit.findText, at: hit.at ?? -1 });
}

const focusedId = ref("");

function tabindexOf(id = "", index = 0) {
  return focusedId.value === id || (focusedId.value === "" && index === 0) ? 0 : -1;
}

async function focusResult(index = 0) {
  const hit = hits.value[index];
  if (!hit) {
    return;
  }
  focusedId.value = hit.id;
  await nextTick();
  const row = document.querySelector(`[data-testid="search-result"][data-doc-id="${hit.id}"]`);
  if (row instanceof HTMLElement) {
    row.focus();
  }
}

function onResultKeydown(event = new KeyboardEvent("keydown"), index = 0) {
  const key = event.key;
  if (key === "ArrowDown" || key === "ArrowRight") {
    event.preventDefault();
    void focusResult(Math.min(hits.value.length - 1, index + 1));
  } else if (key === "ArrowUp" || key === "ArrowLeft") {
    event.preventDefault();
    void focusResult(Math.max(0, index - 1));
  } else if (key === "Enter") {
    event.preventDefault();
    open(hits.value[index]);
  }
}

/** From the search field: ArrowDown enters the results, Enter opens the
 * first (the field itself is the library's, not this surface's). */
function onFieldKeydown(event = new KeyboardEvent("keydown")) {
  const target = event.target;
  if (!showing.value || hits.value.length === 0 || !(target instanceof HTMLElement) || target.dataset.testid !== "library-search") {
    return;
  }
  if (event.key === "ArrowDown") {
    event.preventDefault();
    void focusResult(0);
  } else if (event.key === "Enter") {
    event.preventDefault();
    open(hits.value[0]);
  }
}

onMounted(() => document.addEventListener("keydown", onFieldKeydown));
onBeforeUnmount(() => document.removeEventListener("keydown", onFieldKeydown));
</script>

<template>
  <section
    v-if="everything"
    class="library-search-panel"
    data-testid="library-search-panel"
    aria-label="Search inside documents"
  >
    <div class="search-status-line">
      <ExperimentBadge experiment="library-search" />
      <p
        class="search-progress"
        role="status"
        data-testid="search-index-progress"
        :data-running="progress.running ? 'true' : 'false'"
      >
        {{ progressLabel }}
      </p>
    </div>

    <details
      v-if="unsearched.length > 0"
      class="search-unsearched"
      data-testid="search-unsearched"
    >
      <summary>
        Not searched ({{ noneCount }}){{ unsearched.length > noneCount ? `, partly searched (${unsearched.length - noneCount})` : "" }}
      </summary>
      <ul class="unsearched-list">
        <li
          v-for="status in unsearched"
          :key="status.docId"
          data-testid="search-unsearched-item"
          :data-doc-id="status.docId"
        >
          <span class="unsearched-title">{{ status.title }}</span>
          <span
            class="unsearched-reason"
            data-testid="search-unsearched-reason"
          >{{ status.coverage === "partial" ? `partly searched: ${status.reason}` : status.reason }}</span>
        </li>
      </ul>
    </details>

    <template v-if="showing">
      <p
        v-if="errorText"
        class="search-message search-error"
        role="alert"
        data-testid="search-error"
      >
        {{ errorText }}
      </p>
      <ul
        v-else-if="hits.length > 0"
        class="search-results"
        data-testid="search-results"
        aria-label="Documents containing your search"
      >
        <li
          v-for="(hit, index) in hits"
          :key="hit.id"
          class="search-result"
          data-testid="search-result"
          :data-doc-id="hit.id"
          :tabindex="tabindexOf(hit.id, index)"
          @click="open(hit)"
          @keydown="onResultKeydown($event, index)"
          @focus="focusedId = hit.id"
        >
          <div class="result-head">
            <span
              class="result-title"
              data-testid="search-result-title"
            >
              <HighlightedText
                :segments="hit.title"
                mark-testid="result-title-match"
              />
            </span>
            <span
              class="badge"
              data-testid="search-result-kind"
            >{{ badge(hit.id) }}</span>
            <span
              class="result-count"
              data-testid="search-result-count"
            >{{ matchLabel(hit.matches) }}</span>
          </div>
          <p
            v-for="(segments, snippetIndex) in hit.snippets"
            :key="snippetIndex"
            class="result-snippet"
            data-testid="search-snippet"
          >
            <HighlightedText
              :segments="segments"
              mark-testid="snippet-match"
            />
          </p>
        </li>
      </ul>
      <p
        v-else-if="answered"
        class="search-message"
        data-testid="search-no-results"
      >
        No document in your library contains that.
      </p>
      <p
        v-else
        class="search-message"
        data-testid="search-pending"
      >
        Searching…
      </p>
      <button
        v-if="!errorText && search.total.value > hits.length"
        type="button"
        class="search-show-more"
        data-testid="search-show-more"
        @click="search.showMore()"
      >
        Show more ({{ search.total.value - hits.length }})
      </button>
    </template>
  </section>
</template>

<style scoped>
.library-search-panel {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
}

.search-status-line {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.search-progress,
.search-message {
  margin: 0;
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.search-error {
  color: var(--danger, var(--text-primary));
}

.search-unsearched summary {
  cursor: pointer;
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.unsearched-list {
  margin: var(--space-2) 0 0;
  padding-inline-start: var(--space-5);
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  font-size: var(--text-sm);
}

.unsearched-title {
  color: var(--text-primary);
  margin-inline-end: var(--space-2);
}

.unsearched-reason {
  color: var(--text-secondary);
}

.search-results {
  margin: 0;
  padding: 0;
  list-style: none;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.search-result {
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  cursor: pointer;
}

.search-result:hover,
.search-result:focus-visible {
  border-color: var(--border-strong);
  background: var(--surface-selected);
}

.result-head {
  display: flex;
  align-items: baseline;
  gap: var(--space-3);
  flex-wrap: wrap;
}

.result-title {
  font-weight: var(--weight-bold);
  color: var(--text-primary);
}

.badge {
  font-size: var(--text-sm);
  padding: 0 var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-pill);
  color: var(--text-secondary);
}

.result-count {
  font-size: var(--text-sm);
  color: var(--text-secondary);
}

.result-snippet {
  margin: var(--space-2) 0 0;
  font-size: var(--text-sm);
  color: var(--text-primary);
  line-height: var(--line-normal);
}

.search-show-more {
  align-self: flex-start;
  font: inherit;
  min-height: var(--height-control);
  padding: 0 var(--space-3);
  border: var(--border-width) solid var(--border-strong);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  cursor: pointer;
}
</style>

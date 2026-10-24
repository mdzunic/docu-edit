<script setup lang="ts">
// The deck's find dock (SPEC-098): the find bar — SPEC-017's FindBar, passed in
// the default slot by the view — and, below it, the results list
// (`deck-search-results`) grouped by slide: one `deck-search-group` per slide
// with a hit, its `deck-search-slide-chip` carrying the 1-based slide number,
// then one `deck-search-hit` button per hit showing its ±40-character snippet
// with the match emphasised, the frame's name and a scope tag ("Notes",
// "Table r2c3"). Activating a hit (a click, or Enter/Space on the button)
// shows it exactly as find-next does. `deck-search-progress` stands while a
// big deck's walk is still running, `deck-search-empty` when nothing matched,
// with `deck-search-scope` restating what the search leaves out — exactly
// when it is needed. Reads the useDeckSearch controller like
// TransitionsPanel.vue reads its own; emits nothing. Runtime prop
// declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import { asDeckSearch } from "../../views/presentation/useDeckSearch";
import { asDeckResultGroups } from "../../views/presentation/deckSearchRows";

const props = defineProps({
  controller: { type: Object, required: true },
});

const c = computed(() => asDeckSearch(props.controller));
const groups = computed(() => asDeckResultGroups(c.value.groups.value));
const showResults = computed(() => c.value.query.value !== "" && c.value.error.value === "");
const empty = computed(() => !c.value.indexing.value && c.value.hitCount.value === 0);

/** Escape anywhere in the dock closes it — the bar's own fields already did
 * (preventDefault), so this only acts on the results list — and goes no
 * further, so the stage's Escape does not also drop the selection. */
function onKeydown(event = new KeyboardEvent("keydown")) {
  if (event.key !== "Escape") {
    return;
  }
  event.stopPropagation();
  if (!event.defaultPrevented) {
    event.preventDefault();
    c.value.closeFind();
  }
}
</script>

<template>
  <div
    v-if="c.open.value"
    class="deck-find-dock"
  >
    <div
      class="deck-find-stack"
      @keydown="onKeydown"
    >
      <slot />
      <section
        v-if="showResults"
        class="deck-search-results"
        data-testid="deck-search-results"
        aria-label="Search results"
      >
        <p
          v-if="c.indexing.value"
          class="deck-search-note"
          data-testid="deck-search-progress"
          role="status"
        >
          Searching slides…
        </p>
        <ol
          v-if="groups.length > 0"
          class="deck-search-groups"
        >
          <li
            v-for="group in groups"
            :key="group.slideIndex"
            class="deck-search-group"
            data-testid="deck-search-group"
            :data-slide-number="group.slideNumber"
            :aria-label="`Slide ${group.slideNumber}`"
          >
            <span
              class="deck-search-slide-chip"
              data-testid="deck-search-slide-chip"
              :title="`Slide ${group.slideNumber}`"
            >{{ group.slideNumber }}</span>
            <ul class="deck-search-hits">
              <li
                v-for="entry in group.hits"
                :key="entry.index"
              >
                <button
                  type="button"
                  class="deck-search-hit"
                  :class="{ 'is-current': entry.index === c.currentHitIndex.value }"
                  data-testid="deck-search-hit"
                  :aria-current="entry.index === c.currentHitIndex.value ? 'true' : undefined"
                  @click="c.goToHit(entry.index)"
                >
                  <span class="hit-snippet">{{ entry.before }}<mark>{{ entry.match }}</mark>{{ entry.after }}</span>
                  <span class="hit-meta">
                    <span class="hit-frame">{{ entry.frameName }}</span>
                    <span
                      v-if="entry.scopeTag"
                      class="hit-scope"
                      data-testid="deck-search-hit-scope"
                    >{{ entry.scopeTag }}</span>
                  </span>
                </button>
              </li>
            </ul>
          </li>
        </ol>
        <template v-if="empty">
          <p
            class="deck-search-note"
            data-testid="deck-search-empty"
          >
            No matches in this presentation
          </p>
          <p
            class="deck-search-note deck-search-scope"
            data-testid="deck-search-scope"
          >
            {{ c.scopeTitle.value }}
          </p>
        </template>
      </section>
    </div>
  </div>
</template>

<style scoped>
/* A zero-height strip at the top of the deck body: the stack hangs from its
   right end over the stage, like the text editor's and the sheet's find. */
.deck-find-dock {
  position: relative;
  z-index: 30;
  height: 0;
}

.deck-find-stack {
  position: absolute;
  top: 0.5rem;
  right: 1.5rem;
  display: flex;
  flex-direction: column;
  align-items: stretch;
  gap: var(--space-2);
  max-width: calc(100vw - 3rem);
}

/* FindBar places itself absolutely; in the dock it stacks. */
.deck-find-stack :deep(.find-bar) {
  position: static;
}

.deck-search-results {
  max-height: 18rem;
  overflow: auto;
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-md);
  box-shadow: var(--shadow-2);
  font-size: var(--text-sm);
}

.deck-search-groups,
.deck-search-hits {
  margin: 0;
  padding: 0;
  list-style: none;
}

.deck-search-group {
  display: flex;
  align-items: flex-start;
  gap: var(--space-3);
  padding: var(--space-2) var(--space-3);
  border-top: var(--border-width) solid var(--border-subtle);
}

.deck-search-group:first-child {
  border-top: 0;
}

.deck-search-slide-chip {
  flex: 0 0 auto;
  min-width: 1.75rem;
  padding: var(--space-1) var(--space-2);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--text-secondary);
  font-weight: var(--weight-bold);
  text-align: center;
}

.deck-search-hits {
  flex: 1 1 auto;
  min-width: 0;
}

.deck-search-hit {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  width: 100%;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  background: transparent;
  color: var(--text-primary);
  font: inherit;
  text-align: left;
  cursor: pointer;
}

.deck-search-hit:hover {
  background: var(--state-hover);
}

.deck-search-hit.is-current {
  background: var(--surface-selected);
}

.deck-search-hit:focus-visible {
  outline: var(--space-1) solid var(--focus-ring);
  outline-offset: calc(-1 * var(--space-1));
}

.hit-snippet {
  overflow-wrap: anywhere;
}

.hit-snippet mark {
  background: rgb(255 214 0 / 0.55);
  color: inherit;
  font-weight: var(--weight-bold);
}

.hit-meta {
  display: flex;
  gap: var(--space-2);
  color: var(--text-secondary);
  font-size: var(--text-xs, var(--text-sm));
}

.hit-scope {
  padding: 0 var(--space-1);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.deck-search-note {
  margin: 0;
  padding: var(--space-2) var(--space-3);
  color: var(--text-secondary);
}

.deck-search-scope {
  padding-top: 0;
}
</style>

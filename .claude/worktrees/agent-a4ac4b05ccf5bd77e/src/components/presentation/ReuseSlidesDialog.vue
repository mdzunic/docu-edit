<script setup lang="ts">
// "Reuse slides" (SPEC-099): pick another deck from the library, pick slides
// from it (a listbox of thumbnails), choose whether they take this deck's
// theme or keep their own formatting, and insert them after the current
// selection. Over the shared BaseDialog; every decision lives in the
// useReuseSlides controller. Runtime prop declarations, not the generic forms
// (espree lints .vue scripts).
import { computed } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import SlideThumb from "./SlideThumb.vue";
import { asSlides, filmstripLabel, filmstripThumbScale } from "./slideRender";
import { eventValue } from "../../views/presentation/deckSetupFields";

const props = defineProps({
  /** The useReuseSlides() controller (the sorter's `reuse`). */
  controller: { type: Object, required: true },
});

const c = props.controller;

const slides = computed(() => asSlides(c.sourceSlides.value));
const deck = computed(() => c.source.value);
const scale = computed(() => (deck.value ? filmstripThumbScale(deck.value.slideWidthCm, 132) : 1));

function onKeydown(index = 0, event = new KeyboardEvent("keydown")) {
  if (event.key === " " || event.key === "Enter") {
    event.preventDefault();
    c.toggleSlide(index);
  }
}
</script>

<template>
  <BaseDialog
    title="Reuse slides"
    testid="slides-reuse-dialog"
    width="40rem"
    @dismiss="c.cancel()"
  >
    <div class="reuse-form">
      <p
        v-if="c.decks.value.length === 0"
        class="reuse-empty"
      >
        There are no other presentations in your library.
      </p>
      <label
        v-else
        class="reuse-row"
      >
        <span>Deck</span>
        <select
          data-testid="reuse-deck-select"
          :value="c.deckId.value"
          @change="c.chooseDeck(eventValue($event))"
        >
          <option
            v-for="entry in c.decks.value"
            :key="entry.id"
            :value="entry.id"
          >
            {{ entry.title }}
          </option>
        </select>
      </label>

      <div
        v-if="deck"
        class="reuse-head"
      >
        <span>{{ c.selected.value.length }} of {{ slides.length }} selected</span>
        <UiButton
          data-testid="reuse-select-all"
          @click="c.selectAll()"
        >
          Select all
        </UiButton>
      </div>
      <ul
        v-if="deck"
        class="reuse-slides"
        role="listbox"
        aria-label="Slides to reuse"
        aria-multiselectable="true"
      >
        <li
          v-for="(slide, index) in slides"
          :key="slide.id"
          class="reuse-slide"
          :class="{ selected: c.selected.value.includes(index) }"
          role="option"
          tabindex="0"
          :aria-selected="c.selected.value.includes(index)"
          :aria-label="filmstripLabel(slide, index, slides.length)"
          data-testid="reuse-slide"
          @click="c.toggleSlide(index)"
          @keydown="onKeydown(index, $event)"
        >
          <SlideThumb
            :slide="slide"
            :index="index"
            :width-cm="deck.slideWidthCm"
            :height-cm="deck.slideHeightCm"
            :media="deck.media || []"
            :scale="scale"
            :theme="deck.theme || null"
            stacked
          />
        </li>
      </ul>

      <fieldset class="reuse-formatting">
        <legend>Formatting</legend>
        <label>
          <input
            type="radio"
            name="reuse-formatting"
            value="destination"
            data-testid="reuse-use-destination"
            :checked="c.formatting.value === 'destination'"
            @change="c.setFormatting('destination')"
          >
          Use destination theme
        </label>
        <label>
          <input
            type="radio"
            name="reuse-formatting"
            value="source"
            data-testid="reuse-keep-source"
            :checked="c.formatting.value === 'source'"
            @change="c.setFormatting('source')"
          >
          Keep source formatting
        </label>
      </fieldset>
    </div>

    <template #actions>
      <UiButton
        data-testid="reuse-cancel"
        @click="c.cancel()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="reuse-insert"
        :disabled="c.selected.value.length === 0"
        @click="c.insert()"
      >
        Insert
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.reuse-form {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.reuse-row,
.reuse-head {
  display: flex;
  align-items: center;
  gap: var(--space-2);
}

.reuse-head {
  justify-content: space-between;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.reuse-slides {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2);
  max-height: 22rem;
  overflow-y: auto;
  margin: 0;
  padding: var(--space-1);
  list-style: none;
}

.reuse-slide {
  padding: var(--space-1);
  border: var(--border-width) solid transparent;
  border-radius: var(--radius-sm);
  cursor: pointer;
}

.reuse-slide.selected {
  border-color: var(--accent);
  background: var(--surface-raised);
}

.reuse-slide:focus-visible {
  outline: var(--border-width) solid var(--focus-ring);
}

.reuse-formatting {
  display: flex;
  gap: var(--space-4);
  border: 0;
  padding: 0;
  margin: 0;
}

.reuse-formatting legend {
  padding: 0;
  margin-bottom: var(--space-1);
  font-weight: var(--weight-medium);
}

.reuse-empty {
  margin: 0;
  color: var(--text-secondary);
}
</style>

<script setup lang="ts">
// SPEC-115: the New flow's plan step — a real review gate, not a spinner. The
// validated plan as an editable list: the deck's title, theme and size; per
// slide its layout (a select over the closed vocabulary), title, bullets (one
// per line, two leading spaces per indent level) and note; Alt+↑/Alt+↓ (or the
// arrow buttons) reorder, Delete removes, "Add slide" appends. Build receives
// exactly what this shows. Text is edited as text: nothing a model wrote is
// rendered as HTML here.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { nextTick, useTemplateRef } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import UiSelect from "../../components/ui/UiSelect.vue";
import DesignerDrops from "./DesignerDrops.vue";
import { asDeckDesigner, fieldValue, reorderDelta } from "./designerFields";

const props = defineProps({
  /** The useDeckDesigner() session. */
  controller: { type: Object, required: true },
});

const c = asDeckDesigner(props.controller);
const list = useTemplateRef("list");

function setDeck(field = "title", value = "") {
  if (field === "title" || field === "theme" || field === "size") {
    c.setPlanField(field, value);
  }
}

function setField(index = 0, field = "title", event = new Event("input")) {
  if (field === "title" || field === "bullets" || field === "note") {
    c.setSlideField(index, field, fieldValue(event));
  }
}

/** Moves a row and keeps focus on the same control, now at its new place. */
async function move(index = 0, delta = 0, testid = "") {
  const target = c.moveSlide(index, delta);
  if (target === index) {
    return;
  }
  await nextTick();
  const row = list.value?.querySelectorAll('[data-testid="designer-plan-row"]')[target];
  const control = testid ? row?.querySelector(`[data-testid="${testid}"]`) : null;
  if (control instanceof HTMLElement) {
    control.focus();
  }
}

function onRowKeydown(index = 0, event = new KeyboardEvent("keydown")) {
  const delta = reorderDelta(event);
  if (delta === 0) {
    return;
  }
  event.preventDefault();
  const testid = event.target instanceof HTMLElement ? (event.target.dataset.testid ?? "") : "";
  void move(index, delta, testid);
}
</script>

<template>
  <div class="plan">
    <div class="plan-deck">
      <label class="plan-field plan-field--grow">
        <span>Deck title</span>
        <input
          type="text"
          data-testid="designer-plan-title"
          :value="c.plan.title"
          @input="setDeck('title', fieldValue($event))"
        >
      </label>
      <label class="plan-field">
        <span>Theme</span>
        <UiSelect
          data-testid="designer-plan-theme"
          :model-value="c.plan.theme"
          :options="c.themeNames"
          @update:model-value="setDeck('theme', $event)"
        />
      </label>
      <label class="plan-field">
        <span>Slide size</span>
        <UiSelect
          data-testid="designer-plan-size"
          :model-value="c.plan.size"
          :options="c.sizes"
          @update:model-value="setDeck('size', $event)"
        />
      </label>
    </div>

    <DesignerDrops
      :records="c.dropped.value"
      :pending="c.pendingDrops.value"
    />

    <p class="plan-hint">
      Edit anything below before building. Alt+↑ and Alt+↓ move a slide; bullets are one per line, indented with two
      spaces per level.
    </p>
    <ol
      ref="list"
      class="plan-list"
      data-testid="designer-plan"
      aria-label="Planned slides"
    >
      <li
        v-for="(slide, index) in c.plan.slides"
        :key="slide.key"
        class="plan-row"
        data-testid="designer-plan-row"
        @keydown="onRowKeydown(index, $event)"
      >
        <div class="plan-row-head">
          <span
            class="plan-row-number"
            data-testid="designer-row-number"
          >Slide {{ index + 1 }}</span>
          <UiSelect
            data-testid="designer-row-layout"
            :label="`Slide ${index + 1} layout`"
            :model-value="slide.layout"
            :options="c.layoutOptions"
            @update:model-value="c.setSlideField(index, 'layout', $event)"
          />
          <span class="plan-row-tools">
            <UiButton
              size="sm"
              data-testid="designer-row-up"
              :aria-label="`Move slide ${index + 1} up`"
              :disabled="index === 0"
              @click="move(index, -1, 'designer-row-up')"
            >↑</UiButton>
            <UiButton
              size="sm"
              data-testid="designer-row-down"
              :aria-label="`Move slide ${index + 1} down`"
              :disabled="index === c.plan.slides.length - 1"
              @click="move(index, 1, 'designer-row-down')"
            >↓</UiButton>
            <UiButton
              size="sm"
              data-testid="designer-row-delete"
              :aria-label="`Delete slide ${index + 1}`"
              @click="c.deleteSlide(index)"
            >Delete</UiButton>
          </span>
        </div>
        <input
          type="text"
          class="plan-input"
          data-testid="designer-row-title"
          :aria-label="`Slide ${index + 1} title`"
          :value="slide.title"
          @input="setField(index, 'title', $event)"
        >
        <textarea
          class="plan-input"
          data-testid="designer-row-bullets"
          :aria-label="`Slide ${index + 1} bullets, one per line`"
          rows="3"
          :value="slide.bullets"
          @input="setField(index, 'bullets', $event)"
        />
        <input
          type="text"
          class="plan-input"
          data-testid="designer-row-note"
          :aria-label="`Slide ${index + 1} speaker note`"
          placeholder="Speaker note (optional)"
          :value="slide.note"
          @input="setField(index, 'note', $event)"
        >
      </li>
    </ol>
    <p
      v-if="c.plan.slides.length === 0"
      class="plan-hint"
      data-testid="designer-plan-empty"
    >
      A deck needs at least one slide: add one to build.
    </p>
    <UiButton
      data-testid="designer-add-slide"
      :disabled="!c.canAddSlide.value"
      @click="c.addSlide()"
    >
      Add slide
    </UiButton>
  </div>
</template>

<style scoped>
.plan {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.plan-deck {
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-3);
}

.plan-field {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.plan-field--grow {
  flex: 1 1 14rem;
}

.plan-hint {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.plan-list {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  margin: 0;
  padding: 0;
  list-style: none;
}

.plan-row {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.plan-row:focus-within {
  border-color: var(--border-strong);
}

.plan-row-head {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.plan-row-number {
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
}

.plan-row-tools {
  display: flex;
  gap: var(--space-2);
  margin-left: auto;
}

.plan-input,
.plan-field input {
  width: 100%;
  box-sizing: border-box;
  padding: var(--space-1) var(--space-2);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-base);
  color: var(--text-primary);
  font-family: var(--font-ui);
  font-size: var(--text-control);
}

textarea.plan-input {
  resize: vertical;
}
</style>

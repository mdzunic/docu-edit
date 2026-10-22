<script setup lang="ts">
// The "Deck setup" dialog (SPEC-090): a BaseDialog-shelled two-tab panel (Size
// and Properties) that is a dumb :controller-taking component like
// TransitionsPanel.vue — all state and commands live in the `useDeckSetup`
// controller passed in; this component is presentation + event forwarding only.
// The tabstrip copies PropertiesDialog.vue's roving-tabindex pattern. Runtime
// prop declarations and untyped helpers, not the generic forms (espree lints
// .vue scripts); typed transforms live in the sibling deckSetupFields.ts.
import { nextTick, useTemplateRef } from "vue";
import BaseDialog from "../BaseDialog.vue";
import UiButton from "../ui/UiButton.vue";
import { DECK_PROPERTY_FIELDS, eventValue } from "../../views/presentation/deckSetupFields";

const props = defineProps({
  /** The useDeckSetup() controller object (refs + command functions). */
  controller: { type: Object, required: true },
});

const c = props.controller;

const PRESETS = [
  { value: "16:9", label: "16:9 widescreen" },
  { value: "4:3", label: "4:3 standard" },
  { value: "16:10", label: "16:10" },
  { value: "A4", label: "A4 paper" },
  { value: "Letter", label: "Letter paper" },
  { value: "Custom", label: "Custom" },
];

const TABS = [
  { id: "size", label: "Size", testid: "deck-setup-tab-size" },
  { id: "properties", label: "Properties", testid: "deck-setup-tab-properties" },
];
const tablist = useTemplateRef("tablist");

function focusTab(id = "") {
  void nextTick(() => {
    const el = tablist.value?.querySelector(`[data-testid="deck-setup-tab-${id}"]`);
    if (el instanceof HTMLElement) {
      el.focus();
    }
  });
}

function onTabKeydown(event = new KeyboardEvent("keydown")) {
  const index = TABS.findIndex((tab) => tab.id === c.activeTab.value);
  let nextIndex = -1;
  if (event.key === "ArrowRight") {
    nextIndex = (index + 1) % TABS.length;
  } else if (event.key === "ArrowLeft") {
    nextIndex = (index - 1 + TABS.length) % TABS.length;
  } else if (event.key === "Home") {
    nextIndex = 0;
  } else if (event.key === "End") {
    nextIndex = TABS.length - 1;
  } else {
    return;
  }
  event.preventDefault();
  c.setTab(TABS[nextIndex].id);
  focusTab(TABS[nextIndex].id);
}

function onPreset(event = new Event("change")) {
  c.setPreset(eventValue(event));
}
function onWidth(event = new Event("input")) {
  c.setWidth(eventValue(event));
}
function onHeight(event = new Event("input")) {
  c.setHeight(eventValue(event));
}
function onFirstSlide(event = new Event("input")) {
  c.setFirstSlideNumber(Number(eventValue(event)));
}
function onProperty(key = "", event = new Event("input")) {
  c.setProperty(key, eventValue(event));
}
</script>

<template>
  <BaseDialog
    title="Deck setup"
    testid="deck-setup-dialog"
    width="30rem"
    @dismiss="c.close()"
  >
    <div
      ref="tablist"
      class="tablist"
      role="tablist"
      aria-label="Deck setup tabs"
    >
      <button
        v-for="tab in TABS"
        :id="`deck-setup-tab-${tab.id}`"
        :key="tab.id"
        type="button"
        role="tab"
        :aria-selected="c.activeTab.value === tab.id ? 'true' : 'false'"
        :aria-controls="`deck-setup-panel-${tab.id}`"
        :tabindex="c.activeTab.value === tab.id ? 0 : -1"
        class="tab"
        :class="{ active: c.activeTab.value === tab.id }"
        :data-testid="tab.testid"
        @click="c.setTab(tab.id)"
        @keydown="onTabKeydown"
      >
        {{ tab.label }}
      </button>
    </div>

    <!-- Size tab -->
    <div
      v-show="c.activeTab.value === 'size'"
      id="deck-setup-panel-size"
      role="tabpanel"
      aria-labelledby="deck-setup-tab-size"
    >
      <label class="field">
        <span>Slide size</span>
        <select
          data-testid="deck-size-preset"
          :value="c.preset.value"
          @change="onPreset"
        >
          <option
            v-for="option in PRESETS"
            :key="option.value"
            :value="option.value"
          >
            {{ option.label }}
          </option>
        </select>
      </label>

      <fieldset
        class="orientation"
        data-testid="deck-size-orientation"
      >
        <legend>Orientation</legend>
        <label class="radio">
          <input
            type="radio"
            name="deck-orientation"
            data-testid="deck-size-landscape"
            :checked="c.orientation.value === 'landscape'"
            @change="c.setOrientation('landscape')"
          >
          <span>Landscape</span>
        </label>
        <label class="radio">
          <input
            type="radio"
            name="deck-orientation"
            data-testid="deck-size-portrait"
            :checked="c.orientation.value === 'portrait'"
            @change="c.setOrientation('portrait')"
          >
          <span>Portrait</span>
        </label>
      </fieldset>

      <div class="dimensions">
        <label class="field">
          <span>Width</span>
          <input
            type="text"
            data-testid="deck-size-width"
            :value="c.widthText.value"
            @input="onWidth"
          >
        </label>
        <label class="field">
          <span>Height</span>
          <input
            type="text"
            data-testid="deck-size-height"
            :value="c.heightText.value"
            @input="onHeight"
          >
        </label>
        <span
          class="unit"
          data-testid="deck-size-unit"
        >{{ c.unit.value }}</span>
      </div>
      <p
        v-if="c.sizeError.value"
        class="error"
        data-testid="deck-size-error"
      >
        {{ c.sizeError.value }}
      </p>

      <fieldset
        class="rescale"
        data-testid="deck-size-rescale"
      >
        <legend>Existing content</legend>
        <label class="radio">
          <input
            type="radio"
            name="deck-rescale"
            data-testid="deck-size-rescale-scale"
            :checked="c.rescaleMode.value === 'scale'"
            @change="c.setRescaleMode('scale')"
          >
          <span>Scale content to fit</span>
        </label>
        <label class="radio">
          <input
            type="radio"
            name="deck-rescale"
            data-testid="deck-size-rescale-keep"
            :checked="c.rescaleMode.value === 'keep'"
            @change="c.setRescaleMode('keep')"
          >
          <span>Keep content in place</span>
        </label>
        <p
          v-if="c.rescaleMode.value === 'keep' && c.overflowCount.value > 0"
          class="warning"
          data-testid="deck-size-overflow-warning"
        >
          {{ c.overflowCount.value }} frame(s) would fall outside the slide.
        </p>
      </fieldset>

      <label class="field">
        <span>First slide number</span>
        <input
          type="number"
          min="1"
          max="9999"
          step="1"
          data-testid="deck-first-slide-number"
          :value="c.firstSlideNumber.value"
          @input="onFirstSlide"
        >
      </label>
      <p
        v-if="c.isOdp.value"
        class="hint"
        data-testid="deck-first-slide-number-odp-hint"
      >
        The first slide number isn't saved in .odp files.
      </p>
    </div>

    <!-- Properties tab -->
    <div
      v-show="c.activeTab.value === 'properties'"
      id="deck-setup-panel-properties"
      role="tabpanel"
      aria-labelledby="deck-setup-tab-properties"
    >
      <template
        v-for="field in DECK_PROPERTY_FIELDS"
        :key="field.key"
      >
        <label class="field stacked">
          <span>{{ field.label }}</span>
          <textarea
            v-if="field.multiline"
            :data-testid="field.testid"
            :value="c.propertiesDraft[field.key] ?? ''"
            @input="onProperty(field.key, $event)"
          />
          <input
            v-else
            type="text"
            :data-testid="field.testid"
            :value="c.propertiesDraft[field.key] ?? ''"
            @input="onProperty(field.key, $event)"
          >
        </label>
      </template>

      <dl class="readonly">
        <dt>Created</dt>
        <dd data-testid="deck-properties-created">
          {{ c.created.value }}
        </dd>
        <dt>Modified</dt>
        <dd data-testid="deck-properties-modified">
          {{ c.modified.value }}
        </dd>
      </dl>

      <div
        v-if="c.extraFields.value.length > 0"
        class="extra-fields"
        data-testid="deck-properties-extra-fields"
      >
        <h3>Other properties</h3>
        <p class="hint">
          Read from the file; not editable here, but preserved when you save.
        </p>
        <dl class="readonly">
          <template
            v-for="extra in c.extraFields.value"
            :key="extra.name"
          >
            <dt>{{ extra.name }}</dt>
            <dd>{{ extra.value }}</dd>
          </template>
        </dl>
      </div>
    </div>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="deck-setup-cancel"
        @click="c.close()"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="deck-setup-apply"
        :disabled="!c.canApply.value"
        @click="c.apply()"
      >
        Apply
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.tablist {
  display: flex;
  gap: var(--space-2);
  border-bottom: var(--border-width) solid var(--border-default);
  margin-bottom: var(--space-5);
}

.tab {
  font: inherit;
  padding: var(--space-3) var(--space-4);
  border: none;
  border-bottom: calc(2 * var(--border-width)) solid transparent;
  background: transparent;
  cursor: pointer;
  color: var(--text-secondary);
}

.tab.active {
  color: var(--surface-brand);
  border-bottom-color: var(--surface-brand);
  font-weight: var(--weight-bold);
}

.field {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  margin-bottom: var(--space-4);
}

.field.stacked {
  flex-direction: column;
  align-items: stretch;
  gap: var(--space-2);
}

.field input,
.field select,
.field textarea {
  font: inherit;
  min-width: 10rem;
}

.field.stacked input,
.field.stacked textarea {
  width: 100%;
  box-sizing: border-box;
}

.orientation,
.rescale {
  border: var(--border-width) solid var(--border-subtle);
  border-radius: var(--radius-sm);
  padding: var(--space-2) var(--space-3);
  margin-bottom: var(--space-4);
  display: flex;
  flex-wrap: wrap;
  gap: var(--space-2) var(--space-4);
}

.radio {
  display: inline-flex;
  align-items: center;
  gap: var(--space-1);
}

.dimensions {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}

.dimensions .field {
  flex: 1 1 auto;
  margin-bottom: 0;
}

.unit {
  color: var(--text-secondary);
}

.error {
  margin: var(--space-2) 0 var(--space-4);
  color: var(--text-danger);
  font-size: var(--text-sm);
}

.warning {
  flex-basis: 100%;
  margin: var(--space-1) 0 0;
  color: var(--text-secondary);
}

.hint {
  margin: 0 0 var(--space-3);
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.readonly {
  display: grid;
  grid-template-columns: 8rem 1fr;
  gap: var(--space-2) var(--space-4);
  margin: 0 0 var(--space-4);
}

.readonly dt {
  color: var(--text-secondary);
}

.readonly dd {
  margin: 0;
  color: var(--text-primary);
}

.extra-fields h3 {
  margin: var(--space-4) 0 var(--space-2);
  font-size: var(--text-md);
}
</style>

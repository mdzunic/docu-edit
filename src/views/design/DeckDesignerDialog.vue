<script setup lang="ts">
// SPEC-115: the deck designer (deck-designer) — a three-step flow with
// visible steps: Brief (the prompt and its choices), Plan (the reply streamed,
// checked, then shown as an editable outline — or, without a model, the
// model-free designer's plan), and Build (progress and the result). Back is
// available from every step and nothing is stored before the last one.
//
// One dialog for both entry points: mode "new" (the library's "Presentation
// from a description…") builds a new deck; mode "redesign" (a deck's
// "Redesign…") rearranges the open deck through `host` — the view's model and
// its record/commit — and its final button is "Apply redesign". All of the
// logic is useDeckDesigner's; this file only lays the steps out. Loaded only
// behind the deck-designer experiment's gate, and badged as experimental.
// Runtime prop declarations per the espree constraint (see AppShell.vue).
import { computed } from "vue";
import BaseDialog from "../../components/BaseDialog.vue";
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import UiButton from "../../components/ui/UiButton.vue";
import DesignerBriefStep from "./DesignerBriefStep.vue";
import DesignerPlanStep from "./DesignerPlanStep.vue";
import DesignerRedesignStep from "./DesignerRedesignStep.vue";
import DesignerReply from "./DesignerReply.vue";
import { asRedesignHost } from "./designerFields";
import { useDeckDesigner } from "./useDeckDesigner";

const props = defineProps({
  /** "new" (a deck from a description) or "redesign" (the open deck). */
  mode: { type: String, default: "new", validator: (value) => value === "new" || value === "redesign" },
  /** Redesign only: { model, record, commit, currentIndex } of the open deck. */
  host: { type: Object, default: null },
});

const emit = defineEmits(["close"]);

const c = useDeckDesigner({
  mode: props.mode === "redesign" ? "redesign" : "new",
  host: asRedesignHost(props.host),
  onClose: () => emit("close"),
});

const isNew = c.mode === "new";
const title = isNew ? "Presentation from a description" : "Redesign this deck";
const ready = computed(() => c.step.value === "plan" && c.phase.value === "ready");
const busy = computed(() => c.phase.value === "streaming" || c.phase.value === "repairing");
</script>

<template>
  <BaseDialog
    :title="title"
    testid="deck-designer"
    width="44rem"
    @dismiss="c.cancel()"
  >
    <template #header-extra>
      <ExperimentBadge experiment="deck-designer" />
    </template>
    <ol
      class="designer-steps"
      aria-label="Steps"
      data-testid="designer-steps"
    >
      <li
        v-for="(entry, index) in c.steps"
        :key="entry.id"
        class="designer-step"
        :class="{ 'is-current': entry.id === c.step.value }"
        :data-testid="`designer-step-${entry.id}`"
        :aria-current="entry.id === c.step.value ? 'step' : undefined"
      >
        {{ index + 1 }}. {{ entry.label }}
      </li>
    </ol>

    <div class="designer-body">
      <DesignerBriefStep
        v-if="c.step.value === 'brief'"
        :controller="c"
      />
      <template v-else-if="c.step.value === 'plan'">
        <DesignerReply :controller="c" />
        <template v-if="c.phase.value === 'ready'">
          <DesignerPlanStep
            v-if="isNew"
            :controller="c"
          />
          <DesignerRedesignStep
            v-else
            :controller="c"
          />
        </template>
      </template>
      <div
        v-else
        class="designer-build"
        data-testid="designer-build-step"
      >
        <p
          v-if="c.buildPhase.value === 'building'"
          data-testid="designer-building"
          aria-live="polite"
        >
          Building the deck…
        </p>
        <p
          v-else-if="c.buildPhase.value === 'built' && c.built.value"
          data-testid="designer-built"
          aria-live="polite"
        >
          Created “{{ c.built.value.title }}” — {{ c.built.value.slides }} slides. Opening it…
        </p>
        <template v-else-if="c.buildPhase.value === 'failed'">
          <p
            class="designer-error"
            data-testid="designer-error"
            role="alert"
          >
            {{ c.error.value }}
          </p>
          <p class="designer-note">
            Nothing was stored. Go back to the plan to change it.
          </p>
          <details
            class="designer-raw"
            data-testid="designer-raw"
          >
            <summary>The model's raw reply</summary>
            <pre>{{ c.raw.value }}</pre>
          </details>
        </template>
      </div>
    </div>

    <template #actions>
      <UiButton
        variant="ghost"
        data-testid="designer-cancel"
        @click="c.cancel()"
      >
        Cancel
      </UiButton>
      <UiButton
        data-testid="designer-back"
        @click="c.back()"
      >
        Back
      </UiButton>
      <template v-if="c.step.value === 'brief'">
        <UiButton
          v-if="isNew"
          data-testid="designer-outline"
          :disabled="!c.canOutline.value"
          :title="c.canOutline.value ? undefined : 'Choose a source to outline it without a model'"
          @click="c.outline()"
        >
          Outline this document
        </UiButton>
        <UiButton
          v-else
          data-testid="designer-by-hand"
          @click="c.outline()"
        >
          Rearrange by hand
        </UiButton>
        <UiButton
          v-if="!c.providerOff.value"
          variant="primary"
          data-testid="designer-send"
          :disabled="!c.canSend.value"
          @click="c.send()"
        >
          Send to the model…
        </UiButton>
      </template>
      <template v-else-if="c.step.value === 'plan'">
        <UiButton
          v-if="isNew"
          variant="primary"
          data-testid="designer-build"
          :disabled="!ready || busy || c.plan.slides.length === 0"
          @click="c.build()"
        >
          Build deck
        </UiButton>
        <UiButton
          v-else
          variant="primary"
          data-testid="designer-apply-redesign"
          :disabled="!ready || busy"
          @click="c.applyRedesign()"
        >
          Apply redesign
        </UiButton>
      </template>
    </template>
  </BaseDialog>
</template>

<style scoped>
.designer-steps {
  display: flex;
  gap: var(--space-4);
  margin: 0 0 var(--space-4);
  padding: 0;
  list-style: none;
  color: var(--text-secondary);
  font-family: var(--font-caption);
  font-size: var(--text-sm);
  letter-spacing: var(--tracking-caption);
  text-transform: uppercase;
}

.designer-step.is-current {
  color: var(--text-accent);
  font-weight: var(--weight-bold);
}

/* The plan can be long: the body scrolls, the steps and buttons stay put. */
.designer-body {
  max-height: 60vh;
  overflow-y: auto;
  padding-right: var(--space-2);
}

.designer-error {
  color: var(--text-danger);
  overflow-wrap: anywhere;
}

.designer-note {
  color: var(--text-secondary);
  font-size: var(--text-sm);
}

.designer-raw pre {
  max-height: 16rem;
  overflow: auto;
  padding: var(--space-3);
  background: var(--surface-sunken);
  font-family: var(--font-mono);
  font-size: var(--text-xs);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
}
</style>

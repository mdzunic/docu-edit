<script setup lang="ts">
import { computed, nextTick, onMounted, useTemplateRef, watch } from "vue";
import AppIcon from "../../components/AppIcon.vue";
import ExperimentGate from "../../components/ExperimentGate.vue";
import { useExperiments } from "../../composables/useExperiments";
import { usePreferencesDialog } from "../../composables/usePreferencesDialog";
import { EXPERIMENTS, MODEL_EXPERIMENTS } from "../../lib/experiments";

// SPEC-113: the Preferences dialog's Labs section — every experiment in the
// registry with its title, summary, risk line and toggle, under the standing
// note on what "experimental" means. Toggles write straight through, like
// every other preference. While an experiment that uses a model is on, the
// model provider settings appear below the list; they are that experiment's
// surface, gated and lazily loaded like any other, so with every model
// experiment off the gateway's code is never fetched.
const { flags, isEnabled, setEnabled } = useExperiments();
const preferencesDialog = usePreferencesDialog();
const experiments = EXPERIMENTS;

const modelExperiment = computed(() => MODEL_EXPERIMENTS.find((id) => isEnabled(id)) ?? "");
const loadProviderPanel = () => import("./ModelProviderPanel.vue");

function checked(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement && target.checked;
}

// The experiment badge's "Manage it in Labs" opens the dialog here. Two ticks:
// BaseDialog focuses its own [autofocus] control a tick after mounting, and
// the heading should win that race.
const heading = useTemplateRef("heading");

async function revealIfRequested() {
  if (!preferencesDialog.labsRequested.value) {
    return;
  }
  preferencesDialog.labsRequested.value = false;
  await nextTick();
  await nextTick();
  heading.value?.scrollIntoView?.({ block: "start" });
  heading.value?.focus();
}

onMounted(revealIfRequested);
watch(preferencesDialog.labsRequested, revealIfRequested);
</script>

<template>
  <section
    class="section"
    data-testid="prefs-labs"
    aria-labelledby="preferences-labs-heading"
  >
    <h3
      id="preferences-labs-heading"
      ref="heading"
      class="labs-heading"
      tabindex="-1"
    >
      <AppIcon name="info" />
      Labs
    </h3>
    <p
      class="labs-note"
      data-testid="prefs-labs-note"
    >
      Experimental features may change or be removed at any time, and are not covered by this
      product's round-trip guarantees. Each stays off — absent from the app, its code not even
      loaded — until you turn it on here.
    </p>
    <ul class="labs-list">
      <li
        v-for="experiment in experiments"
        :key="experiment.id"
        class="labs-item"
        :data-testid="`prefs-labs-item-${experiment.id}`"
      >
        <div class="labs-copy">
          <label
            class="labs-title"
            :for="`experiment-${experiment.id}`"
          >{{ experiment.title }}</label>
          <p class="labs-summary">
            {{ experiment.summary }}
          </p>
          <p
            class="labs-risk"
            :data-testid="`experiment-risk-${experiment.id}`"
          >
            Risk: {{ experiment.risk }}
          </p>
        </div>
        <input
          :id="`experiment-${experiment.id}`"
          type="checkbox"
          class="labs-toggle"
          :data-testid="`experiment-${experiment.id}`"
          :checked="flags[experiment.id]"
          @change="(event) => setEnabled(experiment.id, checked(event))"
        >
      </li>
    </ul>
    <ExperimentGate
      v-if="modelExperiment"
      :experiment="modelExperiment"
      :load="loadProviderPanel"
    />
  </section>
</template>

<style scoped>
/* The section shell (border, spacing) comes from PreferencesDialog's own
   .section rule, which reaches this root; the heading mirrors its
   .section-title, which does not reach inside. */
.labs-heading {
  margin: 0;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  color: var(--text-secondary);
  text-transform: uppercase;
}

.labs-note {
  margin: 0;
  padding: var(--space-3);
  border-inline-start: var(--space-1) solid var(--warning);
  background: var(--surface-sunken);
  color: var(--text-primary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.labs-list {
  margin: 0;
  padding: 0;
  list-style: none;
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
}

.labs-item {
  display: grid;
  grid-template-columns: minmax(0, 1fr) auto;
  column-gap: var(--space-5);
  align-items: start;
}

.labs-copy {
  display: flex;
  flex-direction: column;
  gap: var(--space-1);
}

.labs-title {
  color: var(--text-primary);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
}

.labs-summary,
.labs-risk {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.labs-risk {
  color: var(--text-primary);
}

.labs-toggle {
  width: 1rem;
  height: 1rem;
  margin-top: var(--space-1);
  accent-color: var(--accent);
}
</style>

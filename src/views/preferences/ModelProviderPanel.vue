<script setup lang="ts">
import { computed, ref, watch } from "vue";
import ExperimentBadge from "../../components/ExperimentBadge.vue";
import UiButton from "../../components/ui/UiButton.vue";
import UiField from "../../components/ui/UiField.vue";
import UiSelect from "../../components/ui/UiSelect.vue";
import ModelActivityLog from "./ModelActivityLog.vue";
import { useExperiments } from "../../composables/useExperiments";
import { useModelGateway } from "../../composables/useModelGateway";
import { usePreferences } from "../../composables/usePreferences";
import { EXPERIMENTS, MODEL_EXPERIMENTS } from "../../lib/experiments";
import { originOf, resolveEndpoint } from "../../lib/ai/origin";
import {
  MAX_OUTPUT_TOKENS_LIMIT,
  MAX_TEMPERATURE,
  MAX_TIMEOUT_MS,
  MIN_TIMEOUT_MS,
} from "../../lib/ai/providerConfig";

// SPEC-113: the model provider settings — the surface of the experiments that
// use a model, shown in Labs only while one is on (LabsSection gates and
// lazily loads this, and loading it is what loads the gateway). Where model
// requests go and how they are bounded; the API key, masked, in cloud mode
// only; "Test connection"; and this session's activity log (the standing
// consents are LabsSection's, listed even while this panel is not). Every
// edit goes through the gateway, which owns the rules — the key dies with
// the origin, consent with the origin or model.
//
// The .vue script is parsed by espree: no TS-only syntax (see AppShell.vue).
const gateway = useModelGateway();
const { preferences } = usePreferences();
const { isEnabled } = useExperiments();

const config = computed(() => preferences.modelProvider);
const usedBy = computed(() =>
  EXPERIMENTS.filter((experiment) => MODEL_EXPERIMENTS.includes(experiment.id) && isEnabled(experiment.id)),
);

const modeOptions = [
  { value: "off", label: "Off" },
  { value: "local", label: "On this computer" },
  { value: "cloud", label: "Hosted endpoint" },
];

const endpoint = computed(() => resolveEndpoint(config.value.baseUrl));
const destination = computed(() => {
  const resolved = endpoint.value;
  if (!resolved.ok) {
    return "";
  }
  return resolved.endpoint.loopback
    ? `Requests go to ${resolved.endpoint.origin}, on this computer.`
    : `Requests go to ${resolved.endpoint.origin}, a third party that is not part of this product.`;
});
const baseUrlError = computed(() =>
  config.value.baseUrl.trim() === "" || endpoint.value.ok ? "" : endpoint.value.reason,
);

const apiKey = ref(gateway.getApiKey());
const keyNotice = ref("");
const keyError = ref("");

// Test connection: "" (not run), "running", "ok" or "error".
const testState = ref("");
const testMessage = ref("");

/** The line explaining why a key went: it dies with its origin. */
function keyRemovedNotice(from = "", to = "") {
  return (
    `The API key was removed because the endpoint moved from ${from || "no valid address"} ` +
    `to ${to || "no valid address"}. A key is only ever sent to the origin it was entered for — ` +
    "enter the key for the new endpoint if it needs one."
  );
}

// A config change made elsewhere (another tab, a reset) may have discarded
// the key: the field never shows a key the gateway would not send, and says
// why one went.
watch(
  () => originOf(config.value.baseUrl),
  (origin, previous) => {
    const shown = apiKey.value;
    apiKey.value = gateway.getApiKey();
    if (shown !== "" && apiKey.value === "" && origin !== previous) {
      keyNotice.value = keyRemovedNotice(previous ?? "", origin ?? "");
    }
  },
);

function textValue(event = new Event("change")) {
  const target = event.target;
  return target instanceof HTMLInputElement ? target.value : "";
}

/** Puts an input back to the stored value after a rejected edit. */
function revert(event = new Event("change"), value = "") {
  const target = event.target;
  if (target instanceof HTMLInputElement) {
    target.value = value;
  }
}

function modeValue(value = "") {
  return value === "local" ? "local" : value === "cloud" ? "cloud" : "off";
}

function commitBaseUrl(event = new Event("change")) {
  const change = gateway.updateProviderConfig({ baseUrl: textValue(event).trim() });
  if (change.keyDiscarded) {
    apiKey.value = "";
    keyNotice.value = keyRemovedNotice(change.previousOrigin ?? "", change.origin ?? "");
  }
}

function commitKey(event = new Event("change")) {
  keyError.value = "";
  try {
    gateway.setApiKey(textValue(event));
    keyNotice.value = "";
  } catch (error) {
    keyError.value = gateway.describeModelError(error);
  }
  apiKey.value = gateway.getApiKey();
}

function commitTimeout(event = new Event("change")) {
  const ms = Math.round(Number(textValue(event)) * 1000);
  if (Number.isInteger(ms) && ms >= MIN_TIMEOUT_MS && ms <= MAX_TIMEOUT_MS) {
    gateway.updateProviderConfig({ timeoutMs: ms });
  } else {
    revert(event, String(config.value.timeoutMs / 1000));
  }
}

function commitMaxTokens(event = new Event("change")) {
  const tokens = Number(textValue(event));
  if (Number.isInteger(tokens) && tokens >= 1 && tokens <= MAX_OUTPUT_TOKENS_LIMIT) {
    gateway.updateProviderConfig({ maxOutputTokens: tokens });
  } else {
    revert(event, String(config.value.maxOutputTokens));
  }
}

function commitTemperature(event = new Event("change")) {
  const temperature = Number(textValue(event));
  if (Number.isFinite(temperature) && temperature >= 0 && temperature <= MAX_TEMPERATURE) {
    gateway.updateProviderConfig({ temperature });
  } else {
    revert(event, String(config.value.temperature));
  }
}

async function runTest() {
  testState.value = "running";
  testMessage.value = "Testing…";
  try {
    const result = await gateway.testConnection();
    testState.value = "ok";
    testMessage.value = `Reachable — model ${result.model}, ${result.elapsedMs} ms round trip.`;
  } catch (error) {
    testState.value = "error";
    testMessage.value = gateway.describeModelError(error);
  }
}
</script>

<template>
  <div
    class="provider"
    data-testid="ai-provider"
    role="group"
    aria-labelledby="ai-provider-heading"
  >
    <h4
      id="ai-provider-heading"
      class="provider-heading"
    >
      Model provider
      <ExperimentBadge :experiment="usedBy[0]?.id ?? ''" />
    </h4>
    <p class="provider-note">
      Used by {{ usedBy.map((experiment) => experiment.title).join(", ") }}. Requests go straight from this
      browser to the OpenAI-compatible endpoint you name, only when you ask, and only after you have seen
      what will be sent. No model is bundled or downloaded, and none is chosen for you.
    </p>

    <UiField
      label="Provider"
      for="ai-mode"
      hint="Off sends nothing. Default: Off"
    >
      <UiSelect
        id="ai-mode"
        data-testid="ai-mode"
        aria-describedby="ai-mode-hint"
        :model-value="config.mode"
        :options="modeOptions"
        @update:model-value="(value) => gateway.updateProviderConfig({ mode: modeValue(value) })"
      />
    </UiField>

    <template v-if="config.mode !== 'off'">
      <UiField
        class="field-wide"
        label="Base URL"
        for="ai-base-url"
        hint="The endpoint’s OpenAI-compatible base, up to its version path — requests go to {base URL}/chat/completions and {base URL}/models."
        :error="baseUrlError"
      >
        <input
          id="ai-base-url"
          type="url"
          class="ai-text"
          data-testid="ai-base-url"
          spellcheck="false"
          autocomplete="off"
          :aria-describedby="baseUrlError ? 'ai-base-url-error' : 'ai-base-url-hint'"
          :value="config.baseUrl"
          @change="commitBaseUrl"
        >
      </UiField>
      <p
        v-if="destination"
        class="destination"
        data-testid="ai-origin"
      >
        {{ destination }}
      </p>

      <UiField
        class="field-wide"
        label="Model"
        for="ai-model"
        hint="The model id your endpoint serves it under."
      >
        <input
          id="ai-model"
          type="text"
          class="ai-text"
          data-testid="ai-model"
          spellcheck="false"
          autocomplete="off"
          aria-describedby="ai-model-hint"
          :value="config.model"
          @change="(event) => gateway.updateProviderConfig({ model: textValue(event).trim() })"
        >
      </UiField>

      <UiField
        v-if="config.mode === 'cloud'"
        class="field-wide"
        label="API key"
        for="ai-api-key"
        :hint="`Kept in this browser under its own key — never in preferences or a document — and sent only as Authorization: Bearer to ${endpoint.ok ? endpoint.endpoint.origin : 'the endpoint above'}.`"
        :error="keyError"
      >
        <input
          id="ai-api-key"
          type="password"
          class="ai-text"
          data-testid="ai-api-key"
          spellcheck="false"
          autocomplete="off"
          :aria-describedby="keyError ? 'ai-api-key-error' : 'ai-api-key-hint'"
          :value="apiKey"
          @change="commitKey"
        >
      </UiField>
      <p
        v-if="keyNotice"
        class="key-notice"
        role="status"
        data-testid="ai-api-key-cleared"
      >
        {{ keyNotice }}
      </p>

      <UiField
        label="Timeout (seconds)"
        for="ai-timeout"
        hint="A request with no complete answer by then is stopped. Default: 60"
      >
        <input
          id="ai-timeout"
          type="number"
          class="ai-number"
          data-testid="ai-timeout"
          :min="MIN_TIMEOUT_MS / 1000"
          :max="MAX_TIMEOUT_MS / 1000"
          aria-describedby="ai-timeout-hint"
          :value="config.timeoutMs / 1000"
          @change="commitTimeout"
        >
      </UiField>
      <UiField
        label="Max output tokens"
        for="ai-max-tokens"
        hint="Default: 4096"
      >
        <input
          id="ai-max-tokens"
          type="number"
          class="ai-number"
          data-testid="ai-max-tokens"
          min="1"
          :max="MAX_OUTPUT_TOKENS_LIMIT"
          aria-describedby="ai-max-tokens-hint"
          :value="config.maxOutputTokens"
          @change="commitMaxTokens"
        >
      </UiField>
      <UiField
        label="Temperature"
        for="ai-temperature"
        hint="Lower is more predictable. Default: 0.2"
      >
        <input
          id="ai-temperature"
          type="number"
          class="ai-number"
          data-testid="ai-temperature"
          min="0"
          :max="MAX_TEMPERATURE"
          step="0.1"
          aria-describedby="ai-temperature-hint"
          :value="config.temperature"
          @change="commitTemperature"
        >
      </UiField>

      <div class="test-row">
        <UiButton
          data-testid="ai-test"
          :disabled="testState === 'running'"
          @click="runTest"
        >
          Test connection
        </UiButton>
        <p
          class="test-result"
          :class="{ 'is-error': testState === 'error', 'is-ok': testState === 'ok' }"
          role="status"
          data-testid="ai-test-result"
          :data-state="testState"
        >
          {{ testMessage || "Sends one trivial request (no document content) and reports what came back." }}
        </p>
      </div>
    </template>

    <ModelActivityLog />
  </div>
</template>

<style scoped>
.provider {
  display: flex;
  flex-direction: column;
  gap: var(--space-3);
  padding: var(--space-4);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
}

.provider-heading {
  margin: 0;
  display: flex;
  align-items: center;
  gap: var(--space-2);
  color: var(--text-primary);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
}

.provider-note,
.destination,
.test-result {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.destination,
.test-result.is-ok {
  color: var(--text-primary);
}

.test-result.is-error,
.key-notice {
  color: var(--text-danger);
}

.key-notice {
  margin: 0;
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.ai-text,
.ai-number {
  height: var(--height-control);
  padding: 0 var(--space-3);
  color: var(--text-primary);
  background: var(--surface-base);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  font-family: var(--font-ui);
  font-size: var(--text-control);
  box-sizing: border-box;
}

.ai-text {
  width: 100%;
}

.ai-number {
  width: 7rem;
}

.test-row {
  display: flex;
  align-items: center;
  gap: var(--space-3);
}
</style>

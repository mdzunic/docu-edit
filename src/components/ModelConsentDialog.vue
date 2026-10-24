<script setup lang="ts">
import { computed } from "vue";
import BaseDialog from "./BaseDialog.vue";
import UiButton from "./ui/UiButton.vue";

// SPEC-113: the consent asked before a model request reaches an origin with
// no standing grant. It names the exact destination, says plainly whether
// that is this machine or a third party, and shows the exact request body —
// the very string the gateway then sends — in a scrollable, selectable
// region with its byte count. "Send once" sends this request; "Always allow"
// also stores a grant for this origin, model and feature (revocable in
// Labs); Cancel, Escape and the backdrop send nothing.
//
// Hosted by AppShell, loaded only when a request is waiting. Runtime prop
// declarations for the espree constraint AppShell.vue explains.
const props = defineProps({
  /** The gateway's ConsentRequest: origin, local, purpose, model, method,
   * url, body, bytes, sendsKey. */
  request: { type: Object, required: true },
});

const emit = defineEmits(["answer"]);

// Reflect.get reads the untyped Object prop without a TS cast (barred in
// .vue scripts); String()/Number() coerce for display only.
const field = (name = "") => Reflect.get(props.request, name);
const origin = computed(() => String(field("origin")));
const local = computed(() => field("local") === true);
const body = computed(() => String(field("body") ?? ""));
const bytes = computed(() => Number(field("bytes") ?? 0));
const sendsKey = computed(() => field("sendsKey") === true);
</script>

<template>
  <BaseDialog
    title="Send this to a model?"
    testid="ai-consent"
    width="36rem"
    @dismiss="emit('answer', 'cancel')"
  >
    <p
      class="destination"
      :class="{ 'is-local': local }"
      data-testid="ai-consent-destination"
    >
      <template v-if="local">
        This stays on your computer: the request goes to {{ origin }}, a server running on this machine.
      </template>
      <template v-else>
        Your text will be sent to {{ origin }}, which is not part of this product.
      </template>
    </p>
    <dl class="facts">
      <dt>For</dt>
      <dd data-testid="ai-consent-purpose">
        {{ field("purpose") }}
      </dd>
      <dt>Model</dt>
      <dd>{{ field("model") || "—" }}</dd>
      <dt>Request</dt>
      <dd class="request-line">
        {{ field("method") }} {{ field("url") }}
      </dd>
    </dl>
    <p
      v-if="sendsKey"
      class="key-note"
    >
      Your API key goes with it, as the Authorization header, to {{ origin }} only.
    </p>
    <p
      class="payload-label"
      data-testid="ai-consent-bytes"
    >
      Exactly this will be sent — {{ bytes }} bytes:
    </p>
    <pre
      class="payload"
      data-testid="ai-consent-payload"
      tabindex="0"
      aria-label="The exact request body"
    >{{ body === "" ? "(no request body)" : body }}</pre>
    <p class="always-note">
      “Always allow” stops asking for {{ origin }}, this model and this feature until you revoke it in
      Preferences › Labs, or change any of them.
    </p>

    <template #actions>
      <UiButton
        variant="secondary"
        data-testid="ai-consent-cancel"
        autofocus
        @click="emit('answer', 'cancel')"
      >
        Cancel
      </UiButton>
      <UiButton
        variant="secondary"
        data-testid="ai-consent-once"
        @click="emit('answer', 'once')"
      >
        Send once
      </UiButton>
      <UiButton
        variant="primary"
        data-testid="ai-consent-always"
        @click="emit('answer', 'always')"
      >
        Always allow
      </UiButton>
    </template>
  </BaseDialog>
</template>

<style scoped>
.destination {
  margin: 0 0 var(--space-4);
  padding: var(--space-3);
  border: var(--border-width) solid var(--warning);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--text-primary);
  font-size: var(--text-md);
  font-weight: var(--weight-bold);
  line-height: var(--line-normal);
}

.destination.is-local {
  border-color: var(--success);
}

.facts {
  display: grid;
  grid-template-columns: auto minmax(0, 1fr);
  column-gap: var(--space-4);
  row-gap: var(--space-1);
  margin: 0 0 var(--space-3);
  font-size: var(--text-sm);
}

.facts dt {
  color: var(--text-secondary);
}

.facts dd {
  margin: 0;
  color: var(--text-primary);
  overflow-wrap: anywhere;
}

.request-line {
  font-family: var(--font-mono);
}

.key-note,
.payload-label,
.always-note {
  margin: 0 0 var(--space-2);
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.always-note {
  margin: var(--space-3) 0 0;
}

/* Scrollable and selectable: the user can read, select and copy every byte. */
.payload {
  max-height: 16rem;
  overflow: auto;
  margin: 0;
  padding: var(--space-3);
  border: var(--border-width) solid var(--border-default);
  border-radius: var(--radius-sm);
  background: var(--surface-sunken);
  color: var(--text-primary);
  font-family: var(--font-mono);
  font-size: var(--text-sm);
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  user-select: text;
}
</style>

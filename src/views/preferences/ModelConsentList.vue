<script setup lang="ts">
import { onMounted, onUnmounted, ref } from "vue";
import UiButton from "../../components/ui/UiButton.vue";
import { consentGrants, onConsentChange, revokeConsent } from "../../lib/ai/consent";

// SPEC-113: the standing model consents — each "Always allow" the user gave,
// one per origin, model and feature — listed in Labs with a Revoke each, so
// the next such request asks again. Shown by LabsSection while a model
// experiment is on, and whenever a grant is still stored, so a consent can
// be revoked even after its experiment was switched off. Reads only the
// consent store; loads nothing that can send a request.
const props = defineProps({
  /** Render even with no grant stored (while a model experiment is on).
   * Otherwise the list appears only when there is something to revoke. */
  always: { type: Boolean, default: false },
});

const grants = ref(consentGrants());

let stop = () => {};
onMounted(() => {
  grants.value = consentGrants();
  stop = onConsentChange(() => {
    grants.value = consentGrants();
  });
});
onUnmounted(() => stop());
</script>

<template>
  <div
    v-if="props.always || grants.length > 0"
    class="consents"
    data-testid="ai-consents"
  >
    <h4 class="consents-heading">
      Standing model consent
    </h4>
    <p
      v-if="grants.length === 0"
      class="muted"
    >
      None — every model request asks first.
    </p>
    <ul
      v-else
      class="consent-list"
    >
      <li
        v-for="grant in grants"
        :key="`${grant.origin} ${grant.feature}`"
        class="consent-item"
        data-testid="ai-consent-grant"
      >
        <span>{{ grant.origin }} — always allowed for “{{ grant.feature }}” with {{ grant.model || "any model" }}</span>
        <UiButton
          size="sm"
          data-testid="ai-consent-revoke"
          @click="revokeConsent(grant.origin, grant.feature)"
        >
          Revoke
        </UiButton>
      </li>
    </ul>
  </div>
</template>

<style scoped>
.consents {
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.consents-heading {
  margin: 0;
  color: var(--text-primary);
  font-size: var(--text-sm);
  font-weight: var(--weight-bold);
}

.muted {
  margin: 0;
  color: var(--text-secondary);
  font-size: var(--text-sm);
  line-height: var(--line-normal);
}

.consent-list {
  margin: 0;
  padding: 0;
  list-style: none;
  display: flex;
  flex-direction: column;
  gap: var(--space-2);
}

.consent-item {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: var(--space-3);
  font-size: var(--text-sm);
  overflow-wrap: anywhere;
}
</style>

<script setup lang="ts">
import { useToasts } from "../composables/useToasts";

const { toasts } = useToasts();
</script>

<template>
  <div
    class="toast-region"
    aria-live="polite"
    data-testid="toast-region"
  >
    <div
      v-for="toast in toasts"
      :key="toast.id"
      class="toast"
      role="status"
      data-testid="toast"
    >
      {{ toast.message }}
    </div>
  </div>
</template>

<style scoped>
.toast-region {
  position: fixed;
  bottom: var(--space-6);
  left: 50%;
  transform: translateX(-50%);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: var(--space-3);
  z-index: 100;
  pointer-events: none;
}

.toast {
  padding: var(--space-3) var(--space-6);
  border-radius: var(--radius-sm);
  /* An inverted chip: --text-primary and --surface-base flip together across
     themes, so the pair stays the tested ≥4.5:1 body contrast in both (dark
     text on a light chip in dark theme, light text on a dark chip in light).
     Using the fixed --text-on-brand here instead would leave white text on the
     light dark-theme chip — invisible (SPEC-059 AC-40). */
  background: var(--text-primary);
  color: var(--surface-base);
  font-size: var(--text-md);
  box-shadow: var(--shadow-2);
}
</style>

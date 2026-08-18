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
  background: var(--text-primary);
  color: var(--text-on-brand);
  font-size: var(--text-md);
  box-shadow: var(--shadow-2);
}
</style>

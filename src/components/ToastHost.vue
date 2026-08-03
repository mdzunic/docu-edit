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
  bottom: 1.5rem;
  left: 50%;
  transform: translateX(-50%);
  display: flex;
  flex-direction: column;
  align-items: center;
  gap: 0.5rem;
  z-index: 100;
  pointer-events: none;
}

.toast {
  padding: 0.625rem 1.25rem;
  border-radius: 4px;
  background: #202124;
  color: #fff;
  font-size: 0.875rem;
  box-shadow: 0 2px 6px rgb(0 0 0 / 0.3);
}
</style>

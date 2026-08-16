<script setup lang="ts">
// SPEC-048 AC-19: the font-loading skeleton. While the substitute/embedded
// faces load, a soft shimmer covers the page column rather than flashing
// unstyled, wrong-metric text; the pages under it are already in the DOM and
// re-measure the moment the fonts resolve. Never intercepts input
// (pointer-events: none) so it cannot delay editing.
</script>

<template>
  <div
    class="fonts-skeleton print-exclude"
    data-testid="fonts-skeleton"
    aria-hidden="true"
  />
</template>

<style scoped>
.fonts-skeleton {
  position: absolute;
  inset: 2rem 1rem;
  z-index: 15;
  pointer-events: none;
  border-radius: 4px;
  background: linear-gradient(100deg, #f1f3f4 30%, #e8eaed 50%, #f1f3f4 70%);
  background-size: 200% 100%;
  animation: fonts-skeleton-shimmer 1.1s ease-in-out infinite;
}

@keyframes fonts-skeleton-shimmer {
  from {
    background-position: 200% 0;
  }
  to {
    background-position: -200% 0;
  }
}
</style>

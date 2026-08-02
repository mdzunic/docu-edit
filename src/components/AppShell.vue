<script setup lang="ts">
// Runtime prop declaration (not the defineProps<{...}>() generic) because
// ESLint parses .vue scripts with espree, which rejects TS-only syntax;
// vue-tsc still infers `documentTitle: string` from it.
defineProps({
  /** Shown in the app bar next to the product name. */
  documentTitle: { type: String, required: true },
});
</script>

<template>
  <div class="shell">
    <header
      class="app-bar"
      data-testid="app-bar"
    >
      <span
        class="product-name"
        data-testid="product-name"
      >docu-edit</span>
      <h1
        class="document-title"
        data-testid="document-title"
      >
        {{ documentTitle }}
      </h1>
    </header>
    <div
      class="toolbar"
      role="toolbar"
      aria-label="Document toolbar"
      data-testid="toolbar"
    >
      <slot name="toolbar" />
    </div>
    <main
      class="canvas"
      data-testid="canvas"
    >
      <slot />
    </main>
  </div>
</template>

<style scoped>
/* The shell owns the viewport: it fills #app exactly and hides its own
   overflow, so the page never grows a scrollbar — only .canvas scrolls. */
.shell {
  height: 100%;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.app-bar {
  flex: none;
  display: flex;
  align-items: baseline;
  gap: 0.75rem;
  padding: 0.5rem 1rem;
  background: #2c3e50;
  color: #fff;
}

.product-name {
  font-weight: 600;
}

.document-title {
  margin: 0;
  font-size: 1rem;
  font-weight: 400;
}

/* A single non-wrapping row: contents that overflow the viewport width are
   reached by scrolling the strip, never by wrapping or clipping. */
.toolbar {
  flex: none;
  display: flex;
  flex-wrap: nowrap;
  align-items: center;
  gap: 0.25rem;
  min-height: 2.5rem;
  padding: 0.25rem 1rem;
  overflow-x: auto;
  overflow-y: hidden;
  white-space: nowrap;
  background: #f5f5f5;
  border-bottom: 1px solid #d0d0d0;
  box-sizing: border-box;
}

/* min-height: 0 lets the canvas shrink below its content height, so tall
   content scrolls here instead of stretching the shell past the viewport. */
.canvas {
  flex: 1 1 0;
  min-height: 0;
  overflow-y: auto;
  background: #e8eaed;
}
</style>

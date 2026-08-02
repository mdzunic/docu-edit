<script setup lang="ts">
import { computed } from "vue";
import { useRoute } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { documentExists } from "../lib/documents";
import { formatTimestamp } from "../lib/format";

const route = useRoute();
const documentId = computed(() => String(route.params.id ?? ""));
const exists = computed(() => documentExists(documentId.value));

// SPEC-001 decision: with no persistence yet, the document title is simply
// the client-side time at page load, captured once when the view mounts.
const documentTitle = formatTimestamp(new Date());
</script>

<template>
  <AppShell
    v-if="exists"
    :document-title="documentTitle"
  >
    <div class="page-wrap">
      <div
        class="page"
        data-testid="document-page"
      />
    </div>
  </AppShell>
  <AppShell
    v-else
    document-title="No document"
  >
    <div
      class="not-found"
      data-testid="not-found"
    >
      <p class="not-found-message">
        Document not found
      </p>
      <RouterLink to="/">
        Back to home
      </RouterLink>
    </div>
  </AppShell>
</template>

<style scoped>
.page-wrap {
  padding: 2rem 1rem;
}

/* A blank page surface — later specs render document content onto it. */
.page {
  width: min(51rem, 100%);
  min-height: 66rem;
  margin: 0 auto;
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3);
}

.not-found {
  height: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  gap: 0.5rem;
  color: #3c4043;
}

.not-found-message {
  margin: 0;
  font-size: 1.125rem;
}
</style>

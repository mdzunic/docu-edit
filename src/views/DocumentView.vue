<script setup lang="ts">
import { computed, ref, watch } from "vue";
import { useRoute } from "vue-router";
import AppShell from "../components/AppShell.vue";
import { documentExists } from "../lib/documents";
import { getOrCreateDocument } from "../lib/documentStore";
import { DEFAULT_TITLE } from "../lib/documentModel";

const route = useRoute();
const documentId = computed(() => String(route.params.id ?? ""));

// A plain object assigned into a ref is made deeply reactive by Vue, so
// mutating doc.value.title below updates the app bar without extra plumbing.
// The initial value is computed inline (rather than starting from `null`)
// so TS infers `Ref<DocumentModel | null>` from this expression alone —
// `ref<DocumentModel | null>(null)` would need a generic, which ESLint's
// espree parser for .vue scripts can't parse.
const doc = ref(
  documentExists(documentId.value)
    ? getOrCreateDocument(documentId.value)
    : null,
);

watch(documentId, (id) => {
  doc.value = documentExists(id) ? getOrCreateDocument(id) : null;
});

const pageStyle = computed(() => {
  if (!doc.value) {
    return {};
  }
  const { widthCm, heightCm } = doc.value.defaultStyles.pageSize;
  return { width: `${widthCm}cm`, height: `${heightCm}cm` };
});

const marginStyle = computed(() => {
  if (!doc.value) {
    return {};
  }
  const { topCm, rightCm, bottomCm, leftCm } = doc.value.defaultStyles.margins;
  return {
    top: `${topCm}cm`,
    right: `${rightCm}cm`,
    bottom: `${bottomCm}cm`,
    left: `${leftCm}cm`,
  };
});

function handleTitleUpdate(newTitle = "") {
  if (!doc.value) {
    return;
  }
  const trimmed = newTitle.trim();
  doc.value.title = trimmed.length > 0 ? trimmed : DEFAULT_TITLE;
}
</script>

<template>
  <AppShell
    v-if="doc"
    :document-title="doc.title"
    title-editable
    @update:document-title="handleTitleUpdate"
  >
    <div class="page-wrap">
      <div
        class="page"
        data-testid="document-page"
        :style="pageStyle"
      >
        <div
          class="page-margin"
          data-testid="page-margin"
          :style="marginStyle"
        />
      </div>
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

/* A page-shaped surface sized to the document's own defaultStyles (in real
   centimetres via CSS's absolute cm unit) — later specs render content onto
   it. .page-margin marks where the page margins end and content begins. */
.page {
  position: relative;
  margin: 0 auto;
  background: #fff;
  box-shadow: 0 1px 3px rgb(0 0 0 / 0.3);
  box-sizing: border-box;
}

.page-margin {
  position: absolute;
  border: 1px dashed #c8c8c8;
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

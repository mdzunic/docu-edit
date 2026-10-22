// The SPEC-060 returning highlight, extracted from HomeView: when the user
// closes a document, the library scrolls its row into view and briefly
// highlights it (skipped under reduced motion). Returns the id to highlight,
// "" when none — the list and grid take it as their `highlightId`.
import { nextTick, onMounted, ref, type Ref } from "vue";
import { motionBaseMs, prefersReducedMotion } from "../../lib/core/motion";

export function useReturningHighlight(): Ref<string> {
  const highlightId = ref("");

  onMounted(async () => {
    const closedId = window.history.state?.closedDocId;
    if (typeof closedId !== "string" || closedId.length === 0) {
      return;
    }
    await nextTick();
    const row = document.querySelector(`[data-doc-id="${closedId}"]`);
    if (row && typeof row.scrollIntoView === "function") {
      row.scrollIntoView({ block: "center" });
    }
    if (prefersReducedMotion()) {
      return;
    }
    highlightId.value = closedId;
    setTimeout(() => {
      if (highlightId.value === closedId) {
        highlightId.value = "";
      }
    }, motionBaseMs());
  });

  return highlightId;
}

// SPEC-060 — closing the document is a commit boundary. These cover the
// orchestration in useFileActions: the Close control flushes before it
// navigates and only navigates on success, a failed flush stays put with an
// error toast, and an unavailable store raises the confirmation dialog instead
// of losing work silently. The real flush primitive (commitDocument, debounce
// cancellation) is covered in useAutosave.spec.ts; here flushForClose is a
// stub so the branch logic is isolated.
import { defineComponent, h, shallowRef } from "vue";
import { mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { afterEach, describe, expect, it, vi } from "vitest";
import { routes } from "../../router";
import { createDocument, type DocumentModel } from "../../lib/documentModel";

// isStorageAvailable is the one document-store function whose value the close
// path branches on; keep the rest of the module real.
const storage = vi.hoisted(() => ({ available: true }));
vi.mock("../../lib/documentStore", async (importActual) => {
  const actual =
    await importActual<typeof import("../../lib/documentStore")>();
  return { ...actual, isStorageAvailable: () => storage.available };
});

import { useFileActions } from "./useFileActions";

afterEach(() => {
  storage.available = true;
  vi.restoreAllMocks();
});

function mountHarness(flushForClose = vi.fn(async () => true)) {
  const doc = shallowRef<DocumentModel>(createDocument("doc-1"));
  doc.value.title = "My doc";
  const addToast = vi.fn();
  let api: ReturnType<typeof useFileActions> | null = null;
  const Harness = defineComponent({
    setup() {
      api = useFileActions({
        doc,
        addToast,
        exportPdf: vi.fn(async () => {}),
        commitNow: vi.fn(async () => {}),
        flushForClose,
        noteVersionHistoryCommitPoint: vi.fn(),
      } as unknown as Parameters<typeof useFileActions>[0]);
      return () => h("div");
    },
  });
  const router = createRouter({ history: createMemoryHistory(), routes });
  const push = vi.spyOn(router, "push").mockResolvedValue(undefined);
  mount(Harness, { global: { plugins: [router] } });
  return { api: api!, doc, addToast, flushForClose, push };
}

describe("useFileActions — closing the document (SPEC-060)", () => {
  it("flushes before navigating, and navigates with the closed id in router state (AC-3, AC-7, AC-11)", async () => {
    const { api, doc, flushForClose, push } = mountHarness();
    await api.closeOpenDocument();
    expect(flushForClose).toHaveBeenCalledTimes(1);
    expect(push).toHaveBeenCalledWith({
      path: "/",
      state: { closedDocId: doc.value.id },
    });
    // The flush resolved before the navigation was requested (AC-11).
    expect(flushForClose.mock.invocationCallOrder[0]).toBeLessThan(
      push.mock.invocationCallOrder[0],
    );
  });

  it("shows a busy state while the flush is in flight and clears it after (AC-9)", async () => {
    let resolveFlush: (ok: boolean) => void = () => {};
    const flush = vi.fn(
      () => new Promise<boolean>((resolve) => (resolveFlush = resolve)),
    );
    const { api, push } = mountHarness(flush);
    const pending = api.closeOpenDocument();
    await Promise.resolve();
    expect(api.closing.value).toBe(true);
    expect(push).not.toHaveBeenCalled();
    resolveFlush(true);
    await pending;
    expect(api.closing.value).toBe(false);
    expect(push).toHaveBeenCalledTimes(1);
  });

  it("does not navigate when the flush fails; toasts the failure and stays open (AC-12, AC-13, AC-14)", async () => {
    const { api, addToast, push } = mountHarness(vi.fn(async () => false));
    await api.closeOpenDocument();
    expect(push).not.toHaveBeenCalled();
    expect(addToast).toHaveBeenCalledTimes(1);
    expect(addToast.mock.calls[0][0]).toContain("My doc");
    expect(api.closing.value).toBe(false);
  });

  it("raises the confirmation dialog instead of flushing when storage is unavailable (AC-16, AC-17)", async () => {
    storage.available = false;
    const { api, flushForClose, push } = mountHarness();
    await api.closeOpenDocument();
    expect(api.storageWarningOpen.value).toBe(true);
    expect(flushForClose).not.toHaveBeenCalled();
    expect(push).not.toHaveBeenCalled();
  });

  it("'Close anyway' dismisses the dialog and navigates, skipping the flush (AC-17)", async () => {
    storage.available = false;
    const { api, doc, flushForClose, push } = mountHarness();
    await api.closeOpenDocument();
    api.confirmCloseAnyway();
    expect(api.storageWarningOpen.value).toBe(false);
    expect(flushForClose).not.toHaveBeenCalled();
    expect(push).toHaveBeenCalledWith({
      path: "/",
      state: { closedDocId: doc.value.id },
    });
  });
});

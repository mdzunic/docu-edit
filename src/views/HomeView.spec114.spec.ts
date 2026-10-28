// SPEC-114 AC — with the `library-search` flag off the scope control does not
// render, no index is built, no index record is read or written, and the
// library behaves exactly as SPEC-067 specifies. The real lib/idb.ts runs
// over the logging in-memory IndexedDB, holding an index record left by an
// earlier session with the flag on, so "not read, not written" is observed
// rather than assumed. With the flag on, the same library searches inside
// its documents.
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { flushPromises, mount } from "@vue/test-utils";
import { createMemoryHistory, createRouter } from "vue-router";
import { FakeIndexedDb } from "../lib/core/fakeIndexedDb";

let fake: FakeIndexedDb;

beforeEach(() => {
  vi.resetModules();
  localStorage.clear();
  fake = new FakeIndexedDb();
  fake.install();
});

afterEach(() => {
  fake.uninstall();
  vi.restoreAllMocks();
});

/** A stored text document whose body holds `body`. */
async function seedDocument(id: string, title: string, body: string, lastModified: number) {
  const { createDocument } = await import("../lib/documentModel");
  const model = createDocument(id);
  model.title = title;
  model.sections[0].blocks = [{ type: "paragraph", runs: [{ text: body, formatting: {} }] }];
  const documents = fake.stores.get("documents") ?? new Map<string, unknown>();
  fake.stores.set("documents", documents);
  documents.set(id, { id, format: "docx", lastModified, model: JSON.parse(JSON.stringify(model)) });
}

async function mountLibrary(flag: boolean) {
  const { setExperimentEnabled } = await import("../lib/experiments");
  setExperimentEnabled("library-search", flag);
  const { initLibrary } = await import("../lib/documentStore");
  await initLibrary();
  const { routes } = await import("../router");
  const { default: HomeView } = await import("./HomeView.vue");
  const router = createRouter({ history: createMemoryHistory(), routes });
  await router.push("/");
  const wrapper = mount(HomeView, { global: { plugins: [router] }, attachTo: document.body });
  await flushPromises();
  return { wrapper, router };
}

const indexKeysTouched = (): string[] =>
  fake.log.filter((event) => (event.kind === "read" || event.kind === "issue") && event.key?.startsWith("index:")).map((event) => `${event.op ?? event.kind} ${event.key}`);

describe("with the library-search flag off", () => {
  it("renders no scope control, builds no index, reads and writes no index record — the library is SPEC-067's", async () => {
    await seedDocument("a", "Harbour report", "The Rotterdam figures rose.", 2000);
    await seedDocument("b", "Budget", "Nothing about ports.", 1000);
    // An earlier session, with the flag on, left an index record behind.
    fake.stores.get("documents")!.set("index:a", { id: "index:a", docId: "a", indexVersion: 1, lastModified: 2000, tokens: null });
    const { wrapper } = await mountLibrary(false);

    expect(wrapper.find('[data-testid="search-scope"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="library-search-panel"]').exists()).toBe(false);
    expect(wrapper.find('[data-testid="search-index-progress"]').exists()).toBe(false);

    // SPEC-067's title search, untouched: a word only in a document's BODY
    // matches nothing; a title subsequence filters the rows.
    await wrapper.get('[data-testid="library-search"]').setValue("rotterdam");
    await flushPromises();
    expect(wrapper.findAll('[data-testid="row-title"]')).toHaveLength(0);
    expect(wrapper.find('[data-testid="no-results"]').exists()).toBe(true);
    await wrapper.get('[data-testid="library-search"]').setValue("hrbr");
    await flushPromises();
    expect(wrapper.findAll('[data-testid="row-title"]').map((cell) => cell.text())).toEqual(["Harbour report"]);
    expect(wrapper.findAll('[data-testid="title-match"]').length).toBeGreaterThan(0);
    await new Promise((resolve) => setTimeout(resolve, 300));
    await flushPromises();

    expect(indexKeysTouched()).toEqual([]);
    expect(fake.readKeys().some((key) => key.startsWith("index:"))).toBe(false);
    wrapper.unmount();
  });
});

describe("with the library-search flag on", () => {
  it("searches inside the documents in Everything scope, indexing them in the background", async () => {
    await seedDocument("a", "Harbour report", "The Rotterdam figures rose in the third quarter.", 2000);
    await seedDocument("b", "Budget", "Nothing about ports.", 1000);
    const { wrapper } = await mountLibrary(true);

    await vi.waitFor(() => expect(wrapper.find('[data-testid="search-scope"]').exists()).toBe(true));
    const scope = wrapper.get('[data-testid="search-scope"]');
    expect((scope.element as HTMLSelectElement).value).toBe("titles");
    expect(wrapper.find('[data-testid="experiment-badge"]').exists()).toBe(true);

    await scope.setValue("everything");
    await vi.waitFor(async () => {
      await flushPromises();
      expect(wrapper.get('[data-testid="search-index-progress"]').text()).toBe("Indexed 2 of 2");
    });
    expect(wrapper.get('[data-testid="search-index-progress"]').attributes("role")).toBe("status");

    await wrapper.get('[data-testid="library-search"]').setValue('"rotterdam figures"');
    await vi.waitFor(async () => {
      await flushPromises();
      expect(wrapper.findAll('[data-testid="search-result"]')).toHaveLength(1);
    });
    // The results stand in for the rows.
    expect(wrapper.findAll('[data-testid="row-title"]')).toHaveLength(0);
    const result = wrapper.get('[data-testid="search-result"]');
    expect(result.attributes("data-doc-id")).toBe("a");
    expect(result.get('[data-testid="search-result-kind"]').text()).toBe("DOCX");
    expect(result.get('[data-testid="search-result-count"]').text()).toBe("1 match");
    expect(result.findAll('[data-testid="snippet-match"]').map((mark) => mark.text())).toEqual(["Rotterdam figures"]);

    // Index records were written, one per document, under index:<id>.
    expect([...fake.stores.get("documents")!.keys()].filter((key) => key.startsWith("index:")).toSorted()).toEqual(["index:a", "index:b"]);
    wrapper.unmount();
  });
});

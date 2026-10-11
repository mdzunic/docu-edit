import { describe, it, expect, afterEach, beforeAll } from "vitest";
import { mount, flushPromises, type VueWrapper } from "@vue/test-utils";
import { nextTick } from "vue";
import { createMemoryHistory, createRouter } from "vue-router";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import AppShell from "./AppShell.vue";
import { routes } from "../router";
import { setTierWidthForTest } from "../composables/useResponsiveTier";

const fixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "lib", "docx", "fixtures");
const docxBytes = readFileSync(join(fixturesDir, "word-authored.docx"));

/** A window-level drag/drop Event carrying a fake DataTransfer — jsdom has
 * no DragEvent/DataTransfer constructors, so a plain Event is extended with
 * just the `dataTransfer` shape the handlers read (`types`, `files`). */
function fileDragEvent(type: string, files: File[] = []) {
  const event = new Event(type, { bubbles: true, cancelable: true });
  Object.defineProperty(event, "dataTransfer", { value: { types: ["Files"], files } });
  return event;
}

// AppShell listens on `window`, not its own DOM subtree, so a leftover
// mounted instance from a previous test would keep reacting to drag events
// dispatched by later tests; unmounting after each test removes those
// listeners (see AppShell's onUnmounted).
const mountedWrappers: VueWrapper[] = [];
afterEach(() => {
  mountedWrappers.splice(0).forEach((wrapper) => wrapper.unmount());
});

function mountWithRouter(documentTitle = "t") {
  const router = createRouter({ history: createMemoryHistory(), routes });
  const wrapper = mount(AppShell, { props: { documentTitle }, global: { plugins: [router] } });
  mountedWrappers.push(wrapper);
  return { wrapper, router };
}

// A drop's router.push() resolves only once its target route's async
// component has been dynamically import()ed, which needs real event-loop
// time, not just a promise-queue flush (same reason HomeView's library
// spec polls rather than awaiting flushPromises() alone). The window is
// generous (up to 10s) rather than tight: the drop handler also parses the
// dropped .docx before it navigates, and on a CPU-starved gate worker that
// parse plus the lazy editor import — even with the beforeAll cache-warm —
// can outlast a 2s poll, which is exactly the flake this widening removes.
// The window only costs wall-clock when navigation genuinely never happens;
// the assertion that follows still fails a real navigation break.
async function untilNavigated(router: ReturnType<typeof createRouter>) {
  for (let i = 0; i < 1000; i += 1) {
    if (router.currentRoute.value.path !== "/") {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 10));
  }
}

describe("AppShell", () => {
  it("renders the app bar with the product name and document title", () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "My doc" } });
    expect(wrapper.get('[data-testid="app-bar"]').isVisible()).toBe(true);
    expect(wrapper.get('[data-testid="product-name"]').text()).toBe(
      "docu-edit",
    );
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe("My doc");
  });

  it("renders a toolbar strip only when a view supplies one, and a canvas area", () => {
    // No toolbar slot (the library, the presentation editor): no empty strip.
    const bare = mount(AppShell, { props: { documentTitle: "t" } });
    expect(bare.find('[data-testid="toolbar"]').exists()).toBe(false);
    expect(bare.get('[data-testid="canvas"]').element.tagName).toBe("MAIN");

    const wrapper = mount(AppShell, {
      props: { documentTitle: "t" },
      slots: { toolbar: "<button>Bold</button>" },
    });
    expect(wrapper.get('[data-testid="toolbar"]').attributes("role")).toBe(
      "toolbar",
    );
  });

  it("places slotted content into the toolbar and canvas", () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "t" },
      slots: {
        toolbar: "<button>Bold</button>",
        default: "<p>canvas content</p>",
      },
    });
    expect(wrapper.get('[data-testid="toolbar"]').text()).toContain("Bold");
    expect(wrapper.get('[data-testid="canvas"]').text()).toContain(
      "canvas content",
    );
  });

  it("does not enter edit mode on click when titleEditable is unset", async () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    expect(wrapper.find('[data-testid="document-title-input"]').exists()).toBe(
      false,
    );
  });

  it("clicking the title enters edit mode with the current title (AC-6)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get<HTMLInputElement>(
      '[data-testid="document-title-input"]',
    );
    expect(input.element.value).toBe("My doc");
  });

  it("emits update:documentTitle on Enter (AC-7)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("New title");
    await input.trigger("keydown.enter");
    expect(wrapper.emitted("update:documentTitle")).toEqual([["New title"]]);
    expect(wrapper.find('[data-testid="document-title-input"]').exists()).toBe(
      false,
    );
  });

  it("emits update:documentTitle on blur (AC-7)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Blurred title");
    await input.trigger("blur");
    expect(wrapper.emitted("update:documentTitle")).toEqual([
      ["Blurred title"],
    ]);
  });

  it("Escape cancels editing without emitting (AC-8)", async () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "My doc", titleEditable: true },
    });
    await wrapper.get('[data-testid="document-title"]').trigger("click");
    const input = wrapper.get('[data-testid="document-title-input"]');
    await input.setValue("Discarded");
    await input.trigger("keydown.escape");
    expect(wrapper.emitted("update:documentTitle")).toBeUndefined();
    expect(wrapper.get('[data-testid="document-title"]').text()).toBe(
      "My doc",
    );
  });

  it("shows no format badge when formatBadge is unset", () => {
    const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
    expect(wrapper.find('[data-testid="format-badge"]').exists()).toBe(false);
  });

  it("shows the format badge text when set (AC-30)", () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "t", formatBadge: "DOCX" },
    });
    expect(wrapper.get('[data-testid="format-badge"]').text()).toBe("DOCX");
  });

  it("places app-bar-actions slot content inside the app bar (SPEC-006 AC-1)", () => {
    const wrapper = mount(AppShell, {
      props: { documentTitle: "t" },
      slots: { "app-bar-actions": "<button>Download</button>" },
    });
    expect(wrapper.get('[data-testid="app-bar"]').text()).toContain("Download");
  });

  describe("keyboard shortcuts dialog (SPEC-021 AC-6, AC-9, AC-12)", () => {
    it("Ctrl/Cmd+/ opens the dialog, and Escape closes it", async () => {
      const wrapper = mount(AppShell, { props: { documentTitle: "t" }, attachTo: document.body });
      document.dispatchEvent(new KeyboardEvent("keydown", { key: "/", ctrlKey: true, bubbles: true }));
      await nextTick();
      expect(wrapper.find('[data-testid="shortcuts-dialog"]').exists()).toBe(true);

      // Fired from the search input — where focus rests while the dialog is
      // open — bubbling to BaseDialog's panel-level Escape handler (the old
      // handler sat on the backdrop, which a real keypress never targets).
      await wrapper.get('[data-testid="shortcuts-search"]').trigger("keydown", { key: "Escape" });
      expect(wrapper.find('[data-testid="shortcuts-dialog"]').exists()).toBe(false);
      wrapper.unmount();
    });

    it("does not open while focus is on a text input outside the document (AC-12, AC-14)", async () => {
      const wrapper = mount(AppShell, { props: { documentTitle: "t" }, attachTo: document.body });
      const input = document.createElement("input");
      document.body.append(input);
      input.dispatchEvent(new KeyboardEvent("keydown", { key: "/", ctrlKey: true, bubbles: true }));
      await nextTick();
      expect(wrapper.find('[data-testid="shortcuts-dialog"]').exists()).toBe(false);
      input.remove();
      wrapper.unmount();
    });
  });

  describe("drag-and-drop (SPEC-009)", () => {
    // A drop navigates to /doc/:id, whose route component is a lazy
    // `() => import("./views/DocumentView.vue")`. The very first time that
    // module graph is imported in a test worker, Vite transforms the whole
    // (now large) editor dependency tree on demand, which takes a few seconds
    // — longer than untilNavigated's polling window, so the first navigating
    // test would flake. Pre-importing the lazy route components here pays that
    // transform cost once up front, so router.push() resolves from cache and
    // each test's navigation completes promptly.
    beforeAll(async () => {
      await Promise.all([
        import("../views/DocumentView.vue"),
        import("../views/HomeView.vue"),
      ]);
    });

    it("shows a full-window 'Drop to open' overlay while a file is dragged over the window (AC-1)", async () => {
      const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
      mountedWrappers.push(wrapper);
      expect(wrapper.find('[data-testid="drop-overlay"]').exists()).toBe(false);

      window.dispatchEvent(fileDragEvent("dragenter"));
      await nextTick();

      expect(wrapper.get('[data-testid="drop-overlay"]').text()).toBe("Drop to open");
    });

    it("keeps the overlay up across nested dragenter/dragleave and only hides once the drag truly leaves", async () => {
      const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
      mountedWrappers.push(wrapper);
      window.dispatchEvent(fileDragEvent("dragenter"));
      window.dispatchEvent(fileDragEvent("dragenter"));
      window.dispatchEvent(fileDragEvent("dragleave"));
      await nextTick();
      expect(wrapper.find('[data-testid="drop-overlay"]').exists()).toBe(true);

      window.dispatchEvent(fileDragEvent("dragleave"));
      await nextTick();
      expect(wrapper.find('[data-testid="drop-overlay"]').exists()).toBe(false);
    });

    it("dropping a .docx file opens it, exactly as the file picker would (AC-2)", async () => {
      const { wrapper, router } = mountWithRouter();
      const file = new File([docxBytes], "word-authored.docx");

      window.dispatchEvent(fileDragEvent("drop", [file]));
      await flushPromises();
      await untilNavigated(router);

      expect(router.currentRoute.value.path).toMatch(/^\/doc\/.+/);
      expect(wrapper.find('[data-testid="drop-overlay"]').exists()).toBe(false);
    });

    it("dropping an unsupported file type shows the exact error and does not navigate away (AC-4, AC-5)", async () => {
      const { wrapper, router } = mountWithRouter();
      const file = new File([new TextEncoder().encode("hi")], "notes.txt");

      window.dispatchEvent(fileDragEvent("drop", [file]));
      await flushPromises();

      expect(wrapper.get('[data-testid="drop-error"]').text()).toBe(
        "We can only open .docx, .odt, .md, .pptx, .odp, .key, .xlsx, .ods, .csv, .tsv, .pdf and image files",
      );
      expect(router.currentRoute.value.path).toBe("/");
    });

    it("dropping multiple files opens the first and shows a toast naming it (AC-6, AC-7)", async () => {
      const { wrapper, router } = mountWithRouter();
      const first = new File([docxBytes], "First.docx");
      const second = new File([docxBytes], "Second.docx");

      window.dispatchEvent(fileDragEvent("drop", [first, second]));
      await flushPromises();
      await untilNavigated(router);

      expect(router.currentRoute.value.path).toMatch(/^\/doc\/.+/);
      expect(wrapper.get('[data-testid="toast"]').text()).toBe(
        "Opened First — only one file at a time",
      );
    });

    it("dragging over the window while a document is open does not modify it until the drop completes (AC-8)", async () => {
      const { wrapper, router } = mountWithRouter("Existing doc");
      await router.push("/doc/already-open");

      window.dispatchEvent(fileDragEvent("dragenter"));
      await nextTick();
      expect(router.currentRoute.value.path).toBe("/doc/already-open");
      expect(wrapper.get('[data-testid="document-title"]').text()).toBe("Existing doc");

      window.dispatchEvent(fileDragEvent("dragleave"));
      await nextTick();
      expect(router.currentRoute.value.path).toBe("/doc/already-open");
    });
  });

  // SPEC-068 AC-1/2 — the shell exposes the live layout tier as a data
  // attribute so CSS and e2e can key off it. The shared width feeds every
  // consumer of useTier(); here we drive it directly (jsdom has no layout, so
  // the ResizeObserver never measures a real width).
  describe("layout tier (SPEC-068)", () => {
    it("reflects the current tier on the shell root, live", async () => {
      setTierWidthForTest(1440);
      const wrapper = mount(AppShell, { props: { documentTitle: "t" } });
      mountedWrappers.push(wrapper);
      const shell = wrapper.get('[data-testid="app-shell"]');
      expect(shell.attributes("data-tier")).toBe("full");

      setTierWidthForTest(800);
      await nextTick();
      expect(shell.attributes("data-tier")).toBe("medium");

      setTierWidthForTest(390);
      await nextTick();
      expect(shell.attributes("data-tier")).toBe("compact");
    });
  });
});

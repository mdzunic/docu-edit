// SPEC-115 — DeckDesignerDialog (deck-designer) as the user meets it, against
// a stubbed gateway and an in-memory library: three visible steps; the
// Brief's fields — and Redesign's lack of length and source; the New plan's
// editable outline (text, a layout select over the closed vocabulary,
// Alt+arrow reordering, delete and add, and the bullets a title slide will
// not build listed before Build); Redesign's rows with no add, no delete and
// no editable text, and "Apply redesign" in place of "Build deck"; a failed
// reply's error and its raw text — shown as text, never as markup.
import { flushPromises, mount } from "@vue/test-utils";
import { shallowRef } from "vue";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { setExperimentEnabled } from "../../lib/experiments";
import { resetPreferences, savePreference } from "../../lib/preferences";
import { PROVIDER_CONFIG_DEFAULTS } from "../../lib/ai/providerConfig";
import type { PresentationModel } from "../../lib/presentation/model";
import { createBlankPresentation, titleBodyFrames } from "../../lib/presentation/slides";
import DeckDesignerDialog from "./DeckDesignerDialog.vue";

const stub = vi.hoisted(() => ({ replies: [] as string[], requests: [] as Array<{ purpose: string; user: string }>, stored: [] as unknown[] }));

vi.mock("vue-router", () => ({ useRouter: () => ({ push: vi.fn() }) }));

vi.mock("./designerLibrary", async (importOriginal) => {
  const actual = await importOriginal<typeof import("./designerLibrary")>();
  return {
    ...actual,
    storeLibrary: () => ({
      sources: () => [{ id: "doc-1", title: "Migration notes", kind: "document" }],
      outlineOf: () => ({ title: "Migration notes", blocks: [{ kind: "heading", level: 2, text: "Status" }] }),
      titles: () => [],
      store: (model: unknown) => stub.stored.push(model),
      trash: () => {},
    }),
    loadModelGateway: async () => ({
      stream: async (request: { purpose: string; user: string }, onDelta: (chunk: string) => void) => {
        stub.requests.push(request);
        const text = stub.replies.shift() ?? "";
        onDelta(text);
        return { text, model: "stub-model", elapsedMs: 1 };
      },
    }),
  };
});

const PLAN = {
  title: "Rotterdam migration",
  theme: "Corporate blue",
  size: "16:9",
  slides: [
    { layout: "title", title: "Rotterdam migration", bullets: ["Board update"] },
    { layout: "titleAndBody", title: "Status", bullets: ["On track", "Phase 2 done"] },
    { layout: "titleOnly", title: "Questions?" },
  ],
};

beforeEach(() => {
  localStorage.clear();
  resetPreferences();
  setExperimentEnabled("deck-designer", true);
  savePreference("modelProvider", { ...PROVIDER_CONFIG_DEFAULTS, mode: "local", baseUrl: "http://localhost:11434/v1", model: "m" });
  stub.replies = [];
  stub.requests = [];
  stub.stored = [];
});

afterEach(() => {
  document.body.innerHTML = "";
});

function deck(): PresentationModel {
  const model = createBlankPresentation();
  model.slides = ["Alpha", "Beta", ""].map((title, i) => {
    const [titleFrame, body] = titleBodyFrames(model.slideWidthCm, model.slideHeightCm);
    titleFrame.paragraphs = title ? [{ runs: [{ text: title }] }] : [];
    body.paragraphs = [{ runs: [{ text: `${i} body` }] }];
    return { id: `s${i}`, frames: [titleFrame, body] };
  });
  return model;
}

const byTestId = (wrapper: ReturnType<typeof mount>, id: string) => wrapper.findAll(`[data-testid="${id}"]`);

describe("DeckDesignerDialog — the New flow", () => {
  it("shows three steps and the Brief's fields, with a live count of what will be sent", async () => {
    const wrapper = mount(DeckDesignerDialog, { props: { mode: "new" } });
    expect(byTestId(wrapper, "designer-steps")[0].text()).toMatch(/1\. Brief\s*2\. Plan\s*3\. Build/);
    expect(byTestId(wrapper, "designer-step-brief")[0].attributes("aria-current")).toBe("step");
    expect(byTestId(wrapper, "experiment-badge")).toHaveLength(1);
    for (const id of ["designer-brief", "designer-audience", "designer-length", "designer-tone", "designer-source", "designer-keep-prompt"]) {
      expect(byTestId(wrapper, id), id).toHaveLength(1);
    }
    expect(byTestId(wrapper, "designer-source")[0].findAll("option").map((option) => option.text())).toEqual([
      "Nothing",
      "Pasted text",
      "Migration notes (document)",
    ]);
    expect(wrapper.text()).toContain("will travel with the file");
    const count = () => Number(/([\d,]+) bytes/.exec(byTestId(wrapper, "designer-bytes")[0].text())?.[1].replace(/,/g, ""));
    const before = count();
    await byTestId(wrapper, "designer-brief")[0].setValue("Twelve bytes");
    expect(count()).toBe(before + "Twelve bytes".length);
    await byTestId(wrapper, "designer-source")[0].setValue("paste");
    expect(byTestId(wrapper, "designer-source-text")).toHaveLength(1);
    expect(byTestId(wrapper, "designer-back")).toHaveLength(1);
    expect(byTestId(wrapper, "designer-send")[0].attributes("disabled")).toBeUndefined();
  });

  it("the plan step is an editable outline: closed layout select, Alt+arrow, delete, add, and title-slide drops listed before Build", async () => {
    stub.replies = [JSON.stringify(PLAN)];
    const wrapper = mount(DeckDesignerDialog, { props: { mode: "new" }, attachTo: document.body });
    await byTestId(wrapper, "designer-brief")[0].setValue("A board update");
    await byTestId(wrapper, "designer-send")[0].trigger("click");
    await flushPromises();

    expect(stub.requests[0].purpose).toBe("Design a deck");
    expect(byTestId(wrapper, "designer-step-plan")[0].attributes("aria-current")).toBe("step");
    const rows = () => byTestId(wrapper, "designer-plan-row");
    expect(rows()).toHaveLength(3);
    // The reply's unknown theme was replaced, and said so.
    expect(byTestId(wrapper, "designer-drop").map((item) => item.text())).toEqual(['Deck: unknown theme "Corporate blue"; used Office']);
    const layoutSelect = rows()[1].find('[data-testid="designer-row-layout"]');
    expect(layoutSelect.findAll("option").map((option) => option.attributes("value"))).toEqual([
      "title",
      "titleAndBody",
      "sectionHeader",
      "twoContent",
      "titleOnly",
      "blank",
    ]);

    await layoutSelect.setValue("title");
    expect(byTestId(wrapper, "designer-pending-drop").map((item) => item.text())).toEqual([
      "Slide 2 — will not be built: a title slide keeps one line under its title; “Phase 2 done” dropped",
    ]);

    await rows()[2].find('[data-testid="designer-row-title"]').trigger("keydown", { key: "ArrowUp", altKey: true });
    expect(rows().map((row) => (row.find('[data-testid="designer-row-title"]').element as HTMLInputElement).value)).toEqual([
      "Rotterdam migration",
      "Questions?",
      "Status",
    ]);
    await rows()[0].find('[data-testid="designer-row-delete"]').trigger("click");
    await byTestId(wrapper, "designer-add-slide")[0].trigger("click");
    expect(rows()).toHaveLength(3);
    expect(byTestId(wrapper, "designer-build")).toHaveLength(1);
    expect(byTestId(wrapper, "designer-apply-redesign")).toHaveLength(0);

    await byTestId(wrapper, "designer-back")[0].trigger("click");
    expect(byTestId(wrapper, "designer-step-brief")[0].attributes("aria-current")).toBe("step");
    expect(stub.stored).toEqual([]);
    wrapper.unmount();
  });

  it("a reply that fails its repair shows the error and the raw text as text, never as markup", async () => {
    stub.replies = ["<b>not</b> json", "<img src=x onerror=alert(1)> still not"];
    const wrapper = mount(DeckDesignerDialog, { props: { mode: "new" } });
    await byTestId(wrapper, "designer-brief")[0].setValue("A board update");
    await byTestId(wrapper, "designer-send")[0].trigger("click");
    await flushPromises();
    expect(stub.requests).toHaveLength(2);
    expect(byTestId(wrapper, "designer-error")[0].text()).toMatch(/^reply: not valid JSON/);
    const raw = byTestId(wrapper, "designer-raw")[0];
    expect(raw.find("pre").text()).toBe("<img src=x onerror=alert(1)> still not");
    expect(raw.find("img").exists()).toBe(false);
    expect(byTestId(wrapper, "designer-try-again")).toHaveLength(1);
    expect(byTestId(wrapper, "designer-build")[0].attributes("disabled")).toBeDefined();
  });
});

describe("DeckDesignerDialog — Redesign", () => {
  it("has no length and no source; its plan has no add, no delete, no editable text, and Apply redesign", async () => {
    const model = shallowRef<PresentationModel | null>(deck());
    const host = { model, record: vi.fn(), commit: vi.fn() };
    stub.replies = [JSON.stringify({ theme: "Wisp", slides: [{ source: 2 }, { source: 0, layout: "blank" }] })];
    const wrapper = mount(DeckDesignerDialog, { props: { mode: "redesign", host } });
    expect(wrapper.text()).toContain("Redesign this deck");
    expect(byTestId(wrapper, "designer-length")).toHaveLength(0);
    expect(byTestId(wrapper, "designer-source")).toHaveLength(0);
    expect(byTestId(wrapper, "designer-keep-prompt")).toHaveLength(0);
    expect(byTestId(wrapper, "designer-audience")).toHaveLength(1);
    expect(byTestId(wrapper, "designer-tone")).toHaveLength(1);
    await byTestId(wrapper, "designer-send")[0].trigger("click");
    await flushPromises();
    expect(stub.requests[0].purpose).toBe("Redesign a deck");

    const rows = byTestId(wrapper, "designer-redesign-row");
    expect(rows.map((row) => row.find('[data-testid="designer-row-number"]').text())).toEqual(["Slide 3", "Slide 1", "Slide 2"]);
    expect(rows.map((row) => row.find('[data-testid="designer-row-title"]').text())).toEqual(["Slide 3", "Alpha", "Beta"]);
    expect(rows[1].find('[data-testid="designer-row-layout"]').element).toHaveProperty("value", "blank");
    expect(rows[0].find('[data-testid="designer-row-layout"]').findAll("option").map((option) => option.text())).toEqual([
      "Keep current",
      "Title Slide",
      "Title and Content",
      "Section Header",
      "Two Content",
      "Title Only",
      "Blank",
    ]);
    const theme = byTestId(wrapper, "designer-redesign-theme")[0];
    expect(theme.findAll("option").map((option) => option.text())).toEqual(["Keep current theme", "Office", "Facet", "Ion", "Retrospect", "Slice", "Wisp"]);
    expect(theme.element).toHaveProperty("value", "Wisp");
    expect(wrapper.findAll("input[type=text], textarea")).toHaveLength(0);
    expect(byTestId(wrapper, "designer-add-slide")).toHaveLength(0);
    expect(byTestId(wrapper, "designer-row-delete")).toHaveLength(0);
    expect(byTestId(wrapper, "designer-build")).toHaveLength(0);

    // Alt+↓ moves a row; Alt+↑ puts it back.
    const numbers = () => byTestId(wrapper, "designer-redesign-row").map((row) => row.find('[data-testid="designer-row-number"]').text());
    await rows[0].find('[data-testid="designer-row-layout"]').trigger("keydown", { key: "ArrowDown", altKey: true });
    expect(numbers()).toEqual(["Slide 1", "Slide 3", "Slide 2"]);
    await byTestId(wrapper, "designer-redesign-row")[1].find('[data-testid="designer-row-layout"]').trigger("keydown", { key: "ArrowUp", altKey: true });
    expect(numbers()).toEqual(["Slide 3", "Slide 1", "Slide 2"]);

    await byTestId(wrapper, "designer-apply-redesign")[0].trigger("click");
    expect(host.record).toHaveBeenCalledTimes(1);
    expect(host.commit).toHaveBeenCalledTimes(1);
    expect(model.value?.slides.map((slide) => slide.id)).toEqual(["s2", "s0", "s1"]);
    expect(wrapper.emitted("close")).toHaveLength(1);
  });
});

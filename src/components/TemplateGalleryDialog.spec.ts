import { describe, it, expect, beforeEach, vi } from "vitest";
import { mount, flushPromises } from "@vue/test-utils";
import { readFileSync } from "node:fs";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import TemplateGalleryDialog from "./TemplateGalleryDialog.vue";
import { BUILTIN_TEMPLATES } from "../lib/builtinTemplates";
import { deleteTemplate, listUserTemplates, saveAsTemplate } from "../lib/templateStore";
import { createDocument } from "../lib/documentModel";

const docxFixturesDir = join(dirname(fileURLToPath(import.meta.url)), "..", "lib", "docx", "fixtures");

function cleanupUserTemplates() {
  for (const template of listUserTemplates()) {
    deleteTemplate(template.id);
  }
}

describe("TemplateGalleryDialog (SPEC-046 AC-1, AC-2, AC-6, AC-7, AC-8)", () => {
  beforeEach(() => {
    cleanupUserTemplates();
  });

  it("offers a 'Blank document' card alongside the six shipped templates", () => {
    const wrapper = mount(TemplateGalleryDialog);
    expect(wrapper.get('[data-testid="template-card-blank"]').text()).toContain("Blank document");
    const cards = wrapper.findAll('[data-testid="template-card"]');
    expect(cards.length).toBeGreaterThanOrEqual(6);
    const ids = cards.map((card) => card.attributes("data-template-id"));
    expect(ids).toEqual(expect.arrayContaining(BUILTIN_TEMPLATES.map((t) => t.id)));
  });

  it("emits choose-blank when the blank card is clicked", async () => {
    const wrapper = mount(TemplateGalleryDialog);
    await wrapper.get('[data-testid="template-card-blank"]').trigger("click");
    expect(wrapper.emitted("choose-blank")).toHaveLength(1);
  });

  it("emits choose-builtin with the template id when a shipped template card is clicked", async () => {
    const wrapper = mount(TemplateGalleryDialog);
    const card = wrapper.get('[data-template-id="invoice"]');
    await card.trigger("click");
    expect(wrapper.emitted("choose-builtin")?.[0]).toEqual(["invoice"]);
  });

  it("lists user templates and emits choose-user when one is clicked", async () => {
    const record = saveAsTemplate(createDocument("gallery-user-doc"), "My Template");
    const wrapper = mount(TemplateGalleryDialog);
    const card = wrapper.get(`[data-testid="user-template-card"][data-template-id="${record.id}"]`);
    expect(card.text()).toContain("My Template");
    await card.find(".template-card-body").trigger("click");
    expect(wrapper.emitted("choose-user")?.some((call) => call[0] === record.id)).toBe(true);
  });

  it("renames a user template in place", async () => {
    const record = saveAsTemplate(createDocument("gallery-rename-doc"), "Old Name");
    const wrapper = mount(TemplateGalleryDialog);

    await wrapper.get('[data-testid="user-template-rename"]').trigger("click");
    const input = wrapper.get('[data-testid="rename-input"]');
    expect((input.element as HTMLInputElement).value).toBe("Old Name");
    await input.setValue("New Name");
    await wrapper.get('[data-testid="rename-save"]').trigger("click");

    expect(wrapper.get('[data-testid="user-template-card"]').text()).toContain("New Name");
    expect(listUserTemplates().find((t) => t.id === record.id)?.title).toBe("New Name");
  });

  it("deletes a user template after confirmation", async () => {
    const record = saveAsTemplate(createDocument("gallery-delete-doc"), "Delete Me");
    const wrapper = mount(TemplateGalleryDialog);

    await wrapper.get('[data-testid="user-template-delete"]').trigger("click");
    expect(wrapper.get('[data-testid="delete-message"]').text()).toContain("Delete Me");
    await wrapper.get('[data-testid="delete-confirm"]').trigger("click");

    expect(wrapper.find('[data-testid="user-template-card"]').exists()).toBe(false);
    expect(listUserTemplates().find((t) => t.id === record.id)).toBeUndefined();
  });

  it("emits close when the close button is clicked", async () => {
    const wrapper = mount(TemplateGalleryDialog);
    await wrapper.get('[data-testid="template-gallery-close"]').trigger("click");
    expect(wrapper.emitted("close")).toHaveLength(1);
  });

  it("imports a .dotx file and emits choose-file with the opened model (AC-9)", async () => {
    const bytes = readFileSync(join(docxFixturesDir, "word-authored.docx"));
    const file = new File([bytes], "Letter.dotx");
    const wrapper = mount(TemplateGalleryDialog);
    const input = wrapper.get('[data-testid="template-gallery-import-input"]').element as HTMLInputElement;
    Object.defineProperty(input, "files", { value: [file], configurable: true });
    await wrapper.get('[data-testid="template-gallery-import-input"]').trigger("change");
    await flushPromises();
    // parseDocx chunks its body walk on a 4ms time budget (SPEC-047 AC-16),
    // so how many macrotask hops the parse takes depends on machine load —
    // wait for the emit itself, not a fixed number of hops.
    await vi.waitUntil(() => wrapper.emitted("choose-file"), { timeout: 5000 });

    const emitted = wrapper.emitted("choose-file");
    expect(emitted).toHaveLength(1);
    expect(emitted![0][0]).toMatchObject({ format: "docx", title: "Letter" });
  });

  it("shows an error and does not emit when the imported file cannot be parsed", async () => {
    const file = new File([new TextEncoder().encode("not a real template")], "Broken.dotx");
    const wrapper = mount(TemplateGalleryDialog);
    const input = wrapper.get('[data-testid="template-gallery-import-input"]').element as HTMLInputElement;
    Object.defineProperty(input, "files", { value: [file], configurable: true });
    await wrapper.get('[data-testid="template-gallery-import-input"]').trigger("change");
    await flushPromises();
    await flushPromises();
    await new Promise((resolve) => setTimeout(resolve, 20));

    expect(wrapper.emitted("choose-file")).toBeUndefined();
    expect(wrapper.get('[data-testid="template-gallery-import-error"]').text().length).toBeGreaterThan(0);
  });
});

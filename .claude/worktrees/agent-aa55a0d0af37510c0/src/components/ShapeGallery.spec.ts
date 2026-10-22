// SPEC-131: ShapeGallery's `catalogueOnly` hides the Connectors and Freeform
// groups (a sheet inserts catalogue shapes only); without it — the document
// and slide galleries — both groups stay.
import { describe, expect, it } from "vitest";
import { mount } from "@vue/test-utils";
import ShapeGallery from "./ShapeGallery.vue";

describe("ShapeGallery catalogueOnly (SPEC-131)", () => {
  it("shows Connectors and Freeform by default and hides them with catalogueOnly, the catalogue's picks unchanged", async () => {
    const full = mount(ShapeGallery);
    expect(full.find('[data-testid="connector-gallery-group"]').exists()).toBe(true);
    expect(full.find('[data-testid="freeform-gallery-group"]').exists()).toBe(true);
    const sheet = mount(ShapeGallery, { props: { catalogueOnly: true } });
    expect(sheet.find('[data-testid="connector-gallery-group"]').exists()).toBe(false);
    expect(sheet.find('[data-testid="freeform-gallery-group"]').exists()).toBe(false);
    expect(sheet.findAll(".shape-gallery-item")).toHaveLength(full.findAll(".shape-gallery-item").length - 6);
    await sheet.get('[data-testid="shape-gallery-rect"]').trigger("click");
    expect(sheet.emitted("pick")).toEqual([["rect"]]);
  });
});

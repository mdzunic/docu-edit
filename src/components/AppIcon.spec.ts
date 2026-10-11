import { describe, it, expect } from "vitest";
import { mount } from "@vue/test-utils";
import AppIcon from "./AppIcon.vue";
import { ICON_PATHS, hasIcon } from "./icons/paths";

// The command names SPEC-058 says the icon set must cover "at minimum" — every
// control that was a text glyph before the refresh. If a name is dropped from
// paths.ts, this list catches it.
const REQUIRED = [
  "bold", "italic", "underline", "strikethrough", "subscript", "superscript",
  "align-left", "align-center", "align-right", "align-justify",
  "bullet-list", "numbered-list", "multilevel-list",
  "indent-increase", "indent-decrease", "line-spacing",
  "text-color", "highlight", "clear-formatting", "format-painter",
  "undo", "redo", "cut", "copy", "paste",
  "link", "bookmark", "comment", "image", "table", "chart", "shape",
  "text-box", "equation", "special-character", "footnote", "page-break",
  "find", "print", "download", "close",
  "ruler", "outline", "styles", "comments", "track-changes",
  "accessibility", "fonts", "history",
  "zoom-in", "zoom-out",
  "chevron-up", "chevron-down", "chevron-left", "chevron-right",
  "check", "dash", "plus", "minus",
  "more-horizontal", "more-vertical", "search", "settings",
  "warning", "error", "info",
  "conditional-format",
];

describe("AppIcon", () => {
  it("defines a hand-authored path for every required command", () => {
    for (const name of REQUIRED) {
      expect(hasIcon(name), `missing icon: ${name}`).toBe(true);
      expect(ICON_PATHS[name].length).toBeGreaterThan(0);
    }
  });

  it("renders a 20×20 stroke-only inline svg with the AppIcon contract", () => {
    const wrapper = mount(AppIcon, { props: { name: "bold" } });
    const svg = wrapper.get("svg");
    expect(svg.attributes("viewBox")).toBe("0 0 20 20");
    expect(svg.attributes("width")).toBe("20");
    expect(svg.attributes("height")).toBe("20");
    expect(svg.attributes("fill")).toBe("none");
    expect(svg.attributes("stroke")).toBe("currentColor");
    expect(svg.attributes("stroke-width")).toBe("1.5");
    expect(svg.attributes("aria-hidden")).toBe("true");
    expect(svg.attributes("focusable")).toBe("false");
    expect(wrapper.findAll("path").length).toBe(ICON_PATHS.bold.length);
  });

  it("renders nothing for an unknown name (a typo cannot ship as a blank button)", () => {
    const wrapper = mount(AppIcon, { props: { name: "definitely-not-an-icon" } });
    expect(wrapper.find("svg").exists()).toBe(false);
  });

  it("carries no colour literal on the svg itself — it inherits from the button", () => {
    const wrapper = mount(AppIcon, { props: { name: "italic" } });
    const html = wrapper.html();
    // currentColor is the only colour token the icon names.
    expect(html).not.toMatch(/#[0-9a-fA-F]{3,8}/);
    expect(html).toContain('stroke="currentColor"');
  });
});

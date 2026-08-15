import { describe, it, expect } from "vitest";
import { asMenuBarMenus } from "./menuBar";

describe("asMenuBarMenus", () => {
  it("coerces id and label to strings", () => {
    expect(asMenuBarMenus([{ id: "file", label: "File" }])).toEqual([{ id: "file", label: "File" }]);
  });

  it("tolerates missing fields without throwing", () => {
    expect(asMenuBarMenus([{}, null, { id: "edit" }])).toEqual([
      { id: "", label: "" },
      { id: "", label: "" },
      { id: "edit", label: "" },
    ]);
  });
});

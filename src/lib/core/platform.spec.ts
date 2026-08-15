import { afterEach, describe, expect, it, vi } from "vitest";
import { isMacPlatform, modifierKeyLabel } from "./platform";

describe("platform", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("reports Ctrl on a non-Mac platform", () => {
    vi.stubGlobal("navigator", { platform: "Linux x86_64", userAgent: "X11; Linux x86_64" });
    expect(isMacPlatform()).toBe(false);
    expect(modifierKeyLabel()).toBe("Ctrl");
  });

  it("reports Cmd on a Mac platform", () => {
    vi.stubGlobal("navigator", { platform: "MacIntel", userAgent: "Macintosh" });
    expect(isMacPlatform()).toBe(true);
    expect(modifierKeyLabel()).toBe("Cmd");
  });

  it("falls back to userAgent when platform is unavailable", () => {
    vi.stubGlobal("navigator", { platform: "", userAgent: "iPhone" });
    expect(isMacPlatform()).toBe(true);
  });
});

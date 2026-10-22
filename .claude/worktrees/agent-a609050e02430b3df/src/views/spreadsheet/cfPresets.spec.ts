// SPEC-129 "Presets" — the cf-add gallery and the cf-format differential
// presets: every testid is unique and each draft/dxf is well-formed.
import { describe, expect, it } from "vitest";
import { CF_ADD_PRESETS, CF_FORMAT_PRESETS } from "./cfPresets";
import { ICON_SET_NAMES, ODS_ONLY_ICON_SETS } from "../../lib/spreadsheet/conditionalFormats";

describe("CF_ADD_PRESETS (SPEC-129 Presets)", () => {
  it("has a unique testid per preset", () => {
    const testids = CF_ADD_PRESETS.map((preset) => preset.testid);
    expect(new Set(testids).size).toBe(testids.length);
  });

  it("includes the value presets, each pre-setting type/operator with blank values", () => {
    const greaterThan = CF_ADD_PRESETS.find((p) => p.testid === "cf-preset-greater-than")!;
    expect(greaterThan.draft()).toEqual({ kind: "cellIs", operator: "greaterThan", formulas: [""] });
    const between = CF_ADD_PRESETS.find((p) => p.testid === "cf-preset-between")!;
    expect(between.draft()).toEqual({ kind: "cellIs", operator: "between", formulas: ["", ""] });
    const textContains = CF_ADD_PRESETS.find((p) => p.testid === "cf-preset-text-contains")!;
    expect(textContains.draft()).toMatchObject({ kind: "text", operator: "contains" });
  });

  it("includes a gradient and a solid data bar for each of the six colours", () => {
    for (const colour of ["blue", "green", "red", "orange", "light-blue", "purple"]) {
      const gradient = CF_ADD_PRESETS.find((p) => p.testid === `cf-preset-data-bar-${colour}-gradient`);
      const solid = CF_ADD_PRESETS.find((p) => p.testid === `cf-preset-data-bar-${colour}-solid`);
      expect(gradient, colour).toBeDefined();
      expect(solid, colour).toBeDefined();
      expect((gradient!.draft() as { bar: { gradient: boolean } }).bar.gradient).toBe(true);
      expect((solid!.draft() as { bar: { gradient: boolean } }).bar.gradient).toBe(false);
    }
    // Excel 2010's gallery bar: automatic bounds over the full 0–100% length,
    // red negatives, a black axis, a same-colour border on the gradient.
    expect(CF_ADD_PRESETS.find((p) => p.testid === "cf-preset-data-bar-green-gradient")!.draft()).toEqual({
      kind: "dataBar",
      bar: {
        min: { type: "autoMin" },
        max: { type: "autoMax" },
        color: "#63BE7B",
        gradient: true,
        border: "#63BE7B",
        negativeColor: "#FF0000",
        negativeBorderColor: "#FF0000",
        axis: "automatic",
        axisColor: "#000000",
        direction: "context",
        minLengthPct: 0,
        maxLengthPct: 100,
        showValue: true,
      },
    });
  });

  it("includes six colour scales", () => {
    for (let n = 1; n <= 6; n += 1) {
      const preset = CF_ADD_PRESETS.find((p) => p.testid === `cf-preset-color-scale-${n}`);
      expect(preset, String(n)).toBeDefined();
      expect(preset!.draft().kind).toBe("colorScale");
    }
  });

  it("includes every Excel icon set (not the ODS-only smilies)", () => {
    for (const name of ICON_SET_NAMES) {
      const preset = CF_ADD_PRESETS.find((p) => p.testid === `cf-preset-icon-set-${name}`);
      if (ODS_ONLY_ICON_SETS.has(name)) {
        expect(preset, name).toBeUndefined();
      } else {
        expect(preset, name).toBeDefined();
        expect(preset!.draft()).toMatchObject({ kind: "iconSet", icons: { set: name } });
      }
    }
  });

  it("includes New rule…", () => {
    expect(CF_ADD_PRESETS.find((p) => p.testid === "cf-preset-new-rule")).toBeDefined();
  });
});

describe("CF_FORMAT_PRESETS (SPEC-129 Presets)", () => {
  it("has the six named presets with Excel's exact colours", () => {
    const byTestid = Object.fromEntries(CF_FORMAT_PRESETS.map((p) => [p.testid, p]));
    expect(byTestid["cf-format-preset-light-red-fill-dark-red-text"].dxf).toEqual({
      fill: { kind: "solid", color: "#FFC7CE" },
      font: { color: "#9C0006" },
    });
    expect(byTestid["cf-format-preset-yellow-fill-dark-yellow-text"].dxf).toEqual({
      fill: { kind: "solid", color: "#FFEB9C" },
      font: { color: "#9C5700" },
    });
    expect(byTestid["cf-format-preset-green-fill-dark-green-text"].dxf).toEqual({
      fill: { kind: "solid", color: "#C6EFCE" },
      font: { color: "#006100" },
    });
    expect(byTestid["cf-format-preset-light-red-fill"].dxf).toEqual({ fill: { kind: "solid", color: "#FFC7CE" } });
    expect(byTestid["cf-format-preset-red-text"].dxf).toEqual({ font: { color: "#9C0006" } });
    const border = byTestid["cf-format-preset-red-border"].dxf.borders!;
    for (const edge of ["top", "right", "bottom", "left"] as const) {
      expect(border[edge]).toEqual({ style: "thin", color: "#9C0006" });
    }
  });
});

// The rules manager's previews (SPEC-129 "The side panel": `cf-rule-preview`,
// and the rule editor's live `cf-preview`) — THIS FILE is the normative
// source for what a rule's preview draws: a formatting rule the sample text
// "AaBbCc" styled by its dxf (layerDxf over an empty style, then SPEC-117's
// cellCss — the grid's own rendering), a colour scale its gradient strip, a
// data bar its bar (gradient toward 90% white, as the grid draws it) and an
// icon set its icons, lowest first. Framework-free and pure.
import type { DocumentTheme } from "../../lib/core/theme";
import { cellCss, type CssDeclarations } from "../../lib/spreadsheet/cellRender";
import { mixTowardWhite } from "../../lib/spreadsheet/cfEvaluator";
import { iconSetSize, layerDxf, type CfRuleBody, type Dxf, type IconSetName } from "../../lib/spreadsheet/conditionalFormats";
import { cellColorHex } from "../../lib/spreadsheet/styles";

export const PREVIEW_SAMPLE = "AaBbCc";

export type CfPreview =
  | { kind: "format"; box: CssDeclarations; text: CssDeclarations; sample: string }
  | { kind: "scale"; background: string }
  | { kind: "bar"; background: string }
  | { kind: "icons"; set: IconSetName; indices: number[] }
  | { kind: "none" };

/** The preview of a rule (or the editor's pending draft) with its dxf. */
export function rulePreview(rule: CfRuleBody, dxf: Dxf | undefined, theme: DocumentTheme): CfPreview {
  switch (rule.kind) {
    case "colorScale": {
      const colors = rule.stops.map((stop) => cellColorHex(stop.color, theme) || "#ffffff");
      return { kind: "scale", background: `linear-gradient(to right, ${colors.join(", ")})` };
    }
    case "dataBar": {
      const color = cellColorHex(rule.bar.color, theme) || "#638EC6";
      return { kind: "bar", background: rule.bar.gradient ? `linear-gradient(to right, ${color}, ${mixTowardWhite(color, 0.9)})` : color };
    }
    case "iconSet": {
      const count = iconSetSize(rule.icons.set);
      return { kind: "icons", set: rule.icons.set, indices: Array.from({ length: count }, (_, index) => index) };
    }
    case "preserved":
      return { kind: "none" };
    default: {
      const css = cellCss(layerDxf({}, dxf ?? {}), theme);
      return { kind: "format", box: css.box, text: css.text, sample: PREVIEW_SAMPLE };
    }
  }
}

/** Narrows a component's runtime-typed `preview` prop (the .vue parser bars
 * casts). */
export function asCfPreview(value: unknown): CfPreview {
  return value as CfPreview;
}

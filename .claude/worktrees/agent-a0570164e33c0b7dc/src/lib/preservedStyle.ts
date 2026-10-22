// SPEC-073 AC-36 "exactly one source per property": a block's
// `unsupportedStyleXml` holds the styling its source format expressed but the
// model could not — and the writers re-emit a preserved fill/outline/shadow
// in the modelled property's place. When an EDIT gives the model a value for
// one of those properties (including "none": clearing a fill deletes
// `shape.fill`), the preserved slot must be cleared with it, or the next
// same-format save would silently resurrect the old styling over the user's
// visible change — the exact silent-downgrade this spec exists to close.
//
// The string is format-flavoured — a `.odt` block preserves one synthetic
// `style:graphic-properties` element, a `.docx` block a run of `a:*` property
// elements — so this module just sniffs the flavour and delegates to the
// format module that produced it. It is app-level glue (commands call it),
// not part of either format's reader/writer.
import { clearPreservedDocxStyle } from "./docx/shape";
import { clearPreservedOdfGraphicStyle, isPreservedGraphicStyleXml } from "./odf/shape";

export interface PreservedStyleSlots {
  fill?: boolean;
  outline?: boolean;
  shadow?: boolean;
}

/** The block's `unsupportedStyleXml` with the given property slots removed —
 * `undefined` when nothing survives (the caller then deletes the field).
 * Slots not named, and styling in neither slot (a preserved glow, scene3d,
 * a foreign attribute), stay preserved verbatim. */
export function clearPreservedStyleSlots(
  xml: string | undefined,
  slots: PreservedStyleSlots,
): string | undefined {
  if (xml === undefined || (!slots.fill && !slots.outline && !slots.shadow)) {
    return xml;
  }
  return isPreservedGraphicStyleXml(xml)
    ? clearPreservedOdfGraphicStyle(xml, slots)
    : clearPreservedDocxStyle(xml, slots);
}

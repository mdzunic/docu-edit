// Moved: the numeric conversions now live in src/lib/core/units.ts (twips,
// half-points and EMU are OOXML-family units shared with a future .pptx
// writer, not WordprocessingML-specific) and the colour codec in
// src/lib/core/colors.ts. This shim keeps existing import paths working.
export {
  cmToEmu,
  cmToTwips,
  eighthPointsToPt,
  emuToCm,
  halfPointsToPt,
  ptToEighthPoints,
  ptToHalfPoints,
  ptToTwentiethsOfPoint,
  twentiethsOfPointToPt,
  twipsToCm,
} from "../core/units";
export { cssColorToOoxml, ooxmlColorToCss } from "../core/colors";

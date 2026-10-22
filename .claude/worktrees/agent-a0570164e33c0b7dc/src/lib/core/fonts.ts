// The product-neutral half of SPEC-013's font controls: the standard
// font-family catalogue every editing surface offers. The document-walking
// half ("Fonts in this document") stays in lib/fonts.ts.

/** The families the font dropdown always offers (SPEC-013 AC): the classic
 * web-safe set plus the common document fonts — Word's modern defaults
 * (Calibri/Cambria) and their metric-compatible Liberation/Carlito
 * equivalents. Alphabetical, so the list is stable and scannable. */
export const STANDARD_FONT_FAMILIES: string[] = [
  "Arial",
  "Calibri",
  "Cambria",
  "Carlito",
  "Courier New",
  "Garamond",
  "Georgia",
  "Helvetica",
  "Liberation Mono",
  "Liberation Sans",
  "Liberation Serif",
  "Tahoma",
  "Times New Roman",
  "Trebuchet MS",
  "Verdana",
];

// The chart-feature markers: a curated set of XML markers this product's
// chart model (`ChartVisual`) has no field for — trendlines, error bars,
// per-point formatting, custom number formats, secondary axes, gradient/
// pattern/picture fills. SPEC-052 introduced them in `src/lib/opaque.ts` to
// report what an *edited* chart's regenerate-from-model rewrite costs, per
// property; SPEC-092 hoists them here (product-neutral, format-keyed) so both
// the text writers (via `opaque.ts`, which re-imports them) and the
// presentation rewrite-notice (`presentation/charts.ts`, via
// `lostChartFeatures`) share one source of truth.
//
// Not exhaustive — a real chart part can carry far more than this editor will
// ever model (themes, every per-point override, every number format) — this
// is the common, easy-to-name subset; anything outside it is a real loss the
// report stays silent about, the same honest limit every other "which
// sub-feature" heuristic in this repo already accepts.

export interface ChartFeatureMarker {
  test: (xml: string) => boolean;
  reason: string;
}

/** OOXML (`c:chartSpace`) markers — a `.docx` or `.pptx` chart part. */
export const DOCX_CHART_FEATURE_MARKERS: ChartFeatureMarker[] = [
  { test: (xml) => /<c:trendline[\s/>]/.test(xml), reason: "a trendline" },
  { test: (xml) => /<c:errBars[\s/>]/.test(xml), reason: "error bars" },
  { test: (xml) => /<c:dPt[\s>]/.test(xml), reason: "per-point formatting (individual bar/slice colours or markers)" },
  { test: (xml) => /formatCode="(?!General")[^"]*"/.test(xml), reason: "a custom number format" },
  { test: (xml) => (xml.match(/<c:valAx[\s>]/g) ?? []).length > 1, reason: "a secondary axis" },
  { test: (xml) => /<a:gradFill[\s>]|<a:pattFill[\s>]|<a:blipFill[\s>]/.test(xml), reason: "a gradient, pattern or picture fill" },
];

/** ODF (`chart:chart`) markers — an `.odt` or `.odp` embedded chart object. */
export const ODT_CHART_FEATURE_MARKERS: ChartFeatureMarker[] = [
  { test: (xml) => /<chart:regression-curve[\s/>]/.test(xml), reason: "a trendline" },
  { test: (xml) => /<chart:error-indicator[\s/>]/.test(xml), reason: "error bars" },
  { test: (xml) => /<chart:data-point[\s>]/.test(xml), reason: "per-point formatting (individual bar/slice colours or markers)" },
  { test: (xml) => (xml.match(/<chart:axis\b[^>]*dimension="y"/g) ?? []).length > 1, reason: "a secondary axis" },
];

/** Every feature an *edited* chart's regenerated part loses relative to its
 * preserved original `sourceXml`, as human-readable phrases — the raw
 * material for SPEC-092's `chart-rewrite-notice` and SPEC-052's same-format
 * `chartFeature` DropRecords. `format` selects the marker table:
 * `"ooxml"` for a `.docx`/`.pptx` part, `"odf"` for an `.odt`/`.odp` one. */
export function lostChartFeatures(sourceXml: string, format: "ooxml" | "odf"): string[] {
  const markers = format === "ooxml" ? DOCX_CHART_FEATURE_MARKERS : ODT_CHART_FEATURE_MARKERS;
  return markers.filter((marker) => marker.test(sourceXml)).map((marker) => marker.reason);
}

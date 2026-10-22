// SPEC-094 golden: the canonical `de-theme-*` block the odp splice writes into
// styles.xml's `office:styles` for `GOLDEN_THEME` — hand-authored against ODF
// 1.3 §16 (style:style / graphic- and text-properties), re-exporting the same
// scheme the pptx golden pins so both formats are provably writing one theme.
export { GOLDEN_THEME } from "../pptx/themeGolden";

/** The exact block `odpThemeBlockXml(GOLDEN_THEME)` must produce. */
export const ODP_THEME_GOLDEN =
  `<style:style style:name="de-theme" style:family="graphic" style:display-name="Golden"/>` +
  `<style:style style:name="de-theme-dk1" style:family="graphic" style:display-name="Dark 1">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#101010"/></style:style>` +
  `<style:style style:name="de-theme-text-dk1" style:family="text" style:display-name="Dark 1">` +
  `<style:text-properties fo:color="#101010"/></style:style>` +
  `<style:style style:name="de-theme-lt1" style:family="graphic" style:display-name="Light 1">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#fefefe"/></style:style>` +
  `<style:style style:name="de-theme-text-lt1" style:family="text" style:display-name="Light 1">` +
  `<style:text-properties fo:color="#fefefe"/></style:style>` +
  `<style:style style:name="de-theme-dk2" style:family="graphic" style:display-name="Dark 2">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#123456"/></style:style>` +
  `<style:style style:name="de-theme-text-dk2" style:family="text" style:display-name="Dark 2">` +
  `<style:text-properties fo:color="#123456"/></style:style>` +
  `<style:style style:name="de-theme-lt2" style:family="graphic" style:display-name="Light 2">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#edcba9"/></style:style>` +
  `<style:style style:name="de-theme-text-lt2" style:family="text" style:display-name="Light 2">` +
  `<style:text-properties fo:color="#edcba9"/></style:style>` +
  `<style:style style:name="de-theme-accent1" style:family="graphic" style:display-name="Accent 1">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#aa1122"/></style:style>` +
  `<style:style style:name="de-theme-text-accent1" style:family="text" style:display-name="Accent 1">` +
  `<style:text-properties fo:color="#aa1122"/></style:style>` +
  `<style:style style:name="de-theme-accent2" style:family="graphic" style:display-name="Accent 2">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#bb2233"/></style:style>` +
  `<style:style style:name="de-theme-text-accent2" style:family="text" style:display-name="Accent 2">` +
  `<style:text-properties fo:color="#bb2233"/></style:style>` +
  `<style:style style:name="de-theme-accent3" style:family="graphic" style:display-name="Accent 3">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#cc3344"/></style:style>` +
  `<style:style style:name="de-theme-text-accent3" style:family="text" style:display-name="Accent 3">` +
  `<style:text-properties fo:color="#cc3344"/></style:style>` +
  `<style:style style:name="de-theme-accent4" style:family="graphic" style:display-name="Accent 4">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#dd4455"/></style:style>` +
  `<style:style style:name="de-theme-text-accent4" style:family="text" style:display-name="Accent 4">` +
  `<style:text-properties fo:color="#dd4455"/></style:style>` +
  `<style:style style:name="de-theme-accent5" style:family="graphic" style:display-name="Accent 5">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#ee5566"/></style:style>` +
  `<style:style style:name="de-theme-text-accent5" style:family="text" style:display-name="Accent 5">` +
  `<style:text-properties fo:color="#ee5566"/></style:style>` +
  `<style:style style:name="de-theme-accent6" style:family="graphic" style:display-name="Accent 6">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#ff6677"/></style:style>` +
  `<style:style style:name="de-theme-text-accent6" style:family="text" style:display-name="Accent 6">` +
  `<style:text-properties fo:color="#ff6677"/></style:style>` +
  `<style:style style:name="de-theme-hlink" style:family="graphic" style:display-name="Hyperlink">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#0000ee"/></style:style>` +
  `<style:style style:name="de-theme-text-hlink" style:family="text" style:display-name="Hyperlink">` +
  `<style:text-properties fo:color="#0000ee"/></style:style>` +
  `<style:style style:name="de-theme-folHlink" style:family="graphic" style:display-name="Followed Hyperlink">` +
  `<style:graphic-properties draw:fill="solid" draw:fill-color="#551a8b"/></style:style>` +
  `<style:style style:name="de-theme-text-folHlink" style:family="text" style:display-name="Followed Hyperlink">` +
  `<style:text-properties fo:color="#551a8b"/></style:style>` +
  `<style:style style:name="de-theme-major" style:family="text">` +
  `<style:text-properties style:font-name="Georgia" fo:font-family="Georgia"/></style:style>` +
  `<style:style style:name="de-theme-minor" style:family="text">` +
  `<style:text-properties style:font-name="Verdana" fo:font-family="Verdana"/></style:style>`;

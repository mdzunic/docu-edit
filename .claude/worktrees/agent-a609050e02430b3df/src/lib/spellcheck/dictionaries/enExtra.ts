// The curated supplement to the vendored SCOWL word lists (SPEC-035).
//
// `wordlist-english` is a general-purpose English vocabulary and deliberately
// carries no proper nouns, no contractions and no hyphenated compounds. Three
// gaps follow from that, and every word below exists to close one of them —
// each was found by checking the shipped lists, not guessed at:
//
//   1. Proper nouns people type constantly in documents: weekdays, months,
//      and the language names the app's own status bar shows ("English").
//   2. Contractions whose stem is not a word once the clitic is removed, so
//      dictionary.ts's clitic rule cannot reach them ("can't" -> "ca").
//   3. Office, publishing and computing vocabulary newer than the word list.
//
// Inflected forms are spelled out here because the English dictionaries do no
// suffix guessing — SCOWL already carries the real inflections of everything
// it lists, and inventing them for these few extra words would trade a small
// convenience for a checker that accepts "boxs".
export const SHARED_EXTRA: readonly string[] = (
  // Weekdays and months (SCOWL keeps proper nouns in lists this package
  // does not ship; "march", "may" and "august" arrive as common words).
  "monday tuesday wednesday thursday friday saturday sunday mondays " +
  "tuesdays wednesdays thursdays fridays saturdays sundays " +
  "january february april june july september october november december " +
  // Languages and nationalities — "English" is in the editor's status bar.
  "english british american irish scottish welsh french german spanish " +
  "italian portuguese dutch czech polish russian swedish norwegian danish " +
  "finnish greek turkish arabic hebrew chinese japanese korean indian " +
  "canadian australian mexican brazilian european asian african " +
  // Contractions the clitic rule cannot decompose to a real stem.
  "can't won't shan't ain't o'clock " +
  // Office, publishing and document vocabulary.
  "spreadsheet spreadsheets screenshot screenshots subtotal subtotals " +
  "stakeholder stakeholders deliverable deliverables onboarding offboarding " +
  "roadmap roadmaps changelog changelogs workflow workflows " +
  "actionable scalable reusable clickable searchable sortable " +
  "hyperlink hyperlinks watermark watermarks endnote endnotes " +
  "kerning leading tracking ligature ligatures typeface typefaces " +
  "letterhead letterheads pagination gutter gutters " +
  // Computing vocabulary and the formats this product reads and writes.
  "login logins logout username usernames password passwords " +
  "metadata runtime filename filenames hostname hostnames " +
  "timestamp timestamps dataset datasets namespace namespaces " +
  "frontend backend dropdown dropdowns checkbox checkboxes " +
  "plugin plugins sidebar sidebars toolbar toolbars viewport viewports " +
  "tooltip tooltips modal modals autosave resize resizes " +
  "online offline analytics dashboard dashboards " +
  "boolean enum config configs repo repos async await " +
  "docx odt pptx odp pdf html css json xml sql api apis url urls " +
  "cli sdk ui ux svg png jpeg webp markdown unicode " +
  "smartphone smartphones laptop laptops wifi podcast podcasts " +
  "webinar webinars blog blogs"
).split(/\s+/);

// Dialect-specific words newer than the vendored lists. Keep these in step:
// a word added on one side needs its counterpart on the other, or the two
// dictionaries drift apart in coverage rather than only in spelling.
export const US_EXTRA: readonly string[] = "prioritize prioritized prioritizing".split(/\s+/);

export const GB_EXTRA: readonly string[] = "prioritise prioritised prioritising".split(/\s+/);

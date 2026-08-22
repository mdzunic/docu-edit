// The command metadata catalogue (SPEC-062) — the framework-free table that
// gives every id in DocumentView's `editorCommands` registry a user-facing
// name, a menu path and search terms, plus the pure search/ranking the
// command palette renders. No Vue imports: it is unit-testable without
// mounting the view and reusable by any later "what can this app do" surface
// (SPEC-068's small-screen layout is the next).
//
// The completeness test (commandCatalogue.spec.ts) pins the document-scope
// ids to `editorCommands`'s keys in both directions, so this table can never
// silently miss a command or name one that no longer dispatches.

export interface CommandMeta {
  /** Matches an `editorCommands` key exactly (document scope), or a library
   * action id (library scope). */
  id: string;
  /** User-facing name, e.g. "Footnote". Matches the menu bar's own label
   * where the command has a menu item. */
  label: string;
  /** Menu path shown as context, e.g. ["Insert"] renders "Insert → Footnote". */
  path: string[];
  /** SPEC-058 icon name, when the command has one. */
  icon?: string;
  /** Extra search terms: synonyms and the words people actually type. */
  aliases?: string[];
  /** Which surface offers it. */
  scope: "document" | "library";
}

/** The catalogue. Document-scope ids mirror `editorCommands` (asserted by the
 * completeness test); library-scope ids are the actions the library view
 * offers. Listed menu-by-menu, matching the bar's order. */
export const COMMANDS: CommandMeta[] = [
  // --- File -----------------------------------------------------------------
  { id: "file.print", label: "Print", path: ["File"], scope: "document" },
  { id: "file.page-setup", label: "Page setup", path: ["File"], scope: "document", aliases: ["margins", "orientation", "paper size"] },
  { id: "file.properties", label: "Properties…", path: ["File"], scope: "document", aliases: ["metadata", "author"] },
  { id: "file.version-history", label: "Version history", path: ["File"], scope: "document", aliases: ["revisions"] },
  { id: "file.save-as-template", label: "Save as template", path: ["File"], scope: "document" },
  { id: "file.mail-merge", label: "Mail merge…", path: ["File"], scope: "document" },
  { id: "file.delete", label: "Delete document", path: ["File"], scope: "document" },
  { id: "file.close", label: "Close", path: ["File"], scope: "document", aliases: ["return to library", "exit"] },

  // --- Edit -----------------------------------------------------------------
  { id: "edit.undo", label: "Undo", path: ["Edit"], scope: "document" },
  { id: "edit.redo", label: "Redo", path: ["Edit"], scope: "document" },
  { id: "edit.cut", label: "Cut", path: ["Edit"], scope: "document" },
  { id: "edit.copy", label: "Copy", path: ["Edit"], scope: "document" },
  { id: "edit.paste", label: "Paste", path: ["Edit"], scope: "document" },
  { id: "edit.pastePlain", label: "Paste without formatting", path: ["Edit"], scope: "document", aliases: ["paste plain", "paste text"] },
  { id: "edit.selectAll", label: "Select all", path: ["Edit"], scope: "document" },
  { id: "find.openReplace", label: "Find and replace", path: ["Edit"], scope: "document", aliases: ["search", "replace"] },
  { id: "edit.autocorrect-smart-quotes", label: "Smart quotes", path: ["Edit", "Autocorrect"], scope: "document", aliases: ["curly quotes"] },
  { id: "edit.autocorrect-dashes", label: "Em and en dashes", path: ["Edit", "Autocorrect"], scope: "document" },
  { id: "field.update", label: "Update field", path: ["Edit"], scope: "document", aliases: ["refresh field", "F9"] },

  // --- View -----------------------------------------------------------------
  { id: "view.ruler", label: "Ruler", path: ["View"], scope: "document" },
  { id: "view.outline-pane", label: "Outline pane", path: ["View"], scope: "document", aliases: ["navigation"] },
  { id: "view.styles-pane", label: "Styles pane", path: ["View"], scope: "document" },
  { id: "view.comments-pane", label: "Comments pane", path: ["View"], scope: "document" },
  { id: "view.comments-only", label: "Comments-only view", path: ["View"], scope: "document" },
  { id: "view.field-shading", label: "Field shading", path: ["View"], scope: "document" },
  { id: "view.theme-system", label: "System", path: ["View", "Appearance"], scope: "document", aliases: ["theme"] },
  { id: "view.theme-light", label: "Light", path: ["View", "Appearance"], scope: "document", aliases: ["theme", "light mode"] },
  { id: "view.theme-dark", label: "Dark", path: ["View", "Appearance"], scope: "document", aliases: ["theme", "dark mode"] },
  { id: "view.dim-page", label: "Dim page in dark theme", path: ["View", "Appearance"], scope: "document" },
  { id: "edit.toggleTrackChanges", label: "Track changes", path: ["View"], scope: "document", icon: "track-changes" },
  { id: "view.track-changes-pane", label: "Track changes pane", path: ["View"], scope: "document", icon: "track-changes" },
  { id: "view.formattingMarks", label: "Show formatting marks", path: ["View"], scope: "document", aliases: ["pilcrow", "paragraph marks"] },

  // --- Insert ---------------------------------------------------------------
  { id: "insert.table-dialog", label: "Insert table…", path: ["Insert"], scope: "document" },
  { id: "insert.textbox", label: "Text box", path: ["Insert"], scope: "document", icon: "text-box" },
  { id: "insert.chart-from-table", label: "Chart from table", path: ["Insert"], scope: "document" },
  { id: "insert.header", label: "Header", path: ["Insert"], scope: "document" },
  { id: "insert.footer", label: "Footer", path: ["Insert"], scope: "document" },
  { id: "insert.page-number", label: "Page number", path: ["Insert"], scope: "document" },
  { id: "insert.header-footer-options", label: "Header and footer options", path: ["Insert"], scope: "document" },
  { id: "insert.merge-field", label: "Merge field…", path: ["Insert"], scope: "document" },
  { id: "insert.note-settings", label: "Footnote and endnote settings", path: ["Insert"], scope: "document" },
  { id: "insert.field", label: "Field…", path: ["Insert"], scope: "document" },
  { id: "insert.date-time", label: "Date and time", path: ["Insert"], scope: "document" },
  { id: "insert.cross-reference", label: "Cross-reference", path: ["Insert"], scope: "document" },
  { id: "insert.toc", label: "Table of contents", path: ["Insert"], scope: "document", aliases: ["TOC"] },
  { id: "insert.caption", label: "Caption", path: ["Insert"], scope: "document" },
  { id: "insert.table-of-figures", label: "Table of figures", path: ["Insert"], scope: "document" },
  { id: "format.auto-caption-images", label: "Auto-caption new images", path: ["Insert"], scope: "document" },
  { id: "format.auto-caption-tables", label: "Auto-caption new tables", path: ["Insert"], scope: "document" },
  { id: "insert.special-character", label: "Special character", path: ["Insert"], scope: "document", icon: "special-character", aliases: ["symbol", "emoji"] },
  { id: "insert.tab-character", label: "Tab character", path: ["Insert"], scope: "document" },
  { id: "insert.link", label: "Link", path: ["Insert"], scope: "document", aliases: ["hyperlink", "url"] },
  { id: "insert.bookmark", label: "Bookmark", path: ["Insert"], scope: "document" },
  { id: "insert.bookmark-list", label: "Bookmark list", path: ["Insert"], scope: "document" },
  { id: "insert.comment", label: "Comment", path: ["Insert"], scope: "document", aliases: ["annotation"] },
  { id: "insert.pageBreak", label: "Page break", path: ["Insert"], scope: "document", icon: "page-break" },
  { id: "insert.equation", label: "Equation", path: ["Insert"], scope: "document", aliases: ["formula", "math"] },
  { id: "insert.display-equation", label: "Display equation", path: ["Insert"], scope: "document" },
  { id: "insert.numbered-equation", label: "Numbered equation", path: ["Insert"], scope: "document" },
  { id: "insert.footnote", label: "Footnote", path: ["Insert"], scope: "document", aliases: ["ftnt"] },
  { id: "insert.endnote", label: "Endnote", path: ["Insert"], scope: "document" },
  { id: "insert.nonBreakingSpace", label: "Non-breaking space", path: ["Insert"], scope: "document", aliases: ["nbsp"] },

  // --- Format ---------------------------------------------------------------
  { id: "format.styles-pane", label: "Styles pane", path: ["Format"], scope: "document" },
  { id: "format.import-styles", label: "Import styles from another document…", path: ["Format"], scope: "document" },
  { id: "format.clearFormatting", label: "Clear formatting", path: ["Format"], scope: "document", icon: "clear-formatting" },
  { id: "format.tabs", label: "Tabs…", path: ["Format"], scope: "document" },
  { id: "format.columns", label: "Columns", path: ["Format"], scope: "document" },
  { id: "format.multilevel-link-headings", label: "Number headings 1–3", path: ["Format", "Multilevel list"], scope: "document", icon: "multilevel-list" },
  { id: "format.multilevel-define-list", label: "Define new list…", path: ["Format", "Multilevel list"], scope: "document", icon: "multilevel-list" },
  { id: "format.text-color", label: "Text colour", path: ["Format"], scope: "document", icon: "text-color", aliases: ["font color"] },
  { id: "format.highlight", label: "Highlight colour", path: ["Format"], scope: "document" },
  { id: "format.bold", label: "Bold", path: ["Format"], scope: "document" },
  { id: "format.italic", label: "Italic", path: ["Format"], scope: "document" },
  { id: "format.underline", label: "Underline", path: ["Format"], scope: "document" },
  { id: "format.strikethrough", label: "Strikethrough", path: ["Format"], scope: "document" },
  { id: "format.subscript", label: "Subscript", path: ["Format"], scope: "document" },
  { id: "format.superscript", label: "Superscript", path: ["Format"], scope: "document" },
  { id: "paragraph.alignLeft", label: "Align left", path: ["Format"], scope: "document", icon: "align-left" },
  { id: "paragraph.alignCenter", label: "Align centre", path: ["Format"], scope: "document", icon: "align-center" },
  { id: "paragraph.alignRight", label: "Align right", path: ["Format"], scope: "document", icon: "align-right" },
  { id: "paragraph.alignJustify", label: "Justify", path: ["Format"], scope: "document", icon: "align-justify" },
  { id: "style.normal", label: "Apply Normal text", path: ["Format", "Styles"], scope: "document" },
  { id: "style.heading1", label: "Apply Heading 1", path: ["Format", "Styles"], scope: "document" },
  { id: "style.heading2", label: "Apply Heading 2", path: ["Format", "Styles"], scope: "document" },
  { id: "style.heading3", label: "Apply Heading 3", path: ["Format", "Styles"], scope: "document" },
  { id: "style.heading4", label: "Apply Heading 4", path: ["Format", "Styles"], scope: "document" },
  { id: "style.heading5", label: "Apply Heading 5", path: ["Format", "Styles"], scope: "document" },
  { id: "style.heading6", label: "Apply Heading 6", path: ["Format", "Styles"], scope: "document" },

  // --- Tools ----------------------------------------------------------------
  { id: "tools.word-count", label: "Word count", path: ["Tools"], scope: "document", aliases: ["statistics"] },
  { id: "tools.dictionary", label: "Custom dictionary", path: ["Tools"], scope: "document", aliases: ["spell"] },
  { id: "tools.accessibility", label: "Check accessibility", path: ["Tools"], scope: "document", aliases: ["a11y"] },
  { id: "tools.fonts", label: "Fonts", path: ["Tools"], scope: "document" },
  { id: "tools.update-all-fields", label: "Update all fields", path: ["Tools"], scope: "document" },
  { id: "tools.convert-merge-fields", label: "Convert {{…}} text to merge fields", path: ["Tools"], scope: "document" },

  // --- Library (SPEC-062: the palette's second surface) ---------------------
  { id: "library.new-document", label: "New document", path: ["Library"], scope: "library" },
  { id: "library.new-from-template", label: "New from template…", path: ["Library"], scope: "library" },
  { id: "library.open", label: "Open document", path: ["Library"], scope: "library" },
];

// Typed pass-throughs for CommandPalette.vue's untyped Array/Object props: a
// `.vue` script cannot annotate (espree bars TS syntax), so it casts through
// these — the same idiom as styles.ts's asStyleEntries.
export function asCommandMetas(value: unknown[]): CommandMeta[] {
  return value as CommandMeta[];
}
export function asIdList(value: unknown[]): string[] {
  return value as string[];
}
export function asReasonMap(value: object): Record<string, string> {
  return value as Record<string, string>;
}
/** Typed empty seed for the `.vue` script's highlight-ranges default param
 * (espree bars the annotation — same idiom as noStyleEntries). */
export function noRanges(): [number, number][] {
  return [];
}

// --- search & ranking --------------------------------------------------------

/** One matched command with its ranking and the label character ranges to
 * highlight (empty when the query matched only the path or aliases). */
export interface CommandMatch {
  meta: CommandMeta;
  /** Higher is better. Composed of a tier (see the ranking rules) plus a
   * within-tier bonus; never compared across queries. */
  score: number;
  /** `[start, end)` half-open ranges into `meta.label` to highlight. */
  ranges: [number, number][];
}

/** Case- and accent-insensitive fold — the form both the query and every
 * searchable field are compared in. */
export function foldText(text: string): string {
  return text
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

/** The positions of `query` as a subsequence of `target`, greedily matched
 * left to right, or null when it is not a subsequence. Both must already be
 * folded. */
function subsequencePositions(query: string, target: string): number[] | null {
  const positions: number[] = [];
  let qi = 0;
  for (let ti = 0; ti < target.length && qi < query.length; ti++) {
    if (target[ti] === query[qi]) {
      positions.push(ti);
      qi++;
    }
  }
  return qi === query.length ? positions : null;
}

/** Merges sorted single positions into `[start, end)` ranges for highlighting. */
function positionsToRanges(positions: number[]): [number, number][] {
  const ranges: [number, number][] = [];
  for (const pos of positions) {
    const last = ranges[ranges.length - 1];
    if (last && last[1] === pos) {
      last[1] = pos + 1;
    } else {
      ranges.push([pos, pos + 1]);
    }
  }
  return ranges;
}

const TIER_EXACT = 5000;
const TIER_PREFIX = 4000;
const TIER_SUBSTRING = 3000;
const TIER_SUBSEQUENCE = 2000;
const TIER_OTHER = 1000;

/** Scores one command against a folded, non-empty query. Returns null when
 * the query is not even a subsequence of the command's searchable text. */
function scoreCommand(query: string, meta: CommandMeta): CommandMatch | null {
  const label = foldText(meta.label);
  // Subsequence inclusion is decided over label + path + aliases together, so
  // "instab" finds "Insert → Table" and "toc" finds Table of contents by its
  // alias, even when the label alone would miss.
  const haystack = foldText([meta.label, ...meta.path, ...(meta.aliases ?? [])].join(" "));

  if (label === query) {
    return { meta, score: TIER_EXACT, ranges: [[0, meta.label.length]] };
  }
  if (label.startsWith(query)) {
    return { meta, score: TIER_PREFIX, ranges: [[0, query.length]] };
  }
  const substringAt = label.indexOf(query);
  if (substringAt >= 0) {
    // Earlier substring hits score higher within the tier.
    return {
      meta,
      score: TIER_SUBSTRING - substringAt,
      ranges: [[substringAt, substringAt + query.length]],
    };
  }
  const labelPositions = subsequencePositions(query, label);
  if (labelPositions) {
    // Reward matches that start early and stay contiguous.
    const firstIndex = labelPositions[0];
    let adjacency = 0;
    for (let i = 1; i < labelPositions.length; i++) {
      if (labelPositions[i] === labelPositions[i - 1] + 1) {
        adjacency++;
      }
    }
    const score = TIER_SUBSEQUENCE - firstIndex + adjacency * 10;
    return { meta, score, ranges: positionsToRanges(labelPositions) };
  }
  if (subsequencePositions(query, haystack)) {
    // Matched only in the path or aliases — no label characters to highlight.
    return { meta, score: TIER_OTHER, ranges: [] };
  }
  return null;
}

/** Filters and ranks the catalogue for a query (SPEC-062 "Matching and
 * ranking"). An empty query returns every command in catalogue order (the
 * palette's empty state layers recents/common on top itself). Ties break
 * toward more-recently-used (`recentIds`, most-recent-first), then
 * alphabetically by label.
 *
 * Enablement is deliberately not an input: it is a per-render concern the Vue
 * layer owns, so this function stays pure. The palette demotes disabled rows
 * within equal score after calling this. */
export function searchCommands(
  query: string,
  all: CommandMeta[] = COMMANDS,
  recentIds: string[] = [],
): CommandMatch[] {
  const trimmed = foldText(query.trim());
  if (trimmed === "") {
    return all.map((meta) => ({ meta, score: 0, ranges: [] }));
  }
  const matches: CommandMatch[] = [];
  for (const meta of all) {
    const match = scoreCommand(trimmed, meta);
    if (match) {
      matches.push(match);
    }
  }
  const recencyRank = (id: string): number => {
    const index = recentIds.indexOf(id);
    return index === -1 ? Number.POSITIVE_INFINITY : index;
  };
  return matches.sort((a, b) => {
    if (b.score !== a.score) {
      return b.score - a.score;
    }
    const recencyDiff = recencyRank(a.meta.id) - recencyRank(b.meta.id);
    if (recencyDiff !== 0) {
      return recencyDiff;
    }
    return a.meta.label.localeCompare(b.meta.label);
  });
}

// --- recents -----------------------------------------------------------------

/** How many recent commands the palette's empty state shows (SPEC-062). */
export const RECENT_COMMANDS_LIMIT = 8;

const RECENT_COMMANDS_KEY = "docu-edit:recent-commands";

/** The recently run command ids, most recent first — persisted like the
 * other `docu-edit:*` recents (`recent-text-colors`,
 * `recent-special-characters`). Ids no longer in the catalogue are dropped on
 * read. Returns [] where storage is unavailable (private mode, quota). */
export function loadRecentCommands(known: CommandMeta[] = COMMANDS): string[] {
  const ids = new Set(known.map((meta) => meta.id));
  try {
    const raw = localStorage.getItem(RECENT_COMMANDS_KEY);
    if (!raw) {
      return [];
    }
    const parsed: unknown = JSON.parse(raw);
    if (!Array.isArray(parsed)) {
      return [];
    }
    return parsed
      .filter((value): value is string => typeof value === "string" && ids.has(value))
      .slice(0, RECENT_COMMANDS_LIMIT);
  } catch {
    return [];
  }
}

/** Records a run command and returns the updated list: moved to the front,
 * de-duplicated, capped at RECENT_COMMANDS_LIMIT. A blocked `localStorage`
 * degrades to the returned in-session list. */
export function pushRecentCommand(id: string, current: string[]): string[] {
  const next = [id, ...current.filter((value) => value !== id)].slice(0, RECENT_COMMANDS_LIMIT);
  try {
    localStorage.setItem(RECENT_COMMANDS_KEY, JSON.stringify(next));
  } catch {
    // Storage full/unavailable: the in-memory list still updates.
  }
  return next;
}

/** The ids the empty-state "common commands" list shows after recents — a
 * short, curated set of everyday commands, filtered to a given scope. Kept
 * small on purpose (the spec's "short list of common ones"). */
export const COMMON_COMMAND_IDS: string[] = [
  "edit.undo",
  "edit.redo",
  "format.bold",
  "insert.link",
  "insert.comment",
  "insert.table-dialog",
  "find.openReplace",
  "tools.word-count",
];

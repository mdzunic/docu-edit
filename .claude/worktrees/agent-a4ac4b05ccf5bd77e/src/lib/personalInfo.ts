// "Remove personal information on save" (SPEC-041 AC-12a..14): a pure,
// format-agnostic transform the writers apply just before serialising when
// `model.properties?.removePersonalInfoOnSave` is set. It never mutates the
// live/stored model — only the copy that gets written to bytes — so turning
// the toggle back off later still shows the real names/dates.
//
// Comment/tracked-change author and timestamp fields are rewritten by
// structural shape rather than by walking every TrackedChangeInfo-bearing
// type in documentModel.ts by hand (RunTrackedChanges, ParagraphTrackedChanges,
// CellPropertyChange, RowPropertyChange, TablePropertyChange, ...): the model
// is plain JSON-serialisable data (see opaque.ts's header comment), and a
// `{changeId, author, date}` or `{messageId, author, createdAt}` shape
// identifies a tracked-change/comment record wherever it appears in the tree,
// so one generic recursive walk reaches all of them, including any future
// kind added without this module changing.
//
// Author names are replaced with "Author 1"/"Author 2"/... (AC-14), assigned
// by first appearance — comments in thread/message order, then tracked
// changes in document order (collectChanges) — the same original author
// always gets the same number within one save. Dates are replaced with the
// epoch (AC-13 "revision timestamps"): tracked-change dates and a comment
// thread's own date are schema-optional and could be omitted instead, but a
// reply message's timestamp is baked into its plain-text encoding
// (comments.ts's replyParagraphText) and MUST stay a parseable date for the
// thread to survive a reopen (parseReplyParagraphText rejects unparseable
// dates) — the epoch is used uniformly for both cases rather than omitting
// where possible and faking a date only where required, so anonymized output
// is consistent regardless of which kind of record it came from.
import type { DocumentModel } from "./documentModel";
import { collectChanges } from "./trackChanges";

/** The date every stripped revision timestamp is replaced with — matches
 * `docx/properties.ts`'s existing fallback for a tracked change with no
 * `w:date` at all, so an anonymized round trip reads the same as a document
 * that never carried dates for these edits. */
export const ANONYMIZED_DATE = new Date(0).toISOString();

/** Author -> "Author N" (AC-14), by first appearance across comments then
 * tracked changes. A name absent from either (shouldn't happen — every
 * author in the tree came from one of these two collections) falls back to
 * itself rather than throwing. */
function buildAuthorLabels(model: DocumentModel): Map<string, string> {
  const order: string[] = [];
  const seen = new Set<string>();
  const see = (author: string): void => {
    if (!seen.has(author)) {
      seen.add(author);
      order.push(author);
    }
  };
  for (const thread of model.comments ?? []) {
    for (const message of thread.messages) {
      see(message.author);
    }
  }
  for (const change of collectChanges(model)) {
    see(change.author);
  }
  return new Map(order.map((author, index) => [author, `Author ${index + 1}`]));
}

function isTrackedChangeLike(value: Record<string, unknown>): boolean {
  return (
    typeof value.changeId === "string" && typeof value.author === "string" && typeof value.date === "string"
  );
}

function isCommentMessageLike(value: Record<string, unknown>): boolean {
  return (
    typeof value.messageId === "string" && typeof value.author === "string" && typeof value.createdAt === "string"
  );
}

function rewrite(value: unknown, labels: Map<string, string>): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => rewrite(item, labels));
  }
  if (value === null || typeof value !== "object") {
    return value;
  }
  const next: Record<string, unknown> = {};
  for (const [key, child] of Object.entries(value as Record<string, unknown>)) {
    next[key] = rewrite(child, labels);
  }
  if (isTrackedChangeLike(next)) {
    next.author = labels.get(next.author as string) ?? next.author;
    next.date = ANONYMIZED_DATE;
  } else if (isCommentMessageLike(next)) {
    next.author = labels.get(next.author as string) ?? next.author;
    next.createdAt = ANONYMIZED_DATE;
  }
  return next;
}

/** Returns a copy of `model` with personal information removed for saving
 * (AC-12a..14): the document author and last-modified-by properties deleted,
 * every comment/tracked-change author replaced with a consistent "Author N"
 * label, and every comment/tracked-change timestamp replaced with the epoch.
 * `properties.created`/`modified` (the file's own metadata dates, not a
 * revision's) are left untouched, as are custom properties. */
export function stripPersonalInfo(model: DocumentModel): DocumentModel {
  const labels = buildAuthorLabels(model);
  const rewritten = rewrite(model, labels) as DocumentModel;
  if (!rewritten.properties) {
    return rewritten;
  }
  const properties = { ...rewritten.properties };
  delete properties.author;
  delete properties.lastModifiedBy;
  return { ...rewritten, properties };
}

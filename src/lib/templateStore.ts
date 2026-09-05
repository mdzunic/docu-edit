// User templates (SPEC-046): "Save as template" stores the current document
// as a new template; the gallery lists them alongside the shipped set
// (builtinTemplates.ts), with rename and delete. Mirrors documentStore.ts's
// in-memory-map-with-write-through shape, kept as its own module/store the
// same way version history (SPEC-040) is, so listing templates never has to
// read the much larger `documents` population.
import { createDocument, type DocumentModel } from "./documentModel";
import { deleteTemplateRecord, loadAllTemplateRecords, putTemplateRecord } from "./idb";

export interface StoredTemplateRecord {
  id: string;
  title: string;
  model: DocumentModel;
  /** Epoch milliseconds the template was saved. */
  createdAt: number;
}

export const MAX_TEMPLATE_TITLE_LENGTH = 200;

const templates = new Map<string, StoredTemplateRecord>();

function persist(record: StoredTemplateRecord): void {
  putTemplateRecord(record).catch(() => {});
}

let initPromise: Promise<void> | null = null;

/** Loads persisted templates into the in-memory map. Idempotent, and never
 * rejects — a failure just leaves the store empty (built-in templates still
 * work). */
export function initTemplates(): Promise<void> {
  initPromise ??= loadAllTemplateRecords()
    .then((stored) => {
      for (const record of stored) {
        templates.set(record.id, record);
      }
    })
    .catch(() => {});
  return initPromise;
}

/** User templates, newest first. */
export function listUserTemplates(): StoredTemplateRecord[] {
  return [...templates.values()].toSorted((a, b) => b.createdAt - a.createdAt);
}

export function getUserTemplate(id: string): StoredTemplateRecord | undefined {
  return templates.get(id);
}

/** "Save as template" (AC-5): stores an independent deep copy of the
 * document under a fresh id, titled from the requested name (or the
 * document's own title when blank). */
export function saveAsTemplate(model: DocumentModel, requestedTitle: string): StoredTemplateRecord {
  const trimmed = requestedTitle.trim();
  const title = trimmed.length > 0 ? trimmed.slice(0, MAX_TEMPLATE_TITLE_LENGTH) : model.title;
  const templateModel = structuredClone(model);
  const id = crypto.randomUUID();
  templateModel.id = id;
  templateModel.title = title;
  const record: StoredTemplateRecord = { id, title, model: templateModel, createdAt: Date.now() };
  templates.set(id, record);
  persist(record);
  return record;
}

/** Renames a user template (AC-7): trims, rejects empty/over-long titles
 * (returns null). Returns the title actually applied. */
export function renameTemplate(id: string, requestedTitle: string): string | null {
  const record = templates.get(id);
  if (!record) {
    return null;
  }
  const trimmed = requestedTitle.trim();
  if (trimmed.length === 0 || trimmed.length > MAX_TEMPLATE_TITLE_LENGTH) {
    return null;
  }
  record.title = trimmed;
  record.model.title = trimmed;
  persist(record);
  return trimmed;
}

/** Deletes a user template (AC-8). */
export function deleteTemplate(id: string): void {
  if (!templates.delete(id)) {
    return;
  }
  deleteTemplateRecord(id).catch(() => {});
}

/** A fresh, untitled-in-the-library-sense document built from a template: a
 * new id, the template's content — the same "new document from an existing
 * one" shape documentStore.duplicateDocument uses, minus the "(copy)" title
 * suffix (a new document from a template is not a copy of anything the user
 * made). Falls back to a blank document if the template no longer exists. */
export function instantiateTemplate(id: string): DocumentModel {
  const record = templates.get(id);
  if (!record) {
    return createDocument(crypto.randomUUID());
  }
  const model = structuredClone(record.model);
  model.id = crypto.randomUUID();
  return model;
}

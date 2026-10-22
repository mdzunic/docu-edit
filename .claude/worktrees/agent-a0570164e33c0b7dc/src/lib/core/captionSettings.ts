// The "Automatically caption new images and tables" setting (SPEC-055 AC-18,
// AC-19). A per-feature localStorage module, the same shape as
// core/autocorrect.ts — no Vue imports, no model access; the composable owns
// the ref and reads it when inserting an image or table.

export interface CaptionSettings {
  /** Auto-caption every newly inserted image. */
  autoImages: boolean;
  /** Auto-caption every newly inserted table. Independent of autoImages. */
  autoTables: boolean;
}

export const CAPTION_STORAGE_KEY = "docu-edit:captions";

/** Both off by default (SPEC-055 AC-18): auto-captioning is opt-in. */
export const DEFAULT_CAPTION_SETTINGS: CaptionSettings = {
  autoImages: false,
  autoTables: false,
};

/** Load persisted settings, tolerating absent/corrupt storage by falling back
 * to the defaults per field. */
export function loadCaptionSettings(): CaptionSettings {
  try {
    const raw = localStorage.getItem(CAPTION_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_CAPTION_SETTINGS };
    const parsed: unknown = JSON.parse(raw);
    if (typeof parsed !== "object" || parsed === null) {
      return { ...DEFAULT_CAPTION_SETTINGS };
    }
    const obj = parsed as Record<string, unknown>;
    return {
      autoImages: typeof obj.autoImages === "boolean" ? obj.autoImages : DEFAULT_CAPTION_SETTINGS.autoImages,
      autoTables: typeof obj.autoTables === "boolean" ? obj.autoTables : DEFAULT_CAPTION_SETTINGS.autoTables,
    };
  } catch {
    return { ...DEFAULT_CAPTION_SETTINGS };
  }
}

export function saveCaptionSettings(settings: CaptionSettings): void {
  try {
    localStorage.setItem(CAPTION_STORAGE_KEY, JSON.stringify(settings));
  } catch {
    // Persistence is best-effort; the in-memory setting still applies.
  }
}

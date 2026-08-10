// Custom document property validation (SPEC-041 AC-7, AC-9, AC-11): pure
// helpers the Custom Properties tab uses to validate a name/value before
// committing it to the model, and to canonicalise a value into the string
// encoding `CustomProperty.value` uses for its type (see documentModel.ts's
// `CustomPropertyType` doc comment).
import type { CustomProperty, CustomPropertyType } from "./documentModel";

const DATE_PATTERN = /^\d{4}-\d{2}-\d{2}$/;

/** Whether `value` is a real calendar date in "YYYY-MM-DD" form (rejects
 * e.g. "2024-02-30"). */
function isValidDate(value: string): boolean {
  if (!DATE_PATTERN.test(value)) {
    return false;
  }
  const [year, month, day] = value.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

/** A custom property name's validation error (AC-7), or "" when valid:
 * non-empty and unique among the document's other custom property names
 * (case-sensitive, matching how `.docx`/`.odt` compare property names). */
export function customPropertyNameError(name: string, existingNames: string[]): string {
  const trimmed = name.trim();
  if (trimmed.length === 0) {
    return "Property name cannot be empty";
  }
  if (existingNames.includes(trimmed)) {
    return "A custom property with this name already exists";
  }
  return "";
}

/** A typed custom property value's validation error (AC-11), or "" when
 * valid — e.g. non-numeric text for a "number" type. */
export function customPropertyValueError(type: CustomPropertyType, rawValue: string): string {
  const trimmed = rawValue.trim();
  switch (type) {
    case "text":
      return "";
    case "number":
      return trimmed.length > 0 && Number.isFinite(Number(trimmed)) ? "" : "Enter a number";
    case "date":
      return isValidDate(trimmed) ? "" : "Enter a date as YYYY-MM-DD";
    case "boolean":
      return trimmed === "true" || trimmed === "false" ? "" : "Choose Yes or No";
  }
}

/** `rawValue` normalised into `CustomProperty.value`'s canonical encoding
 * for `type` — the caller checks `customPropertyValueError` first; passing
 * an invalid value returns it unchanged. */
export function canonicalCustomPropertyValue(type: CustomPropertyType, rawValue: string): string {
  const trimmed = rawValue.trim();
  if (type === "number" && customPropertyValueError(type, trimmed) === "") {
    return String(Number(trimmed));
  }
  return trimmed;
}

/** The default (empty) value a fresh row of `type` starts from — "false"
 * for boolean so its Yes/No control always has a valid selection. */
export function defaultCustomPropertyValue(type: CustomPropertyType): string {
  return type === "boolean" ? "false" : "";
}

/** A fresh, empty custom property of `type`, ready to add to
 * `DocumentProperties.customProperties` once its name is filled in. */
export function createCustomProperty(name: string, type: CustomPropertyType, value: string): CustomProperty {
  return { name: name.trim(), type, value: canonicalCustomPropertyValue(type, value) };
}

// --- typed seeds for PropertiesDialog.vue -----------------------------------
// Declared here rather than inline in the `.vue` file: its script is parsed
// by plain espree (see AppShell.vue's espree comment), which rejects the
// generics/annotations a literal's precise type would otherwise need.

/** The Custom Properties tab's "add a property" form state. */
export interface NewCustomPropertyDraft {
  name: string;
  type: CustomPropertyType;
  value: string;
}

/** Typed empty seed for the add-property form. */
export function emptyNewCustomPropertyDraft(): NewCustomPropertyDraft {
  return { name: "", type: "text", value: "" };
}

/** One existing row's in-progress text, kept separate from the committed
 * `CustomProperty` until it validates (AC-11). */
export interface CustomPropertyRowDraft {
  name: string;
  value: string;
}

/** Typed empty seed for the Custom Properties tab's per-row drafts, keyed by
 * array index. */
export function emptyCustomPropertyRowDrafts(): Record<number, CustomPropertyRowDraft> {
  return {};
}

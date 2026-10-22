---
id: SPEC-066
title: Preferences
status: draft
area: shell
depends_on: [SPEC-058, SPEC-059, SPEC-064]
---
## Why
The product has preferences. It has no Preferences. Settings that outlive a
session are scattered across eight `localStorage` keys written by whichever
spec needed one —

```
docu-edit:autocorrect              docu-edit:outline-pane
docu-edit:comment-author           docu-edit:recent-text-colors
docu-edit:comments-pane            docu-edit:recent-special-characters
docu-edit:custom-dictionary        docu-edit:version-history-pane
```

— each read and written by its own module (`src/lib/core/autocorrect.ts`,
`src/lib/outlinePane.ts`, `src/lib/commentsPane.ts`,
`src/lib/customDictionary.ts`, `src/lib/core/colors.ts`,
`src/lib/core/commentIdentity.ts`), with its own encoding, its own default
and its own handling of a blocked `localStorage`. Others are stored *in the
document* even though they describe the app, not the file:
`DocumentModel.measurementUnit`, `showRuler`, `showFormattingMarks`,
`suppressConversionReport`. And some settings a user would expect simply
cannot be expressed — there is no way to say "new documents should be
Times New Roman 12" or "autosave every 5 seconds" or "always show
formatting marks".

The immediate trigger is that SPEC-059 adds a theme, SPEC-063 adds a
default zoom and view mode, and SPEC-064 makes the measurement unit
app-wide. Each of those either needs a settings home or invents a ninth
ad-hoc key. This spec builds the home once, gives the existing settings one
storage shape and one defaulting rule, and makes "app preference" a thing
with a place — the way SPEC-049 made "a command has a menu" a thing with a
rule.

## Acceptance criteria
- [ ] A Preferences dialog opens from `File → Preferences…` and, on macOS, from the conventional Cmd+, binding registered in `src/lib/shortcuts.ts` as `app.preferences`
- [ ] The dialog has the sections and settings in "The settings", each with its current value, a description where the name is not self-explanatory, and its default marked
- [ ] Changing a setting applies **immediately and live** — no OK button, no restart, no "apply" step — and the dialog's close button just closes it
- [ ] Every preference is stored under one key, `docu-edit:preferences`, as a single JSON object; the eight legacy keys are migrated on first read and then removed, per "Storage and migration"
- [ ] A missing, corrupt or partially-populated stored object falls back per-field to the declared default; one bad field never discards the rest, asserted by a unit test over a truncated and a malformed payload
- [ ] Preferences are read through one typed module (`src/lib/preferences.ts`) with a declared schema, defaults and validators; no component reads `localStorage` for a preference after this spec
- [ ] A blocked or full `localStorage` (SPEC-003) degrades to in-session preferences with the storage warning already shown, and never throws
- [ ] Settings that are genuinely per-document stay per-document, per "App preference or document property?", and the dialog does not offer them
- [ ] The document-level app-preference fields that are really app defaults — `measurementUnit`, `showRuler`, `showFormattingMarks` — gain an app-level default that a **new** document inherits, while an existing document keeps its own stored value
- [ ] "Reset all preferences" restores every default after a confirmation, and does not touch documents, drafts, the custom dictionary's word list, or version history
- [ ] The dialog is built on `BaseDialog.vue` with SPEC-058's `UiField` primitives, is keyboard-navigable throughout, groups settings under headed sections, and each control is labelled and described by its help text
- [ ] The Preferences dialog is reachable from the library view as well as the editor, since most of what it holds is not document-specific
- [ ] An e2e test changes the theme, the measurement unit and the default font, reloads, and asserts all three survived

## Out of scope
Any new *feature* behind a preference — every setting here toggles
something that already exists or is specced elsewhere. Per-document
settings UI (Page setup and Document properties already own those).
Syncing preferences between devices or exporting them; there is no account
(SPEC-001). Customising keyboard shortcuts, menus or toolbars. Enterprise
policy / admin-locked settings. A "workspace" or profile concept.

## Reference

### App preference or document property?
The rule this spec establishes, because the current split is accidental:

> A setting belongs to **preferences** when it describes how *this person
> on this device* wants the application to behave, and to the **document**
> when it describes the file or when another application opening the file
> would need it.

By that rule `measurementUnit`, `showRuler` and `showFormattingMarks` are
preferences that the document happens to remember per file — a reasonable
behaviour, since you may want marks on in one document and off in another.
So they get both: an app-level default used for new documents and for
documents with no stored value, and a per-document override that the
existing fields already provide. Nothing about their model storage or
round-tripping changes (none of the three maps to anything in `.docx` or
`.odt`).

`suppressConversionReport` stays document-only: it is a fact about one
file's import, not a preference.

Page size, margins, language, styles and every field in Document
properties stay document properties, and the dialog does not show them.

### The settings

**Appearance**
| Setting | Values | Default |
| --- | --- | --- |
| Theme | System / Light / Dark | System |
| Dim page in dark theme | on / off | off |
| Measurement unit | cm / mm / in / pt | cm |
| Default zoom for new documents | the SPEC-063 ladder, Fit width, Fit page | 100% |
| Default view mode | Print / Web | Print |
| Show ruler by default | on / off | on |
| Show formatting marks by default | on / off | off |

**Editing**
| Setting | Values | Default |
| --- | --- | --- |
| Smart quotes | on / off | on |
| Em and en dashes | on / off | on |
| Check spelling as you type | on / off | on |
| Default document language | the SPEC-035 language list | en-US |
| Author name for comments and tracked changes | text | the SPEC-038 identity |

**New documents**
| Setting | Values | Default |
| --- | --- | --- |
| Default font family | the SPEC-013 font list | Calibri |
| Default font size | points | 11 |
| Default page size | the SPEC-027 sizes | Letter or A4 by locale |
| Default margins | the SPEC-027 presets | Normal |

**Saving**
| Setting | Values | Default |
| --- | --- | --- |
| Autosave | on / off | on |
| Autosave delay | 1 s / 3 s / 5 s / 10 s | the SPEC-010 debounce |
| Keep version history | on / off | on |
| Default download format | Same as source / .docx / .odt | Same as source |

**Advanced**
| Setting | Values | Default |
| --- | --- | --- |
| Show the conversion report after a cross-format save | on / off | on |
| Reset all preferences | action | — |

Each row names a setting that already has an implementation to drive
(SPEC-010's debounce, SPEC-013's font list, SPEC-035's language set,
SPEC-040's history) or a spec that adds it (SPEC-059's theme, SPEC-063's
zoom and mode, SPEC-064's unit). "Autosave: off" leaves crash-recovery
drafts working — it stops the *committed* write, not the safety net, so no
edit is ever unrecoverable.

### The module
```ts
// src/lib/preferences.ts — framework-free, unit-testable
export interface Preferences {
  theme: "system" | "light" | "dark";
  dimPage: boolean;
  measurementUnit: MeasurementUnit;
  defaultZoom: number | "fit-width" | "fit-page";
  defaultViewMode: "print" | "web";
  showRuler: boolean;
  showFormattingMarks: boolean;
  smartQuotes: boolean;
  smartDashes: boolean;
  spellCheck: boolean;
  defaultLanguage: string;
  authorName: string;
  defaultFontFamily: string;
  defaultFontSizePt: number;
  defaultPageSizeId: string;
  defaultMarginsId: string;
  autosave: boolean;
  autosaveDelayMs: number;
  keepVersionHistory: boolean;
  defaultDownloadFormat: "source" | "docx" | "odt";
  showConversionReport: boolean;
}

export const PREFERENCE_DEFAULTS: Preferences;
export function loadPreferences(): Preferences;
export function savePreference<K extends keyof Preferences>(
  key: K, value: Preferences[K]
): void;
export function resetPreferences(): void;
export function onPreferencesChange(listener: (prefs: Preferences) => void): () => void;
```

Reads validate per field: a value of the wrong type, or outside its
declared set, is replaced by that field's default and the rest is kept.
`onPreferencesChange` is how the theme composable, the ruler and the
autosave scheduler react without polling — and it fires for changes made in
another tab (via the `storage` event), so two open tabs stay consistent.

A thin Vue wrapper (`src/composables/usePreferences.ts`) exposes the object
as reactive refs; components use that, never `localStorage`.

### Storage and migration
One key, `docu-edit:preferences`, holding the JSON object. On first load
after this spec ships, the legacy keys are folded in and deleted:

| Legacy key | Becomes |
| --- | --- |
| `docu-edit:autocorrect` | `smartQuotes`, `smartDashes` |
| `docu-edit:comment-author` | `authorName` |
| `docu-edit:theme`, `docu-edit:dim-page` (SPEC-059) | `theme`, `dimPage` |

Three keys are **not** migrated, because they are not preferences:
`docu-edit:custom-dictionary` (user data — a word list),
`docu-edit:recent-text-colors` and `docu-edit:recent-special-characters`
(MRU state). `docu-edit:outline-pane`, `docu-edit:comments-pane` and
`docu-edit:version-history-pane` are SPEC-065's to migrate into
`docu-edit:side-panel`, not this spec's — layout state, not preferences.
Recording both lists here is the point: the boundary is stated once so the
next spec does not have to guess.

### Specs this changes
- `specs/027-page-setup.md` — the measurement unit and ruler visibility
  gain app-level defaults; their per-document storage is unchanged.
- `specs/028-page-and-section-breaks.md` — formatting-mark visibility gains
  an app-level default; unchanged otherwise.
- `specs/010-autosave-and-crash-recovery.md` — the debounce interval
  becomes configurable and autosave becomes disableable, with crash
  recovery unaffected either way. Its mechanism is untouched.
- `specs/002-create-blank-document.md` — a new blank document takes its
  font, size, page size and margins from preferences instead of from
  hard-coded defaults. The *model* defaults it documents stay the fallback
  when no preference is set, so nothing in that spec's normative model
  changes.
- `specs/038-comments.md` / `specs/039-track-changes.md` — the author name
  moves into preferences; identity behaviour is unchanged.

# docu-edit

Browser-based word processor that opens, edits and saves `.docx` (OOXML) and
`.odt` (ODF) documents. Built spec-by-spec by an automated factory; `specs/`
holds the queue and `specs/000-roadmap.md` the intended order.

## Stack

- **Vue 3** with `<script setup lang="ts">` SFCs — Composition API, not Options API
- **TypeScript** strict mode
- **Vite** dev server and build (port 5173)
- **vue-router** — routes declared in `src/router.ts`
- **HTML Canvas** for complex shapes and page rendering where DOM is a poor fit
- **Vitest** + `@vue/test-utils` for unit tests
- **Playwright** for e2e — pinned to **1.49.1** to match the CI container's
  preinstalled browsers. Do not change this version.
- **OXLint** for `.ts`, **ESLint** (`eslint-plugin-vue`) for `.vue` templates

## Commands

| Task | Command |
| --- | --- |
| Dev server | `npm run dev` |
| Unit tests | `npm test` |
| E2E tests | `npx playwright test` |
| Typecheck | `npm run typecheck` |
| Lint | `npm run lint` |

## Layout

```
src/
  main.ts          app entry — mounts to #app, installs the router
  router.ts        route table; add routes here, not inline
  App.vue          root shell — renders <RouterView />
  views/           one component per route
    library/       the landing page's library module (product-neutral UI)
    document/      the text editor's composables
  components/      reusable UI
  lib/             framework-free logic (parsers, formatters, converters)
    core/          product-neutral layer (may import only core siblings)
    docx/ odf/ md/ the TEXT product's format readers/writers
    presentation/  the PRESENTATION product (.pptx/.odp); imports core only
e2e/               Playwright specs; e2e/SPEC-NNN.spec.ts per feature
```

Keep document parsing and serialisation in `src/lib/` as plain TypeScript with
no Vue imports — it must be unit-testable without mounting a component, and
`.docx`/`.odt`/`.pptx`/`.odp` handling is the part most likely to be reused
headlessly.

Two products share this repo: the word processor (documentModel.ts +
docx/odf/md, `/doc/:id`) and presentations (lib/presentation, `/slides/:id`).
They must never import each other — anything both need is format/product-
neutral by definition and belongs in `src/lib/core`; the app layer
(`documentKinds.ts`, `openDocument.ts`, `documentStore.ts`, the router, the
views) is where they compose. Enforced by `core/boundary.spec.ts` and
`presentation/boundary.spec.ts`; the fuller map is
`docs/presentations-architecture.md`.

## Conventions

- One component per file, `PascalCase.vue`; composables `useThing.ts`
- `src/views/DocumentView.vue` is the editor's wiring hub, not a home for
  feature code: a feature's logic goes in a `src/views/document/` composable
  (its `*Deps` interface extends `EditorContext` from
  `src/views/document/context.ts`) and its UI in components; DocumentView
  gains only a `use*({ ...editorContext, ...extras })` wiring call and the
  template hookup. `DocumentView.size.spec.ts` enforces a line budget that
  only ratchets down — when it fails, extract; never raise the budget.
- `src/views/HomeView.vue` is the landing page's wiring hub under the same
  rule (`HomeView.size.spec.ts`): library UI belongs in a product-neutral
  `src/views/library/` component; HomeView keeps only product wiring
  (create/open actions, kind routing via `lib/documentKinds.ts`, dialogs).
- Props and emits fully typed; no `any`
- Scoped styles by default; global styles only in `App.vue`
- Prefer semantic HTML and ARIA roles — e2e tests select by role and test-id,
  so `data-testid` on anything a test needs to find
- Unit tests live next to their subject as `*.spec.ts`

## Testing rules

- `src/harness.spec.ts` and `e2e/smoke.spec.ts` are baseline checks that must
  keep passing; they assert nothing about product behaviour, so a feature
  should never need to change them.
- Add unit tests for logic in `src/lib/`, e2e tests for user-visible behaviour.
- Never weaken or delete an existing test to make a change pass.

## Constraints

- `package.json` is protected: dependencies are not added by feature work. If a
  spec genuinely cannot be built with the current dependency set, stop and say
  so rather than working around it. Two sanctioned exceptions, both pinned by
  `src/lib/math/noDependencies.spec.ts`:
  - `pdfjs-dist` (Apache-2.0, zero runtime deps), approved 2026-08-30 for PDF
    first-page previews — loaded lazily in `src/lib/pdfRaster.ts` only (the app
    layer; `lib/media` stays framework- and package-free).
  - `wordlist-english` (MIT, zero runtime deps), approved 2026-09-04 for
    SPEC-035's English dictionaries. Data only — SCOWL word lists as JSON, no
    engine — imported lazily by `src/lib/spellcheck/dictionaries/enBase.ts`,
    `enUS.ts` and `enGB.ts`. The hand-curated lists it replaced recognised 659
    words and underlined roughly half of ordinary prose. It is English-only on
    purpose: the equivalent packages for German, Czech and Spanish are GPL, so
    `deDE`/`frFR`/`esES`/`csCZ` stay curated lists.
- Round-trip fidelity matters more than feature count: never discard document
  content you did not parse — preserve unknown parts verbatim (see SPEC-008).

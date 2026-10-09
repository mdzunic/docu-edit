// The checked-in golden for pptx deck sections (SPEC-099, SPEC-086's process
// applying unchanged): the wire detail this spec is least certain of — the
// `p:ext` uri GUID, the `p14` namespace URI and the element names — pinned by
// a literal rather than by memory. `SECTIONS_GOLDEN` is the exact `p:ext`
// block `sectionsExtXml` emits for `GOLDEN_SECTIONS`, and the block the reader
// models back into it; PowerPoint verification is recorded in the PR, not CI.
// The `xmlns:p14` declaration rides on the `p14:sectionLst` element itself, as
// PowerPoint writes it, so the block is well-formed wherever it is spliced.
import type { DeckSection } from "../model";

export const SECTIONS_GOLDEN =
  `<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">` +
  `<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">` +
  `<p14:section name="Intro" id="{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}">` +
  `<p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst>` +
  `</p14:section>` +
  `<p14:section name="Second" id="{B1D0F3A2-7E4C-4A9B-8C21-6F3E5D7A9B10}">` +
  `<p14:sldIdLst><p14:sldId id="258"/></p14:sldIdLst>` +
  `</p14:section>` +
  `</p14:sectionLst></p:ext>`;

/** The model `SECTIONS_GOLDEN` encodes: two sections over a three-slide deck
 * (`sectionKeys[i]` is slide i's `sectionKey`). */
export const GOLDEN_SECTIONS: { sections: DeckSection[]; sectionKeys: string[] } = {
  sections: [
    { key: "{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}", name: "Intro" },
    { key: "{B1D0F3A2-7E4C-4A9B-8C21-6F3E5D7A9B10}", name: "Second" },
  ],
  sectionKeys: [
    "{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}",
    "{5C7C9E4A-1C6B-4E2E-9A7F-2E5B1D0F3A11}",
    "{B1D0F3A2-7E4C-4A9B-8C21-6F3E5D7A9B10}",
  ],
};

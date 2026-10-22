// A browser always has a global `DOMParser`/`XMLSerializer`/`Node`; plain
// Node.js never does. The existing readers (docx/reader.ts, odf/reader.ts)
// parse XML with `DOMParser` and walk it with the DOM `Node` interface's
// constants (e.g. odf/reader.ts's `Node.TEXT_NODE`) — reused unchanged here
// per SPEC-045's Reference section ("adds no new serialisation code") — so
// a genuine Node.js caller of `readDocx`/`readOdt` (AC-17) needs these
// supplied. This lazily borrows them from the project's own `jsdom` test
// dependency, already installed for `npm test`'s jsdom environment; a
// browser never takes this branch since `typeof DOMParser` is never
// "undefined" there.
let ensured: Promise<void> | undefined;

export function ensureNodeXmlGlobals(): Promise<void> {
  if (typeof DOMParser !== "undefined") {
    return Promise.resolve();
  }
  if (!ensured) {
    ensured = import("jsdom").then(({ JSDOM }) => {
      const { window } = new JSDOM("");
      globalThis.DOMParser = window.DOMParser;
      globalThis.XMLSerializer = window.XMLSerializer;
      globalThis.Node = window.Node;
    });
  }
  return ensured;
}

// Minimal ambient typing for the one jsdom export nodeDom.ts uses. jsdom
// ships no types of its own and this project does not depend on
// `@types/jsdom` (package.json is protected — see CLAUDE.md); jsdom is
// already a real devDependency (vitest's jsdom test environment), just an
// untyped one, so this declares only the shape actually called.
declare module "jsdom" {
  export class JSDOM {
    constructor(html?: string);
    window: {
      DOMParser: typeof DOMParser;
      XMLSerializer: typeof XMLSerializer;
      Node: typeof Node;
    };
  }
}

// SPEC-114: the word processor's `find` deep link. /doc/:id?find=<text>
// (&n=<k>) — how a library search result opens AT its match — opens
// SPEC-017's find bar with the text compiled and the nth match selected
// (the first by default). It changes nothing about the find bar: the caret
// moves to the start of the match the link names, the bar opens, and the
// text goes into it — the bar then selects the first match at or after the
// caret, exactly as it does for anything typed into it.
//
// Applied once per document and link, as soon as the session can search
// (the edit lock is settled); a reload of the same link applies it again.
import { watch, type Ref } from "vue";
import { useRoute } from "vue-router";
import { compileQuery, findMatches, isQueryError, matchSelection } from "../../lib/findReplace";
import { findParams } from "../../lib/findLink";
import type { EditorContext } from "./context";

export interface FindLinkDeps extends EditorContext {
  /** Whether Find may open (useFindReplace's own gate). */
  canSearch: Readonly<Ref<boolean>>;
  findQuery: Ref<string>;
  findMatchCase: Ref<boolean>;
  findWholeWords: Ref<boolean>;
  findUseRegex: Ref<boolean>;
  openFind: (expand?: boolean) => void;
}

export function useFindLink(deps: FindLinkDeps): void {
  const route = useRoute();
  let applied = "";
  watch(
    [deps.canSearch, deps.documentId, () => route.query.find, () => route.query.n],
    () => {
      const model = deps.doc.value;
      const { find, n } = findParams(route.query);
      const key = `${deps.documentId.value}\n${find}\n${n}`;
      if (!deps.canSearch.value || !model || find === "" || key === applied) {
        return;
      }
      applied = key;
      // The nth match as the bar will count it, under the bar's own options.
      const compiled = compileQuery(find, {
        matchCase: deps.findMatchCase.value,
        wholeWords: deps.findWholeWords.value,
        regex: deps.findUseRegex.value,
      });
      const matches = compiled !== null && !isQueryError(compiled) ? findMatches(model, compiled) : [];
      const target = matches[Math.min(n, matches.length) - 1];
      if (target) {
        deps.applyCaretMove(matchSelection(target).anchor);
      }
      deps.openFind(false);
      deps.findQuery.value = find;
    },
    { immediate: true },
  );
}

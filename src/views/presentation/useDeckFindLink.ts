// SPEC-114: the deck's `find` deep link. /slides/:id?find=<text>(&n=<k>) —
// how a library search result opens a deck AT its match — opens SPEC-098's
// deck find with the text and, once its walk settles, makes the nth hit
// current (the first by default): its slide becomes current and the match is
// selected, exactly as a click on that result row does. Nothing about the
// deck find changes. Applied once per deck and link.
import { watch, type ShallowRef } from "vue";
import type { RouteLocationNormalizedLoaded } from "vue-router";
import type { PresentationModel } from "../../lib/presentation/model";
import { findParams } from "../../lib/findLink";
import type { DeckSearch } from "./useDeckSearch";

export interface DeckFindLinkDeps {
  route: RouteLocationNormalizedLoaded;
  presentation: ShallowRef<PresentationModel | null>;
  search: DeckSearch;
}

export function useDeckFindLink(deps: DeckFindLinkDeps): void {
  const { search } = deps;
  let applied = "";
  watch(
    [() => deps.presentation.value?.id, () => deps.route.query.find, () => deps.route.query.n],
    ([id]) => {
      const { find, n } = findParams(deps.route.query);
      const key = `${id}\n${find}\n${n}`;
      if (!id || find === "" || key === applied) {
        return;
      }
      applied = key;
      search.query.value = find;
      search.openFind();
      // The walk is debounced and sliced: select once it has settled.
      let settledOnce = false;
      const stop = watch(
        [search.pending, search.indexing, search.hits],
        () => {
          if (settledOnce || search.pending.value || search.indexing.value || search.query.value !== find) {
            return;
          }
          settledOnce = true;
          stop();
          const hits = search.hits.value;
          if (hits.length > 0) {
            search.goToHit(Math.min(n, hits.length) - 1);
          }
        },
        { flush: "post" },
      );
    },
    { immediate: true },
  );
}

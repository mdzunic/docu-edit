import type { RouteLocationNormalized, RouteRecordRaw } from "vue-router";
import { documentKindOf } from "./lib/documentStore";
import { editorPathFor, type DocumentKind } from "./lib/documentKinds";

// Route table for the app. SPEC-001 adds the shell at `/` and the
// document workspace at `/doc/:id`; later specs extend this list.
//
// The editor routes guard each other's kind (see documentKinds.ts): a
// stored presentation id navigated to /doc/:id lands on /slides/:id, a media
// id on /media/:id (SPEC-107), and so on in every direction, so no product
// ever loads another's model. The guards run only in the mounted app
// (direct mounts in tests bypass them); main.ts awaits initLibrary() before
// mounting, so the store can answer.

/** The guard for a route owned by `kind`: an id stored as any other kind is
 * sent to that kind's own route; an unknown id stays put. */
function kindGuard(kind: DocumentKind) {
  return (to: RouteLocationNormalized) => {
    const id = String(to.params.id ?? "");
    const stored = documentKindOf(id);
    return stored === undefined || stored === kind ? true : editorPathFor(stored, id);
  };
}

export const routes: RouteRecordRaw[] = [
  {
    path: "/",
    name: "home",
    component: () => import("./views/HomeView.vue"),
  },
  {
    path: "/doc/:id",
    name: "document",
    component: () => import("./views/DocumentView.vue"),
    beforeEnter: kindGuard("text"),
  },
  {
    path: "/slides/:id",
    name: "slides",
    component: () => import("./views/PresentationView.vue"),
    beforeEnter: kindGuard("presentation"),
  },
  {
    // SPEC-084: full-screen, chrome-free slideshow playback (1-based `n`). The
    // kind guard mirrors /slides/:id's so a text id hands off to /doc/:id.
    path: "/slides/:id/present/:n",
    name: "slides-present",
    component: () => import("./views/PresentationPlayView.vue"),
    beforeEnter: kindGuard("presentation"),
  },
  {
    // SPEC-107: the media viewer, kind-guarded in both directions exactly as
    // /doc/:id and /slides/:id are.
    path: "/media/:id",
    name: "media",
    component: () => import("./views/MediaView.vue"),
    beforeEnter: kindGuard("media"),
  },
  {
    path: "/open",
    name: "open-from-url",
    component: () => import("./views/OpenFromUrlView.vue"),
  },
];

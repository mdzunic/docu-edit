import type { RouteLocationNormalized, RouteRecordRaw } from "vue-router";
import { documentKindOf } from "./lib/documentStore";
import { editorPathFor, type DocumentKind } from "./lib/documentKinds";
import { isExperimentEnabled, type ExperimentId } from "./lib/experiments";

// Route table for the app. SPEC-001 adds the shell at `/` and the
// document workspace at `/doc/:id`; later specs extend this list.
//
// The editor routes guard each other's kind (see documentKinds.ts) through
// ONE helper, never pairwise: an id stored as any other kind than the
// route's is sent to editorPathFor(its kind, id) — a presentation id at
// /doc/:id lands on /slides/:id, a media id on /media/:id (SPEC-107), a
// workbook id on /sheets/:id (SPEC-116), and so on in every direction — so
// no product ever loads another's model, and a fifth kind needs no edit
// here beyond its own route. The guards run only in the mounted app (direct
// mounts in tests bypass them); main.ts awaits initLibrary() before
// mounting, so the store can answer.

/** The guard for a route owned by `expected`: an id stored as any other
 * kind is sent to that kind's own route; an unknown id stays put. */
export function guardKind(expected: DocumentKind) {
  return (to: RouteLocationNormalized) => {
    const id = String(to.params.id ?? "");
    const stored = documentKindOf(id);
    return stored === undefined || stored === expected ? true : editorPathFor(stored, id);
  };
}

// SPEC-113: an experimental feature's route is absent while its Labs flag is
// off — the guard sends it to the library before the route's component (a
// lazy import(), like every route here) is ever loaded. A guard only reads
// the flag; it never makes a request.

/** The guard for a route that belongs to experiment `id`: "/" while the
 * experiment is off. */
export function guardExperiment(id: ExperimentId) {
  return () => (isExperimentEnabled(id) ? true : "/");
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
    beforeEnter: guardKind("text"),
  },
  {
    path: "/slides/:id",
    name: "slides",
    component: () => import("./views/PresentationView.vue"),
    beforeEnter: guardKind("presentation"),
  },
  {
    // SPEC-084: full-screen, chrome-free slideshow playback (1-based `n`). The
    // kind guard mirrors /slides/:id's so any other kind's id hands off.
    path: "/slides/:id/present/:n",
    name: "slides-present",
    component: () => import("./views/PresentationPlayView.vue"),
    beforeEnter: guardKind("presentation"),
  },
  {
    // SPEC-100: presenter view (1-based `n`) — this tab drives the show and
    // opens /present/:n?follow=… as the audience window. Kind-guarded like the
    // other two slide routes, so a text id lands on /doc/:id.
    path: "/slides/:id/presenter/:n",
    name: "slides-presenter",
    component: () => import("./views/PresenterView.vue"),
    beforeEnter: guardKind("presentation"),
  },
  {
    // SPEC-107: the media viewer, kind-guarded like every editor route.
    path: "/media/:id",
    name: "media",
    component: () => import("./views/MediaView.vue"),
    beforeEnter: guardKind("media"),
  },
  {
    // SPEC-116: the read-only workbook view.
    path: "/sheets/:id",
    name: "sheets",
    component: () => import("./views/SpreadsheetView.vue"),
    beforeEnter: guardKind("spreadsheet"),
  },
  {
    // SPEC-109: the assembly workspace. It takes no id, so no kind guard.
    path: "/assemble",
    name: "assemble",
    component: () => import("./views/AssembleView.vue"),
  },
  {
    path: "/open",
    name: "open-from-url",
    component: () => import("./views/OpenFromUrlView.vue"),
  },
];

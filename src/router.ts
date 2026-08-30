import type { RouteRecordRaw } from "vue-router";
import { documentKindOf } from "./lib/documentStore";

// Route table for the app. SPEC-001 adds the shell at `/` and the
// document workspace at `/doc/:id`; later specs extend this list.
//
// The two editor routes guard each other's kind (see documentKinds.ts): a
// stored presentation id navigated to /doc/:id lands on /slides/:id and
// vice versa, so neither product ever loads the other's model. The guards
// run only in the mounted app (direct mounts in tests bypass them);
// main.ts awaits initLibrary() before mounting, so the store can answer.
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
    beforeEnter: (to) => {
      const id = String(to.params.id ?? "");
      const kind = documentKindOf(id);
      if (kind === "presentation") {
        return `/slides/${id}`;
      }
      return kind === "media" ? `/media/${id}` : true;
    },
  },
  {
    path: "/slides/:id",
    name: "slides",
    component: () => import("./views/PresentationView.vue"),
    beforeEnter: (to) => {
      const id = String(to.params.id ?? "");
      const kind = documentKindOf(id);
      if (kind === "text") {
        return `/doc/${id}`;
      }
      return kind === "media" ? `/media/${id}` : true;
    },
  },
  {
    // SPEC-084: full-screen, chrome-free slideshow playback (1-based `n`). The
    // kind guard mirrors /slides/:id's so a text id hands off to /doc/:id.
    path: "/slides/:id/present/:n",
    name: "slides-present",
    component: () => import("./views/PresentationPlayView.vue"),
    beforeEnter: (to) => {
      const id = String(to.params.id ?? "");
      const kind = documentKindOf(id);
      if (kind === "text") {
        return `/doc/${id}`;
      }
      return kind === "media" ? `/media/${id}` : true;
    },
  },
  {
    // SPEC-107: the media viewer, kind-guarded in both directions exactly as
    // /doc/:id and /slides/:id are.
    path: "/media/:id",
    name: "media",
    component: () => import("./views/MediaView.vue"),
    beforeEnter: (to) => {
      const id = String(to.params.id ?? "");
      const kind = documentKindOf(id);
      if (kind === "text") {
        return `/doc/${id}`;
      }
      return kind === "presentation" ? `/slides/${id}` : true;
    },
  },
  {
    path: "/open",
    name: "open-from-url",
    component: () => import("./views/OpenFromUrlView.vue"),
  },
];

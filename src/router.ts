import type { RouteRecordRaw } from "vue-router";

// Route table for the app. SPEC-001 adds the shell at `/` and the
// document workspace at `/doc/:id`; later specs extend this list.
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
  },
  {
    path: "/open",
    name: "open-from-url",
    component: () => import("./views/OpenFromUrlView.vue"),
  },
];

import { createApp } from "vue";
import { createRouter, createWebHistory } from "vue-router";
import App from "./App.vue";
import { routes } from "./router";
import { initLibrary } from "./lib/documentStore";

const router = createRouter({ history: createWebHistory(), routes });

// Hydrate the document library before first render so views can read it
// synchronously. initLibrary never rejects — when storage is unavailable the
// app still mounts, in-memory, with the warning banner shown.
initLibrary().then(() => {
  createApp(App).use(router).mount("#app");
});

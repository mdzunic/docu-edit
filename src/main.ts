import { createApp } from "vue";
import { createRouter, createWebHistory } from "vue-router";
import App from "./App.vue";
import { routes } from "./router";
import { initLibrary } from "./lib/documentStore";
import { initTemplates } from "./lib/templateStore";

const router = createRouter({ history: createWebHistory(), routes });

// Hydrate the document library and the user template gallery before first
// render so views can read them synchronously. Neither ever rejects — when
// storage is unavailable the app still mounts, in-memory.
Promise.all([initLibrary(), initTemplates()]).then(() => {
  createApp(App).use(router).mount("#app");
});

import { createApp } from "vue";
import { createRouter, createWebHistory } from "vue-router";
import App from "./App.vue";
import { routes } from "./router";
import { initLibrary } from "./lib/documentStore";
import { initTemplates } from "./lib/templateStore";
import { installProviderGuard } from "./lib/ai/providerGuard";

const router = createRouter({ history: createWebHistory(), routes });

// SPEC-113: an API key never outlives the origin it was entered for, and a
// standing model consent never outlives its origin or model — whichever path
// changes the stored provider config. Loads no network code.
installProviderGuard();

// Hydrate the document library and the user template gallery before first
// render so views can read them synchronously. Neither ever rejects — when
// storage is unavailable the app still mounts, in-memory.
Promise.all([initLibrary(), initTemplates()]).then(() => {
  createApp(App).use(router).mount("#app");
});

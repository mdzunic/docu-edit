// defineConfig comes from vitest/config (not vite) so the `test` block below
// is type-checked rather than rejected as an unknown property.
import { defineConfig } from "vitest/config";
import vue from "@vitejs/plugin-vue";
import { fileURLToPath, URL } from "node:url";

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: { "@": fileURLToPath(new URL("./src", import.meta.url)) },
  },
  server: {
    port: 5173,
    // The factory runs the dev server inside a container; binding all
    // interfaces keeps it reachable if the harness ever probes from outside
    // the network namespace.
    host: true,
  },
  test: {
    environment: "jsdom",
    include: ["src/**/*.spec.ts"],
  },
});

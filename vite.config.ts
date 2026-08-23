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
    // The DocumentView specs mount the whole editor, and the first test in each
    // file pays for transforming its entire dependency graph: ~3-4 s on an idle
    // machine, and past Vitest's 5 s default under the load of the full suite
    // running in the factory's container (eight files timed out together on
    // 2026-08-22 and took 68 tests down with them). A timeout is a hang
    // detector, not a performance budget — give it room.
    testTimeout: 30_000,
  },
});

import pluginVue from "eslint-plugin-vue";

// ESLint handles .vue templates only; OXLint covers .ts (see package.json).
export default [
  ...pluginVue.configs["flat/recommended"],
  {
    files: ["src/**/*.vue"],
    rules: {
      "vue/multi-word-component-names": "off",
    },
  },
];

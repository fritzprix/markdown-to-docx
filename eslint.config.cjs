const js = require("@eslint/js");
const ts = require("typescript-eslint");
const globals = require("globals");

module.exports = [
  {
    ignores: ["dist/", "node_modules/", "build/", "*.docx"],
  },
  {
    files: ["src/**/*.ts", "examples/**/*.ts"],
    languageOptions: {
      parser: ts.parser,
      parserOptions: {
        ecmaVersion: 2020,
        sourceType: "module",
      },
      globals: globals.node,
    },
    plugins: {
      "@typescript-eslint": ts.plugin,
    },
    rules: {
      ...js.configs.recommended.rules,
      ...ts.configs.recommended.rules,
      "@typescript-eslint/no-explicit-any": "warn",
      "@typescript-eslint/explicit-function-return-types": "off",
      "no-console": "off",
    },
  },
];

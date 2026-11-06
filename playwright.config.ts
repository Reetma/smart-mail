import { defineConfig } from "@playwright/test";

// Smart Mail (extension/) contre les fausses pages Outlook des deux extensions
// d'origine, servies par page.route sous https://outlook.office.com.
//   node node_modules/@playwright/test/cli.js test
export default defineConfig({
  testDir: "tests",
  testMatch: "**/*.spec.ts",
  workers: 1,
  fullyParallel: false,
  timeout: 180000,
  reporter: [["list"]],
  use: { screenshot: "only-on-failure" },
});

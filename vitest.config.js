import { defineConfig } from "vitest/config";

const SETUP = "./tests/helpers/setup.js";

export default defineConfig({
    test: {
        maxWorkers: "50%",
        projects: [
            { test: { name: "unit", include: ["tests/e2e/unit/**/*.test.js"], setupFiles: [SETUP] } },
            {
                test: {
                    name: "workbooks",
                    include: ["tests/e2e/workbooks/**/*.test.js", "tests/e2e/ai/**/*.test.js"],
                    setupFiles: [SETUP],
                },
            },
            { test: { name: "cli", include: ["tests/cli/**/*.test.js"] } },
            { test: { name: "performance", include: ["tests/performance/**/*.test.js"], setupFiles: [SETUP] } },
        ],
    },
});

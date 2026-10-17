#!/usr/bin/env node
import fs from "node:fs/promises";

import { runXlsxContextCli } from "./lib/cli-runner.mjs";

const PACKAGE_JSON_URL = new URL("../package.json", import.meta.url);

async function readPackageVersion() {
    const packageJson = JSON.parse(await fs.readFile(PACKAGE_JSON_URL, "utf8"));
    return packageJson.version || "0.0.0";
}

runXlsxContextCli({
    argv: process.argv.slice(2),
    readPackageVersion,
})
    .then((exitCode) => {
        process.exit(exitCode);
    })
    .catch((error) => {
        console.error(error instanceof Error ? error.message : String(error));
        process.exit(1);
    });

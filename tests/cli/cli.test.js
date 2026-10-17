// @vitest-environment node

import { mkdtempSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

import { describe, expect, it } from "vitest";
import { BIN, ROOT, fixture } from "../helpers/paths.js";

const execFileAsync = promisify(execFile);

describe("xlsx-context cli", () => {
    it("finishes a small fixture conversion within a reasonable time budget", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("edge", "edge-empty-sample01.xlsx");
        const outputPath = path.join(workspace, "edge-empty.md");

        try {
            const startedAt = performance.now();
            await execFileAsync(process.execPath, [BIN, fixturePath, "--out", outputPath], {
                cwd: ROOT,
            });
            const elapsedMs = performance.now() - startedAt;

            expect(readFileSync(outputPath, "utf8")).toContain("# ");
            expect(elapsedMs).toBeLessThan(5000);
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    }, 10000);

    it("writes combined markdown from a workbook", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("table", "table-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "result.md");

        try {
            await execFileAsync(process.execPath, [BIN, fixturePath, "--out", outputPath], {
                cwd: ROOT,
            });

            const outputText = readFileSync(outputPath, "utf8");
            expect(outputText).toContain("# ");
            expect(outputText).toContain("## ");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("writes a zip export from a workbook", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("table", "table-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "result.zip");

        try {
            await execFileAsync(process.execPath, [BIN, fixturePath, "--zip", outputPath], {
                cwd: ROOT,
            });

            const outputBytes = readFileSync(outputPath);
            expect(outputBytes.length).toBeGreaterThan(0);
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("reports the input workbook when the command fails", async () => {
        await expect(
            execFileAsync(process.execPath, [BIN, fixture("does-not-exist.xlsx")], {
                cwd: ROOT,
            })
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("[does-not-exist.xlsx] read failed:"),
        });
    });

    it("prints help and exits successfully", async () => {
        const result = await execFileAsync(process.execPath, [BIN, "--help"], {
            cwd: ROOT,
        });

        expect(result.stdout).toContain("Usage:");
        expect(result.stdout).toContain("--shape-details");
        expect(result.stdout).toContain("--include-shape-details");
        expect(result.stdout).toContain("--encoding");
        expect(result.stdout).toContain("--bom");
        expect(result.stdout).toContain("--formatting-mode");
        expect(result.stdout).toContain("--table-detection-mode");
        expect(result.stdout).toContain("--front-matter");
        expect(result.stdout).toContain("GUI-aligned defaults:");
        expect(result.stdout).toContain("formatting-mode=gfm");
        expect(result.stdout).toContain("shape-details=exclude");
        expect(result.stdout).toContain("front-matter=include");
        expect(result.stdout).toContain("Output contract for agents:");
        expect(result.stdout).toContain("YAML front matter");
        expect(result.stdout).toContain("Front matter fields:");
        expect(result.stdout).toContain("title, type, conversion");
        expect(result.stdout).toContain("Conversion fields:");
        expect(result.stdout).toContain("shape_details");
        expect(result.stdout).toContain("Exit codes:");
    });

    it("prints version and exits successfully", async () => {
        const result = await execFileAsync(process.execPath, [BIN, "--version"], {
            cwd: ROOT,
        });

        expect(result.stdout.trim()).toBe("1.3.0");
        expect(result.stderr).toBe("");
    });

    it("fails for an unknown option", async () => {
        await expect(
            execFileAsync(process.execPath, [BIN, fixture("table", "table-basic-sample01.xlsx"), "--unknown-option"], {
                cwd: ROOT,
            })
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Unknown option: --unknown-option"),
        });
    });

    it("prints a summary when requested", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));

        try {
            const result = await execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--summary"],
                {
                    cwd: workspace,
                }
            );

            expect(result.stdout).toContain("[workbook] table-basic-sample01.xlsx");
            expect(result.stdout).toContain("Output file:");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("accepts the legacy github formatting mode as gfm", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));

        try {
            const result = await execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--summary", "--formatting-mode", "github"],
                {
                    cwd: workspace,
                }
            );

            expect(result.stdout).toContain("Formatting mode: gfm");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("accepts output mode and shape detail options", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("shape", "shape-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "shape.md");

        try {
            await execFileAsync(
                process.execPath,
                [
                    BIN,
                    fixturePath,
                    "--out",
                    outputPath,
                    "--output-mode",
                    "both",
                    "--formatting-mode",
                    "gfm",
                    "--table-detection-mode",
                    "border",
                    "--shape-details",
                    "include",
                ],
                {
                    cwd: ROOT,
                }
            );

            const outputText = readFileSync(outputPath, "utf8");
            expect(outputText).toContain("# ");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("shows shape text and a Mermaid diagram by default and adds raw details on request", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("shape", "shape-flowchart-sample01.xlsx");
        const defaultPath = path.join(workspace, "default.md");
        const detailPath = path.join(workspace, "detail.md");

        try {
            await execFileAsync(process.execPath, [BIN, fixturePath, "--out", defaultPath], { cwd: ROOT });
            await execFileAsync(
                process.execPath,
                [BIN, fixturePath, "--out", detailPath, "--shape-details", "include"],
                {
                    cwd: ROOT,
                }
            );

            const defaultText = readFileSync(defaultPath, "utf8");
            const detailText = readFileSync(detailPath, "utf8");
            expect(defaultText).toContain("shape_details: exclude");
            expect(defaultText).toContain("  - Shape 001 (Shape (flowChartTerminator)): 開始");
            expect(defaultText).toContain("```mermaid\nflowchart LR\n");
            expect(defaultText).not.toContain("#### Shape:");
            expect(detailText).toContain("```mermaid\nflowchart LR\n");
            expect(detailText).toContain("#### Shape: 001 (H3)");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("omits front matter when requested", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("table", "table-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "front-matter-exclude.md");

        try {
            await execFileAsync(
                process.execPath,
                [BIN, fixturePath, "--out", outputPath, "--front-matter", "exclude"],
                {
                    cwd: ROOT,
                }
            );

            const outputText = readFileSync(outputPath, "utf8");
            expect(outputText.startsWith("---\n")).toBe(false);
            expect(outputText.startsWith("# Book: table-basic-sample01.xlsx")).toBe(true);
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("keeps border-priority as a compatibility alias", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("table", "table-border-priority-sample01.xlsx");
        const outputPath = path.join(workspace, "border-priority.md");

        try {
            await execFileAsync(
                process.execPath,
                [BIN, fixturePath, "--out", outputPath, "--table-detection-mode", "border-priority"],
                {
                    cwd: ROOT,
                }
            );

            const outputText = readFileSync(outputPath, "utf8");
            expect(outputText).toContain("# ");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("writes UTF-16BE markdown with BOM when requested", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("table", "table-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "utf16be.md");

        try {
            await execFileAsync(
                process.execPath,
                [BIN, fixturePath, "--out", outputPath, "--encoding", "utf-16be", "--bom", "on"],
                {
                    cwd: ROOT,
                }
            );

            const outputBytes = readFileSync(outputPath);
            expect(Array.from(outputBytes.slice(0, 2))).toEqual([0xfe, 0xff]);
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("writes Shift_JIS markdown when requested", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("table", "table-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "sjis.md");

        try {
            await execFileAsync(process.execPath, [BIN, fixturePath, "--out", outputPath, "--encoding", "shift_jis"], {
                cwd: ROOT,
            });

            const outputBytes = readFileSync(outputPath);
            const outputText = new TextDecoder("shift_jis").decode(outputBytes);
            expect(outputText).toContain("# ");
            expect(outputText).toContain("## ");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    });

    it("keeps --include-shape-details as a compatibility alias", async () => {
        const workspace = mkdtempSync(path.join(tmpdir(), "xlsx-context-cli-"));
        const fixturePath = fixture("shape", "shape-basic-sample01.xlsx");
        const outputPath = path.join(workspace, "shape-alias.md");

        try {
            await execFileAsync(process.execPath, [BIN, fixturePath, "--out", outputPath, "--include-shape-details"], {
                cwd: ROOT,
            });

            const outputText = readFileSync(outputPath, "utf8");
            expect(outputText).toContain("# ");
        } finally {
            rmSync(workspace, { recursive: true, force: true });
        }
    }, 15000);

    it("fails for an invalid output mode", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--output-mode", "invalid"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid output mode: invalid"),
        });
    });

    it("fails for an invalid formatting mode", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--formatting-mode", "invalid"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid formatting mode: invalid"),
        });
    });

    it("fails for an invalid shape details mode", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--shape-details", "invalid"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid shape details mode: invalid"),
        });
    });

    it("fails for an invalid front matter mode", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--front-matter", "invalid"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid front matter mode: invalid"),
        });
    });

    it("fails for an invalid table detection mode", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--table-detection-mode", "invalid"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid table detection mode: invalid"),
        });
    });

    it("fails for an invalid encoding", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--encoding", "invalid"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid encoding: invalid"),
        });
    });

    it("fails for an invalid BOM mode", async () => {
        await expect(
            execFileAsync(process.execPath, [BIN, fixture("table", "table-basic-sample01.xlsx"), "--bom", "invalid"], {
                cwd: ROOT,
            })
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Invalid BOM mode: invalid"),
        });
    });

    it("fails when a value option is missing its value", async () => {
        await expect(
            execFileAsync(process.execPath, [BIN, fixture("table", "table-basic-sample01.xlsx"), "--encoding"], {
                cwd: ROOT,
            })
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("Missing value for --encoding"),
        });
    });

    it("fails when BOM is enabled for shift_jis", async () => {
        await expect(
            execFileAsync(
                process.execPath,
                [BIN, fixture("table", "table-basic-sample01.xlsx"), "--encoding", "shift_jis", "--bom", "on"],
                {
                    cwd: ROOT,
                }
            )
        ).rejects.toMatchObject({
            stderr: expect.stringContaining("BOM cannot be enabled for shift_jis."),
        });
    });
});

import { describe, expect, it } from "vitest";
import { toMarkdownLinkTarget } from "../../../../../src/structure/formatting/link-target.ts";

describe("markdown link target", () => {
    it("keeps ordinary targets", () => {
        expect(toMarkdownLinkTarget("https://example.com/a?b=1#c")).toBe("https://example.com/a?b=1#c");
        expect(toMarkdownLinkTarget("mailto:a@example.com")).toBe("mailto:a@example.com");
        expect(toMarkdownLinkTarget("..\\docs\\report.xlsx")).toBe("..\\docs\\report.xlsx");
    });

    it("encodes characters that would close or break the destination", () => {
        expect(toMarkdownLinkTarget("https://x.test/a b)(c<d>")).toBe("https://x.test/a%20b%29%28c%3Cd%3E");
        expect(toMarkdownLinkTarget("https://x.test/)[evil](javascript:alert(1)")).toBe(
            "https://x.test/%29[evil]%28javascript:alert%281%29"
        );
    });

    it("rejects empty targets and schemes that run code", () => {
        expect(toMarkdownLinkTarget("  ")).toBeNull();
        expect(toMarkdownLinkTarget("javascript:alert(1)")).toBeNull();
        expect(toMarkdownLinkTarget(" JavaScript:alert(1)")).toBeNull();
        expect(toMarkdownLinkTarget("java\tscript:alert(1)")).toBeNull();
        expect(toMarkdownLinkTarget("vbscript:msgbox")).toBeNull();
        expect(toMarkdownLinkTarget("data:text/html,<script>")).toBeNull();
    });
});

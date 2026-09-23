import { describe, it, expect } from "vitest";
import {
  escapeHtml,
  buildEventBody,
} from "../public/outlook-addin/eventbody.js";

const source = (over = {}) => ({
  restId: "AAMkADhMGAAA=",
  subject: "After Hours in Chicago",
  webLink: "https://outlook.office365.com/owa/?ItemID=AAMkADMGAAA%3D&exvsurl=1",
  ...over,
});

describe("escapeHtml", () => {
  it("escapes every HTML-significant character", () => {
    expect(escapeHtml(`&<>"'`)).toBe("&amp;&lt;&gt;&quot;&#39;");
  });

  it("escapes ampersands before the entities it introduces", () => {
    expect(escapeHtml("a & <b>")).toBe("a &amp; &lt;b&gt;");
  });

  it("returns empty for non-strings", () => {
    expect(escapeHtml(null)).toBe("");
    expect(escapeHtml(undefined)).toBe("");
  });
});

describe("buildEventBody without a link", () => {
  it("returns the plain-text body unchanged when there is no source", () => {
    expect(buildEventBody("Annual offsite.", null)).toEqual({
      contentType: "Text",
      content: "Annual offsite.",
    });
  });

  it("returns plain text when the webLink could not be resolved", () => {
    expect(buildEventBody("Annual offsite.", source({ webLink: null }))).toEqual({
      contentType: "Text",
      content: "Annual offsite.",
    });
  });

  it("does not escape the plain-text body", () => {
    const body = buildEventBody("R&D <all> sync", null);
    expect(body.content).toBe("R&D <all> sync");
  });

  it("handles a missing description", () => {
    expect(buildEventBody(null, null)).toEqual({
      contentType: "Text",
      content: "",
    });
  });
});

describe("buildEventBody with a link", () => {
  it("switches to HTML and embeds the link", () => {
    const body = buildEventBody("Career reception.", source());
    expect(body.contentType).toBe("HTML");
    expect(body.content).toContain("<p>Career reception.</p>");
    expect(body.content).toContain("After Hours in Chicago</a>");
  });

  it("keeps the URL intact in the href", () => {
    const body = buildEventBody("x", source());
    expect(body.content).toContain(
      'href="https://outlook.office365.com/owa/?ItemID=AAMkADMGAAA%3D&amp;exvsurl=1"',
    );
  });

  it("escapes a description containing markup", () => {
    const body = buildEventBody("R&D <all> sync", source());
    expect(body.content).toContain("R&amp;D &lt;all&gt; sync");
    expect(body.content).not.toContain("<all>");
  });

  it("escapes a subject containing markup", () => {
    const body = buildEventBody("x", source({ subject: '<img src=x onerror=1>' }));
    expect(body.content).toContain("&lt;img src=x onerror=1&gt;");
    expect(body.content).not.toContain("<img");
  });

  it("falls back to a generic label when the subject is empty", () => {
    const body = buildEventBody("x", source({ subject: "   " }));
    expect(body.content).toContain(">Open the email</a>");
  });

  it("still produces a valid body when the description is empty", () => {
    const body = buildEventBody("", source());
    expect(body.contentType).toBe("HTML");
    expect(body.content).toContain("After Hours in Chicago</a>");
    expect(body.content).not.toContain("<p></p>");
  });

  it("refuses a javascript: URL and falls back to plain text", () => {
    const body = buildEventBody("x", source({ webLink: "javascript:alert(1)" }));
    expect(body).toEqual({ contentType: "Text", content: "x" });
  });

  it("refuses a data: URL and falls back to plain text", () => {
    const body = buildEventBody("x", source({ webLink: "data:text/html,<p>" }));
    expect(body.contentType).toBe("Text");
  });
});

import { describe, it, expect, vi } from "vitest";
import handler from "../api/extract-event.js";
import { ConfigError } from "../lib/errors.js";

vi.mock("../lib/router.js", async () => {
  const actual =
    await vi.importActual<typeof import("../lib/router.js")>(
      "../lib/router.js",
    );
  return {
    ...actual,
    extractEvents: vi.fn(),
  };
});

import { extractEvents } from "../lib/router.js";

type ApiResponse = { error?: string; events?: unknown[] };

function makeRes() {
  const res: {
    statusCode: number;
    headers: Record<string, string>;
    body: ApiResponse | null;
    setHeader: (k: string, v: string) => typeof res;
    status: (code: number) => typeof res;
    json: (b: ApiResponse) => typeof res;
  } = {
    statusCode: 200,
    headers: {},
    body: null,
    setHeader(k, v) {
      this.headers[k] = v;
      return this;
    },
    status(code) {
      this.statusCode = code;
      return this;
    },
    json(b) {
      this.body = b;
      return this;
    },
  };
  return res;
}

describe("POST /api/extract-event — input validation (12 cases)", () => {
  it("T1 GET -> 405 Method Not Allowed with Allow: POST header", async () => {
    const req = { method: "GET", body: {} } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(405);
    expect(res.headers.Allow).toBe("POST");
    expect(res.body).toEqual({ error: "Method Not Allowed" });
  });

  it("T2 PUT -> 405", async () => {
    const req = { method: "PUT", body: {} } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(405);
  });

  it("T3 empty body -> 400 about email", async () => {
    const req = { method: "POST", body: {} } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
    expect(res.body?.error).toMatch(/"email" string field/);
  });

  it("T4 missing model -> 400 about model", async () => {
    const req = { method: "POST", body: { email: "hi" } } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
    expect(res.body?.error).toMatch(/Supported values: gpt-4o-mini, MiniMax-M3/);
  });

  it("T5 unknown model -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: "hi", model: "claude-3-opus" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
    expect(res.body?.error).toMatch(/Supported values/);
  });

  it("T6 model is a number -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: "hi", model: 42 },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
  });

  it("T7 model is null -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: "hi", model: null },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
  });

  it("T8 empty email string -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: "", model: "MiniMax-M3" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
    expect(res.body?.error).toMatch(/"email" string field/);
  });

  it("T9 whitespace-only email -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: "   ", model: "MiniMax-M3" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
  });

  it("T10 email is a number -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: 42, model: "MiniMax-M3" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
  });

  it("T11 email is null -> 400", async () => {
    const req = {
      method: "POST",
      body: { email: null, model: "MiniMax-M3" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
  });

  it("T12 non-object body -> 400 about email", async () => {
    const req = { method: "POST", body: "not-json" } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(400);
  });
});

describe("POST /api/extract-event — error wrapping (Fix 6)", () => {
  it("ConfigError message is passed through verbatim", async () => {
    vi.mocked(extractEvents).mockRejectedValueOnce(
      new ConfigError("MINIMAX_API_KEY is not configured"),
    );
    const req = {
      method: "POST",
      body: { email: "x", model: "MiniMax-M3" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(500);
    expect(res.body?.error).toBe("MINIMAX_API_KEY is not configured");
  });

  it("non-ConfigError is hidden behind a generic message and logged", async () => {
    const consoleSpy = vi
      .spyOn(console, "error")
      .mockImplementation(() => {});
    vi.mocked(extractEvents).mockRejectedValueOnce(
      new Error(
        "Upstream leaked: api.openai.com returned 401 with internal request id abc-def-ghi",
      ),
    );
    const req = {
      method: "POST",
      body: { email: "x", model: "gpt-4o-mini" },
    } as never;
    const res = makeRes();
    await handler(req, res as never);
    expect(res.statusCode).toBe(500);
    expect(res.body?.error).toBe("Extraction failed");
    expect(res.body?.error).not.toMatch(/abc-def-ghi/);
    expect(consoleSpy).toHaveBeenCalledOnce();
    consoleSpy.mockRestore();
  });
});

import { describe, it, expect, vi } from "vitest";

vi.mock("openai", () => {
  return {
    default: vi.fn().mockImplementation(() => ({})),
  };
});

describe("lib/openai.ts — config check", () => {
  it("throws ConfigError when OPENAI_API_KEY is missing", async () => {
    const prev = process.env.OPENAI_API_KEY;
    delete process.env.OPENAI_API_KEY;
    vi.resetModules();
    const { extractEventsWithOpenAI } = await import("../lib/openai.js");
    await expect(
      extractEventsWithOpenAI("hi", "gpt-4o-mini"),
    ).rejects.toThrow(/OPENAI_API_KEY is not configured/);
    if (prev !== undefined) process.env.OPENAI_API_KEY = prev;
    vi.resetModules();
  });
});

describe("lib/minimax.ts — config check", () => {
  it("throws ConfigError when MINIMAX_API_KEY is missing", async () => {
    const prev = process.env.MINIMAX_API_KEY;
    delete process.env.MINIMAX_API_KEY;
    vi.resetModules();
    const { extractEventsWithMiniMax } = await import("../lib/minimax.js");
    await expect(
      extractEventsWithMiniMax("hi", "MiniMax-M3"),
    ).rejects.toThrow(/MINIMAX_API_KEY is not configured/);
    if (prev !== undefined) process.env.MINIMAX_API_KEY = prev;
    vi.resetModules();
  });
});

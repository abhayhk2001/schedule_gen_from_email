import { describe, it, expect, vi, beforeEach } from "vitest";

vi.mock("../lib/openai.js", () => ({
  extractEventsWithOpenAI: vi.fn(),
}));
vi.mock("../lib/minimax.js", () => ({
  extractEventsWithMiniMax: vi.fn(),
}));

import { extractEvents } from "../lib/router.js";
import { extractEventsWithOpenAI } from "../lib/openai.js";
import { extractEventsWithMiniMax } from "../lib/minimax.js";

describe("lib/router.ts — dispatch", () => {
  beforeEach(() => {
    vi.mocked(extractEventsWithOpenAI).mockReset();
    vi.mocked(extractEventsWithMiniMax).mockReset();
  });

  it("routes gpt-4o-mini to the OpenAI provider", async () => {
    vi.mocked(extractEventsWithOpenAI).mockResolvedValueOnce({ events: [] });
    await extractEvents("hi", "gpt-4o-mini");
    expect(extractEventsWithOpenAI).toHaveBeenCalledWith("hi", "gpt-4o-mini");
    expect(extractEventsWithMiniMax).not.toHaveBeenCalled();
  });

  it("routes MiniMax-M3 to the MiniMax provider", async () => {
    vi.mocked(extractEventsWithMiniMax).mockResolvedValueOnce({ events: [] });
    await extractEvents("hi", "MiniMax-M3");
    expect(extractEventsWithMiniMax).toHaveBeenCalledWith("hi", "MiniMax-M3");
    expect(extractEventsWithOpenAI).not.toHaveBeenCalled();
  });
});

import { describe, it, expect, vi } from "vitest";
import OpenAI from "openai";
import { runExtraction } from "../lib/extract.js";

function makeMockClient(content: string | null = "{}") {
  return {
    chat: {
      completions: {
        create: vi.fn().mockResolvedValue({
          choices: [{ message: { content } }],
        }),
      },
    },
  } as unknown as OpenAI;
}

describe("lib/extract.ts — runExtraction", () => {
  it("returns { events: [] } when content is null", async () => {
    const client = makeMockClient(null);
    const result = await runExtraction(client, "m", "hi");
    expect(result).toEqual({ events: [] });
  });

  it("returns { events: [] } when content is empty string", async () => {
    const client = makeMockClient("");
    const result = await runExtraction(client, "m", "hi");
    expect(result).toEqual({ events: [] });
  });

  it("parses valid JSON and returns events array", async () => {
    const client = makeMockClient(
      JSON.stringify({
        events: [
          {
            date: "2025-10-03",
            time: "13:00",
            timezone: "America/Chicago",
            event_name: "Career Fair",
            description: "Bring resumes.",
          },
        ],
      }),
    );
    const result = await runExtraction(client, "m", "hi");
    expect(result.events).toHaveLength(1);
    expect(result.events[0].event_name).toBe("Career Fair");
  });

  it("returns { events: [] } on malformed JSON instead of throwing", async () => {
    const client = makeMockClient("{not valid json");
    const result = await runExtraction(client, "m", "hi");
    expect(result).toEqual({ events: [] });
  });

  it("returns { events: [] } when parsed.events is not an array", async () => {
    const client = makeMockClient(JSON.stringify({ events: "oops" }));
    const result = await runExtraction(client, "m", "hi");
    expect(result).toEqual({ events: [] });
  });

  it("forwards extra_body when provided", async () => {
    const client = makeMockClient();
    await runExtraction(client, "m", "hi", {
      thinking: { type: "disabled" },
    });
    const args = (client.chat.completions.create as ReturnType<typeof vi.fn>).mock
      .calls[0][0] as Record<string, unknown>;
    expect(args.extra_body).toEqual({ thinking: { type: "disabled" } });
  });

  it("does not include extra_body when omitted", async () => {
    const client = makeMockClient();
    await runExtraction(client, "m", "hi");
    const args = (client.chat.completions.create as ReturnType<typeof vi.fn>).mock
      .calls[0][0] as Record<string, unknown>;
    expect(args.extra_body).toBeUndefined();
  });

  it("sends system + user messages with today's ISO date", async () => {
    const client = makeMockClient();
    await runExtraction(client, "m", "the email body");
    const args = (client.chat.completions.create as ReturnType<typeof vi.fn>).mock
      .calls[0][0] as { messages: Array<{ role: string; content: string }> };
    expect(args.messages).toHaveLength(2);
    expect(args.messages[0].role).toBe("system");
    expect(args.messages[1].role).toBe("user");
    expect(args.messages[1].content).toMatch(/Today's date is \d{4}-\d{2}-\d{2}\./);
    expect(args.messages[1].content).toContain("the email body");
  });

  it("uses strict json_schema response_format", async () => {
    const client = makeMockClient();
    await runExtraction(client, "m", "hi");
    const args = (client.chat.completions.create as ReturnType<typeof vi.fn>).mock
      .calls[0][0] as {
      response_format: {
        type: string;
        json_schema: { strict: boolean; schema: unknown };
      };
    };
    expect(args.response_format.type).toBe("json_schema");
    expect(args.response_format.json_schema.strict).toBe(true);
  });
});

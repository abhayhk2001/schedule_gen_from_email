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

  it("parses valid JSON and returns events array (incl. new fields)", async () => {
    const client = makeMockClient(
      JSON.stringify({
        events: [
          {
            date: "2025-10-03",
            time: "13:00",
            timezone: "America/Chicago",
            event_name: "Career Fair",
            location: "Illini Union, 1401 W Green St, Urbana",
            description: "Bring resumes.",
            whole_day: false,
            end_date: null,
            end_time: "17:00",
          },
        ],
      }),
    );
    const result = await runExtraction(client, "m", "hi");
    expect(result.events).toHaveLength(1);
    expect(result.events[0].event_name).toBe("Career Fair");
    expect(result.events[0].location).toBe("Illini Union, 1401 W Green St, Urbana");
    expect(result.events[0].end_time).toBe("17:00");
    expect(result.events[0].whole_day).toBe(false);
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

  it("uses strict json_schema response_format with whole_day, end_date, end_time", async () => {
    const client = makeMockClient();
    await runExtraction(client, "m", "hi");
    const args = (client.chat.completions.create as ReturnType<typeof vi.fn>).mock
      .calls[0][0] as {
      response_format: {
        type: string;
        json_schema: { strict: boolean; schema: { properties: { events: { items: { required: string[] } } } } };
      };
    };
    expect(args.response_format.type).toBe("json_schema");
    expect(args.response_format.json_schema.strict).toBe(true);
    expect(args.response_format.json_schema.schema.properties.events.items.required).toEqual(
      expect.arrayContaining([
        "whole_day",
        "end_date",
        "end_time",
        "date",
        "time",
        "timezone",
        "event_name",
        "location",
        "description",
      ]),
    );
  });

  it("parses whole_day=true with time/end_time=null", async () => {
    const client = makeMockClient(
      JSON.stringify({
        events: [
          {
            date: "2026-10-05",
            time: null,
            timezone: null,
            event_name: "Offsite",
            description: "Annual offsite.",
            whole_day: true,
            end_date: null,
            end_time: null,
          },
        ],
      }),
    );
    const result = await runExtraction(client, "m", "hi");
    expect(result.events[0].whole_day).toBe(true);
    expect(result.events[0].time).toBeNull();
    expect(result.events[0].end_time).toBeNull();
  });

  it("parses location='Virtual' verbatim for online-only events", async () => {
    const client = makeMockClient(
      JSON.stringify({
        events: [
          {
            date: "2026-11-02",
            time: "09:00",
            timezone: "America/New_York",
            event_name: "Planning sync",
            location: "Virtual",
            description: "Weekly planning over Teams.",
            whole_day: false,
            end_date: null,
            end_time: "09:30",
          },
        ],
      }),
    );
    const result = await runExtraction(client, "m", "hi");
    expect(result.events[0].location).toBe("Virtual");
  });
});

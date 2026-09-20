export interface Event {
  date: string | null;
  time: string | null;
  timezone: string | null;
  event_name: string | null;
  description: string | null;
  whole_day: boolean;
  end_date: string | null;
  end_time: string | null;
}

export const SUPPORTED_MODELS = ["gpt-4o-mini", "MiniMax-M3"] as const;
export type SupportedModel = (typeof SUPPORTED_MODELS)[number];

export interface ExtractResponse {
  events: Event[];
}

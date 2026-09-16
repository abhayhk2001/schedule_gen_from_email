export interface Event {
  date: string | null;
  time: string | null;
  timezone: string | null;
  event_name: string | null;
  description: string | null;
}

export interface ExtractRequest {
  email: string;
}

export interface ExtractResponse {
  events: Event[];
}

export interface ErrorResponse {
  error: string;
}

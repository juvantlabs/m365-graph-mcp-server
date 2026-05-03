import type { Client } from "@microsoft/microsoft-graph-client";
import { describe, expect, it, vi } from "vitest";

import { listCalendarsTool, summarizeCalendar } from "../../src/tools/list_calendars.js";

describe("summarizeCalendar", () => {
  it("extracts core fields", () => {
    const cal = {
      id: "cal-1",
      name: "Calendar",
      color: "lightBlue",
      owner: { name: "Alice", address: "alice@example.com" },
      isDefaultCalendar: true,
      canEdit: true,
      canShare: false,
    };
    expect(summarizeCalendar(cal)).toEqual({
      id: "cal-1",
      name: "Calendar",
      color: "lightBlue",
      owner_name: "Alice",
      owner_email: "alice@example.com",
      is_default: true,
      can_edit: true,
      can_share: false,
    });
  });

  it("defaults missing color to 'auto'", () => {
    expect(summarizeCalendar({ id: "x", name: "x" }).color).toBe("auto");
  });

  it("returns null owner fields when no owner", () => {
    const s = summarizeCalendar({ id: "x", name: "x" });
    expect(s.owner_name).toBeNull();
    expect(s.owner_email).toBeNull();
  });

  it("coerces missing booleans to false", () => {
    const s = summarizeCalendar({ id: "x", name: "x" });
    expect(s.is_default).toBe(false);
    expect(s.can_edit).toBe(false);
    expect(s.can_share).toBe(false);
  });
});

function mockClient(response: unknown): Client {
  const top = vi.fn().mockReturnValue({ get: vi.fn().mockResolvedValue(response) });
  const api = vi.fn().mockReturnValue({ top });
  return { api } as unknown as Client;
}

describe("listCalendarsTool handler", () => {
  it("returns count + calendars in response", async () => {
    const client = mockClient({
      value: [
        { id: "c1", name: "Personal", isDefaultCalendar: true },
        { id: "c2", name: "Work" },
      ],
    });
    const resp = await listCalendarsTool.handler(client, {});
    const parsed = JSON.parse((resp.content[0] as { type: string; text: string }).text);
    expect(parsed.count).toBe(2);
    expect(parsed.calendars[0].is_default).toBe(true);
  });

  it("validates limit out of range", async () => {
    const client = mockClient({ value: [] });
    await expect(listCalendarsTool.handler(client, { limit: 999 })).rejects.toThrow();
  });
});

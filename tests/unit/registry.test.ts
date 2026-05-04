import { describe, expect, it } from "vitest";

import { ALL_TOOLS, buildHandlerMap } from "../../src/tools/index.js";

describe("ALL_TOOLS registry", () => {
  it("includes all read + write tools (files + calendars)", () => {
    const names = ALL_TOOLS.map((t) => t.definition.name);
    // Files — read
    expect(names).toContain("m365-graph:list_drives");
    expect(names).toContain("m365-graph:list_items");
    expect(names).toContain("m365-graph:search_files");
    expect(names).toContain("m365-graph:download_file");
    // Files — write
    expect(names).toContain("m365-graph:upload_file");
    expect(names).toContain("m365-graph:copy_file");
    expect(names).toContain("m365-graph:move_file");
    expect(names).toContain("m365-graph:delete_file");
    // Calendars — read
    expect(names).toContain("m365-graph:list_calendars");
    expect(names).toContain("m365-graph:list_events");
    expect(names).toContain("m365-graph:search_events");
    expect(names).toContain("m365-graph:get_event");
    // Calendars — write
    expect(names).toContain("m365-graph:create_event");
    expect(names).toContain("m365-graph:update_event");
    expect(names).toContain("m365-graph:cancel_event");
  });

  it("each tool has a non-empty name + description", () => {
    for (const t of ALL_TOOLS) {
      expect(t.definition.name).toMatch(/^m365-graph:/);
      expect(t.definition.description.length).toBeGreaterThan(20);
    }
  });

  it("each tool has an inputSchema with type=object", () => {
    for (const t of ALL_TOOLS) {
      expect(t.definition.inputSchema.type).toBe("object");
      expect(t.definition.inputSchema.required).toBeDefined();
    }
  });
});

describe("buildHandlerMap", () => {
  it("returns a Map with one entry per tool", () => {
    const map = buildHandlerMap(ALL_TOOLS);
    expect(map.size).toBe(ALL_TOOLS.length);
  });

  it("handlers are invokable from the map by tool name", () => {
    const map = buildHandlerMap(ALL_TOOLS);
    for (const t of ALL_TOOLS) {
      expect(typeof map.get(t.definition.name)).toBe("function");
    }
  });

  it("returns undefined for unknown tool names", () => {
    const map = buildHandlerMap(ALL_TOOLS);
    expect(map.get("m365-graph:nonexistent")).toBeUndefined();
  });
});

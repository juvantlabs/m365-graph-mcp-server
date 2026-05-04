import { describe, expect, it } from "vitest";

import { classifyMonitorStatus, summarizeCopiedItem } from "../../src/tools/copy_file.js";

describe("classifyMonitorStatus", () => {
  it("maps 'completed' to completed status", () => {
    const r = classifyMonitorStatus({
      status: "completed",
      resourceLocation: "/items/abc",
      percentageComplete: 100,
    });
    expect(r.status).toBe("completed");
    expect(r.resource_location).toBe("/items/abc");
    expect(r.percent_complete).toBe(100);
  });

  it("maps 'failed' to failed status with error payload", () => {
    const r = classifyMonitorStatus({
      status: "failed",
      error: { code: "rejected", message: "permission denied" },
    });
    expect(r.status).toBe("failed");
    expect(r.error).toEqual({ code: "rejected", message: "permission denied" });
  });

  it("maps everything else to in_progress", () => {
    expect(classifyMonitorStatus({ status: "inProgress", percentageComplete: 50 }).status).toBe(
      "in_progress",
    );
    expect(classifyMonitorStatus({ status: "notStarted" }).status).toBe("in_progress");
    expect(classifyMonitorStatus({}).status).toBe("in_progress");
  });

  it("is case-insensitive on status comparison", () => {
    expect(classifyMonitorStatus({ status: "Completed" }).status).toBe("completed");
    expect(classifyMonitorStatus({ status: "FAILED" }).status).toBe("failed");
  });
});

describe("summarizeCopiedItem", () => {
  it("extracts canonical fields", () => {
    const item = {
      id: "i1",
      name: "report.pdf",
      size: 2048,
      webUrl: "https://example/i1",
      parentReference: { id: "parent-1" },
    };
    expect(summarizeCopiedItem(item)).toEqual({
      id: "i1",
      name: "report.pdf",
      size: 2048,
      webUrl: "https://example/i1",
      parent_id: "parent-1",
    });
  });

  it("returns null parent_id when no parentReference", () => {
    expect(summarizeCopiedItem({ id: "x", name: "x", size: 0, webUrl: "" }).parent_id).toBeNull();
  });
});

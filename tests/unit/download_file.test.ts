import path from "node:path";
import os from "node:os";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, describe, expect, it, vi } from "vitest";

import {
  deriveSafeLocalPath,
  downloadFileTool,
  getSandboxRoot,
} from "../../src/tools/download_file.js";

describe("getSandboxRoot env-var precedence", () => {
  const originalEnv = { ...process.env };

  afterEach(() => {
    process.env = { ...originalEnv };
  });

  it("uses M365_DOWNLOAD_DIR when set", () => {
    process.env.M365_DOWNLOAD_DIR = "/custom/dir";
    expect(getSandboxRoot("tenant-1")).toBe(path.resolve("/custom/dir", "tenant-1"));
  });

  it("falls back to XDG_CACHE_HOME", () => {
    delete process.env.M365_DOWNLOAD_DIR;
    process.env.XDG_CACHE_HOME = "/xdg-cache";
    expect(getSandboxRoot("tenant-1")).toBe(
      path.resolve("/xdg-cache", "m365-graph-mcp-server", "tenant-1"),
    );
  });

  it("falls back to ~/.cache when neither override is set", () => {
    delete process.env.M365_DOWNLOAD_DIR;
    delete process.env.XDG_CACHE_HOME;
    expect(getSandboxRoot("tenant-1")).toBe(
      path.resolve(os.homedir(), ".cache", "m365-graph-mcp-server", "tenant-1"),
    );
  });

  it("scopes path per-tenant (different tenants → different roots)", () => {
    process.env.M365_DOWNLOAD_DIR = "/custom";
    expect(getSandboxRoot("a")).not.toBe(getSandboxRoot("b"));
  });
});

describe("deriveSafeLocalPath", () => {
  const sandbox = "/sandbox/tenant-x";

  it("constructs filename as <hash>-<sanitized name>", () => {
    const localPath = deriveSafeLocalPath(sandbox, "item-id-123", "report.pdf");
    expect(localPath.startsWith(sandbox + "/")).toBe(true);
    // The hash is deterministic for a given item_id, but we don't lock
    // its value — just verify the structural shape.
    expect(localPath).toMatch(/^.+\/[0-9a-f]{16}-report\.pdf$/);
  });

  it("sanitizes path-traversal payloads in filename", () => {
    const localPath = deriveSafeLocalPath(sandbox, "any-id", "../../etc/passwd");
    expect(localPath.startsWith(sandbox + "/")).toBe(true);
    expect(localPath).not.toContain("../");
    expect(localPath).not.toContain("/etc/passwd");
  });

  it("never escapes the sandbox root even with adversarial filenames", () => {
    for (const name of [
      "../malicious",
      "../../../../etc/passwd",
      "/absolute/path",
      "name\\with\\backslash",
    ]) {
      const localPath = deriveSafeLocalPath(sandbox, "id", name);
      expect(localPath.startsWith(sandbox + "/")).toBe(true);
    }
  });

  it("produces a different filename for the same name but different item_ids", () => {
    const a = deriveSafeLocalPath(sandbox, "id-a", "shared.pdf");
    const b = deriveSafeLocalPath(sandbox, "id-b", "shared.pdf");
    expect(a).not.toBe(b);
  });
});

describe("downloadFileTool handler — pre-flight error paths", () => {
  function mockClientWithItem(item: unknown): Client {
    const get = vi.fn().mockResolvedValue(item);
    const api = vi.fn().mockReturnValue({ get });
    return { api } as unknown as Client;
  }

  it("requires item_id", async () => {
    const client = mockClientWithItem({});
    await expect(downloadFileTool.handler(client, {})).rejects.toThrow(
      "'item_id' must be a non-empty string",
    );
  });

  it("rejects folders (no content to download)", async () => {
    const folderItem = { id: "f1", name: "Documents", folder: { childCount: 5 }, size: 0 };
    const client = mockClientWithItem(folderItem);
    await expect(
      downloadFileTool.handler(client, { item_id: "f1" }),
    ).rejects.toThrow("is a folder");
  });

  it("rejects files exceeding the 200 MB cap", async () => {
    const bigFile = { id: "big", name: "huge.bin", size: 250 * 1024 * 1024 };
    const client = mockClientWithItem(bigFile);
    await expect(
      downloadFileTool.handler(client, { item_id: "big" }),
    ).rejects.toThrow("exceeds the 200 MB cap");
  });

  it("rejects items missing @microsoft.graph.downloadUrl", async () => {
    const item = { id: "x", name: "x.txt", size: 100 };
    const client = mockClientWithItem(item);
    await expect(
      downloadFileTool.handler(client, { item_id: "x" }),
    ).rejects.toThrow("@microsoft.graph.downloadUrl");
  });
});

import path from "node:path";
import os from "node:os";
import { mkdtemp, readdir, readFile, rm, stat } from "node:fs/promises";
import { PassThrough, Readable } from "node:stream";
import { pipeline } from "node:stream/promises";

import type { Client } from "@microsoft/microsoft-graph-client";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import {
  buildConvertedContentPath,
  cappedCounter,
  checkConvertible,
  convertedFilename,
  deriveSafeLocalPath,
  DOWNLOAD_FORMATS,
  downloadFileTool,
  fileExtension,
  getSandboxRoot,
  PDF_CONVERTIBLE_EXTENSIONS,
} from "../../src/tools/download_file.js";
import { validateHttpsUrl, validateOptionalStrictEnum } from "../../src/types/validators.js";

const validateDownloadFormat = (v: unknown): string | undefined =>
  validateOptionalStrictEnum(v, "format", DOWNLOAD_FORMATS);

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

// ----------------------------------------------------------------------
// format: "pdf" — server-side conversion via /content?format=pdf
// ----------------------------------------------------------------------

describe("download_file format parameter — schema", () => {
  it("exposes `format` as an optional strict enum ['pdf']", () => {
    const schema = downloadFileTool.definition.inputSchema;
    const format = schema.properties.format as { type: string; enum: string[] };
    expect(format.type).toBe("string");
    expect(format.enum).toEqual(["pdf"]);
    expect(schema.required).toEqual(["item_id"]);
    expect([...DOWNLOAD_FORMATS]).toEqual(["pdf"]);
  });
});

describe("validateDownloadFormat", () => {
  it("returns undefined when absent or null (plain download)", () => {
    expect(validateDownloadFormat(undefined)).toBeUndefined();
    expect(validateDownloadFormat(null)).toBeUndefined();
  });

  it("accepts 'pdf'", () => {
    expect(validateDownloadFormat("pdf")).toBe("pdf");
  });

  it.each([["PDF"], ["docx"], ["pdf&foo=bar"], [""], [1], [true], [["pdf"]]])(
    "rejects %j (never passed through)",
    (value) => {
      expect(() => validateDownloadFormat(value)).toThrow("'format' must be one of [\"pdf\"]");
    },
  );
});

describe("buildConvertedContentPath", () => {
  it("builds the /content?format=pdf path for a SharePoint drive", () => {
    expect(buildConvertedContentPath("/drives/b!abc", "01ITEM", "pdf")).toBe(
      "/drives/b!abc/items/01ITEM/content?format=pdf",
    );
  });

  it("builds the path for the primary OneDrive", () => {
    expect(buildConvertedContentPath("/me/drive", "01ITEM", "pdf")).toBe(
      "/me/drive/items/01ITEM/content?format=pdf",
    );
  });

  it("URL-encodes the item id", () => {
    expect(buildConvertedContentPath("/me/drive", "a/b?c", "pdf")).toBe(
      "/me/drive/items/a%2Fb%3Fc/content?format=pdf",
    );
  });
});

describe("fileExtension / checkConvertible", () => {
  it("extracts lower-cased extensions", () => {
    expect(fileExtension("Contract.DOCX")).toBe("docx");
    expect(fileExtension("a.b.xlsx")).toBe("xlsx");
    expect(fileExtension("README")).toBeNull();
    expect(fileExtension(".profile")).toBeNull();
    expect(fileExtension("trailing.")).toBeNull();
  });

  it.each([["NDA.docx"], ["deck.PPTX"], ["sheet.xlsx"], ["notes.odt"], ["memo.rtf"], ["scan.tiff"], ["mail.msg"]])(
    "accepts %s",
    (name) => {
      expect(() => checkConvertible(name, "pdf")).not.toThrow();
    },
  );

  it("rejects an unsupported extension and lists the supported ones", () => {
    expect(() => checkConvertible("archive.zip", "pdf")).toThrow(
      /Cannot convert 'archive\.zip' to PDF: the file has extension '\.zip'.*Supported source types: .*docx/,
    );
  });

  it("rejects a name with no extension", () => {
    expect(() => checkConvertible("LICENSE", "pdf")).toThrow("has no file extension");
  });

  it("rejects a file that is already a PDF with a hint to omit format", () => {
    expect(() => checkConvertible("signed.pdf", "pdf")).toThrow(
      "already a PDF — call download_file without 'format'",
    );
  });

  it("has a documented allowlist covering the common office formats", () => {
    for (const ext of ["doc", "docx", "xls", "xlsx", "ppt", "pptx", "odt", "rtf", "html", "md"]) {
      expect(PDF_CONVERTIBLE_EXTENSIONS.has(ext)).toBe(true);
    }
    expect(PDF_CONVERTIBLE_EXTENSIONS.has("pdf")).toBe(false);
    expect(PDF_CONVERTIBLE_EXTENSIONS.has("png")).toBe(false);
  });
});

describe("convertedFilename", () => {
  it("swaps the final extension for .pdf", () => {
    expect(convertedFilename("Contract v2.docx", "pdf")).toBe("Contract v2.pdf");
    expect(convertedFilename("a.b.xlsx", "pdf")).toBe("a.b.pdf");
  });

  it("never lets the sanitizer's 200-char cap cut the extension", () => {
    const long = `${"x".repeat(300)}.docx`;
    const out = convertedFilename(long, "pdf");
    expect(out.length).toBeLessThanOrEqual(200);
    expect(out.endsWith(".pdf")).toBe(true);
    const local = deriveSafeLocalPath("/sandbox/t", "id", out);
    expect(local.endsWith(".pdf")).toBe(true);
  });

  it("falls back to 'file' for an empty stem", () => {
    expect(convertedFilename("", "pdf")).toBe("file.pdf");
  });

  it("yields <hash>-<stem>.pdf under the sandbox, distinct from the original download", () => {
    const sandbox = "/sandbox/tenant-x";
    const original = deriveSafeLocalPath(sandbox, "item-1", "NDA.docx");
    const converted = deriveSafeLocalPath(sandbox, "item-1", convertedFilename("NDA.docx", "pdf"));
    expect(converted).toMatch(/^\/sandbox\/tenant-x\/[0-9a-f]{16}-NDA\.pdf$/);
    expect(converted).not.toBe(original);
    // Same item → same hash prefix (deterministic overwrite policy).
    expect(path.basename(converted).slice(0, 16)).toBe(path.basename(original).slice(0, 16));
  });

  it("keeps adversarial names inside the sandbox", () => {
    const local = deriveSafeLocalPath(
      "/sandbox/t",
      "id",
      convertedFilename("../../etc/passwd.docx", "pdf"),
    );
    expect(local.startsWith("/sandbox/t/")).toBe(true);
    expect(local).not.toContain("../");
  });
});

describe("validateHttpsUrl (redirect target)", () => {
  it("accepts an https URL", () => {
    expect(validateHttpsUrl("https://cdn.example/x?sig=1", "Location")).toBe(
      "https://cdn.example/x?sig=1",
    );
  });

  it.each([
    [null, "Location is missing"],
    ["", "Location is missing"],
    ["/relative/path", "not an absolute URL"],
    ["http://cdn.example/x", "non-https"],
    ["file:///etc/passwd", "non-https"],
  ])("rejects %j", (value, msg) => {
    expect(() => validateHttpsUrl(value, "Location")).toThrow(msg);
  });
});

describe("cappedCounter", () => {
  it("counts bytes and errors past the limit", async () => {
    const ok = cappedCounter(10);
    await pipeline(Readable.from([Buffer.from("12345")]), ok.stream, new PassThrough().resume());
    expect(ok.bytes()).toBe(5);

    const capped = cappedCounter(4);
    await expect(
      pipeline(Readable.from([Buffer.from("12345")]), capped.stream, new PassThrough().resume()),
    ).rejects.toThrow("exceeds the 200 MB cap");
  });
});

describe("downloadFileTool handler — format: 'pdf'", () => {
  const originalEnv = { ...process.env };
  let sandboxBase: string;

  interface ConvertSpec {
    item: Record<string, unknown>;
    contentResponse?: Response;
  }

  function mockClient(spec: ConvertSpec): {
    client: Client;
    calls: Array<{ path: string; middlewareOptions?: unknown[]; responseType?: unknown }>;
  } {
    const calls: Array<{ path: string; middlewareOptions?: unknown[]; responseType?: unknown }> = [];
    const api = vi.fn((p: string) => {
      const call: (typeof calls)[number] = { path: p };
      calls.push(call);
      const req = {
        middlewareOptions: vi.fn((opts: unknown[]) => {
          call.middlewareOptions = opts;
          return req;
        }),
        responseType: vi.fn((rt: unknown) => {
          call.responseType = rt;
          return req;
        }),
        get: vi.fn(async () => {
          if (p.includes("/content")) {
            if (!spec.contentResponse) throw new Error("unexpected content call");
            return spec.contentResponse;
          }
          return spec.item;
        }),
      };
      return req;
    });
    return { client: { api } as unknown as Client, calls };
  }

  const docx = {
    id: "01DOC",
    name: "NDA.docx",
    size: 12_345,
    file: { mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" },
    "@microsoft.graph.downloadUrl": "https://cdn.example/original",
  };

  beforeEach(async () => {
    sandboxBase = await mkdtemp(path.join(os.tmpdir(), "m365-dl-test-"));
    process.env.M365_DOWNLOAD_DIR = sandboxBase;
    process.env.M365_TENANT_ID = "tenant-t";
  });

  afterEach(async () => {
    process.env = { ...originalEnv };
    vi.unstubAllGlobals();
    await rm(sandboxBase, { recursive: true, force: true });
  });

  it("rejects an invalid format before any Graph call", async () => {
    const { client } = mockClient({ item: docx });
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "html" }),
    ).rejects.toThrow("'format' must be one of");
    expect((client.api as ReturnType<typeof vi.fn>)).not.toHaveBeenCalled();
  });

  it("rejects an unsupported source type after metadata only", async () => {
    const { client, calls } = mockClient({
      item: { ...docx, name: "photo.png", file: { mimeType: "image/png" } },
    });
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow("Cannot convert 'photo.png' to PDF");
    expect(calls.map((c) => c.path)).toEqual(["/me/drive/items/01DOC"]);
  });

  it("follows Graph's 302 itself, without auth, and writes <hash>-NDA.pdf (0600)", async () => {
    const redirect = new Response(null, {
      status: 302,
      headers: { Location: "https://cdn.example/converted?sig=abc" },
    });
    const { client, calls } = mockClient({ item: docx, contentResponse: redirect });
    const pdfBytes = Buffer.from("%PDF-1.7\nfake\n%%EOF");
    const fetchMock = vi.fn(async () => new Response(pdfBytes, { status: 200 }));
    vi.stubGlobal("fetch", fetchMock);

    const res = await downloadFileTool.handler(client, {
      item_id: "01DOC",
      drive_id: "b!lib",
      format: "pdf",
    });
    const out = JSON.parse((res.content[0] as { text: string }).text);

    expect(calls.map((c) => c.path)).toEqual([
      "/drives/b!lib/items/01DOC",
      "/drives/b!lib/items/01DOC/content?format=pdf",
    ]);
    // SDK redirect-following disabled for the conversion request.
    const opts = calls[1].middlewareOptions as Array<{ shouldRedirect: () => boolean }>;
    expect(opts).toHaveLength(1);
    expect(opts[0].shouldRedirect()).toBe(false);
    expect(calls[1].responseType).toBe("raw");

    // The pre-authenticated URL is fetched with no init (no bearer token).
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(fetchMock.mock.calls[0]).toEqual(["https://cdn.example/converted?sig=abc"]);

    const sandbox = path.join(sandboxBase, "tenant-t");
    expect(out.local_path).toMatch(new RegExp(`^${sandbox}/[0-9a-f]{16}-NDA\\.pdf$`));
    expect(await readFile(out.local_path)).toEqual(pdfBytes);
    expect((await stat(out.local_path)).mode & 0o777).toBe(0o600);
    expect(out).toMatchObject({
      item_id: "01DOC",
      drive_id: "b!lib",
      size_bytes: pdfBytes.length,
      name: "NDA.docx",
      content_type: "application/pdf",
      format: "pdf",
      source_size_bytes: 12_345,
      source_content_type: docx.file.mimeType,
    });
  });

  it("accepts a direct 200 with the converted body", async () => {
    const body = Buffer.from("%PDF-direct");
    const { client } = mockClient({ item: docx, contentResponse: new Response(body, { status: 200 }) });
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);

    const res = await downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" });
    const out = JSON.parse((res.content[0] as { text: string }).text);
    expect(fetchMock).not.toHaveBeenCalled();
    expect(await readFile(out.local_path)).toEqual(body);
  });

  it("refuses a non-https redirect target", async () => {
    const redirect = new Response(null, {
      status: 302,
      headers: { Location: "http://cdn.example/converted" },
    });
    const { client } = mockClient({ item: docx, contentResponse: redirect });
    const fetchMock = vi.fn();
    vi.stubGlobal("fetch", fetchMock);
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow("non-https");
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it("maps Graph's 406 notSupported to a clear conversion error", async () => {
    const graphErr = new Response(
      JSON.stringify({ error: { code: "notSupported", message: "Conversion not supported" } }),
      { status: 406 },
    );
    const { client } = mockClient({ item: docx, contentResponse: graphErr });
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow(
      "Microsoft Graph could not convert 'NDA.docx' to PDF: HTTP 406 (notSupported: Conversion not supported)",
    );
  });

  it("surfaces other Graph errors with status and code", async () => {
    const graphErr = new Response(
      JSON.stringify({ error: { code: "accessDenied", message: "nope" } }),
      { status: 403 },
    );
    const { client } = mockClient({ item: docx, contentResponse: graphErr });
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow("PDF conversion HTTP 403 (accessDenied: nope)");
  });

  it("refuses a converted file whose Content-Length exceeds the 200 MB cap", async () => {
    const redirect = new Response(null, {
      status: 302,
      headers: { Location: "https://cdn.example/converted" },
    });
    const { client } = mockClient({ item: docx, contentResponse: redirect });
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response("x", {
        status: 200,
        headers: { "content-length": String(201 * 1024 * 1024) },
      })),
    );
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow("exceeds the 200 MB cap");
  });

  it("reports a failed download of the converted bytes and leaves no file", async () => {
    const redirect = new Response(null, {
      status: 302,
      headers: { Location: "https://cdn.example/converted" },
    });
    const { client } = mockClient({ item: docx, contentResponse: redirect });
    vi.stubGlobal(
      "fetch",
      vi.fn(async () => new Response("gone", { status: 404, statusText: "Not Found" })),
    );
    await expect(
      downloadFileTool.handler(client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow("PDF conversion download HTTP 404");
    const sandbox = path.join(sandboxBase, "tenant-t");
    expect(await readdir(sandbox)).toEqual([]);
  });

  it("still applies the folder and source-size pre-flight checks", async () => {
    const folder = mockClient({ item: { id: "f", name: "Docs.docx", folder: {} } });
    await expect(
      downloadFileTool.handler(folder.client, { item_id: "f", format: "pdf" }),
    ).rejects.toThrow("is a folder");

    const big = mockClient({ item: { ...docx, size: 250 * 1024 * 1024 } });
    await expect(
      downloadFileTool.handler(big.client, { item_id: "01DOC", format: "pdf" }),
    ).rejects.toThrow("exceeds the 200 MB cap");
  });
});

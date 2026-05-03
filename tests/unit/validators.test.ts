import { describe, expect, it } from "vitest";

import {
  sanitizeFilename,
  validateOptionalInteger,
  validateOptionalString,
  validateRequiredISODate,
  validateRequiredString,
} from "../../src/types/validators.js";

describe("validateRequiredString", () => {
  it("returns the string when valid", () => {
    expect(validateRequiredString("foo", "field")).toBe("foo");
  });

  it("throws on empty string", () => {
    expect(() => validateRequiredString("", "field")).toThrow(
      "'field' must be a non-empty string",
    );
  });

  it("throws on undefined", () => {
    expect(() => validateRequiredString(undefined, "field")).toThrow();
  });

  it("throws on null", () => {
    expect(() => validateRequiredString(null, "field")).toThrow();
  });

  it.each([
    [42, "number"],
    [true, "boolean"],
    [{}, "object"],
    [[], "array"],
  ])("throws on non-string (%s = %s)", (value) => {
    expect(() => validateRequiredString(value, "field")).toThrow();
  });
});

describe("validateOptionalString", () => {
  it("returns undefined for undefined input", () => {
    expect(validateOptionalString(undefined, "field")).toBeUndefined();
  });

  it("returns undefined for null input", () => {
    expect(validateOptionalString(null, "field")).toBeUndefined();
  });

  it("returns the string when valid", () => {
    expect(validateOptionalString("foo", "field")).toBe("foo");
  });

  it("throws on empty string (delegates to required)", () => {
    expect(() => validateOptionalString("", "field")).toThrow();
  });

  it("throws on non-string non-nullish input", () => {
    expect(() => validateOptionalString(42, "field")).toThrow();
  });
});

describe("validateOptionalInteger", () => {
  const opts = { min: 1, max: 10, default: 5 };

  it("returns default for undefined", () => {
    expect(validateOptionalInteger(undefined, "n", opts)).toBe(5);
  });

  it("returns default for null", () => {
    expect(validateOptionalInteger(null, "n", opts)).toBe(5);
  });

  it("returns the integer when in range", () => {
    expect(validateOptionalInteger(7, "n", opts)).toBe(7);
  });

  it("accepts the min boundary", () => {
    expect(validateOptionalInteger(1, "n", opts)).toBe(1);
  });

  it("accepts the max boundary", () => {
    expect(validateOptionalInteger(10, "n", opts)).toBe(10);
  });

  it("throws below min", () => {
    expect(() => validateOptionalInteger(0, "n", opts)).toThrow(
      "'n' must be an integer between 1 and 10",
    );
  });

  it("throws above max", () => {
    expect(() => validateOptionalInteger(11, "n", opts)).toThrow();
  });

  it("throws on non-integer number", () => {
    expect(() => validateOptionalInteger(3.5, "n", opts)).toThrow();
  });

  it("throws on non-number", () => {
    expect(() => validateOptionalInteger("5", "n", opts)).toThrow();
  });
});

describe("sanitizeFilename", () => {
  it("replaces forward slashes with underscores", () => {
    expect(sanitizeFilename("foo/bar/baz")).toBe("foo_bar_baz");
  });

  it("replaces backslashes with underscores", () => {
    expect(sanitizeFilename("foo\\bar")).toBe("foo_bar");
  });

  it("replaces null bytes with underscores", () => {
    expect(sanitizeFilename("foo\0bar")).toBe("foo_bar");
  });

  it("strips leading dots (path traversal defense)", () => {
    expect(sanitizeFilename("..hidden")).toBe("_hidden");
    expect(sanitizeFilename("...etc")).toBe("_etc");
  });

  it("preserves dots elsewhere in the name", () => {
    expect(sanitizeFilename("doc.v1.pdf")).toBe("doc.v1.pdf");
  });

  it("caps length at 200 chars", () => {
    const long = "a".repeat(500);
    const result = sanitizeFilename(long);
    expect(result.length).toBe(200);
    expect(result).toBe("a".repeat(200));
  });

  it("rejects classic path-traversal payload structurally", () => {
    // After sanitization, ../../etc/passwd cannot be a path component
    // because '/' becomes '_' and leading '.' becomes '_'.
    const result = sanitizeFilename("../../etc/passwd");
    expect(result).not.toContain("/");
    expect(result.startsWith(".")).toBe(false);
  });
});

describe("validateRequiredISODate", () => {
  it("accepts a date-only ISO string", () => {
    expect(validateRequiredISODate("2026-05-04", "f")).toBe("2026-05-04");
  });

  it("accepts a datetime with Z timezone", () => {
    expect(validateRequiredISODate("2026-05-04T10:00:00Z", "f")).toBe("2026-05-04T10:00:00Z");
  });

  it("accepts a datetime with offset timezone", () => {
    expect(validateRequiredISODate("2026-05-04T10:00:00+01:00", "f")).toBe(
      "2026-05-04T10:00:00+01:00",
    );
  });

  it("accepts a datetime with milliseconds", () => {
    const s = "2026-05-04T10:00:00.123Z";
    expect(validateRequiredISODate(s, "f")).toBe(s);
  });

  it.each([
    "yesterday",
    "2026/05/04",
    "May 4 2026",
    "abc",
    "",
  ])("rejects malformed input: %s", (bad) => {
    expect(() => validateRequiredISODate(bad, "f")).toThrow();
  });

  it("delegates to validateRequiredString for non-string inputs", () => {
    expect(() => validateRequiredISODate(undefined, "f")).toThrow(
      "'f' must be a non-empty string",
    );
  });
});

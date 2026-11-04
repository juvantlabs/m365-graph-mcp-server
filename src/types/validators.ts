/**
 * Input validators for MCP tool handlers.
 *
 * Tool handlers receive `Record<string, unknown>` from the MCP SDK and
 * must validate each field before passing it to Graph. These helpers
 * centralize the validation logic + give consistent error messages
 * the agent can act on.
 *
 * Naming convention: every exported helper starts with `validate*`,
 * which the CI dead-code grep enforces is imported elsewhere in src/.
 * Defense-in-depth code that's never wired into a real handler is a
 * security smell (handbook anti-pattern S1).
 */

export function validateRequiredString(value: unknown, fieldName: string): string {
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`'${fieldName}' must be a non-empty string`);
  }
  return value;
}

export function validateOptionalString(
  value: unknown,
  fieldName: string,
): string | undefined {
  if (value === undefined || value === null) return undefined;
  return validateRequiredString(value, fieldName);
}

export function validateOptionalInteger(
  value: unknown,
  fieldName: string,
  options: { min: number; max: number; default: number },
): number {
  if (value === undefined || value === null) return options.default;
  if (
    typeof value !== "number" ||
    !Number.isInteger(value) ||
    value < options.min ||
    value > options.max
  ) {
    throw new Error(
      `'${fieldName}' must be an integer between ${options.min} and ${options.max}`,
    );
  }
  return value;
}

const ISO_DATE_LOOKS_LIKE = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}(:\d{2}(\.\d+)?)?(Z|[+-]\d{2}:?\d{2})?)?$/;

export function validateOptionalEnum<T extends string>(
  value: unknown,
  fieldName: string,
  allowed: ReadonlyArray<T>,
  defaultValue: T,
): T {
  if (value === undefined || value === null) return defaultValue;
  if (typeof value !== "string" || !(allowed as ReadonlyArray<string>).includes(value)) {
    throw new Error(
      `'${fieldName}' must be one of ${JSON.stringify(allowed)}; got ${JSON.stringify(value)}`,
    );
  }
  return value as T;
}

/**
 * Optional strict enum with no default: absent/null → undefined;
 * anything else must be one of `allowed` exactly (case-sensitive).
 * Used where "unset" has its own meaning (e.g. download_file `format`
 * absent = original bytes), so a default value would be wrong.
 */
export function validateOptionalStrictEnum<T extends string>(
  value: unknown,
  fieldName: string,
  allowed: ReadonlyArray<T>,
): T | undefined {
  if (value === undefined || value === null) return undefined;
  if (typeof value !== "string" || !(allowed as ReadonlyArray<string>).includes(value)) {
    throw new Error(
      `'${fieldName}' must be one of ${JSON.stringify(allowed)}; got ${JSON.stringify(value)}`,
    );
  }
  return value as T;
}

/**
 * Require an absolute https URL (e.g. a server-issued redirect target
 * that will be fetched). Returns the normalized URL string.
 */
export function validateHttpsUrl(value: unknown, label: string): string {
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`${label} is missing`);
  }
  let url: URL;
  try {
    url = new URL(value);
  } catch {
    throw new Error(`${label} is not an absolute URL`);
  }
  if (url.protocol !== "https:") {
    throw new Error(`${label} must be https; refusing non-https target (${url.protocol})`);
  }
  return url.toString();
}

/**
 * Lightweight ISO 8601 date/datetime check. Accepts:
 *   - YYYY-MM-DD
 *   - YYYY-MM-DDTHH:MM
 *   - YYYY-MM-DDTHH:MM:SS
 *   - YYYY-MM-DDTHH:MM:SS.sss
 * with optional Z or ±HH:MM timezone offset.
 *
 * Microsoft Graph does the strict parsing; we just guard against
 * obviously-wrong inputs so the agent gets a near-the-input error
 * rather than a cryptic Graph 400.
 */
export function validateRequiredISODate(value: unknown, fieldName: string): string {
  const s = validateRequiredString(value, fieldName);
  if (!ISO_DATE_LOOKS_LIKE.test(s)) {
    throw new Error(
      `'${fieldName}' must be an ISO 8601 date or datetime (e.g. '2026-05-04' or '2026-05-04T10:00:00Z'); got ${JSON.stringify(s)}`,
    );
  }
  return s;
}

/**
 * Strip path-traversal-dangerous characters from a filename. Used
 * before constructing a local FS path that includes a server-supplied
 * filename (e.g. download_file caches the Graph item.name on disk).
 *
 * Defense-in-depth: combined with prefix check on path.resolve(), even
 * a malicious filename ('../../etc/passwd') cannot escape the sandbox.
 */
export function sanitizeFilename(name: string): string {
  return name
    .replace(/[/\\\0]/g, "_")
    .replace(/^\.+/, "_") // no leading dots
    .slice(0, 200);
}

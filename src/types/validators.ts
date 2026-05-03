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

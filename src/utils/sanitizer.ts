/**
 * Sanitize a string for safe Excel output.
 * Only removes invalid XML 1.0 control characters.
 * NOTE: Do NOT escape XML entities (&amp; &lt; etc.) — ExcelJS handles that internally.
 */
export function sanitizeForExcel(input: string): string {
  if (!input || typeof input !== 'string') return input;

  // Remove invalid XML 1.0 characters (allow valid ranges)
  let cleaned = '';
  for (let i = 0; i < input.length; i++) {
    const code = input.charCodeAt(i);
    const isValid =
      code === 0x09 ||
      code === 0x0a ||
      code === 0x0d ||
      (code >= 0x20 && code <= 0xd7ff) ||
      (code >= 0xe000 && code <= 0xfffd);
    if (isValid) cleaned += input[i];
  }

  return cleaned;
}

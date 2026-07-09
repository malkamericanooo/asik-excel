/**
 * Sanitize a string for safe XML/Excel output.
 * Removes invalid XML characters and escapes XML special characters.
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

  // Escape XML special characters
  return cleaned
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

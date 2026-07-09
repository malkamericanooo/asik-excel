/**
 * Convert a YYYY-MM-DD string to an Excel serial date number.
 * Excel epoch: 1899-12-30 UTC. Uses UTC to avoid timezone/DST skew.
 * For modern dates (post 1900-02-28) this formula is exact because Excel's
 * erroneous 1900 leap day shifts all serials after serial 60 by +1, which is
 * precisely accounted for by using Dec 30 as the epoch instead of Dec 31.
 */
export function dateStringToExcelSerial(dateStr: string): number {
  const parts = dateStr.split('-');
  if (parts.length !== 3) return 0;
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);
  const day = parseInt(parts[2], 10);
  if (isNaN(year) || isNaN(month) || isNaN(day)) return 0;
  const dateMs = Date.UTC(year, month - 1, day);
  const epochMs = Date.UTC(1899, 11, 30); // 1899-12-30 UTC
  return Math.round((dateMs - epochMs) / 86_400_000);
}

/**
 * Convert an Excel serial date to a JS Date object (UTC midnight).
 */
export function excelSerialToDate(serial: number): Date {
  const epochMs = Date.UTC(1899, 11, 30); // 1899-12-30 UTC
  return new Date(epochMs + serial * 86_400_000);
}

/**
 * Convert an Excel serial date to "YYYY-MM-DD" string.
 */
export function excelSerialToDateStr(serial: number): string {
  const d = excelSerialToDate(serial);
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, '0');
  const day = String(d.getUTCDate()).padStart(2, '0');
  return `${y}-${m}-${day}`;
}

/**
 * Check if an Excel serial date falls within the given month/year (1-indexed month).
 */
export function isInMonthYear(serial: number, month: number, year: number): boolean {
  if (!serial) return false;
  const d = excelSerialToDate(serial);
  return d.getUTCFullYear() === year && d.getUTCMonth() + 1 === month;
}

/** Indonesian month names (index 1-12) */
export const BULAN_INDONESIA = [
  '', 'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember',
];

/**
 * Parse a raw cell value that might be a date string, Excel serial, or Date object.
 * Returns a "YYYY-MM-DD" string, or '' if unparseable.
 */
export function parseDateCell(cell: unknown): string {
  if (cell == null) return '';

  // Numeric Excel serial
  if (typeof cell === 'number') {
    if (cell <= 0) return '';
    return excelSerialToDateStr(cell);
  }

  if (typeof cell === 'string') {
    const s = cell.trim();
    if (!s) return '';

    // YYYY-MM-DD (ISO) — most common from ASIK
    if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);

    // DD/MM/YYYY
    const dmy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (dmy) return `${dmy[3]}-${dmy[2].padStart(2, '0')}-${dmy[1].padStart(2, '0')}`;

    // DD-MM-YYYY
    const dmy2 = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})/);
    if (dmy2) return `${dmy2[3]}-${dmy2[2].padStart(2, '0')}-${dmy2[1].padStart(2, '0')}`;

    // Try JS Date parsing as last resort (e.g. "17 Jun 2026")
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) {
      const y = parsed.getUTCFullYear();
      const mo = String(parsed.getUTCMonth() + 1).padStart(2, '0');
      const d = String(parsed.getUTCDate()).padStart(2, '0');
      return `${y}-${mo}-${d}`;
    }
  }

  // Date object
  if (cell instanceof Date) {
    const y = cell.getUTCFullYear();
    const mo = String(cell.getUTCMonth() + 1).padStart(2, '0');
    const d = String(cell.getUTCDate()).padStart(2, '0');
    return `${y}-${mo}-${d}`;
  }

  return '';
}

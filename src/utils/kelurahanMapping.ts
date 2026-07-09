import type { SheetName } from '../types';

/** Normalize for exact comparison: trim, lowercase, remove apostrophes */
function normalizeKelurahan(s: string): string {
  return s
    .trim()
    .toLowerCase()
    .replace(/[''`]/g, '')
    .replace(/\s+/g, ' ');
}

const VALID_KELURAHAN: Record<string, Exclude<SheetName, 'LUAR WILAYAH' | 'Kejar'>> = {
  mabuun: 'MABUUN',
  pembataan: 'PEMBATAAN',
  kasiau: 'KASIAU',
  maburai: 'MABURAI',
  sulingan: 'SULINGAN',
};

/**
 * Classify a kelurahan string to a master sheet name.
 * Exact match only (case-insensitive, trimmed). Call after ruling out Status = "kejar".
 */
export function classifyKelurahan(kelurahan: string): SheetName {
  if (!kelurahan || !kelurahan.trim()) return 'LUAR WILAYAH';
  const key = normalizeKelurahan(kelurahan);
  return VALID_KELURAHAN[key] ?? 'LUAR WILAYAH';
}

export const AREA_SHEETS: Exclude<SheetName, 'LUAR WILAYAH' | 'Kejar'>[] = [
  'MABUUN', 'KASIAU', 'PEMBATAAN', 'SULINGAN', 'MABURAI',
];

export const SHEET_LABELS: Record<SheetName, string> = {
  MABUUN: 'Mabuun',
  KASIAU: 'Kasiau',
  PEMBATAAN: 'Pembataan',
  SULINGAN: 'Sulingan',
  MABURAI: 'Maburai',
  'LUAR WILAYAH': 'Luar Wilayah',
  Kejar: 'Kejar',
};

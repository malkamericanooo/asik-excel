import type { VaccineDatabase, CumulativeChild } from './vaccineDatabase';
import { DB_SCHEMA_VERSION, createEmptyDatabase } from './vaccineDatabase';
import { ALL_SHEETS } from '../types';

/**
 * Penyimpanan database kumulatif — seluruhnya di perangkat nakes.
 * Tidak ada server, tidak ada pengiriman data pasien ke mana pun.
 *
 * localStorage dipilih karena cukup untuk skala Puskesmas (±2.000 anak
 * ≈ 1 MB JSON) dan tidak butuh setup. Kalau nanti datanya jauh lebih besar,
 * ganti isi loadDatabase/saveDatabase dengan IndexedDB — pemanggilnya tidak
 * perlu berubah.
 */

export const STORAGE_KEY = 'puskesmas-mabuun:vaccine-db:v1';

/** Batas aman sebelum localStorage biasanya menolak (±5 MB). */
const SIZE_WARN_BYTES = 4_000_000;

export interface LoadResult {
  database: VaccineDatabase;
  /** Diisi kalau data tersimpan ada tapi tidak bisa dipakai. */
  error?: string;
}

function isValidChild(v: unknown): v is CumulativeChild {
  if (!v || typeof v !== 'object') return false;
  const c = v as Record<string, unknown>;
  return (
    typeof c.id === 'string' &&
    typeof c.nama === 'string' &&
    (c.jk === 'L' || c.jk === 'P') &&
    typeof c.tanggalLahirStr === 'string' &&
    typeof c.tanggalLahirSerial === 'number' &&
    typeof c.sheet === 'string' &&
    (ALL_SHEETS as string[]).includes(c.sheet as string) &&
    typeof c.firstSeen === 'string' &&
    typeof c.lastSeen === 'string' &&
    !!c.vaccines && typeof c.vaccines === 'object'
  );
}

/**
 * Validasi bentuk database. Dipakai untuk data dari localStorage MAUPUN
 * dari file JSON yang diimpor pengguna — file impor tidak bisa dipercaya
 * begitu saja, jadi bentuknya diperiksa sebelum masuk aplikasi.
 */
export function parseDatabase(raw: unknown): { database?: VaccineDatabase; error?: string } {
  if (!raw || typeof raw !== 'object') return { error: 'Isi file bukan objek database.' };
  const d = raw as Record<string, unknown>;

  if (typeof d.schemaVersion !== 'number') return { error: 'File tidak punya schemaVersion.' };
  if (d.schemaVersion > DB_SCHEMA_VERSION) {
    return { error: `File dibuat aplikasi versi lebih baru (schema ${d.schemaVersion}), aplikasi ini hanya sampai ${DB_SCHEMA_VERSION}.` };
  }
  if (!Array.isArray(d.children)) return { error: 'File tidak punya daftar children.' };
  if (!Array.isArray(d.periods)) return { error: 'File tidak punya daftar periods.' };

  const children = d.children.filter(isValidChild);
  const dropped = d.children.length - children.length;
  if (children.length === 0 && d.children.length > 0) {
    return { error: 'Semua baris di file tidak sesuai format database.' };
  }

  const database: VaccineDatabase = {
    schemaVersion: DB_SCHEMA_VERSION,
    updatedAt: typeof d.updatedAt === 'string' ? d.updatedAt : new Date().toISOString(),
    periods: (d.periods as unknown[]).filter((p): p is string => typeof p === 'string').sort(),
    children: children.map((c) => ({ ...c, vaccineSources: c.vaccineSources ?? {} })),
  };

  return dropped > 0
    ? { database, error: `${dropped} baris dilewati karena formatnya tidak sesuai.` }
    : { database };
}

export function loadDatabase(storage: Storage = localStorage): LoadResult {
  let text: string | null = null;
  try {
    text = storage.getItem(STORAGE_KEY);
  } catch {
    return { database: createEmptyDatabase(), error: 'Penyimpanan browser tidak bisa dibaca.' };
  }
  if (!text) return { database: createEmptyDatabase() };

  try {
    const parsed = parseDatabase(JSON.parse(text));
    if (!parsed.database) {
      return { database: createEmptyDatabase(), error: parsed.error };
    }
    return { database: parsed.database, error: parsed.error };
  } catch {
    return { database: createEmptyDatabase(), error: 'Data tersimpan rusak dan tidak bisa dibaca.' };
  }
}

export function saveDatabase(
  db: VaccineDatabase,
  storage: Storage = localStorage,
): { ok: boolean; error?: string; bytes: number } {
  const text = JSON.stringify(db);
  const bytes = text.length;
  try {
    storage.setItem(STORAGE_KEY, text);
  } catch {
    return {
      ok: false,
      bytes,
      error: `Database terlalu besar untuk penyimpanan browser (${(bytes / 1_000_000).toFixed(1)} MB). Unduh sebagai file JSON lalu kosongkan penyimpanan.`,
    };
  }
  return bytes > SIZE_WARN_BYTES
    ? { ok: true, bytes, error: `Database sudah ${(bytes / 1_000_000).toFixed(1)} MB — sebaiknya diunduh sebagai cadangan.` }
    : { ok: true, bytes };
}

export function clearDatabase(storage: Storage = localStorage): void {
  try {
    storage.removeItem(STORAGE_KEY);
  } catch {
    /* penyimpanan tidak tersedia — tidak ada yang perlu dibersihkan */
  }
}

/** Database sebagai file JSON untuk cadangan / dipindah ke web lanjutan. */
export function databaseToBlob(db: VaccineDatabase): Blob {
  return new Blob([JSON.stringify(db, null, 2)], { type: 'application/json' });
}

export function databaseFileName(db: VaccineDatabase): string {
  const last = db.periods[db.periods.length - 1] ?? 'kosong';
  return `database-imunisasi-kumulatif_${last}.json`;
}

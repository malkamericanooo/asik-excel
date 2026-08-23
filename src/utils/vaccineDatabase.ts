import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { childKey } from './childIdentity';
import { VACCINE_ORDER } from './vaccineMapping';

/**
 * DATABASE IMUNISASI KUMULATIF
 *
 * Laporan bulanan ke pemerintah tetap dibangun ulang tiap bulan dari file
 * ASIK bulan itu saja (lihat masterExcel.ts) — itu memang yang diminta.
 * Database ini terpisah: ia MENUMPUK hasil tiap bulan menjadi satu riwayat
 * imunisasi per anak, supaya web lanjutan (dashboard prioritas, sweeping
 * kader, pengingat WA) punya data historis untuk dihitung.
 *
 * Tetap 100% di browser — tidak ada data pasien yang dikirim ke mana pun.
 */

export const DB_SCHEMA_VERSION = 1;

export interface CumulativeChild extends ChildRecord {
  /** Kunci identitas stabil (nama ternormalisasi + tanggal lahir). */
  id: string;
  /** Sheet wilayah terakhir yang diketahui. */
  sheet: SheetName;
  /** Periode "YYYY-MM" saat anak ini pertama & terakhir terlihat. */
  firstSeen: string;
  lastSeen: string;
  /** Periode asal tiap tanggal vaksin — untuk audit kalau ada koreksi. */
  vaccineSources: Partial<Record<VaccineKey, string>>;
}

export interface VaccineDatabase {
  schemaVersion: number;
  /** ISO timestamp perubahan terakhir. */
  updatedAt: string;
  /** Periode "YYYY-MM" yang sudah pernah digabung, terurut menaik. */
  periods: string[];
  children: CumulativeChild[];
}

/** Tanggal vaksin yang sama tercatat berbeda antar periode. */
export interface MergeConflict {
  id: string;
  nama: string;
  vaccine: VaccineKey;
  existingSerial: number;
  incomingSerial: number;
  existingPeriod: string;
  incomingPeriod: string;
}

export interface MergeStats {
  period: string;
  /** True kalau periode ini sudah pernah digabung sebelumnya (upload ulang). */
  alreadyMerged: boolean;
  childrenAdded: number;
  childrenUpdated: number;
  vaccinesAdded: number;
  vaccinesChanged: number;
  sheetChanges: number;
  conflicts: MergeConflict[];
}

/** Periode laporan sebagai "YYYY-MM" (bulan 1-indexed). */
export function periodKey(month: number, year: number): string {
  return `${year}-${String(month).padStart(2, '0')}`;
}

export function createEmptyDatabase(): VaccineDatabase {
  return {
    schemaVersion: DB_SCHEMA_VERSION,
    updatedAt: new Date().toISOString(),
    periods: [],
    children: [],
  };
}

function emptyStats(period: string, alreadyMerged: boolean): MergeStats {
  return {
    period,
    alreadyMerged,
    childrenAdded: 0,
    childrenUpdated: 0,
    vaccinesAdded: 0,
    vaccinesChanged: 0,
    sheetChanges: 0,
    conflicts: [],
  };
}

/**
 * Gabungkan satu periode laporan bulanan ke dalam database kumulatif.
 *
 * Aturan gabung:
 *   - Anak dikenali dengan kunci yang sama persis dengan merger bulanan
 *     (nama ternormalisasi + tanggal lahir), jadi tidak akan terpecah dua.
 *   - Tanggal vaksin baru ditambahkan; kalau vaksin yang sama sudah ada
 *     dengan tanggal BERBEDA, nilai baru dipakai (dianggap koreksi) dan
 *     selisihnya dicatat di `stats.conflicts` supaya bisa ditinjau nakes.
 *   - Identitas (NIK, nama ortu, alamat, wilayah) diperbarui kalau data
 *     baru lebih lengkap; nilai lama tidak ditimpa oleh nilai kosong.
 *
 * Murni — database lama tidak dimutasi, versi baru dikembalikan.
 */
export function mergeMonthlyIntoDatabase(
  db: VaccineDatabase,
  masterData: MasterData,
  month: number,
  year: number,
): { database: VaccineDatabase; stats: MergeStats } {
  const period = periodKey(month, year);
  const stats = emptyStats(period, db.periods.includes(period));

  const byId = new Map<string, CumulativeChild>();
  for (const child of db.children) {
    byId.set(child.id, { ...child, vaccines: { ...child.vaccines }, vaccineSources: { ...child.vaccineSources } });
  }

  for (const sheet of ALL_SHEETS) {
    for (const incoming of masterData[sheet]) {
      const id = childKey(incoming.nama, incoming.tanggalLahirStr);
      const existing = byId.get(id);

      if (!existing) {
        const vaccineSources: Partial<Record<VaccineKey, string>> = {};
        for (const vk of Object.keys(incoming.vaccines) as VaccineKey[]) {
          if (incoming.vaccines[vk]) {
            vaccineSources[vk] = period;
            stats.vaccinesAdded++;
          }
        }
        byId.set(id, {
          ...incoming,
          vaccines: { ...incoming.vaccines },
          id,
          sheet,
          firstSeen: period,
          lastSeen: period,
          vaccineSources,
        });
        stats.childrenAdded++;
        continue;
      }

      let touched = false;

      for (const vk of Object.keys(incoming.vaccines) as VaccineKey[]) {
        const incomingSerial = incoming.vaccines[vk];
        if (!incomingSerial) continue;
        const existingSerial = existing.vaccines[vk];

        if (!existingSerial) {
          existing.vaccines[vk] = incomingSerial;
          existing.vaccineSources[vk] = period;
          stats.vaccinesAdded++;
          touched = true;
        } else if (existingSerial !== incomingSerial) {
          stats.conflicts.push({
            id,
            nama: existing.nama,
            vaccine: vk,
            existingSerial,
            incomingSerial,
            existingPeriod: existing.vaccineSources[vk] ?? existing.lastSeen,
            incomingPeriod: period,
          });
          existing.vaccines[vk] = incomingSerial;
          existing.vaccineSources[vk] = period;
          stats.vaccinesChanged++;
          touched = true;
        }
      }

      // Identitas: isi yang kosong, jangan timpa isi dengan kosong.
      if (!existing.nik && incoming.nik) { existing.nik = incoming.nik; touched = true; }
      if (!existing.namaOrangTua && incoming.namaOrangTua) { existing.namaOrangTua = incoming.namaOrangTua; touched = true; }
      if (incoming.alamat && incoming.alamat !== existing.alamat) { existing.alamat = incoming.alamat; touched = true; }
      if (incoming.jk && incoming.jk !== existing.jk) { existing.jk = incoming.jk; touched = true; }

      if (sheet !== existing.sheet) {
        existing.sheet = sheet;
        stats.sheetChanges++;
        touched = true;
      }

      if (period > existing.lastSeen) existing.lastSeen = period;
      if (period < existing.firstSeen) existing.firstSeen = period;
      if (touched) stats.childrenUpdated++;
    }
  }

  const periods = db.periods.includes(period)
    ? [...db.periods]
    : [...db.periods, period].sort();

  return {
    database: {
      schemaVersion: DB_SCHEMA_VERSION,
      updatedAt: new Date().toISOString(),
      periods,
      children: [...byId.values()],
    },
    stats,
  };
}

// ─── Query helpers (bahan untuk web lanjutan) ────────────────────────

/** Jumlah anak yang sudah menerima tiap vaksin, dipecah L/P. */
export function getCoverage(
  db: VaccineDatabase,
): Record<VaccineKey, { L: number; P: number; total: number }> {
  const out = {} as Record<VaccineKey, { L: number; P: number; total: number }>;
  for (const vk of VACCINE_ORDER) out[vk] = { L: 0, P: 0, total: 0 };
  for (const child of db.children) {
    for (const vk of VACCINE_ORDER) {
      if (child.vaccines[vk]) {
        out[vk][child.jk]++;
        out[vk].total++;
      }
    }
  }
  return out;
}

/** Anak per wilayah/sheet. */
export function getChildrenBySheet(db: VaccineDatabase, sheet: SheetName): CumulativeChild[] {
  return db.children.filter((c) => c.sheet === sheet);
}

export function findChildById(db: VaccineDatabase, id: string): CumulativeChild | undefined {
  return db.children.find((c) => c.id === id);
}

/** Vaksin yang BELUM tercatat untuk seorang anak — dasar antrean tindak lanjut. */
export function getMissingVaccines(child: CumulativeChild): VaccineKey[] {
  return VACCINE_ORDER.filter((vk) => !child.vaccines[vk]);
}

export function getDatabaseSummary(db: VaccineDatabase): {
  totalChildren: number;
  totalVaccineRecords: number;
  periods: string[];
  perSheet: Record<SheetName, number>;
} {
  const perSheet = {} as Record<SheetName, number>;
  for (const s of ALL_SHEETS) perSheet[s] = 0;
  let totalVaccineRecords = 0;
  for (const child of db.children) {
    perSheet[child.sheet]++;
    for (const vk of VACCINE_ORDER) if (child.vaccines[vk]) totalVaccineRecords++;
  }
  return {
    totalChildren: db.children.length,
    totalVaccineRecords,
    periods: [...db.periods],
    perSheet,
  };
}

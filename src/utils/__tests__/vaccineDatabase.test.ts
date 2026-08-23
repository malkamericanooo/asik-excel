import { describe, it, expect } from 'vitest';
import {
  createEmptyDatabase, mergeMonthlyIntoDatabase, periodKey,
  getCoverage, getMissingVaccines, getDatabaseSummary, findChildById,
} from '../vaccineDatabase';
import { createEmptyMasterData } from '../asikParser';
import { dateStringToExcelSerial } from '../dateUtils';
import { childKey } from '../childIdentity';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../../types';

function child(nama: string, tglLahir: string, vaccines: Partial<Record<VaccineKey, string>> = {}, over: Partial<ChildRecord> = {}): ChildRecord {
  const v: Partial<Record<VaccineKey, number>> = {};
  for (const [k, d] of Object.entries(vaccines)) v[k as VaccineKey] = dateStringToExcelSerial(d as string);
  return {
    nama, jk: 'L',
    tanggalLahirSerial: dateStringToExcelSerial(tglLahir),
    tanggalLahirStr: tglLahir,
    nik: '', namaOrangTua: '', alamat: 'Mabuun',
    vaccines: v,
    ...over,
  };
}

function md(entries: Partial<Record<SheetName, ChildRecord[]>>): MasterData {
  const base = createEmptyMasterData();
  for (const [s, rows] of Object.entries(entries)) base[s as SheetName] = rows as ChildRecord[];
  return base;
}

describe('periodKey', () => {
  it('memformat bulan jadi dua digit', () => {
    expect(periodKey(6, 2026)).toBe('2026-06');
    expect(periodKey(12, 2026)).toBe('2026-12');
  });
});

describe('mergeMonthlyIntoDatabase', () => {
  it('menambah anak baru dari periode pertama', () => {
    const { database, stats } = mergeMonthlyIntoDatabase(
      createEmptyDatabase(),
      md({ MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-05' })] }),
      6, 2026,
    );
    expect(stats.childrenAdded).toBe(1);
    expect(stats.vaccinesAdded).toBe(1);
    expect(database.children).toHaveLength(1);
    expect(database.periods).toEqual(['2026-06']);
    expect(database.children[0].firstSeen).toBe('2026-06');
    expect(database.children[0].vaccineSources.DPT_1).toBe('2026-06');
  });

  it('MENUMPUK vaksin lintas bulan untuk anak yang sama, bukan menduplikasi', () => {
    const juni = mergeMonthlyIntoDatabase(
      createEmptyDatabase(),
      md({ MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-05' })] }),
      6, 2026,
    );
    const juli = mergeMonthlyIntoDatabase(
      juni.database,
      md({ MABUUN: [child('Budi', '2025-01-10', { DPT_2: '2026-07-09' })] }),
      7, 2026,
    );

    expect(juli.database.children).toHaveLength(1);
    expect(juli.stats.childrenAdded).toBe(0);
    expect(juli.stats.vaccinesAdded).toBe(1);
    const budi = juli.database.children[0];
    expect(budi.vaccines.DPT_1).toBe(dateStringToExcelSerial('2026-06-05'));
    expect(budi.vaccines.DPT_2).toBe(dateStringToExcelSerial('2026-07-09'));
    expect(budi.firstSeen).toBe('2026-06');
    expect(budi.lastSeen).toBe('2026-07');
    expect(juli.database.periods).toEqual(['2026-06', '2026-07']);
  });

  it('mencatat konflik kalau tanggal vaksin yang sama berubah', () => {
    const a = mergeMonthlyIntoDatabase(
      createEmptyDatabase(),
      md({ MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-05' })] }),
      6, 2026,
    );
    const b = mergeMonthlyIntoDatabase(
      a.database,
      md({ MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-07' })] }),
      7, 2026,
    );
    expect(b.stats.conflicts).toHaveLength(1);
    expect(b.stats.conflicts[0].vaccine).toBe('DPT_1');
    expect(b.stats.conflicts[0].existingPeriod).toBe('2026-06');
    expect(b.stats.conflicts[0].incomingPeriod).toBe('2026-07');
    // koreksi dipakai
    expect(b.database.children[0].vaccines.DPT_1).toBe(dateStringToExcelSerial('2026-06-07'));
  });

  it('memakai kunci identitas yang sama dengan merger bulanan', () => {
    const { database } = mergeMonthlyIntoDatabase(
      createEmptyDatabase(),
      md({ MABUUN: [child('  BUDI   Santoso ', '2025-01-10')] }),
      6, 2026,
    );
    const id = childKey('budi santoso', '2025-01-10');
    expect(findChildById(database, id)).toBeDefined();
  });

  it('mencatat perpindahan wilayah', () => {
    const a = mergeMonthlyIntoDatabase(
      createEmptyDatabase(),
      md({ MABUUN: [child('Budi', '2025-01-10')] }), 6, 2026,
    );
    const b = mergeMonthlyIntoDatabase(
      a.database,
      md({ KASIAU: [child('Budi', '2025-01-10')] }), 7, 2026,
    );
    expect(b.stats.sheetChanges).toBe(1);
    expect(b.database.children[0].sheet).toBe('KASIAU');
    expect(b.database.children).toHaveLength(1);
  });

  it('mengisi identitas kosong tanpa menimpa yang sudah terisi', () => {
    const a = mergeMonthlyIntoDatabase(
      createEmptyDatabase(),
      md({ MABUUN: [child('Budi', '2025-01-10', {}, { nik: '1234', namaOrangTua: 'Siti' })] }),
      6, 2026,
    );
    const b = mergeMonthlyIntoDatabase(
      a.database,
      md({ MABUUN: [child('Budi', '2025-01-10', {}, { nik: '', namaOrangTua: '' })] }),
      7, 2026,
    );
    expect(b.database.children[0].nik).toBe('1234');
    expect(b.database.children[0].namaOrangTua).toBe('Siti');
  });

  it('menandai upload ulang periode yang sama', () => {
    const a = mergeMonthlyIntoDatabase(createEmptyDatabase(), md({ MABUUN: [child('Budi', '2025-01-10')] }), 6, 2026);
    const b = mergeMonthlyIntoDatabase(a.database, md({ MABUUN: [child('Budi', '2025-01-10')] }), 6, 2026);
    expect(b.stats.alreadyMerged).toBe(true);
    expect(b.database.periods).toEqual(['2026-06']);
    expect(b.database.children).toHaveLength(1);
  });

  it('tidak memutasi database lama', () => {
    const a = mergeMonthlyIntoDatabase(createEmptyDatabase(), md({ MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-05' })] }), 6, 2026);
    const before = JSON.stringify(a.database);
    mergeMonthlyIntoDatabase(a.database, md({ MABUUN: [child('Budi', '2025-01-10', { DPT_2: '2026-07-09' })] }), 7, 2026);
    expect(JSON.stringify(a.database)).toBe(before);
  });

  it('mengurutkan periode meski digabung tidak berurutan', () => {
    const a = mergeMonthlyIntoDatabase(createEmptyDatabase(), md({ MABUUN: [child('Budi', '2025-01-10')] }), 8, 2026);
    const b = mergeMonthlyIntoDatabase(a.database, md({ MABUUN: [child('Budi', '2025-01-10')] }), 6, 2026);
    expect(b.database.periods).toEqual(['2026-06', '2026-08']);
    expect(b.database.children[0].firstSeen).toBe('2026-06');
    expect(b.database.children[0].lastSeen).toBe('2026-08');
  });
});

describe('query helpers', () => {
  it('menghitung cakupan per vaksin dipecah L/P', () => {
    const { database } = mergeMonthlyIntoDatabase(createEmptyDatabase(), md({
      MABUUN: [
        child('Budi', '2025-01-10', { DPT_1: '2026-06-05' }, { jk: 'L' }),
        child('Ani', '2025-02-11', { DPT_1: '2026-06-06' }, { jk: 'P' }),
        child('Cici', '2025-03-12', {}, { jk: 'P' }),
      ],
    }), 6, 2026);
    const cov = getCoverage(database);
    expect(cov.DPT_1).toEqual({ L: 1, P: 1, total: 2 });
    expect(cov.BCG).toEqual({ L: 0, P: 0, total: 0 });
  });

  it('mendaftar vaksin yang belum tercatat', () => {
    const { database } = mergeMonthlyIntoDatabase(createEmptyDatabase(),
      md({ MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-05' })] }), 6, 2026);
    const missing = getMissingVaccines(database.children[0]);
    expect(missing).not.toContain('DPT_1');
    expect(missing).toContain('DPT_2');
  });

  it('meringkas isi database', () => {
    const { database } = mergeMonthlyIntoDatabase(createEmptyDatabase(), md({
      MABUUN: [child('Budi', '2025-01-10', { DPT_1: '2026-06-05', DPT_2: '2026-06-06' })],
      KASIAU: [child('Ani', '2025-02-11', { BCG: '2026-06-07' })],
    }), 6, 2026);
    const s = getDatabaseSummary(database);
    expect(s.totalChildren).toBe(2);
    expect(s.totalVaccineRecords).toBe(3);
    expect(s.perSheet.MABUUN).toBe(1);
    expect(s.perSheet.KASIAU).toBe(1);
    expect(s.periods).toEqual(['2026-06']);
  });
});

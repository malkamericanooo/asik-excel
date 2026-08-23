import { describe, it, expect, beforeEach } from 'vitest';
import {
  STORAGE_KEY, loadDatabase, saveDatabase, clearDatabase,
  parseDatabase, databaseFileName,
} from '../vaccineDatabaseStorage';
import { createEmptyDatabase, mergeMonthlyIntoDatabase } from '../vaccineDatabase';
import { createEmptyMasterData } from '../asikParser';
import { dateStringToExcelSerial } from '../dateUtils';
import type { ChildRecord } from '../../types';

/** localStorage tiruan — test jalan di environment node. */
class MemoryStorage implements Storage {
  private map = new Map<string, string>();
  limitBytes = Infinity;
  get length() { return this.map.size; }
  clear() { this.map.clear(); }
  key(i: number) { return [...this.map.keys()][i] ?? null; }
  getItem(k: string) { return this.map.get(k) ?? null; }
  removeItem(k: string) { this.map.delete(k); }
  setItem(k: string, v: string) {
    if (v.length > this.limitBytes) throw new DOMException('QuotaExceededError');
    this.map.set(k, v);
  }
}

function sampleChild(nama = 'Budi'): ChildRecord {
  return {
    nama, jk: 'L',
    tanggalLahirSerial: dateStringToExcelSerial('2025-01-10'),
    tanggalLahirStr: '2025-01-10',
    nik: '3201', namaOrangTua: 'Siti', alamat: 'Mabuun',
    vaccines: { DPT_1: dateStringToExcelSerial('2026-06-05') },
  };
}

function sampleDb() {
  const master = createEmptyMasterData();
  master.MABUUN = [sampleChild()];
  return mergeMonthlyIntoDatabase(createEmptyDatabase(), master, 6, 2026).database;
}

let storage: MemoryStorage;
beforeEach(() => { storage = new MemoryStorage(); });

describe('save / load round-trip', () => {
  it('menyimpan lalu membaca database yang sama', () => {
    const db = sampleDb();
    const saved = saveDatabase(db, storage);
    expect(saved.ok).toBe(true);

    const loaded = loadDatabase(storage);
    expect(loaded.error).toBeUndefined();
    expect(loaded.database.children).toHaveLength(1);
    expect(loaded.database.children[0].nama).toBe('Budi');
    expect(loaded.database.children[0].vaccines.DPT_1).toBe(dateStringToExcelSerial('2026-06-05'));
    expect(loaded.database.periods).toEqual(['2026-06']);
  });

  it('mengembalikan database kosong kalau belum ada yang tersimpan', () => {
    const loaded = loadDatabase(storage);
    expect(loaded.database.children).toEqual([]);
    expect(loaded.error).toBeUndefined();
  });

  it('tidak crash kalau data tersimpan rusak', () => {
    storage.setItem(STORAGE_KEY, '{ ini bukan json');
    const loaded = loadDatabase(storage);
    expect(loaded.database.children).toEqual([]);
    expect(loaded.error).toMatch(/rusak/i);
  });

  it('melaporkan error kalau penyimpanan penuh', () => {
    storage.limitBytes = 10;
    const res = saveDatabase(sampleDb(), storage);
    expect(res.ok).toBe(false);
    expect(res.error).toMatch(/terlalu besar/i);
  });

  it('clearDatabase menghapus data tersimpan', () => {
    saveDatabase(sampleDb(), storage);
    clearDatabase(storage);
    expect(loadDatabase(storage).database.children).toEqual([]);
  });
});

describe('parseDatabase — file impor tidak dipercaya begitu saja', () => {
  it('menolak yang bukan objek', () => {
    expect(parseDatabase('halo').error).toBeTruthy();
    expect(parseDatabase(null).error).toBeTruthy();
  });

  it('menolak file tanpa schemaVersion', () => {
    expect(parseDatabase({ children: [], periods: [] }).error).toMatch(/schemaVersion/);
  });

  it('menolak schema dari aplikasi versi lebih baru', () => {
    const res = parseDatabase({ schemaVersion: 99, children: [], periods: [] });
    expect(res.error).toMatch(/versi lebih baru/);
    expect(res.database).toBeUndefined();
  });

  it('membuang baris yang bentuknya salah dan memberi tahu', () => {
    const good = sampleDb().children[0];
    const res = parseDatabase({
      schemaVersion: 1,
      updatedAt: new Date().toISOString(),
      periods: ['2026-06'],
      children: [good, { nama: 'tanpa id' }, 42, null],
    });
    expect(res.database?.children).toHaveLength(1);
    expect(res.error).toMatch(/3 baris dilewati/);
  });

  it('menolak sheet di luar daftar wilayah yang sah', () => {
    const bad = { ...sampleDb().children[0], sheet: 'WILAYAH PALSU' };
    const res = parseDatabase({ schemaVersion: 1, periods: [], children: [bad] });
    expect(res.error).toMatch(/tidak sesuai/);
  });

  it('menerima database yang sah dan mengurutkan periode', () => {
    const db = sampleDb();
    const res = parseDatabase({ ...db, periods: ['2026-08', '2026-06'] });
    expect(res.error).toBeUndefined();
    expect(res.database?.periods).toEqual(['2026-06', '2026-08']);
  });
});

describe('databaseFileName', () => {
  it('memakai periode terakhir', () => {
    expect(databaseFileName(sampleDb())).toBe('database-imunisasi-kumulatif_2026-06.json');
  });
  it('menangani database kosong', () => {
    expect(databaseFileName(createEmptyDatabase())).toBe('database-imunisasi-kumulatif_kosong.json');
  });
});

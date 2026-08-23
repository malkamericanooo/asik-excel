import { describe, it, expect } from 'vitest';
import { sortChildrenByBirthDate, sortChildrenGroups } from '../sortChildren';
import { dateStringToExcelSerial } from '../dateUtils';
import type { ChildRecord } from '../../types';

function child(nama: string, tglLahir: string, jk: 'L' | 'P' = 'L'): ChildRecord {
  return {
    nama,
    jk,
    tanggalLahirSerial: tglLahir ? dateStringToExcelSerial(tglLahir) : 0,
    tanggalLahirStr: tglLahir,
    nik: '',
    namaOrangTua: '',
    alamat: 'Mabuun',
    vaccines: {},
  };
}

const names = (rows: ChildRecord[]) => rows.map((r) => r.nama);

describe('sortChildrenByBirthDate', () => {
  it('mengurutkan tanggal lahir kronologis, paling tua di atas', () => {
    const rows = [
      child('Citra', '2023-05-10'),
      child('Ayu', '2021-12-31'),
      child('Budi', '2022-01-01'),
    ];
    expect(names(sortChildrenByBirthDate(rows))).toEqual(['Ayu', 'Budi', 'Citra']);
  });

  it('mengurutkan berdasarkan tanggal penuh, bukan angka hari saja', () => {
    // Kalau hanya angka hari yang dipakai, "Lahir tgl 01" akan naik ke atas.
    const rows = [
      child('Lahir 05 Januari 2020', '2020-01-05'),
      child('Lahir 01 Desember 2023', '2023-12-01'),
      child('Lahir 28 Juni 2021', '2021-06-28'),
    ];
    expect(names(sortChildrenByBirthDate(rows))).toEqual([
      'Lahir 05 Januari 2020',
      'Lahir 28 Juni 2021',
      'Lahir 01 Desember 2023',
    ]);
  });

  it('tanggal lahir sama diurutkan nama A→Z, tanpa peduli besar-kecil huruf', () => {
    const rows = [
      child('zulfa', '2022-03-04'),
      child('Ahmad', '2022-03-04'),
      child('budi', '2022-03-04'),
    ];
    expect(names(sortChildrenByBirthDate(rows))).toEqual(['Ahmad', 'budi', 'zulfa']);
  });

  it('menaruh tanggal lahir kosong paling atas', () => {
    const rows = [
      child('Punya Tanggal', '2022-03-04'),
      child('Tanpa Tanggal', ''),
      child('Punya Tanggal Lama', '2019-01-01'),
    ];
    expect(names(sortChildrenByBirthDate(rows))).toEqual([
      'Tanpa Tanggal',
      'Punya Tanggal Lama',
      'Punya Tanggal',
    ]);
  });

  it('beberapa tanggal kosong tetap diurut nama A→Z di antara mereka', () => {
    const rows = [child('Sinta', ''), child('Rina', ''), child('Ali', '2020-02-02')];
    expect(names(sortChildrenByBirthDate(rows))).toEqual(['Rina', 'Sinta', 'Ali']);
  });

  it('tidak memutasi array asli', () => {
    const rows = [child('Citra', '2023-05-10'), child('Ayu', '2021-12-31')];
    const before = names(rows);
    sortChildrenByBirthDate(rows);
    expect(names(rows)).toEqual(before);
  });

  it('stabil — dua kali sort menghasilkan urutan identik', () => {
    const rows = [
      child('Dewi', '2022-01-01'),
      child('Dewi', '2022-01-01'),
      child('Candra', '2020-06-06'),
    ];
    const once = sortChildrenByBirthDate(rows);
    expect(names(sortChildrenByBirthDate(once))).toEqual(names(once));
  });

  it('daftar kosong tidak error', () => {
    expect(sortChildrenByBirthDate([])).toEqual([]);
  });
});

describe('sortChildrenGroups', () => {
  it('mengurutkan tiap sheet secara terpisah, urutan sheet tetap', () => {
    const mabuun = [child('Bayu', '2023-01-01'), child('Anisa', '2021-01-01')];
    const kasiau = [child('Dinda', '2024-01-01'), child('Cahya', '2020-01-01')];
    expect(names(sortChildrenGroups([mabuun, kasiau]))).toEqual([
      'Anisa', 'Bayu', 'Cahya', 'Dinda',
    ]);
  });
});

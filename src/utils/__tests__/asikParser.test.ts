import { describe, it, expect, beforeEach } from 'vitest';
import * as XLSX from 'xlsx';
import { parseAndMergeAsikFile, createEmptyMasterData } from '../asikParser';
import type { MasterData, ProcessResult } from '../../types';

function makeAsikBuffer(rows: Record<string, string | number>[]) {
  const header = [
    'ID', 'NIK Anak', 'Nama Anak', 'Tanggal Lahir Anak', 'Jenis Kelamin Anak',
    'NIK Orang Tua', 'Nama Orang Tua', 'Provinsi', 'Kabupaten atau Kota',
    'Kecamatan', 'Kelurahan atau Desa', 'Kode Puskesmas', 'Puskesmas',
    'Nama Antigen', 'Tanggal Imunisasi', 'Tanggal Input Data', 'Nomor Batch',
    'Pos Imunisasi', 'Status Imunisasi', 'Sumber Pencatatan Imunisasi',
  ];
  const aoa = [header, ...rows.map((r) => header.map((h) => r[h] ?? ''))];
  const ws = XLSX.utils.aoa_to_sheet(aoa);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
}

function makeRow(overrides: Partial<Record<string, string | number>> = {}): Record<string, string | number> {
  return {
    ID: 1, 'NIK Anak': '0000000000000000', 'Nama Anak': 'Budi Santoso',
    'Tanggal Lahir Anak': '2026-04-01', 'Jenis Kelamin Anak': 'Laki-laki',
    'NIK Orang Tua': '0000000000000000', 'Nama Orang Tua': 'Siti Aminah',
    Provinsi: 'KALIMANTAN SELATAN', 'Kabupaten atau Kota': 'KAB. TABALONG',
    Kecamatan: 'MURUNG PUDAK', "Kelurahan atau Desa": "MABU'UN",
    'Kode Puskesmas': '1060598', Puskesmas: 'MABUUN',
    'Nama Antigen': 'DPT-HB-Hib - 1', 'Tanggal Imunisasi': '2026-06-17',
    'Tanggal Input Data': '2026-06-17', 'Nomor Batch': '0',
    'Pos Imunisasi': 'PUSKESMAS MABUUN DALAM GEDUNG',
    'Status Imunisasi': 'ideal', 'Sumber Pencatatan Imunisasi': 'Imunisasi Langsung',
    ...overrides,
  };
}

let masterData: MasterData;
let result: ProcessResult;

beforeEach(() => {
  masterData = createEmptyMasterData();
  result = { added: 0, updated: 0, moved: 0, skipped: 0, logs: [] };
});

describe('new child', () => {
  it('adds to MABUUN sheet with correct name and gender', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow()]), masterData, result);
    expect(masterData.MABUUN).toHaveLength(1);
    expect(result.added).toBe(1);
    expect(masterData.MABUUN[0].nama).toBe('Budi Santoso');
    expect(masterData.MABUUN[0].jk).toBe('L');
  });
  it('assigns correct vaccine date (DPT_1)', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow()]), masterData, result);
    expect(masterData.MABUUN[0].vaccines.DPT_1).toBeGreaterThan(0);
  });
  it('maps Perempuan gender to P', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Jenis Kelamin Anak': 'Perempuan', 'Nama Anak': 'Sari' })]), masterData, result);
    expect(masterData.MABUUN[0].jk).toBe('P');
  });
  it('routes Status Kejar to Kejar sheet', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Status Imunisasi': 'kejar' })]), masterData, result);
    expect(masterData.Kejar).toHaveLength(1);
    expect(masterData.MABUUN).toHaveLength(0);
  });
  it('routes unknown kelurahan to LUAR WILAYAH', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Kelurahan atau Desa': 'TANTA HULU' })]), masterData, result);
    expect(masterData['LUAR WILAYAH']).toHaveLength(1);
  });
  it('routes KASIAU kelurahan to KASIAU sheet', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Kelurahan atau Desa': 'KASIAU' })]), masterData, result);
    expect(masterData.KASIAU).toHaveLength(1);
  });
});

describe('deduplication', () => {
  it('updates existing child without creating duplicate', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Antigen': 'DPT-HB-Hib - 1', 'Tanggal Imunisasi': '2026-06-03' })]), masterData, result);
    result = { added: 0, updated: 0, moved: 0, skipped: 0, logs: [] };
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Antigen': 'DPT-HB-Hib - 2', 'Tanggal Imunisasi': '2026-06-17' })]), masterData, result);
    expect(masterData.MABUUN).toHaveLength(1);
    expect(result.updated).toBe(1);
    expect(masterData.MABUUN[0].vaccines.DPT_1).toBeGreaterThan(0);
    expect(masterData.MABUUN[0].vaccines.DPT_2).toBeGreaterThan(0);
  });
  it('dedup is case-insensitive for Nama Anak', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Anak': 'budi santoso' })]), masterData, result);
    result = { added: 0, updated: 0, moved: 0, skipped: 0, logs: [] };
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Anak': 'BUDI SANTOSO', 'Nama Antigen': 'IPV - 1' })]), masterData, result);
    expect(masterData.MABUUN).toHaveLength(1);
    expect(result.updated).toBe(1);
  });
  it('moves child from area sheet to Kejar when status changes', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Status Imunisasi': 'ideal' })]), masterData, result);
    result = { added: 0, updated: 0, moved: 0, skipped: 0, logs: [] };
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Status Imunisasi': 'kejar', 'Nama Antigen': 'IPV - 1' })]), masterData, result);
    expect(masterData.MABUUN).toHaveLength(0);
    expect(masterData.Kejar).toHaveLength(1);
    expect(result.moved).toBe(1);
  });
  it('two different children are not deduped', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Anak': 'Budi Santoso' }), makeRow({ 'Nama Anak': 'Ahmad Fauzi', ID: 2 })]), masterData, result);
    expect(masterData.MABUUN).toHaveLength(2);
    expect(result.added).toBe(2);
  });
});

describe('validation', () => {
  it('throws on file missing required column Nama Antigen', () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ['Nama Anak', 'Tanggal Lahir Anak', 'Nama Orang Tua', 'Tanggal Imunisasi'],
      ['Budi', '2026-04-01', 'Siti', '2026-06-01'],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' }) as ArrayBuffer;
    expect(() => parseAndMergeAsikFile(buf, masterData, result)).toThrow(/Nama Antigen/);
  });
  it('skips rows with unknown antigen', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Antigen': 'ANTIGEN TIDAK DIKENAL' })]), masterData, result);
    expect(result.skipped).toBe(1);
    expect(result.added).toBe(0);
  });
  it('skips rows with missing Nama Anak', () => {
    parseAndMergeAsikFile(makeAsikBuffer([makeRow({ 'Nama Anak': '' })]), masterData, result);
    expect(result.skipped).toBe(1);
  });
});

import { describe, it, expect, beforeAll, beforeEach } from 'vitest';
import * as XLSX from 'xlsx';
import { readFileSync } from 'fs';
import { resolve } from 'path';
import { buildMasterExcel, findSummaryStartRow, getUploadedVaccines } from '../masterExcel';
import { createEmptyMasterData } from '../asikParser';
import { dateStringToExcelSerial } from '../dateUtils';
import type { MasterData, ChildRecord } from '../../types';
import { ALL_SHEETS } from '../../types';

const TEMPLATE_PATH = resolve(__dirname, '../../../public/templates/master-template.xlsx');
let templateBuffer: ArrayBuffer;

beforeAll(() => {
  const buf = readFileSync(TEMPLATE_PATH);
  templateBuffer = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
});

function makeChild(overrides: Partial<ChildRecord> = {}): ChildRecord {
  return {
    nama: 'Test Anak', jk: 'L',
    tanggalLahirSerial: dateStringToExcelSerial('2026-04-01'),
    tanggalLahirStr: '2026-04-01', nik: '0000000000000000',
    namaOrangTua: 'Test Ortu', alamat: 'MABUUN', vaccines: {},
    ...overrides,
  };
}

let masterData: MasterData;
beforeEach(() => { masterData = createEmptyMasterData(); });

describe('buildMasterExcel', () => {
  it('returns a Blob', async () => {
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    expect(blob).toBeInstanceOf(Blob);
    expect(blob.size).toBeGreaterThan(0);
  });

  it('produces 7 sheets in correct order', async () => {
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    expect(wb.SheetNames).toEqual(ALL_SHEETS);
  });

  it('title row text is correct', async () => {
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['MABUUN'];
    expect(ws['A1']?.v).toBe('LAPORAN BULANAN HASIL IMUNISASI RUTIN BAYI PUSKESMAS');
    expect(ws['C3']?.v).toBe(': Juni 2026');
  });

  it('data row written correctly for a male child', async () => {
    masterData.MABUUN.push(makeChild({ vaccines: { DPT_1: dateStringToExcelSerial('2026-06-17') } }));
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['MABUUN'];
    expect(ws['A7']?.v).toBe(1);
    expect(ws['B7']?.v).toBe('Test Anak');
    expect(ws['C7']?.v).toBe('L');
    expect(ws['P7']?.v).toBeGreaterThan(0);
  });

  it('female child DPT_1 goes to P column (Q7)', async () => {
    masterData.KASIAU.push(makeChild({ jk: 'P', alamat: 'KASIAU', vaccines: { DPT_1: dateStringToExcelSerial('2026-06-17') } }));
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['KASIAU'];
    expect(ws['P7']?.v).toBeFalsy();
    expect(ws['Q7']?.v).toBeGreaterThan(0);
  });

  it('summary counts only dates in the target month', async () => {
    masterData.MABUUN.push(makeChild({ vaccines: { BCG: dateStringToExcelSerial('2026-06-17') } }));
    masterData.MABUUN.push(makeChild({ nama: 'Anak Dua', vaccines: { BCG: dateStringToExcelSerial('2026-05-10') } }));
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['MABUUN'];
    const summaryStart = findSummaryStartRow(ws);
    const countRow = summaryStart + 2;
    expect(ws[XLSX.utils.encode_cell({ r: countRow, c: 11 })]?.v).toBe(1);
    expect(ws[XLSX.utils.encode_cell({ r: countRow + 1, c: 11 })]?.v).toBe(1);
  });

  it('empty sheet still has header rows', async () => {
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['KASIAU'];
    expect(ws['A5']?.v).toBe('NO');
    expect(ws['H6']?.v).toBe('L');
    expect(ws['H5']?.v).toBe('HB0 (<24 JAM)');
  });

  it('summary block follows data with blank separator row', async () => {
    masterData.MABUUN.push(makeChild({ vaccines: { DPT_1: dateStringToExcelSerial('2026-06-17') } }));
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['MABUUN'];
    const summaryStart = findSummaryStartRow(ws);
    expect(summaryStart).toBe(8);
    expect(ws[XLSX.utils.encode_cell({ r: summaryStart, c: 6 })]?.v).toBe('Jumlah Imunisasi Bulan Juni 2026');
    expect(ws[XLSX.utils.encode_cell({ r: summaryStart, c: 7 })]?.v).toBe('HB0 (<24 JAM)');
  });

  it('exports successfully with only one vaccine uploaded (partial)', async () => {
    masterData.MABUUN.push(makeChild({ vaccines: { BCG: dateStringToExcelSerial('2026-06-17') } }));
    await expect(buildMasterExcel(masterData, 6, 2026, templateBuffer)).resolves.not.toThrow();
  });

  it('output file size is reasonable', async () => {
    masterData.MABUUN.push(makeChild({ vaccines: { DPT_1: dateStringToExcelSerial('2026-06-17') } }));
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    expect(blob.size).toBeGreaterThan(0);
  });

  it('columns are within reasonable range in every sheet', async () => {
    const blob = await buildMasterExcel(masterData, 6, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    for (const sheetName of ALL_SHEETS) {
      const ws = wb.Sheets[sheetName];
      if (!ws || !ws['!ref']) continue;
      const range = XLSX.utils.decode_range(ws['!ref']);
      // exceljs may produce larger column ranges due to template, just check it's not insane
      expect(range.e.c).toBeLessThanOrEqual(100);
    }
  });

  it('summary label includes month and year', async () => {
    const blob = await buildMasterExcel(masterData, 3, 2026, templateBuffer);
    const buf = await blob.arrayBuffer();
    const ws = XLSX.read(buf, { type: 'array' }).Sheets['KASIAU'];
    const summaryStart = findSummaryStartRow(ws);
    const label = ws[XLSX.utils.encode_cell({ r: summaryStart, c: 6 })]?.v;
    expect(label).toBe('Jumlah Imunisasi Bulan Maret 2026');
  });
});

describe('getUploadedVaccines', () => {
  it('returns empty set for empty data', () => expect(getUploadedVaccines(masterData).size).toBe(0));
  it('detects BCG and DPT_1 when present', () => {
    masterData.MABUUN.push(makeChild({ vaccines: { BCG: 46190, DPT_1: 46190 } }));
    const u = getUploadedVaccines(masterData);
    expect(u.has('BCG')).toBe(true);
    expect(u.has('DPT_1')).toBe(true);
    expect(u.has('POLIO_1')).toBe(false);
  });
  it('detects across multiple sheets', () => {
    masterData.MABUUN.push(makeChild({ vaccines: { BCG: 46190 } }));
    masterData.KASIAU.push(makeChild({ vaccines: { IPV_1: 46190 } }));
    const u = getUploadedVaccines(masterData);
    expect(u.has('BCG')).toBe(true);
    expect(u.has('IPV_1')).toBe(true);
  });
  it('does not count zero vaccine values', () => {
    masterData.MABUUN.push(makeChild({ vaccines: { BCG: 0 } }));
    expect(getUploadedVaccines(masterData).has('BCG')).toBe(false);
  });
});

import * as XLSX from 'xlsx';
import type { ChildRecord, MasterData, ProcessResult, SheetName, VaccineKey } from '../types';
import { classifyKelurahan } from './kelurahanMapping';
import { dateStringToExcelSerial, parseDateCell } from './dateUtils';
import { resolveVaccineKey } from './vaccineResolver';

function normalizeKey(s: string): string {
  return (s || '').toLowerCase().trim().replace(/\s+/g, ' ');
}

function childKey(nama: string, tglLahir: string, namaOrtu: string): string {
  return `${normalizeKey(nama)}|${tglLahir}|${normalizeKey(namaOrtu)}`;
}

function findChild(
  masterData: MasterData,
  key: string,
): { sheet: SheetName; index: number } | null {
  for (const [sheetName, rows] of Object.entries(masterData)) {
    const idx = rows.findIndex(
      (r) => childKey(r.nama, r.tanggalLahirStr, r.namaOrangTua) === key,
    );
    if (idx !== -1) return { sheet: sheetName as SheetName, index: idx };
  }
  return null;
}

function formatAlamat(kelurahanRaw: string, isKejar: boolean): string {
  if (isKejar) {
    return kelurahanRaw ? `Kejar - ${kelurahanRaw}` : 'Kejar';
  }
  return kelurahanRaw || 'Luar Wilayah';
}

export interface ParseOptions {
  /** Used when input file has no Nama Antigen column (simple format). */
  selectedVaccine?: VaccineKey;
}

/** Parse a single ASIK Excel file and merge into masterData in-place */
export function parseAndMergeAsikFile(
  fileBuffer: ArrayBuffer,
  masterData: MasterData,
  result: ProcessResult,
  options: ParseOptions = {},
): void {
  const wb = XLSX.read(fileBuffer, { type: 'array', cellDates: false });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    defval: null,
  }) as unknown[][];

  if (rows.length < 2) {
    result.logs.push('File kosong atau tidak memiliki data.');
    return;
  }

  const headerRow = rows[0] as string[];
  const colIdx: Record<string, number> = {};
  headerRow.forEach((h, i) => {
    if (h) colIdx[String(h).trim().toLowerCase()] = i;
  });

  // Normalize search keys to lowercase for fuzzy matching
  function findColHeuristic(candidates: string[]): number {
    for (const c of candidates) {
      const key = c.toLowerCase();
      if (key in colIdx) return colIdx[key];
    }
    // Fuzzy search: try partial match if exact failed
    for (const c of candidates) {
      const key = c.toLowerCase();
      for (const [colKey, colIdxVal] of Object.entries(colIdx)) {
        if (colKey.includes(key) || key.includes(colKey)) return colIdxVal;
      }
    }
    return -1;
  }

  const col = {
    nama: findColHeuristic(['Nama Anak', 'NAMA ANAK', 'Nama']),
    tglLahir: findColHeuristic(['Tanggal Lahir Anak', 'Tanggal Lahir', 'TGL LAHIR', 'TTL']),
    jk: findColHeuristic(['Jenis Kelamin Anak', 'Jenis Kelamin', 'JK', 'KELAMIN']),
    namaOrtu: findColHeuristic(['Nama Orang Tua', 'NAMA ORANG TUA', 'Ortu', 'ORANG TUA']),
    kelurahan: findColHeuristic([
      'Klasifikasi Kelurahan',
      'Kelurahan atau Desa',
      'Kelurahan',
      'KELURAHAN',
      'DESA',
      'Alamat',
    ]),
    antigen: findColHeuristic(['Nama Antigen', 'ANTIGEN', 'Antigen', 'Jenis Vaksin', 'Vaksin']),
    tglImunisasi: findColHeuristic(['Tanggal Imunisasi', 'TGL IMUNISASI', 'Tgl Imunisasi', 'TANGGAL']),
    status: findColHeuristic(['Status Imunisasi', 'Status', 'STATUS']),
    nik: findColHeuristic(['NIK Anak', 'NIK', 'nik', 'N I K']),
  };

  const missing: string[] = [];
  if (col.nama === -1) missing.push('Nama Anak');
  if (col.tglLahir === -1) missing.push('Tanggal Lahir');
  if (col.namaOrtu === -1) missing.push('Nama Orang Tua');
  if (col.tglImunisasi === -1) missing.push('Tanggal Imunisasi');
  if (col.antigen === -1 && !options.selectedVaccine) {
    missing.push('Nama Antigen (atau pilih jenis vaksin di dropdown)');
  }
  if (missing.length > 0) {
    throw new Error(
      `Kolom wajib tidak ditemukan: ${missing.join(', ')}. Pastikan file dari ASIK yang benar.`,
    );
  }

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i] as unknown[];
    if (!row || row.every((c) => c == null || c === '')) continue;

    const nama = String(row[col.nama] ?? '').trim();
    const tglLahirRaw = parseDateCell(row[col.tglLahir]);
    const namaOrtu = String(row[col.namaOrtu] ?? '').trim();
    const antigenRaw = col.antigen >= 0 ? String(row[col.antigen] ?? '').trim() : '';
    const tglImunisasiRaw = parseDateCell(row[col.tglImunisasi]);
    const statusRaw = col.status >= 0
      ? String(row[col.status] ?? 'ideal').toLowerCase().trim()
      : 'ideal';
    const kelurahanRaw = col.kelurahan >= 0 ? String(row[col.kelurahan] ?? '').trim() : '';
    const jkRaw = col.jk >= 0 ? String(row[col.jk] ?? '').trim() : '';
    const nikRaw = col.nik >= 0 ? String(row[col.nik] ?? '').replace(/^'/, '').trim() : '';

    if (!nama || !tglLahirRaw || !namaOrtu) {
      result.skipped++;
      result.logs.push(`Baris ${i + 1}: data identitas tidak lengkap, dilewati.`);
      continue;
    }

    // NIK kadang diisi placeholder (000...), jadi hanya warning jika ada karakter non-digit
    if (nikRaw && /[^0-9]/.test(nikRaw)) {
      result.logs.push(`Baris ${i + 1}: NIK "${nikRaw}" mengandung karakter tidak dikenal, akan tetap diproses.`);
    }

    const vaccineKey = resolveVaccineKey(antigenRaw, options.selectedVaccine);
    if (!vaccineKey) {
      result.skipped++;
      result.logs.push(`Baris ${i + 1}: antigen "${antigenRaw}" tidak dikenal, dilewati.`);
      continue;
    }

    const tglImunisasiSerial = dateStringToExcelSerial(tglImunisasiRaw);
    const tglLahirSerial = dateStringToExcelSerial(tglLahirRaw);
    const jk: 'L' | 'P' = jkRaw.toLowerCase().includes('per') ? 'P' : 'L';

    const isKejar = statusRaw === 'kejar';
    const targetSheet: SheetName = isKejar ? 'Kejar' : classifyKelurahan(kelurahanRaw);
    const alamat = formatAlamat(kelurahanRaw, isKejar);

    const key = childKey(nama, tglLahirRaw, namaOrtu);
    const existing = findChild(masterData, key);

    if (existing) {
      const { sheet: existingSheet, index } = existing;
      const child = masterData[existingSheet][index];
      child.vaccines[vaccineKey] = tglImunisasiSerial;

      if (existingSheet !== targetSheet) {
        masterData[existingSheet].splice(index, 1);
        child.alamat = alamat;
        masterData[targetSheet].push(child);
        result.moved++;
        result.logs.push(`Dipindahkan: "${nama}" dari ${existingSheet} → ${targetSheet}`);
      } else {
        child.alamat = alamat;
        result.updated++;
      }
    } else {
      const newChild: ChildRecord = {
        nama,
        jk,
        tanggalLahirSerial: tglLahirSerial,
        tanggalLahirStr: tglLahirRaw,
        nik: nikRaw,
        namaOrangTua: namaOrtu,
        alamat,
        vaccines: { [vaccineKey]: tglImunisasiSerial },
      };
      masterData[targetSheet].push(newChild);
      result.added++;
    }
  }
}

function findCol(colIdx: Record<string, number>, candidates: string[]): number {
  for (const c of candidates) {
    if (c in colIdx) return colIdx[c];
  }
  return -1;
}

export function createEmptyMasterData(): MasterData {
  return {
    MABUUN: [],
    KASIAU: [],
    PEMBATAAN: [],
    SULINGAN: [],
    MABURAI: [],
    'LUAR WILAYAH': [],
    Kejar: [],
  };
}

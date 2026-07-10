import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { VACCINE_ORDER, VACCINE_COLUMN_INDEX } from './vaccineMapping';
import { isInMonthYear, BULAN_INDONESIA } from './dateUtils';
import { sanitizeForExcel } from './sanitizer';

const FIRST_DATA_ROW = 7;
const FIXED_TABLE_ROWS = 200;
const DATA_AREA_END = FIRST_DATA_ROW + FIXED_TABLE_ROWS - 1; // row 206
const SUMMARY_LABEL_COL = 7; // Column G (1-based)
const DATE_FMT = 'dd-mmm-yy';

function countVaccines(
  children: ChildRecord[],
  month: number,
  year: number,
): { nL: Record<VaccineKey, number>; nP: Record<VaccineKey, number> } {
  const nL = {} as Record<VaccineKey, number>;
  const nP = {} as Record<VaccineKey, number>;
  for (const vk of VACCINE_ORDER) {
    nL[vk] = 0;
    nP[vk] = 0;
    for (const child of children) {
      const s = child.vaccines[vk];
      if (s && isInMonthYear(s, month, year)) {
        if (child.jk === 'L') nL[vk]++;
        else nP[vk]++;
      }
    }
  }
  return { nL, nP };
}

const SHEET_KELURAHAN_LABEL: Record<SheetName, string> = {
  MABUUN: 'Mabuun',
  KASIAU: 'Kasiau',
  PEMBATAAN: 'Pembataan',
  SULINGAN: 'Sulingan',
  MABURAI: 'Maburai',
  'LUAR WILAYAH': 'LUAR WILAYAH',
  Kejar: '',
};

async function stripSharedFormulas(templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetFiles = Object.keys(zip.files).filter(
    (f) => f.startsWith('xl/worksheets/sheet') && f.endsWith('.xml'),
  );
  for (const file of sheetFiles) {
    let content = await zip.file(file)!.async('string');
    // Remove all formula tags: <f>...</f> and <f ... /> (self-closing)
    content = content.replace(/<f[^>]*>[\s\S]*?<\/f>/g, '');
    content = content.replace(/<f[^>]*\/>/g, '');
    zip.file(file, content);
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

function clearOldSummary(ws: ExcelJS.Worksheet): void {
  const maxScan = Math.min(ws.rowCount, 2000);
  for (let r = FIRST_DATA_ROW; r <= maxScan; r++) {
    const cell = ws.getRow(r).getCell(SUMMARY_LABEL_COL);
    const val = cell.value;
    if (val && typeof val === 'string' && val.startsWith('Jumlah')) {
      cell.value = null;
      // Also clear cells in the block (rows 724-728 etc.)
      for (let dr = 1; dr <= 4; dr++) {
        const row = ws.getRow(r + dr);
        for (let c = 1; c <= 49; c++) {
          row.getCell(c).value = null;
        }
      }
      break;
    }
  }
}

function writeSummary(
  ws: ExcelJS.Worksheet,
  startRow: number,
  month: number,
  year: number,
  nL: Record<VaccineKey, number>,
  nP: Record<VaccineKey, number>,
): void {
  ws.getRow(startRow).getCell(SUMMARY_LABEL_COL).value =
    `Jumlah Imunisasi Bulan ${BULAN_INDONESIA[month]} ${year}`;

  const countRow = ws.getRow(startRow + 2);
  for (const vk of VACCINE_ORDER) {
    const cL = VACCINE_COLUMN_INDEX[vk] + 1;
    countRow.getCell(cL).value = nL[vk];
    countRow.getCell(cL).numFmt = undefined;
    countRow.getCell(cL + 1).value = nP[vk];
    countRow.getCell(cL + 1).numFmt = undefined;
  }

  const totalRow = ws.getRow(startRow + 3);
  for (const vk of VACCINE_ORDER) {
    const cL = VACCINE_COLUMN_INDEX[vk] + 1;
    totalRow.getCell(cL).value = nL[vk] + nP[vk];
    totalRow.getCell(cL).numFmt = undefined;
    totalRow.getCell(cL + 1).value = null;
  }
}

export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  // Strip shared formulas before exceljs loads (avoids "Shared Formula master" crash)
  const cleanedBuffer = await stripSharedFormulas(templateBuffer);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cleanedBuffer);

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) continue;
    const children = masterData[sheetName];

    // --- Update header ---
    ws.getCell('C3').value = `: ${BULAN_INDONESIA[month]} ${year}`;
    ws.getCell('A4').value = sheetName === 'MABUUN' ? '``' : 'Kelurahan/Desa';
    ws.getCell('C4').value = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;

    // --- Clear old summary from template ---
    clearOldSummary(ws);

    const lastChildRow = FIRST_DATA_ROW + children.length - 1;

    // --- Write children data ---
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);
      row.getCell(1).value = idx + 1;
      row.getCell(1).numFmt = undefined; // Force number, not date
      row.getCell(2).value = sanitizeForExcel(child.nama);
      row.getCell(2).numFmt = undefined;
      row.getCell(3).value = child.jk;
      row.getCell(3).numFmt = undefined;
      if (child.tanggalLahirSerial) {
        row.getCell(4).value = child.tanggalLahirSerial;
        row.getCell(4).numFmt = DATE_FMT;
      }
      row.getCell(5).value = child.nik || null;
      row.getCell(5).numFmt = undefined;
      row.getCell(6).value = sanitizeForExcel(child.namaOrangTua);
      row.getCell(6).numFmt = undefined;
      row.getCell(7).value = sanitizeForExcel(child.alamat);
      row.getCell(7).numFmt = undefined;
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        const s = child.vaccines[vk];
        if (s) {
          if (child.jk === 'L') {
            row.getCell(cL).value = s;
            row.getCell(cL).numFmt = DATE_FMT;
            row.getCell(cL + 1).value = null;
          } else {
            row.getCell(cL).value = null;
            row.getCell(cL + 1).value = s;
            row.getCell(cL + 1).numFmt = DATE_FMT;
          }
        } else {
          row.getCell(cL).value = null;
          row.getCell(cL + 1).value = null;
        }
      }
    });

    // --- Write summary with 2-row gap ---
    const { nL, nP } = countVaccines(children, month, year);

    if (lastChildRow <= DATA_AREA_END) {
      // Children fit in 200-row table → summary at DATA_AREA_END + 3 (row 209)
      writeSummary(ws, DATA_AREA_END + 3, month, year, nL, nP);
    } else {
      // Children overflow (>200) → summary right after data with 2-row gap
      writeSummary(ws, lastChildRow + 3, month, year, nL, nP);
    }
  }

  const buf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

export function getUploadedVaccines(masterData: MasterData): Set<VaccineKey> {
  const uploaded = new Set<VaccineKey>();
  for (const sheet of ALL_SHEETS) {
    for (const child of masterData[sheet]) {
      for (const [vk, val] of Object.entries(child.vaccines)) {
        if (val) uploaded.add(vk as VaccineKey);
      }
    }
  }
  return uploaded;
}

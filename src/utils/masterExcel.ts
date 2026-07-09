import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { VACCINE_ORDER, VACCINE_COLUMN_INDEX } from './vaccineMapping';
import { isInMonthYear, BULAN_INDONESIA } from './dateUtils';
import { sanitizeForExcel } from './sanitizer';

// ExcelJS uses 1-based indexing for both rows and columns
const FIRST_DATA_ROW = 7;
const TOTAL_COLS = 49;
const SUMMARY_LABEL_COL = 7; // Column G (1-based)

/** Find summary start row in xlsx WorkSheet (used by tests). */
export function findSummaryStartRow(ws: XLSX.WorkSheet): number {
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1:A1');
  for (let r = 6; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: 6 })];
    if (cell?.v && String(cell.v).startsWith('Jumlah')) return r;
  }
  return range.e.r - 3;
}

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

/**
 * Strip shared formulas from the template workbook via JSZip before loading with exceljs.
 * exceljs cannot handle shared formulas, causing "Shared Formula master must exist" errors.
 */
async function stripSharedFormulas(templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetFiles = Object.keys(zip.files).filter(
    (f) => f.startsWith('xl/worksheets/sheet') && f.endsWith('.xml'),
  );

  for (const file of sheetFiles) {
    let content = await zip.file(file)!.async('string');
    // Remove all formula tags: <f>...</f> and self-closing <f ... />
    content = content.replace(/<f[^>]*>.*?<\/f>/gs, '');
    content = content.replace(/<f[^>]*\/>/g, '');
    zip.file(file, content);
  }

  const buf = await zip.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
  });
  return buf;
}

/**
 * Build Master Excel using exceljs to preserve all template styles.
 * Strips shared formulas from template first via JSZip to avoid exceljs errors,
 * then modifies cell values in-place to preserve colors, fills, borders, fonts.
 */
export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  // 1. Strip shared formulas from template XML before loading with exceljs
  const cleanedBuffer = await stripSharedFormulas(templateBuffer);

  // 2. Load cleaned template with exceljs (preserves all styles)
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cleanedBuffer);

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) continue;

    const children = masterData[sheetName];

    // --- Update header cells ---
    ws.getCell('C3').value = `: ${BULAN_INDONESIA[month]} ${year}`;
    const kelLabel = sheetName === 'MABUUN' ? '``' : 'Kelurahan/Desa';
    ws.getCell('A4').value = kelLabel;
    ws.getCell('C4').value = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;

    // --- Find template summary position ---
    let templateSummaryStart = -1;
    const maxScan = Math.min(ws.rowCount, 1000);
    for (let r = FIRST_DATA_ROW; r <= maxScan; r++) {
      const row = ws.getRow(r);
      const val = row.getCell(SUMMARY_LABEL_COL).value;
      if (val && typeof val === 'string' && val.startsWith('Jumlah')) {
        templateSummaryStart = r;
        break;
      }
    }
    if (templateSummaryStart === -1) {
      templateSummaryStart = ws.rowCount - 3;
    }

    // --- Save summary template styles from original position ---
    const summaryTemplate: { style: ExcelJS.Style; value: ExcelJS.CellValue }[][] = [];
    for (let dr = 0; dr < 4; dr++) {
      const row = ws.getRow(templateSummaryStart + dr);
      const rowData: { style: ExcelJS.Style; value: ExcelJS.CellValue }[] = [];
      for (let c = 1; c <= TOTAL_COLS; c++) {
        const cell = row.getCell(c);
        rowData.push({ style: { ...cell.style }, value: cell.value as ExcelJS.CellValue });
      }
      summaryTemplate.push(rowData);
    }

    // --- Clear ALL data rows ---
    const clearEnd = Math.max(templateSummaryStart + 4, FIRST_DATA_ROW + children.length + 10);
    for (let r = FIRST_DATA_ROW; r <= clearEnd; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        row.getCell(c).value = null;
      }
      row.commit();
    }

    // --- Write children at FIRST_DATA_ROW ---
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);

      row.getCell(1).value = idx + 1;
      row.getCell(2).value = sanitizeForExcel(child.nama);
      row.getCell(3).value = child.jk;
      row.getCell(4).value = child.tanggalLahirSerial;
      row.getCell(5).value = child.nik || null;
      row.getCell(6).value = sanitizeForExcel(child.namaOrangTua);
      row.getCell(7).value = sanitizeForExcel(child.alamat);

      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        const serial = child.vaccines[vk];
        if (serial) {
          if (child.jk === 'L') {
            row.getCell(cL).value = serial;
            row.getCell(cL + 1).value = null;
          } else {
            row.getCell(cL).value = null;
            row.getCell(cL + 1).value = serial;
          }
        } else {
          row.getCell(cL).value = null;
          row.getCell(cL + 1).value = null;
        }
      }
      row.commit();
    });

    // --- Write summary block at new position ---
    const newSummaryStart = FIRST_DATA_ROW + children.length + 1;
    const { nL, nP } = countVaccines(children, month, year);

    for (let dr = 0; dr < 4; dr++) {
      const row = ws.getRow(newSummaryStart + dr);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        const tmpl = summaryTemplate[dr][c - 1];
        row.getCell(c).style = tmpl.style;
        row.getCell(c).value = tmpl.value;
      }

      if (dr === 0) {
        row.getCell(SUMMARY_LABEL_COL).value = `Jumlah Imunisasi Bulan ${BULAN_INDONESIA[month]} ${year}`;
      } else if (dr === 2) {
        for (const vk of VACCINE_ORDER) {
          const cL = VACCINE_COLUMN_INDEX[vk] + 1;
          row.getCell(cL).value = nL[vk];
          row.getCell(cL + 1).value = nP[vk];
        }
      } else if (dr === 3) {
        for (const vk of VACCINE_ORDER) {
          const cL = VACCINE_COLUMN_INDEX[vk] + 1;
          row.getCell(cL).value = nL[vk] + nP[vk];
          row.getCell(cL + 1).value = null;
        }
      }
      row.commit();
    }

    // --- Trim worksheet ---
    const lastRowToKeep = newSummaryStart + 3;
    if (ws.rowCount > lastRowToKeep) {
      ws.spliceRows(lastRowToKeep + 1, ws.rowCount - lastRowToKeep);
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

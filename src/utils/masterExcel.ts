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
const DATE_FMT = 'dd-mmm-yy';

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

/** Deep-clone an ExcelJS style (fill, font, border, alignment, numFmt, protection). */
function cloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
  if (!style || Object.keys(style).length === 0) return {};
  return JSON.parse(JSON.stringify(style));
}

/** Save a row's styles (for columns 1..TOTAL_COLS). */
function saveRowStyles(ws: ExcelJS.Worksheet, rowNum: number): Partial<ExcelJS.Style>[] {
  const styles: Partial<ExcelJS.Style>[] = [];
  const row = ws.getRow(rowNum);
  for (let c = 1; c <= TOTAL_COLS; c++) {
    styles.push(cloneStyle(row.getCell(c).style));
  }
  return styles;
}

/** Save a 4-row summary block: styles + values for each cell. */
function saveSummaryBlock(
  ws: ExcelJS.Worksheet,
  startRow: number,
): { style: Partial<ExcelJS.Style>; value: ExcelJS.CellValue }[][] {
  const block: { style: Partial<ExcelJS.Style>; value: ExcelJS.CellValue }[][] = [];
  for (let dr = 0; dr < 4; dr++) {
    const row = ws.getRow(startRow + dr);
    const rowData: { style: Partial<ExcelJS.Style>; value: ExcelJS.CellValue }[] = [];
    for (let c = 1; c <= TOTAL_COLS; c++) {
      const cell = row.getCell(c);
      rowData.push({
        style: cloneStyle(cell.style),
        value: cell.value as ExcelJS.CellValue,
      });
    }
    block.push(rowData);
  }
  return block;
}

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

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

/**
 * Collect merge ranges in the summary area (between startRow and startRow+3),
 * saving them as offsets relative to startRow.
 * Then unmerge them from the worksheet.
 */
function extractSummaryMerges(
  ws: ExcelJS.Worksheet,
  summaryStart: number,
): { sRow: number; sCol: number; eRow: number; eCol: number }[] {
  const offsets: { sRow: number; sCol: number; eRow: number; eCol: number }[] = [];

  // ExcelJS stores merges internally — access via model
  const model = ws.model as { merges?: string[] };
  if (!model.merges) return offsets;

  const toRemove: string[] = [];
  for (const merge of model.merges) {
    // Parse "G725:G726" format
    const match = merge.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/);
    if (!match) continue;
    const r1 = parseInt(match[2]);
    const r2 = parseInt(match[4]);
    if (r1 >= summaryStart && r2 <= summaryStart + 3) {
      toRemove.push(merge);
      offsets.push({
        sRow: r1 - summaryStart,
        sCol: colLetterToNum(match[1]),
        eRow: r2 - summaryStart,
        eCol: colLetterToNum(match[3]),
      });
    }
  }

  // Remove the old summary merges
  for (const m of toRemove) {
    try { ws.unMergeCells(m); } catch { /* ignore */ }
  }

  return offsets;
}

/** Convert column letters (A=1, B=2, ..., AW=49) to 1-based number. */
function colLetterToNum(letters: string): number {
  let result = 0;
  for (let i = 0; i < letters.length; i++) {
    result = result * 26 + (letters.charCodeAt(i) - 64);
  }
  return result;
}

/**
 * Build Master Excel using exceljs to preserve all template styles (colors, fills, borders).
 * Flow:
 * 1. Strip shared formulas via JSZip (exceljs can't handle them)
 * 2. Load cleaned template with exceljs (preserves all styles)
 * 3. For each sheet: save styles, clear old data, write new data with styles, write summary
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
      const val = ws.getRow(r).getCell(SUMMARY_LABEL_COL).value;
      if (val && typeof val === 'string' && val.startsWith('Jumlah')) {
        templateSummaryStart = r;
        break;
      }
    }
    if (templateSummaryStart === -1) templateSummaryStart = ws.rowCount - 3;

    // --- Save styles BEFORE clearing ---
    // Data row style template (from the first data row in the template)
    const dataRowStyles = saveRowStyles(ws, FIRST_DATA_ROW);
    // Summary block (4 rows of styles + values)
    const summaryTemplate = saveSummaryBlock(ws, templateSummaryStart);
    // Summary merges (relative offsets)
    const summaryMergeOffsets = extractSummaryMerges(ws, templateSummaryStart);

    // --- Clear ALL old data rows + old summary ---
    const clearEnd = Math.max(templateSummaryStart + 4, FIRST_DATA_ROW + children.length + 10);
    for (let r = FIRST_DATA_ROW; r <= clearEnd; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        row.getCell(c).value = null;
        // Also clear style for rows that will be beyond our data+summary
        // (data row styles will be re-applied below for actual data rows)
      }
      row.commit();
    }

    // --- Write children at FIRST_DATA_ROW with preserved styles ---
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);

      // Apply template data row styles to ALL columns first
      for (let c = 1; c <= TOTAL_COLS; c++) {
        row.getCell(c).style = cloneStyle(dataRowStyles[c - 1]);
      }

      // Write fixed columns
      row.getCell(1).value = idx + 1;
      row.getCell(2).value = sanitizeForExcel(child.nama);
      row.getCell(3).value = child.jk;
      if (child.tanggalLahirSerial) {
        row.getCell(4).value = child.tanggalLahirSerial;
        row.getCell(4).numFmt = DATE_FMT;
      }
      row.getCell(5).value = child.nik || null;
      row.getCell(6).value = sanitizeForExcel(child.namaOrangTua);
      row.getCell(7).value = sanitizeForExcel(child.alamat);

      // Write vaccine columns
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1; // Convert 0-based to 1-based
        const serial = child.vaccines[vk];
        if (serial) {
          if (child.jk === 'L') {
            row.getCell(cL).value = serial;
            row.getCell(cL).numFmt = DATE_FMT;
            row.getCell(cL + 1).value = null;
          } else {
            row.getCell(cL).value = null;
            row.getCell(cL + 1).value = serial;
            row.getCell(cL + 1).numFmt = DATE_FMT;
          }
        } else {
          row.getCell(cL).value = null;
          row.getCell(cL + 1).value = null;
        }
      }
      row.commit();
    });

    // --- Blank separator row ---
    // (just skip one row — it stays empty with no style, which is fine)

    // --- Write summary block at new position ---
    const newSummaryStart = FIRST_DATA_ROW + children.length + 1;
    const { nL, nP } = countVaccines(children, month, year);

    for (let dr = 0; dr < 4; dr++) {
      const row = ws.getRow(newSummaryStart + dr);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        const tmpl = summaryTemplate[dr][c - 1];
        row.getCell(c).style = cloneStyle(tmpl.style);
        row.getCell(c).value = tmpl.value;
      }

      if (dr === 0) {
        // Overwrite summary label with month/year
        row.getCell(SUMMARY_LABEL_COL).value =
          `Jumlah Imunisasi Bulan ${BULAN_INDONESIA[month]} ${year}`;
      } else if (dr === 2) {
        // Counts per L/P
        for (const vk of VACCINE_ORDER) {
          const cL = VACCINE_COLUMN_INDEX[vk] + 1;
          row.getCell(cL).value = nL[vk];
          row.getCell(cL + 1).value = nP[vk];
        }
      } else if (dr === 3) {
        // Total L + P
        for (const vk of VACCINE_ORDER) {
          const cL = VACCINE_COLUMN_INDEX[vk] + 1;
          row.getCell(cL).value = nL[vk] + nP[vk];
          row.getCell(cL + 1).value = null;
        }
      }
      row.commit();
    }

    // --- Re-add summary merges at new position ---
    for (const offset of summaryMergeOffsets) {
      try {
        ws.mergeCells(
          newSummaryStart + offset.sRow, offset.sCol,
          newSummaryStart + offset.eRow, offset.eCol,
        );
      } catch { /* ignore merge conflicts */ }
    }

    // --- Trim: remove excess rows after summary ---
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

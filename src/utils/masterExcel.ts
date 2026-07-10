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
const DATA_COLUMNS = 49; // A = 1, AW = 49

/**
 * Pre-process template via JSZip:
 * 1. Strip ALL cell `<v>` values (clear data, preserve cell styles/borders)
 * 2. Strip ALL `<f>` formula tags (prevent shared formula errors)
 * 3. Strip ALL `<mergeCell>` elements (prevent stale merge references)
 */
async function cleanTemplate(templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetFiles = Object.keys(zip.files).filter(
    (f) => f.startsWith('xl/worksheets/sheet') && f.endsWith('.xml'),
  );
  for (const file of sheetFiles) {
    let content = await zip.file(file)!.async('string');

    // 1. Remove cell values <v>...</v> only from DATA rows (r >= 7)
    //    This preserves header values (rows 1-6) while clearing old data
    content = content.replace(/<row r="(\d+)"[^>]*>[\s\S]*?<\/row>/g, (match, rStr) => {
      const r = parseInt(rStr, 10);
      if (r >= 7) {
        return match.replace(/<v>[\s\S]*?<\/v>/g, '');
      }
      return match;
    });

    // 2. Remove ALL formulas: <f>...</f> and <f ... />
    content = content.replace(/<f[^>]*>[\s\S]*?<\/f>/g, '');
    content = content.replace(/<f[^>]*\/>/g, '');

    // 3. Remove ALL mergeCell definitions (they reference old positions)
    content = content.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/g, '');

    zip.file(file, content);
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
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

function writeChildCell(
  row: ExcelJS.Row,
  col: number,
  value: string | number | null | undefined,
  numFmt?: string,
): void {
  const cell = row.getCell(col);
  cell.value = value ?? null;
  if (numFmt) {
    cell.numFmt = numFmt;
  } else {
    cell.numFmt = undefined;
  }
}

export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  // Pre-clean template: strip values, formulas, merge cells (preserves styles/borders)
  const cleanedBuffer = await cleanTemplate(templateBuffer);
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

    const lastChildRow = FIRST_DATA_ROW + children.length - 1;

    // --- Write children data ---
    for (let idx = 0; idx < children.length; idx++) {
      const child = children[idx];
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);
      writeChildCell(row, 1, idx + 1);
      writeChildCell(row, 2, sanitizeForExcel(child.nama));
      writeChildCell(row, 3, child.jk);
      writeChildCell(row, 4, child.tanggalLahirSerial, DATE_FMT);
      writeChildCell(row, 5, child.nik || null);
      writeChildCell(row, 6, sanitizeForExcel(child.namaOrangTua));
      writeChildCell(row, 7, sanitizeForExcel(child.alamat));

      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        const s = child.vaccines[vk];
        if (s) {
          if (child.jk === 'L') {
            writeChildCell(row, cL, s, DATE_FMT);
            writeChildCell(row, cL + 1, null);
          } else {
            writeChildCell(row, cL, null);
            writeChildCell(row, cL + 1, s, DATE_FMT);
          }
        } else {
          writeChildCell(row, cL, null);
          writeChildCell(row, cL + 1, null);
        }
      }
    }

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

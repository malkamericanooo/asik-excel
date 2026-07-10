import ExcelJS from 'exceljs';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { VACCINE_ORDER, VACCINE_COLUMN_INDEX } from './vaccineMapping';
import { isInMonthYear, BULAN_INDONESIA } from './dateUtils';
import { sanitizeForExcel } from './sanitizer';

const FIRST_DATA_ROW = 7;
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

function findSummaryRow(ws: ExcelJS.Worksheet): number {
  const maxScan = Math.min(ws.rowCount, 1000);
  for (let r = FIRST_DATA_ROW; r <= maxScan; r++) {
    const val = ws.getRow(r).getCell(SUMMARY_LABEL_COL).value;
    if (val && typeof val === 'string' && val.startsWith('Jumlah')) return r;
  }
  return ws.rowCount - 4;
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
    countRow.getCell(cL + 1).value = nP[vk];
  }

  const totalRow = ws.getRow(startRow + 3);
  for (const vk of VACCINE_ORDER) {
    const cL = VACCINE_COLUMN_INDEX[vk] + 1;
    totalRow.getCell(cL).value = nL[vk] + nP[vk];
    totalRow.getCell(cL + 1).value = null;
  }
}

/**
 * Convert all formula cells in a workbook to their cached values BEFORE modifications.
 * This avoids exceljs "Shared Formula master" crash during writeBuffer().
 */
function convertFormulasToValues(wb: ExcelJS.Workbook): void {
  for (const ws of wb.worksheets) {
    if (!ws) continue;
    ws.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if ((cell as any).isFormula) {
          const result = (cell as any).result;
          cell.value = (result !== undefined && result !== null) ? result : null;
        }
      });
    });
  }
}

export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBuffer);
  convertFormulasToValues(wb);

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) continue;
    const children = masterData[sheetName];

    // --- Update header ---
    ws.getCell('C3').value = `: ${BULAN_INDONESIA[month]} ${year}`;
    ws.getCell('A4').value = sheetName === 'MABUUN' ? '``' : 'Kelurahan/Desa';
    ws.getCell('C4').value = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;

    // --- Find existing summary position ---
    const summaryRow = findSummaryRow(ws);
    const dataAreaEnd = summaryRow - 1;
    const lastChildRow = FIRST_DATA_ROW + children.length - 1;

    // --- Write children data ---
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);
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

    if (lastChildRow < summaryRow) {
      writeSummary(ws, summaryRow, month, year, nL, nP);
    } else {
      // Children overflow: table ends at lastChildRow, 2 empty rows gap, then summary
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

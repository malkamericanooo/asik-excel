import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { VACCINE_ORDER, VACCINE_COLUMN_INDEX } from './vaccineMapping';
import { isInMonthYear, BULAN_INDONESIA } from './dateUtils';
import { sanitizeForExcel } from './sanitizer';

const FIRST_DATA_ROW = 7;
const TOTAL_COLS = 49;
const SUMMARY_LABEL_COL = 7; // Column G (1-based)

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

async function stripSharedFormulas(templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetFiles = Object.keys(zip.files).filter(
    (f) => f.startsWith('xl/worksheets/sheet') && f.endsWith('.xml'),
  );
  for (const file of sheetFiles) {
    let content = await zip.file(file)!.async('string');
    content = content.replace(/<f[^>]*>.*?<\/f>/gs, '');
    content = content.replace(/<f[^>]*\/>/g, '');
    zip.file(file, content);
  }
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  const cleanedBuffer = await stripSharedFormulas(templateBuffer);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cleanedBuffer);

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) continue;
    const children = masterData[sheetName];

    // Update header
    ws.getCell('C3').value = `: ${BULAN_INDONESIA[month]} ${year}`;
    ws.getCell('A4').value = sheetName === 'MABUUN' ? '``' : 'Kelurahan/Desa';
    ws.getCell('C4').value = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;

    // Find existing summary position
    let summaryRow = -1;
    const maxScan = Math.min(ws.rowCount, 1000);
    for (let r = FIRST_DATA_ROW; r <= maxScan; r++) {
      const val = ws.getRow(r).getCell(SUMMARY_LABEL_COL).value;
      if (val && typeof val === 'string' && val.startsWith('Jumlah')) {
        summaryRow = r;
        break;
      }
    }
    if (summaryRow === -1) summaryRow = ws.rowCount - 4;

    // Write children at FIRST_DATA_ROW (overwrite existing values, don't create extra cells)
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);
      // Only touch the columns we need - don't loop through all 49
      row.getCell(1).value = idx + 1;
      row.getCell(2).value = sanitizeForExcel(child.nama);
      row.getCell(3).value = child.jk;
      row.getCell(4).value = child.tanggalLahirSerial;
      row.getCell(5).value = child.nik || null;
      row.getCell(6).value = sanitizeForExcel(child.namaOrangTua);
      row.getCell(7).value = sanitizeForExcel(child.alamat);
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        const s = child.vaccines[vk];
        if (s) {
          row.getCell(cL).value = child.jk === 'L' ? s : null;
          row.getCell(cL + 1).value = child.jk === 'P' ? s : null;
        } else {
          row.getCell(cL).value = null;
          row.getCell(cL + 1).value = null;
        }
      }
    });

    // Update summary at the existing position
    const { nL, nP } = countVaccines(children, month, year);

    ws.getRow(summaryRow).getCell(SUMMARY_LABEL_COL).value = 
      `Jumlah Imunisasi Bulan ${BULAN_INDONESIA[month]} ${year}`;

    const countRow = ws.getRow(summaryRow + 2);
    for (const vk of VACCINE_ORDER) {
      const cL = VACCINE_COLUMN_INDEX[vk] + 1;
      countRow.getCell(cL).value = nL[vk];
      countRow.getCell(cL + 1).value = nP[vk];
    }

    const totalRow = ws.getRow(summaryRow + 3);
    for (const vk of VACCINE_ORDER) {
      const cL = VACCINE_COLUMN_INDEX[vk] + 1;
      totalRow.getCell(cL).value = nL[vk] + nP[vk];
      totalRow.getCell(cL + 1).value = null;
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

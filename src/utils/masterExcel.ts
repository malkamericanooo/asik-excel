import * as XLSX from 'xlsx';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { VACCINE_ORDER, VACCINE_COLUMN_INDEX } from './vaccineMapping';
import { isInMonthYear, BULAN_INDONESIA } from './dateUtils';

export const FIRST_DATA_ROW = 6;
export const TOTAL_COLS = 49;
const DATE_FMT = 'dd-mmm-yy';
const SUMMARY_LABEL_COL = 6;

/** Deep-clone a worksheet cell (value + style), without formulas. */
function cloneCell(cell: XLSX.CellObject | undefined): XLSX.CellObject | undefined {
  if (!cell) return undefined;
  const cloned: XLSX.CellObject = { ...cell };
  delete cloned.f;
  if (cell.s) cloned.s = JSON.parse(JSON.stringify(cell.s));
  return cloned;
}

/** Find the row index where the summary block starts ("Jumlah" in column G). */
export function findSummaryStartRow(ws: XLSX.WorkSheet): number {
  const range = XLSX.utils.decode_range(ws['!ref'] ?? 'A1:A1');
  for (let r = FIRST_DATA_ROW; r <= range.e.r; r++) {
    const cell = ws[XLSX.utils.encode_cell({ r, c: SUMMARY_LABEL_COL })];
    if (cell?.v && String(cell.v).startsWith('Jumlah')) return r;
  }
  return range.e.r - 3;
}

/** Extract the 4-row summary block (with styles) from a template sheet. */
export function extractSummaryBlock(ws: XLSX.WorkSheet, startRow: number): XLSX.CellObject[][] {
  const block: XLSX.CellObject[][] = [];
  for (let dr = 0; dr < 4; dr++) {
    const row: XLSX.CellObject[] = [];
    for (let c = 0; c < TOTAL_COLS; c++) {
      const addr = XLSX.utils.encode_cell({ r: startRow + dr, c });
      row.push(cloneCell(ws[addr]) ?? { v: '', t: 'z' });
    }
    block.push(row);
  }
  return block;
}

/** Clone style reference row from template (row 6) for new data rows. */
function getDataRowStyleTemplate(ws: XLSX.WorkSheet): XLSX.CellObject[] {
  const template: XLSX.CellObject[] = [];
  for (let c = 0; c < TOTAL_COLS; c++) {
    const addr = XLSX.utils.encode_cell({ r: FIRST_DATA_ROW, c });
    template.push(cloneCell(ws[addr]) ?? { v: '', t: 'z' });
  }
  return template;
}

function setCell(
  ws: XLSX.WorkSheet,
  r: number,
  c: number,
  value: string | number | null | undefined,
  styleTemplate?: XLSX.CellObject,
  numFmt?: string,
): void {
  const addr = XLSX.utils.encode_cell({ r, c });
  const t = typeof value === 'number' ? 'n' : typeof value === 'string' && value !== '' ? 's' : 'z';
  const base = styleTemplate ? { ...styleTemplate } : (ws[addr] ?? { v: '', t: 'z' });
  delete base.f;
  ws[addr] = { ...base, v: value ?? '', t };
  if (numFmt) ws[addr].z = numFmt;
}

function writeChildRow(
  ws: XLSX.WorkSheet,
  row: number,
  child: ChildRecord,
  no: number,
  styleRow: XLSX.CellObject[],
): void {
  setCell(ws, row, 0, no, styleRow[0]);
  setCell(ws, row, 1, child.nama, styleRow[1]);
  setCell(ws, row, 2, child.jk, styleRow[2]);
  setCell(ws, row, 3, child.tanggalLahirSerial, styleRow[3], DATE_FMT);
  setCell(ws, row, 4, child.nik || null, styleRow[4]);
  setCell(ws, row, 5, child.namaOrangTua, styleRow[5]);
  setCell(ws, row, 6, child.alamat, styleRow[6]);

  for (const vk of VACCINE_ORDER) {
    const cL = VACCINE_COLUMN_INDEX[vk];
    const serial = child.vaccines[vk];
    if (serial) {
      if (child.jk === 'L') {
        setCell(ws, row, cL, serial, styleRow[cL], DATE_FMT);
        setCell(ws, row, cL + 1, null, styleRow[cL + 1]);
      } else {
        setCell(ws, row, cL, null, styleRow[cL]);
        setCell(ws, row, cL + 1, serial, styleRow[cL + 1], DATE_FMT);
      }
    } else {
      setCell(ws, row, cL, null, styleRow[cL]);
      setCell(ws, row, cL + 1, null, styleRow[cL + 1]);
    }
  }
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

function writeSummaryBlock(
  ws: XLSX.WorkSheet,
  startRow: number,
  template: XLSX.CellObject[][],
  children: ChildRecord[],
  month: number,
  year: number,
): void {
  const { nL, nP } = countVaccines(children, month, year);

  for (let dr = 0; dr < 4; dr++) {
    for (let c = 0; c < TOTAL_COLS; c++) {
      const addr = XLSX.utils.encode_cell({ r: startRow + dr, c });
      const tmpl = template[dr][c];
      ws[addr] = cloneCell(tmpl) ?? { v: '', t: 'z' };
    }
  }

  // Overwrite summary label with month/year
  const labelAddr = XLSX.utils.encode_cell({ r: startRow, c: SUMMARY_LABEL_COL });
  const labelCell = ws[labelAddr] ?? { t: 's' };
  labelCell.v = `Jumlah Imunisasi Bulan ${BULAN_INDONESIA[month]} ${year}`;
  labelCell.t = 's';
  delete labelCell.f;
  ws[labelAddr] = labelCell;

  // Row 2 (counts) — overwrite vaccine count cells
  const countRow = startRow + 2;
  for (const vk of VACCINE_ORDER) {
    const cL = VACCINE_COLUMN_INDEX[vk];
    setCell(ws, countRow, cL, nL[vk], template[2][cL]);
    setCell(ws, countRow, cL + 1, nP[vk], template[2][cL + 1]);
  }

  // Row 3 (Total L + P)
  const totalRow = startRow + 3;
  for (const vk of VACCINE_ORDER) {
    const cL = VACCINE_COLUMN_INDEX[vk];
    setCell(ws, totalRow, cL, nL[vk] + nP[vk], template[3][cL]);
    setCell(ws, totalRow, cL + 1, null, template[3][cL + 1]);
  }
}

function clearRows(ws: XLSX.WorkSheet, fromRow: number, toRow: number): void {
  for (let r = fromRow; r <= toRow; r++) {
    for (let c = 0; c < TOTAL_COLS; c++) {
      const addr = XLSX.utils.encode_cell({ r, c });
      delete ws[addr];
    }
  }
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

function updateSheetHeader(
  ws: XLSX.WorkSheet,
  sheetName: SheetName,
  month: number,
  year: number,
): void {
  const bulanAddr = 'C3';
  if (ws[bulanAddr]) ws[bulanAddr].v = `: ${BULAN_INDONESIA[month]} ${year}`;

  const kelLabel = sheetName === 'MABUUN' ? '``' : 'Kelurahan/Desa';
  if (ws['A4']) ws['A4'].v = kelLabel;
  if (ws['C4']) ws['C4'].v = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;
}

/** Remove cells outside the active range and drop merges that fall beyond it. */
function trimWorksheet(ws: XLSX.WorkSheet, lastRow: number): void {
  ws['!ref'] = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: lastRow, c: TOTAL_COLS - 1 });

  for (const addr of Object.keys(ws)) {
    if (addr.startsWith('!')) continue;
    const cell = XLSX.utils.decode_cell(addr);
    if (cell.r > lastRow || cell.c >= TOTAL_COLS) delete ws[addr];
  }

  if (ws['!merges']) {
    ws['!merges'] = ws['!merges'].filter(
      (m) => m.s.r <= lastRow && m.e.r <= lastRow && m.s.c < TOTAL_COLS && m.e.c < TOTAL_COLS,
    );
  }

  if (ws['!rows']) ws['!rows'] = ws['!rows'].slice(0, lastRow + 1);
  if (ws['!cols']) ws['!cols'] = ws['!cols'].slice(0, TOTAL_COLS);
}

/**
 * Build Master Excel from template + merged data.
 * @param templateBuffer - optional uploaded template; uses embedded default when omitted in browser
 */
export function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Blob {
  const wb = XLSX.read(templateBuffer, { type: 'array', cellStyles: true });

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.Sheets[sheetName];
    if (!ws) continue;

    const children = masterData[sheetName];
    const summaryStart = findSummaryStartRow(ws);
    const summaryTemplate = extractSummaryBlock(ws, summaryStart);
    const dataStyleRow = getDataRowStyleTemplate(ws);

    updateSheetHeader(ws, sheetName, month, year);

    // Remove old data + old summary
    clearRows(ws, FIRST_DATA_ROW, summaryStart + 3);

    // Write data rows
    let row = FIRST_DATA_ROW;
    children.forEach((child, idx) => {
      writeChildRow(ws, row, child, idx + 1, dataStyleRow);
      row++;
    });

    // Blank separator row
    row++;

    // Write summary block
    writeSummaryBlock(ws, row, summaryTemplate, children, month, year);

    // Strip formulas that would corrupt the output after rows are trimmed
    stripWorksheetFormulas(ws);
    trimWorksheet(ws, row + 3);
  }

  // Write WITHOUT cellStyles to prevent XML corruption in the output.
  // cellStyles is only needed for reading to preserve internal format refs.
  const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
  return new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}

/** Strip stale template formulas that would corrupt Excel after rows are trimmed. */
export function stripWorksheetFormulas(ws: XLSX.WorkSheet): void {
  for (const addr of Object.keys(ws)) {
    if (addr.startsWith('!')) continue;
    const cell = ws[addr];
    if (cell?.f) delete cell.f;
  }
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

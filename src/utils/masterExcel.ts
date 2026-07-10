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
const TABLE_SIZE = 200; // Fixed table size: rows 7-206

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

/** Deep-clone an ExcelJS style. */
function cloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
  if (!style || Object.keys(style).length === 0) return {};
  return JSON.parse(JSON.stringify(style));
}

/** Save a row's styles (1-based cols 1..TOTAL_COLS). */
function saveRowStyles(ws: ExcelJS.Worksheet, rowNum: number): Partial<ExcelJS.Style>[] {
  const styles: Partial<ExcelJS.Style>[] = [];
  const row = ws.getRow(rowNum);
  for (let c = 1; c <= TOTAL_COLS; c++) {
    styles.push(cloneStyle(row.getCell(c).style));
  }
  return styles;
}

/** Save a 4-row summary block: styles + values. */
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

// ─── Pre-processing: cleanTemplate ───────────────────────────────────
/**
 * Clean template via JSZip BEFORE exceljs load:
 * 1. Strip <v> values from rows >= 7 (keep styles via s= attribute)
 * 2. Strip formula tags (<f>) to prevent exceljs "Shared Formula" errors
 * 3. Strip <mergeCells> to prevent stale merge references
 */
async function cleanTemplate(templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);
  const sheetFiles = Object.keys(zip.files).filter(
    (f) => f.startsWith('xl/worksheets/sheet') && f.endsWith('.xml'),
  );

  for (const file of sheetFiles) {
    let xml = await zip.file(file)!.async('string');

    // Strip formulas: <f>...</f> and <f ... />
    xml = xml.replace(/<f[^>]*>.*?<\/f>/gs, '');
    xml = xml.replace(/<f[^>]*\/>/g, '');

    // Strip values from data rows (row >= 7) but keep style references.
    // Each cell looks like: <c r="B7" s="5"><v>...</v></c>
    // We want to keep <c r="B7" s="5"></c> (preserves style)
    xml = xml.replace(/<c\s+([^>]*?)>\s*<v>[^<]*<\/v>\s*<\/c>/g, (match, attrs) => {
      // Parse row number from r="XX7" attribute
      const rMatch = attrs.match(/r="[A-Z]+(\d+)"/);
      if (rMatch && parseInt(rMatch[1]) >= FIRST_DATA_ROW) {
        // Strip value but keep cell with style
        return `<c ${attrs}></c>`;
      }
      return match; // Keep header rows as-is
    });

    // Strip mergeCells section entirely to prevent stale references
    xml = xml.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/g, '');
    xml = xml.replace(/<mergeCells[^>]*\/>/g, '');

    zip.file(file, xml);
  }

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

// ─── Post-processing: stripPhantomSheets ─────────────────────────────
/**
 * Remove phantom sheets from ExcelJS output via JSZip.
 * ExcelJS creates empty sheets to fill gaps in non-sequential sheetIds.
 * E.g., template has ids [12,13,14,15,16,17,19] → ExcelJS fills 1-11,18.
 */
async function stripPhantomSheets(outputBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(outputBuffer);

  // 1. Read workbook.xml to find which sheets are real (have names)
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  if (!wbXml) return outputBuffer;

  // Extract real sheet references: <sheet name="MABUUN" sheetId="14" r:id="rId1"/>
  const sheetRefs: { name: string; rId: string }[] = [];
  const sheetRegex = /<sheet\s+[^>]*name="([^"]*)"[^>]*r:id="([^"]*)"[^>]*\/>/g;
  let match;
  while ((match = sheetRegex.exec(wbXml)) !== null) {
    sheetRefs.push({ name: match[1], rId: match[2] });
  }

  // 2. Read workbook.xml.rels to map rId → sheet file path
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (!relsXml) return outputBuffer;

  const rIdToFile: Record<string, string> = {};
  const relRegex = /<Relationship\s+[^>]*Id="([^"]*)"[^>]*Target="([^"]*)"[^>]*\/>/g;
  while ((match = relRegex.exec(relsXml)) !== null) {
    rIdToFile[match[1]] = match[2];
  }

  // 3. Determine which sheet files are real
  const realSheetFiles = new Set<string>();
  for (const ref of sheetRefs) {
    const target = rIdToFile[ref.rId];
    if (target) {
      // Target is relative to xl/, e.g., "worksheets/sheet14.xml"
      realSheetFiles.add(`xl/${target}`);
    }
  }

  // 4. Find and remove phantom sheet files
  const allSheetFiles = Object.keys(zip.files).filter(
    (f) => f.match(/^xl\/worksheets\/sheet\d+\.xml$/) && !f.endsWith('.rels'),
  );

  let removedCount = 0;
  for (const sheetFile of allSheetFiles) {
    if (!realSheetFiles.has(sheetFile)) {
      zip.remove(sheetFile);
      // Also remove its .rels if it exists
      const relsFile = sheetFile.replace('worksheets/', 'worksheets/_rels/') + '.rels';
      if (zip.files[relsFile]) zip.remove(relsFile);
      removedCount++;
    }
  }

  if (removedCount === 0) return outputBuffer;

  // 5. Fix [Content_Types].xml — remove overrides for deleted sheets
  const ctXml = await zip.file('[Content_Types].xml')?.async('string');
  if (ctXml) {
    const fixedCt = ctXml.replace(
      /<Override[^>]*PartName="\/xl\/worksheets\/sheet\d+\.xml"[^>]*\/>/g,
      (overrideMatch) => {
        const partMatch = overrideMatch.match(/PartName="\/([^"]+)"/);
        if (partMatch && !realSheetFiles.has(partMatch[1])) {
          return ''; // Remove override for phantom sheet
        }
        return overrideMatch;
      },
    );
    zip.file('[Content_Types].xml', fixedCt);
  }

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

// ─── Main build function ─────────────────────────────────────────────
/**
 * Build Master Excel using exceljs to preserve all template styles.
 * Pipeline:
 * 1. cleanTemplate() — JSZip strip values/formulas/mergeCells from rows >= 7
 * 2. ExcelJS load — preserves all styles (fills, borders, fonts)
 * 3. Write data + summary per sheet
 * 4. stripPhantomSheets() — JSZip remove phantom sheet XML files
 */
export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  // 1. Clean template (strip old data, formulas, merges)
  const cleanedBuffer = await cleanTemplate(templateBuffer);

  // 2. Load with exceljs (all styles preserved)
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cleanedBuffer);

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) continue;

    const children = masterData[sheetName];

    // --- Save data row style from template row 7 BEFORE we clear ---
    const dataRowStyles = saveRowStyles(ws, FIRST_DATA_ROW);

    // --- Find & save summary template (row where "Jumlah" was) ---
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
    const summaryTemplate = saveSummaryBlock(ws, templateSummaryStart);

    // --- Update header cells ---
    ws.getCell('C3').value = `: ${BULAN_INDONESIA[month]} ${year}`;
    const kelLabel = sheetName === 'MABUUN' ? '' : 'Kelurahan/Desa';
    ws.getCell('A4').value = kelLabel;
    ws.getCell('C4').value = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;

    // --- Clear all rows from FIRST_DATA_ROW onwards ---
    // cleanTemplate already stripped values, but exceljs may have loaded some
    for (let r = FIRST_DATA_ROW; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        row.getCell(c).value = null;
      }
      row.commit();
    }

    // --- Write children starting at FIRST_DATA_ROW ---
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);

      // Apply template data row styles
      for (let c = 1; c <= TOTAL_COLS; c++) {
        row.getCell(c).style = cloneStyle(dataRowStyles[c - 1]);
      }

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

    // --- Calculate summary position (fixed 200-row table or overflow) ---
    const dataEndRow = FIRST_DATA_ROW + children.length - 1;
    let summaryRow: number;
    if (children.length <= TABLE_SIZE) {
      // Fixed layout: summary at row 209 (7 + 200 + 2 blank rows)
      summaryRow = FIRST_DATA_ROW + TABLE_SIZE + 2;
    } else {
      // Overflow: summary 3 rows after last data row
      summaryRow = dataEndRow + 3;
    }

    // --- Write summary block ---
    const { nL, nP } = countVaccines(children, month, year);

    for (let dr = 0; dr < 4; dr++) {
      const row = ws.getRow(summaryRow + dr);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        const tmpl = summaryTemplate[dr][c - 1];
        row.getCell(c).style = cloneStyle(tmpl.style);
        row.getCell(c).value = tmpl.value;
      }

      if (dr === 0) {
        row.getCell(SUMMARY_LABEL_COL).value =
          `Jumlah Imunisasi Bulan ${BULAN_INDONESIA[month]} ${year}`;
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

    // --- Re-add header merges (A5:A6, B5:B6, ..., G5:G6 + vaccine header pairs) ---
    // These were stripped by cleanTemplate, re-add them
    try {
      // Fixed column header merges (rows 5-6)
      for (let c = 1; c <= 7; c++) {
        ws.mergeCells(5, c, 6, c);
      }
      // Vaccine header merges (each vaccine spans 2 cols in row 5)
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        ws.mergeCells(5, cL, 5, cL + 1);
      }
    } catch { /* ignore if merges already exist */ }

    // --- Add summary merges ---
    try {
      // "Jumlah" label spans 2 rows
      ws.mergeCells(summaryRow, SUMMARY_LABEL_COL, summaryRow + 1, SUMMARY_LABEL_COL);
      // Vaccine headers in summary row 0 (each spans 2 cols)
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        ws.mergeCells(summaryRow, cL, summaryRow, cL + 1);
      }
      // Total L+P row: each vaccine pair spans 2 cols
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        ws.mergeCells(summaryRow + 3, cL, summaryRow + 3, cL + 1);
      }
    } catch { /* ignore merge conflicts */ }
  }

  // 3. Write ExcelJS output
  const ejsBuf = (await wb.xlsx.writeBuffer()) as ArrayBuffer;

  // 4. Post-process: strip phantom sheets
  const finalBuf = await stripPhantomSheets(ejsBuf);

  return new Blob([finalBuf], {
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

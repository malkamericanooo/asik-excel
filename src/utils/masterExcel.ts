import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import type { ChildRecord, MasterData, SheetName, VaccineKey } from '../types';
import { ALL_SHEETS } from '../types';
import { VACCINE_ORDER, VACCINE_COLUMN_INDEX } from './vaccineMapping';
import { isInMonthYear, BULAN_INDONESIA } from './dateUtils';
import { sanitizeForExcel } from './sanitizer';

// ExcelJS uses 1-based indexing
const FIRST_DATA_ROW = 7;
const TOTAL_COLS = 49;
const SUMMARY_LABEL_COL = 7; // Column G
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

function cloneStyle(style: Partial<ExcelJS.Style>): Partial<ExcelJS.Style> {
  if (!style || Object.keys(style).length === 0) return {};
  return JSON.parse(JSON.stringify(style));
}

function saveRowStyles(ws: ExcelJS.Worksheet, rowNum: number): Partial<ExcelJS.Style>[] {
  const styles: Partial<ExcelJS.Style>[] = [];
  const row = ws.getRow(rowNum);
  for (let c = 1; c <= TOTAL_COLS; c++) {
    styles.push(cloneStyle(row.getCell(c).style));
  }
  return styles;
}

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

// ─── Pre-processing via JSZip ────────────────────────────────────────
/**
 * Prepare template for ExcelJS:
 * 1. Renumber sheetIds to sequential 1,2,3... (prevents phantom sheets)
 * 2. Rename sheet XML files to sheet1.xml, sheet2.xml, ...
 * 3. Strip formulas (<f> tags) — ExcelJS can't handle shared formulas
 * 4. Strip cell values from rows >= 7 (keep styles via s= attribute)
 * 5. Strip mergeCells — we re-create them programmatically
 */
async function prepareTemplate(templateBuffer: ArrayBuffer): Promise<ArrayBuffer> {
  const zip = await JSZip.loadAsync(templateBuffer);

  // ── Step 1: Renumber sheetIds & rename files ──
  const wbXml = await zip.file('xl/workbook.xml')?.async('string');
  const relsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string');
  if (wbXml && relsXml) {
    // Parse sheet entries: <sheet name="MABUUN" sheetId="14" r:id="rId1"/>
    const sheetEntries: { name: string; oldId: string; rId: string }[] = [];
    const sheetRegex = /<sheet\s+([^>]*?)\/>/g;
    let m;
    while ((m = sheetRegex.exec(wbXml)) !== null) {
      const attrs = m[1];
      const nameMatch = attrs.match(/name="([^"]*)"/);
      const idMatch = attrs.match(/sheetId="(\d+)"/);
      const rIdMatch = attrs.match(/r:id="([^"]*)"/);
      if (nameMatch && idMatch && rIdMatch) {
        sheetEntries.push({ name: nameMatch[1], oldId: idMatch[1], rId: rIdMatch[1] });
      }
    }

    // Map rId → old file path from rels
    const rIdToTarget: Record<string, string> = {};
    const relRegex = /<Relationship\s+([^>]*?)\/>/g;
    while ((m = relRegex.exec(relsXml)) !== null) {
      const attrs = m[1];
      const idMatch = attrs.match(/Id="([^"]*)"/);
      const targetMatch = attrs.match(/Target="([^"]*)"/);
      if (idMatch && targetMatch) {
        rIdToTarget[idMatch[1]] = targetMatch[1];
      }
    }

    // Rename sheet files: sheetOLD.xml → sheetNEW.xml (sequential 1-based)
    let newWbXml = wbXml;
    let newRelsXml = relsXml;

    for (let i = 0; i < sheetEntries.length; i++) {
      const entry = sheetEntries[i];
      const newId = String(i + 1);
      const oldTarget = rIdToTarget[entry.rId]; // e.g., "worksheets/sheet14.xml"
      const newTarget = `worksheets/sheet${newId}.xml`;

      if (oldTarget && oldTarget !== newTarget) {
        const oldPath = `xl/${oldTarget}`;
        const newPath = `xl/${newTarget}`;

        // Copy XML file content to new name
        const content = await zip.file(oldPath)?.async('string');
        if (content) {
          zip.file(newPath, content);
          if (oldPath !== newPath) zip.remove(oldPath);
        }

        // Copy rels file if exists
        const oldRels = oldPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
        const newRels = newPath.replace('worksheets/', 'worksheets/_rels/') + '.rels';
        const relsContent = await zip.file(oldRels)?.async('string');
        if (relsContent) {
          zip.file(newRels, relsContent);
          if (oldRels !== newRels) zip.remove(oldRels);
        }

        // Update rels: Target="worksheets/sheet14.xml" → Target="worksheets/sheet1.xml"
        newRelsXml = newRelsXml.replace(
          `Target="${oldTarget}"`,
          `Target="${newTarget}"`,
        );
      }

      // Update workbook.xml: sheetId="14" → sheetId="1"
      newWbXml = newWbXml.replace(
        `sheetId="${entry.oldId}"`,
        `sheetId="${newId}"`,
      );
    }

    zip.file('xl/workbook.xml', newWbXml);
    zip.file('xl/_rels/workbook.xml.rels', newRelsXml);

    // Fix [Content_Types].xml — update PartNames for renamed sheets
    const ctXml = await zip.file('[Content_Types].xml')?.async('string');
    if (ctXml) {
      let newCtXml = ctXml;
      for (let i = 0; i < sheetEntries.length; i++) {
        const oldTarget = rIdToTarget[sheetEntries[i].rId];
        const newTarget = `worksheets/sheet${i + 1}.xml`;
        if (oldTarget && oldTarget !== newTarget) {
          newCtXml = newCtXml.replace(
            `/xl/${oldTarget}`,
            `/xl/${newTarget}`,
          );
        }
      }
      zip.file('[Content_Types].xml', newCtXml);
    }
  }

  // ── Step 2: Clean sheet XML content ──
  const sheetFiles = Object.keys(zip.files).filter(
    (f) => f.match(/^xl\/worksheets\/sheet\d+\.xml$/),
  );

  for (const file of sheetFiles) {
    let xml = await zip.file(file)!.async('string');

    // Strip formulas
    xml = xml.replace(/<f[^>]*>.*?<\/f>/gs, '');
    xml = xml.replace(/<f[^>]*\/>/g, '');

    // Strip values from rows >= 7 (keep cell element with style attribute)
    xml = xml.replace(/<c\s+([^>]*?)>\s*<v>[^<]*<\/v>\s*<\/c>/g, (match, attrs) => {
      const rMatch = attrs.match(/r="[A-Z]+(\d+)"/);
      if (rMatch && parseInt(rMatch[1]) >= FIRST_DATA_ROW) {
        return `<c ${attrs}></c>`;
      }
      return match;
    });

    // Strip mergeCells — we re-create them programmatically
    xml = xml.replace(/<mergeCells[^>]*>[\s\S]*?<\/mergeCells>/g, '');
    xml = xml.replace(/<mergeCells[^>]*\/>/g, '');

    zip.file(file, xml);
  }

  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

// ─── Main build function ─────────────────────────────────────────────
export async function buildMasterExcel(
  masterData: MasterData,
  month: number,
  year: number,
  templateBuffer: ArrayBuffer,
): Promise<Blob> {
  // 1. Prepare template (renumber IDs, strip data/formulas/merges)
  const cleanedBuffer = await prepareTemplate(templateBuffer);

  // 2. Load with ExcelJS (styles preserved, no phantom sheets)
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(cleanedBuffer);

  for (const sheetName of ALL_SHEETS) {
    const ws = wb.getWorksheet(sheetName);
    if (!ws) continue;

    const children = masterData[sheetName];

    // Save styles BEFORE clearing
    const dataRowStyles = saveRowStyles(ws, FIRST_DATA_ROW);

    // Find & save summary template
    let templateSummaryStart = -1;
    for (let r = FIRST_DATA_ROW; r <= Math.min(ws.rowCount, 1000); r++) {
      const val = ws.getRow(r).getCell(SUMMARY_LABEL_COL).value;
      if (val && typeof val === 'string' && val.startsWith('Jumlah')) {
        templateSummaryStart = r;
        break;
      }
    }
    if (templateSummaryStart === -1) templateSummaryStart = ws.rowCount - 3;
    const summaryTemplate = saveSummaryBlock(ws, templateSummaryStart);

    // Update header
    ws.getCell('C3').value = `: ${BULAN_INDONESIA[month]} ${year}`;
    ws.getCell('A4').value = sheetName === 'MABUUN' ? '' : 'Kelurahan/Desa';
    ws.getCell('C4').value = `: ${SHEET_KELURAHAN_LABEL[sheetName]}`;

    // Clear ALL rows from data start onwards (values AND styles)
    for (let r = FIRST_DATA_ROW; r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      for (let c = 1; c <= TOTAL_COLS; c++) {
        row.getCell(c).value = null;
        row.getCell(c).style = {};
      }
      row.commit();
    }

    // Write children
    children.forEach((child, idx) => {
      const r = FIRST_DATA_ROW + idx;
      const row = ws.getRow(r);

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
        }
      }
      row.commit();
    });

    // Summary position: right after data + 1 blank separator row
    const summaryRow = FIRST_DATA_ROW + children.length + 1;

    // Write summary block
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

    // Re-add merges (headers + summary)
    try {
      for (let c = 1; c <= 7; c++) ws.mergeCells(5, c, 6, c);
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        ws.mergeCells(5, cL, 5, cL + 1);
      }
      ws.mergeCells(summaryRow, SUMMARY_LABEL_COL, summaryRow + 1, SUMMARY_LABEL_COL);
      for (const vk of VACCINE_ORDER) {
        const cL = VACCINE_COLUMN_INDEX[vk] + 1;
        ws.mergeCells(summaryRow, cL, summaryRow, cL + 1);
        ws.mergeCells(summaryRow + 3, cL, summaryRow + 3, cL + 1);
      }
    } catch { /* ignore */ }
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

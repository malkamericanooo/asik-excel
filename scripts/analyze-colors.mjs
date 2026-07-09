#!/usr/bin/env node
// Analyze colors in the template using ExcelJS (since SheetJS doesn't preserve styles well)
import ExcelJS from 'exceljs';
import { readFileSync } from 'fs';
import { resolve } from 'path';

const TEMPLATE_KOSONG = '/Users/malkaarif/Desktop/IMUNISASI JUNI 2026 kosong.xlsx';
const MASTER_OUTPUT = '/Users/malkaarif/Desktop/Master_Imunisasi_Juni_2026 (5).xlsx';

async function analyzeColors(path, label) {
  console.log(`\n=== ${label} ===`);
  console.log(`Path: ${path}`);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path);
  
  for (const ws of wb.worksheets) {
    console.log(`\n--- Sheet: "${ws.name}" ---`);
    // Check first 8 rows for fills/colors
    for (let r = 1; r <= Math.min(8, ws.rowCount); r++) {
      const row = ws.getRow(r);
      const fills = [];
      row.eachCell({ includeEmpty: false }, (cell, colNum) => {
        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
          const color = cell.fill.fgColor.argb || cell.fill.fgColor.theme || JSON.stringify(cell.fill.fgColor);
          fills.push(`${cell.address}=${color}`);
        }
      });
      if (fills.length > 0) {
        console.log(`  Row ${r} fills: ${fills.slice(0, 10).join(', ')}${fills.length > 10 ? ` ... +${fills.length - 10} more` : ''}`);
      }
    }
    
    // Check summary rows for fills
    for (let r = Math.max(1, ws.rowCount - 6); r <= ws.rowCount; r++) {
      const row = ws.getRow(r);
      const fills = [];
      row.eachCell({ includeEmpty: false }, (cell, colNum) => {
        if (cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) {
          const color = cell.fill.fgColor.argb || cell.fill.fgColor.theme || JSON.stringify(cell.fill.fgColor);
          fills.push(`${cell.address}=${color}`);
        }
      });
      if (fills.length > 0) {
        console.log(`  Row ${r} fills: ${fills.slice(0, 10).join(', ')}${fills.length > 10 ? ` ... +${fills.length - 10} more` : ''}`);
      }
    }
    
    // Sample data row fills (row 7)
    const dataRow = ws.getRow(7);
    const dataFills = [];
    dataRow.eachCell({ includeEmpty: true }, (cell, colNum) => {
      if (colNum <= 48) {
        const fill = cell.fill;
        const border = cell.border;
        if (fill && fill.type === 'pattern' && fill.fgColor) {
          dataFills.push(`${cell.address}: fill=${fill.fgColor.argb || JSON.stringify(fill.fgColor)}`);
        }
        if (border && Object.keys(border).length > 0) {
          if (dataFills.length === 0 || !dataFills[dataFills.length-1].startsWith(cell.address)) {
            dataFills.push(`${cell.address}: hasBorder`);
          }
        }
      }
    });
    if (dataFills.length > 0) {
      console.log(`  Row 7 styling (sample): ${dataFills.slice(0, 8).join(', ')}`);
    }
    
    // Check header row 5 styling
    const headerRow = ws.getRow(5);
    const headerFills = [];
    headerRow.eachCell({ includeEmpty: false }, (cell, colNum) => {
      if (colNum <= 48) {
        const fill = cell.fill;
        if (fill && fill.type === 'pattern' && fill.fgColor) {
          headerFills.push(`${cell.address}=${fill.fgColor.argb || JSON.stringify(fill.fgColor)}`);
        }
      }
    });
    if (headerFills.length > 0) {
      console.log(`  Header row 5 fills: ${headerFills.slice(0, 10).join(', ')}${headerFills.length > 10 ? ` ... +${headerFills.length - 10} more` : ''}`);
    }
    
    // Only check first sheet in detail
    break;
  }
}

async function main() {
  try {
    await analyzeColors(TEMPLATE_KOSONG, 'TEMPLATE KOSONG (Desktop)');
  } catch (e) { console.log('Error reading template kosong:', e.message); }
  
  try {
    await analyzeColors(MASTER_OUTPUT, 'MASTER OUTPUT (5) (Desktop)');
  } catch (e) { console.log('Error reading master output:', e.message); }
}

main().catch(console.error);

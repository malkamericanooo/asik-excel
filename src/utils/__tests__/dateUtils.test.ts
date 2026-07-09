import { describe, it, expect } from 'vitest';
import { dateStringToExcelSerial, excelSerialToDateStr, isInMonthYear, parseDateCell } from '../dateUtils';

describe('dateStringToExcelSerial', () => {
  it('converts 2000-01-01 to known Excel serial 36526', () => {
    expect(dateStringToExcelSerial('2000-01-01')).toBe(36526);
  });
  it('converts 2026-06-17 correctly (36526 + 9664)', () => {
    expect(dateStringToExcelSerial('2026-06-17')).toBe(46190);
  });
  it('converts 2026-01-01 correctly (36526 + 9497)', () => {
    expect(dateStringToExcelSerial('2026-01-01')).toBe(46023);
  });
  it('returns 0 for invalid input', () => {
    expect(dateStringToExcelSerial('invalid')).toBe(0);
  });
});

describe('excelSerialToDateStr', () => {
  it('round-trips 2026-06-17', () => {
    const serial = dateStringToExcelSerial('2026-06-17');
    expect(serial).toBe(46190);
    expect(excelSerialToDateStr(serial)).toBe('2026-06-17');
  });
  it('round-trips 2024-12-14', () => {
    const serial = dateStringToExcelSerial('2024-12-14');
    expect(excelSerialToDateStr(serial)).toBe('2024-12-14');
  });
  it('converts known serial 36526 to 2000-01-01', () => {
    expect(excelSerialToDateStr(36526)).toBe('2000-01-01');
  });
});

describe('isInMonthYear', () => {
  const serial = dateStringToExcelSerial('2026-06-17');
  it('returns true for matching month/year', () => expect(isInMonthYear(serial, 6, 2026)).toBe(true));
  it('returns false for wrong month', () => expect(isInMonthYear(serial, 5, 2026)).toBe(false));
  it('returns false for wrong year', () => expect(isInMonthYear(serial, 6, 2025)).toBe(false));
  it('returns false for zero serial', () => expect(isInMonthYear(0, 6, 2026)).toBe(false));
});

describe('parseDateCell', () => {
  it('parses YYYY-MM-DD string', () => expect(parseDateCell('2026-06-17')).toBe('2026-06-17'));
  it('parses DD/MM/YYYY string', () => expect(parseDateCell('17/06/2026')).toBe('2026-06-17'));
  it('parses numeric Excel serial', () => expect(parseDateCell(46190)).toBe('2026-06-17'));
  it('returns empty string for null', () => expect(parseDateCell(null)).toBe(''));
  it('returns empty string for zero', () => expect(parseDateCell(0)).toBe(''));
  it('returns empty string for empty string', () => expect(parseDateCell('')).toBe(''));
  it('parses Date object', () => expect(parseDateCell(new Date(Date.UTC(2026, 5, 17)))).toBe('2026-06-17'));
  it('strips time from YYYY-MM-DDTHH:mm:ss', () => expect(parseDateCell('2026-06-17T00:00:00')).toBe('2026-06-17'));
});

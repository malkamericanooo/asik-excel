import { describe, it, expect } from 'vitest';
import { classifyKelurahan } from '../kelurahanMapping';

describe('classifyKelurahan', () => {
  it("maps MABU'UN exactly (apostrophe stripped)", () => expect(classifyKelurahan("MABU'UN")).toBe('MABUUN'));
  it('maps MABUUN uppercase', () => expect(classifyKelurahan('MABUUN')).toBe('MABUUN'));
  it('maps mabuun lowercase', () => expect(classifyKelurahan('mabuun')).toBe('MABUUN'));
  it('maps kasiau with trailing space', () => expect(classifyKelurahan('kasiau ')).toBe('KASIAU'));
  it('maps KASIAU', () => expect(classifyKelurahan('KASIAU')).toBe('KASIAU'));
  it('maps PEMBATAAN', () => expect(classifyKelurahan('PEMBATAAN')).toBe('PEMBATAAN'));
  it('maps SULINGAN', () => expect(classifyKelurahan('SULINGAN')).toBe('SULINGAN'));
  it('maps MABURAI', () => expect(classifyKelurahan('MABURAI')).toBe('MABURAI'));
  it('maps Mabuun Rt.2 to Luar Wilayah (not exact match)', () => expect(classifyKelurahan('Mabuun Rt.2')).toBe('LUAR WILAYAH'));
  it('maps BELIMBING RAYA to Luar Wilayah', () => expect(classifyKelurahan('BELIMBING RAYA')).toBe('LUAR WILAYAH'));
  it('maps mabun typo to Luar Wilayah', () => expect(classifyKelurahan('mabun')).toBe('LUAR WILAYAH'));
  it('maps Pembatan typo to Luar Wilayah', () => expect(classifyKelurahan('PEMBATAN')).toBe('LUAR WILAYAH'));
  it('maps unknown kelurahan to Luar Wilayah', () => expect(classifyKelurahan('JANGKUNG')).toBe('LUAR WILAYAH'));
  it('maps TANTA HULU to Luar Wilayah', () => expect(classifyKelurahan('TANTA HULU')).toBe('LUAR WILAYAH'));
  it('maps empty string to Luar Wilayah', () => expect(classifyKelurahan('')).toBe('LUAR WILAYAH'));

  // Additional edge cases for Belimbing & partial matches
  it('maps "Belimbing" to Luar Wilayah', () => expect(classifyKelurahan('Belimbing')).toBe('LUAR WILAYAH'));
  it('maps "belimbing" lowercase to Luar Wilayah', () => expect(classifyKelurahan('belimbing')).toBe('LUAR WILAYAH'));
  it('maps "Belimbing Raya Rt.14" to Luar Wilayah', () => expect(classifyKelurahan('Belimbing Raya Rt.14')).toBe('LUAR WILAYAH'));
  it('maps "MABUUN RT.10" to Luar Wilayah (not exact)', () => expect(classifyKelurahan('MABUUN RT.10')).toBe('LUAR WILAYAH'));
  it('maps null/undefined-like to Luar Wilayah', () => expect(classifyKelurahan('  ')).toBe('LUAR WILAYAH'));
  it('maps "Tanjung" to Luar Wilayah', () => expect(classifyKelurahan('Tanjung')).toBe('LUAR WILAYAH'));
  it('maps "Maburai Rt.5" to Luar Wilayah (not exact)', () => expect(classifyKelurahan('Maburai Rt.5')).toBe('LUAR WILAYAH'));
  it('maps mixed case "mAbUrAi" to MABURAI', () => expect(classifyKelurahan('mAbUrAi')).toBe('MABURAI'));
  it('maps "Pembataan " with trailing spaces to PEMBATAAN', () => expect(classifyKelurahan('  Pembataan  ')).toBe('PEMBATAAN'));
});

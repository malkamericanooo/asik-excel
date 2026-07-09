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
});

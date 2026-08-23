/**
 * Identitas anak — dipakai bersama oleh merger bulanan (asikParser) dan
 * database kumulatif (vaccineDatabase) supaya aturan deduplikasi tidak
 * pernah berbeda di antara keduanya.
 */

/** Normalisasi teks untuk pencocokan: huruf kecil, spasi ganda dirapikan. */
export function normalizeKey(s: string): string {
  return (s || '').toLowerCase().trim().replace(/\s+/g, ' ');
}

/**
 * Kunci identitas anak: nama ternormalisasi + tanggal lahir (YYYY-MM-DD).
 * Nama orang tua sengaja TIDAK dipakai — di ekspor ASIK kolom itu sering
 * kosong atau berbeda penulisannya untuk anak yang sama.
 */
export function childKey(nama: string, tglLahirStr: string): string {
  return `${normalizeKey(nama)}|${tglLahirStr}`;
}

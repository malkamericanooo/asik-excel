import type { ChildRecord } from '../types';

/**
 * Urutkan daftar anak dalam satu sheet master sesuai permintaan Puskesmas.
 *
 * Aturan urut:
 *   1. Tanggal lahir kronologis — yang paling tua di atas (serial Excel menaik).
 *   2. Tanggal lahir kosong ditaruh paling atas supaya langsung terlihat
 *      dan bisa dilengkapi nakes. (Parser sebenarnya sudah menolak baris
 *      tanpa tanggal lahir, ini jaring pengaman kalau data masuk lewat
 *      jalur lain.)
 *   3. Tanggal lahir sama → nama anak A→Z (abaikan besar-kecil huruf & aksen).
 *
 * Tidak memutasi array input — nomor urut kolom A, blok ringkasan, dan
 * posisi merge di master Excel semuanya dihitung ulang dari hasil sort ini.
 */
export function sortChildrenByBirthDate(children: ChildRecord[]): ChildRecord[] {
  return [...children].sort((a, b) => {
    const sa = a.tanggalLahirSerial || 0;
    const sb = b.tanggalLahirSerial || 0;

    if (sa !== sb) {
      if (!sa) return -1; // tanggal kosong → paling atas
      if (!sb) return 1;
      return sa - sb;
    }

    return (a.nama || '').localeCompare(b.nama || '', 'id', { sensitivity: 'base' });
  });
}

/** Terapkan urutan yang sama ke gabungan semua sheet (dipakai pratinjau UI). */
export function sortChildrenGroups(groups: ChildRecord[][]): ChildRecord[] {
  return groups.flatMap((g) => sortChildrenByBirthDate(g));
}

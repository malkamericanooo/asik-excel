# Imunisasi Master Merger

Aplikasi web untuk menggabungkan file-file imunisasi ASIK (per jenis vaksin) menjadi satu file Master Excel — laporan bulanan Puskesmas Mabuun.

## Deploy ke Vercel

1. Push folder ini ke repository GitHub
2. Import repo di [vercel.com](https://vercel.com)
3. Framework: **Vite** — Vercel auto-detect
4. Deploy

Tidak perlu env var apapun — aplikasi 100% client-side.

## Menjalankan secara lokal

```bash
npm install
npm run dev
```

## Menjalankan test

```bash
npm test
```

## Struktur

```
src/
├── App.tsx                       # UI utama
├── types.ts                      # Tipe data
├── index.css                     # Tailwind CSS
├── main.tsx                      # Entry point
└── utils/
    ├── vaccineMapping.ts         # Nama Antigen ASIK → kolom vaksin
    ├── kelurahanMapping.ts       # Kelurahan → sheet (MABUUN/KASIAU/dll)
    ├── dateUtils.ts              # Konversi tanggal Excel serial ↔ string
    ├── asikParser.ts             # Parse file ASIK + merge + deduplikasi
    ├── masterExcel.ts            # Generate file Master Excel (.xlsx)
    └── __tests__/                # 76 unit tests (Vitest)
```

## Business rules

- **Deduplikasi**: Nama Anak (case-insensitive) + Tanggal Lahir + Nama Orang Tua
- **Klasifikasi sheet**: `Status = kejar` → Kejar; sisanya → kelurahan fuzzy-match ke 5 wilayah; tidak cocok → Luar Wilayah
- **L/P per vaksin**: JK Laki-laki → kolom L; Perempuan → kolom P
- **IBL (Baduta Lengkap)**: dilewati, tidak ada kolom master
- **Ringkasan bulanan**: hitung per kolom L/P yang tanggal imunisasinya cocok bulan/tahun laporan

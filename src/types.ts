export type VaccineKey =
  | 'HB0_24JAM' | 'HB0_7HARI'
  | 'BCG'
  | 'POLIO_1' | 'POLIO_2' | 'POLIO_3' | 'POLIO_4'
  | 'DPT_1' | 'DPT_2' | 'DPT_3' | 'DPT_4'
  | 'PCV_1' | 'PCV_2' | 'PCV_3'
  | 'ROTA_1' | 'ROTA_2' | 'ROTA_3'
  | 'IPV_1' | 'IPV_2'
  | 'MR_1' | 'BOOSTER_MR';

export type SheetName = 'MABUUN' | 'KASIAU' | 'PEMBATAAN' | 'SULINGAN' | 'MABURAI' | 'LUAR WILAYAH' | 'Kejar';

export const ALL_SHEETS: SheetName[] = ['MABUUN', 'KASIAU', 'PEMBATAAN', 'SULINGAN', 'MABURAI', 'LUAR WILAYAH', 'Kejar'];

export interface ChildRecord {
  nama: string;
  jk: 'L' | 'P';
  tanggalLahirSerial: number;   // Excel serial
  tanggalLahirStr: string;       // YYYY-MM-DD
  nik: string;
  namaOrangTua: string;
  alamat: string;                // Kelurahan / Desa value
  vaccines: Partial<Record<VaccineKey, number>>; // VaccineKey → Excel serial date
}

export type MasterData = Record<SheetName, ChildRecord[]>;

export interface UploadLogEntry {
  id: string;
  fileName: string;
  antigen: string;
  processedAt: string;
  dataCount: number;
  status: 'success' | 'error' | 'warning';
  message?: string;
}

export interface ProcessResult {
  added: number;
  updated: number;
  moved: number;
  skipped: number;
  logs: string[];
}

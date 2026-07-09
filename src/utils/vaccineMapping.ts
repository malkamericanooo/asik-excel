import type { VaccineKey } from '../types';

export const VACCINE_COLUMN_INDEX: Record<VaccineKey, number> = {
  HB0_24JAM: 7,
  HB0_7HARI: 9,
  BCG: 11,
  POLIO_1: 13,
  DPT_1: 15,
  POLIO_2: 17,
  PCV_1: 19,
  ROTA_1: 21,
  DPT_2: 23,
  POLIO_3: 25,
  PCV_2: 27,
  ROTA_2: 29,
  DPT_3: 31,
  POLIO_4: 33,
  IPV_1: 35,
  ROTA_3: 37,
  MR_1: 39,
  IPV_2: 41,
  PCV_3: 43,
  DPT_4: 45,
  BOOSTER_MR: 47,
};

export const VACCINE_COLUMN_LABELS: Record<VaccineKey, string> = {
  HB0_24JAM: 'HB0 (<24 JAM)',
  HB0_7HARI: 'HB0 (1-7 HARI)',
  BCG: 'BCG',
  POLIO_1: 'POLIO (1)',
  DPT_1: 'DPT/HB-Hib (1)',
  POLIO_2: 'POLIO (2)',
  PCV_1: 'PCV (1)',
  ROTA_1: 'ROTA (1)',
  DPT_2: 'DPT/HB-Hib (2)',
  POLIO_3: 'POLIO (3)',
  PCV_2: 'PCV (2)',
  ROTA_2: 'ROTA (2)',
  DPT_3: 'DPT/HB-Hib (3)',
  POLIO_4: 'POLIO (4)',
  IPV_1: 'IPV (1)',
  ROTA_3: 'ROTA (3)',
  MR_1: 'Campak-Rubella (MR)',
  IPV_2: 'Ipv (2)',
  PCV_3: 'PCV (3)',
  DPT_4: 'Booster DPT/HB/Hib',
  BOOSTER_MR: 'Booster MR',
};

export const VACCINE_ORDER: VaccineKey[] = [
  'HB0_24JAM', 'HB0_7HARI', 'BCG', 'POLIO_1', 'DPT_1', 'POLIO_2',
  'PCV_1', 'ROTA_1', 'DPT_2', 'POLIO_3', 'PCV_2', 'ROTA_2',
  'DPT_3', 'POLIO_4', 'IPV_1', 'ROTA_3', 'MR_1', 'IPV_2',
  'PCV_3', 'DPT_4', 'BOOSTER_MR',
];

/** Display names for the vaccine dropdown selector */
export const VACCINE_DISPLAY_NAMES: Record<VaccineKey, string> = {
  HB0_24JAM: 'HB0 < 24 Jam',
  HB0_7HARI: 'HB0 1-7 Hari',
  BCG: 'BCG',
  POLIO_1: 'Polio 1',
  DPT_1: 'DPT/HB/Hib 1',
  POLIO_2: 'Polio 2',
  PCV_1: 'PCV 1',
  ROTA_1: 'Rotavirus 1',
  DPT_2: 'DPT/HB/Hib 2',
  POLIO_3: 'Polio 3',
  PCV_2: 'PCV 2',
  ROTA_2: 'Rotavirus 2',
  DPT_3: 'DPT/HB/Hib 3',
  POLIO_4: 'Polio 4',
  IPV_1: 'IPV 1',
  ROTA_3: 'Rotavirus 3',
  MR_1: 'Campak-Rubella (MR)',
  IPV_2: 'IPV 2 (Baduta)',
  PCV_3: 'PCV 3',
  DPT_4: 'Booster DPT/HB/Hib',
  BOOSTER_MR: 'Booster MR',
};

/**
 * Map ASIK "Nama Antigen" string to our VaccineKey.
 * Returns null if the antigen cannot be mapped (e.g. IBL).
 */
export function mapAntigenToVaccineKey(antigen: string): VaccineKey | null {
  const a = antigen.toUpperCase().replace(/\s+/g, ' ').trim();

  // BCG
  if (a.startsWith('BCG')) return 'BCG';

  // HB0 — distinguish <24 jam (dose 1) vs 1-7 hari (dose 2)
  if (a.includes('HB0') || a.includes('HB 0') || a.includes('HEPATITIS B - 0') || a.includes('HEPATITIS B-0')) {
    if (a.includes('- 2') || a.endsWith(' 2') || a.includes('1-7') || a.includes('7 HARI')) return 'HB0_7HARI';
    return 'HB0_24JAM';
  }

  // DPT/HB-Hib (check highest number first)
  if (a.includes('DPT') || a.includes('DPTHB')) {
    if (a.includes('- 4') || a.endsWith(' 4')) return 'DPT_4';
    if (a.includes('- 3') || a.endsWith(' 3')) return 'DPT_3';
    if (a.includes('- 2') || a.endsWith(' 2')) return 'DPT_2';
    if (a.includes('- 1') || a.endsWith(' 1')) return 'DPT_1';
    return 'DPT_1';
  }

  // IPV
  if (a.includes('IPV')) {
    if (a.includes('- 2') || a.endsWith(' 2')) return 'IPV_2';
    return 'IPV_1';
  }

  // PCV
  if (a.includes('PCV')) {
    if (a.includes('- 3') || a.endsWith(' 3')) return 'PCV_3';
    if (a.includes('- 2') || a.endsWith(' 2')) return 'PCV_2';
    return 'PCV_1';
  }

  // Rotavirus / ROTA
  if (a.includes('ROTA') || a.includes('ROTAVIRUS')) {
    if (a.includes('- 3') || a.endsWith(' 3')) return 'ROTA_3';
    if (a.includes('- 2') || a.endsWith(' 2')) return 'ROTA_2';
    return 'ROTA_1';
  }

  // POLIO
  if (a.includes('POLIO')) {
    if (a.includes('- 4') || a.endsWith(' 4')) return 'POLIO_4';
    if (a.includes('- 3') || a.endsWith(' 3')) return 'POLIO_3';
    if (a.includes('- 2') || a.endsWith(' 2')) return 'POLIO_2';
    return 'POLIO_1';
  }

  // MR / Campak-Rubella
  if (a.includes('CAMPAK') || a.includes('RUBELLA')) {
    if (a.includes('BOOSTER') || a.includes('- 2') || a.endsWith(' 2')) return 'BOOSTER_MR';
    return 'MR_1';
  }
  if (a.match(/\bMR\b/)) {
    if (a.includes('BOOSTER') || a.includes('- 2') || a.endsWith(' 2')) return 'BOOSTER_MR';
    return 'MR_1';
  }

  // IBL (Imunisasi Baduta Lengkap) – skip
  if (a.includes('BADUTA') || a.includes('IBL')) return null;

  return null;
}

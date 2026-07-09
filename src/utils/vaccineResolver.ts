import type { VaccineKey } from '../types';
import { mapAntigenToVaccineKey } from './vaccineMapping';

/** Map dropdown vaccine key to ASIK antigen label for simple-format files. */
const VACCINE_TO_ANTIGEN: Partial<Record<VaccineKey, string>> = {
  HB0_24JAM: 'Hepatitis B - 0 - 1',
  HB0_7HARI: 'Hepatitis B - 0 - 2',
  BCG: 'BCG',
  POLIO_1: 'Polio - 1',
  DPT_1: 'DPT-HB-Hib - 1',
  POLIO_2: 'Polio - 2',
  PCV_1: 'PCV - 1',
  ROTA_1: 'Rotavirus - 1',
  DPT_2: 'DPT-HB-Hib - 2',
  POLIO_3: 'Polio - 3',
  PCV_2: 'PCV - 2',
  ROTA_2: 'Rotavirus - 2',
  DPT_3: 'DPT-HB-Hib - 3',
  POLIO_4: 'Polio - 4',
  IPV_1: 'IPV - 1',
  ROTA_3: 'Rotavirus - 3',
  MR_1: 'Campak-Rubella (MR)',
  IPV_2: 'IPV - 2',
  PCV_3: 'PCV - 3',
  DPT_4: 'DPT-HB-Hib - 4',
  BOOSTER_MR: 'Campak-Rubella (MR) - 2',
};

export function antigenFromVaccineKey(vk: VaccineKey): string {
  return VACCINE_TO_ANTIGEN[vk] ?? vk;
}

export function resolveVaccineKey(
  antigenRaw: string,
  selectedVaccine?: VaccineKey,
): VaccineKey | null {
  if (antigenRaw) {
    return mapAntigenToVaccineKey(antigenRaw);
  }
  return selectedVaccine ?? null;
}

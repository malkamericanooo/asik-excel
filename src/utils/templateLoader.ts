const TEMPLATE_PATH = '/templates/master-template.xlsx';

/** Load the default master template bundled in /public/templates/ */
export async function loadDefaultTemplate(): Promise<ArrayBuffer> {
  const res = await fetch(TEMPLATE_PATH, { cache: 'force-cache' });
  if (!res.ok) {
    throw new Error(
      `Template Master tidak ditemukan (${res.status}). Upload template manual di langkah 2.`,
    );
  }
  return res.arrayBuffer();
}

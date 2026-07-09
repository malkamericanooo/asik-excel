/** Load the default master template bundled in /public/templates/ */
export async function loadDefaultTemplate(): Promise<ArrayBuffer> {
  const res = await fetch('/templates/master-template.xlsx');
  if (!res.ok) throw new Error('Gagal memuat template Master default.');
  return res.arrayBuffer();
}

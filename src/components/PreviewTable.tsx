import type { ChildRecord } from '../types';

interface PreviewTableProps {
  children: ChildRecord[];
  maxRows?: number;
  title?: string;
}

const COLUMNS = [
  { key: 'nama', label: 'Nama Anak' },
  { key: 'jk', label: 'JK' },
  { key: 'tanggalLahirStr', label: 'Tgl Lahir' },
  { key: 'nik', label: 'NIK' },
  { key: 'namaOrangTua', label: 'Orang Tua' },
  { key: 'alamat', label: 'Alamat' },
] as const;

export function PreviewTable({ children, maxRows = 5, title }: PreviewTableProps) {
  if (!children || children.length === 0) return null;

  const displayRows = children.slice(0, maxRows);
  const totalSheets = new Set(children.map((c) => c.alamat)).size;

  return (
    <section className="bagian">
      <div className="bagian-kepala">
        <h2>{title ?? 'Pratinjau Data'}</h2>
        <span className="bagian-info">
          Menampilkan {Math.min(displayRows.length, maxRows)} dari {children.length} anak
          {totalSheets > 1 && ` · ${totalSheets} wilayah`}
        </span>
      </div>

      <div className="tabel-wadah">
        <table className="tabel">
          <thead>
            <tr>
              <th className="w-8">#</th>
              {COLUMNS.map((col) => (
                <th key={col.key}>{col.label}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {displayRows.map((child, idx) => (
              <tr key={`${child.nama}-${child.tanggalLahirStr}`}>
                <td className="redup angka">{idx + 1}</td>
                {COLUMNS.map((col) => {
                  const val = child[col.key as keyof ChildRecord];
                  const display = val != null && val !== '' ? String(val) : '—';
                  return (
                    <td
                      key={col.key}
                      className={`max-w-[180px] truncate ${display === '—' ? 'redup' : ''} ${col.key === 'nik' ? 'nik' : ''}`}
                      title={display !== '—' ? display : undefined}
                    >
                      {display}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {children.length > maxRows && (
        <p className="kecil mt-2">{children.length - maxRows} data lainnya tidak ditampilkan, semuanya ikut di file master.</p>
      )}
    </section>
  );
}

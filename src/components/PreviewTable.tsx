import { Eye, ChevronRight } from 'lucide-react';
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
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
      <div className="flex items-center gap-2 px-4 py-3 border-b border-gray-100 bg-gray-50/50">
        <Eye className="w-4 h-4 text-blue-500" />
        <h3 className="text-sm font-semibold text-gray-700">
          {title ?? 'Pratinjau Data'}
        </h3>
        <span className="ml-auto text-xs text-gray-400">
          Menampilkan {Math.min(displayRows.length, maxRows)} dari {children.length} anak
          {totalSheets > 1 && ` · ${totalSheets} wilayah`}
        </span>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-xs">
          <thead>
            <tr className="bg-gray-50 border-b border-gray-100">
              <th className="text-left px-3 py-2 font-medium text-gray-500 w-8">#</th>
              {COLUMNS.map((col) => (
                <th key={col.key} className="text-left px-3 py-2 font-medium text-gray-500 whitespace-nowrap">
                  {col.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {displayRows.map((child, idx) => (
              <tr key={`${child.nama}-${child.tanggalLahirStr}`} className="hover:bg-blue-50/30 transition-colors">
                <td className="px-3 py-2 text-gray-400">{idx + 1}</td>
                {COLUMNS.map((col) => {
                  const val = child[col.key as keyof ChildRecord];
                  const display = val != null && val !== '' ? String(val) : '—';
                  return (
                    <td
                      key={col.key}
                      className={`px-3 py-2 text-gray-700 max-w-[180px] truncate ${
                        display === '—' ? 'text-gray-300 italic' : ''
                      }`}
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
        <div className="px-4 py-2 border-t border-gray-100 bg-gray-50/30">
          <p className="text-xs text-gray-400 flex items-center gap-1">
            <ChevronRight className="w-3 h-3" />
            {children.length - maxRows} data lainnya tidak ditampilkan
          </p>
        </div>
      )}
    </div>
  );
}

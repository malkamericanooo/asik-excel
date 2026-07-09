import { CheckCircle2, AlertTriangle, XCircle, ClipboardList } from 'lucide-react';

export type ValidationStatus = 'valid' | 'empty' | 'invalid';

export interface ColumnValidation {
  column: string;
  emptyCount: number;
  invalidCount: number;
  status: ValidationStatus;
}

interface ValidationTableProps {
  validations: ColumnValidation[];
  totalRows: number;
  validRows: number;
  title?: string;
}

function StatusBadge({ status }: { status: ValidationStatus }) {
  const config = {
    valid: {
      label: '✅ Lengkap & Valid',
      bg: 'bg-emerald-100 text-emerald-700',
      dot: 'bg-emerald-500',
    },
    empty: {
      label: '⚠️ Ada Data Kosong',
      bg: 'bg-amber-100 text-amber-700',
      dot: 'bg-amber-500',
    },
    invalid: {
      label: '❌ Format Salah',
      bg: 'bg-red-100 text-red-700',
      dot: 'bg-red-500',
    },
  };

  const c = config[status];
  return (
    <span className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-lg text-xs font-medium ${c.bg}`}>
      <span className={`w-1.5 h-1.5 rounded-full ${c.dot}`} />
      {c.label}
    </span>
  );
}

export function ValidationTable({ validations, totalRows, validRows, title }: ValidationTableProps) {
  if (validations.length === 0) return null;

  const hasIssues = validations.some((v) => v.status !== 'valid');

  return (
    <div className="bg-white rounded-xl border border-gray-100 shadow-sm overflow-hidden">
      <div className="flex items-center gap-2 px-4 py-3 border-b border-gray-100 bg-gray-50/50">
        <ClipboardList className="w-4 h-4 text-blue-500" />
        <h3 className="text-sm font-semibold text-gray-700">
          {title ?? 'Validasi Data'}
        </h3>
        <span className="ml-auto text-xs text-gray-400">
          {validRows}/{totalRows} baris valid
          {hasIssues && (
            <span className="ml-1 text-amber-600">· perlu diperiksa</span>
          )}
        </span>
      </div>

      <div className="overflow-x-auto">
        <table className="w-full text-xs">
          <thead>
            <tr className="bg-gray-50 border-b border-gray-100">
              <th className="text-left px-3 py-2 font-medium text-gray-500">Kolom</th>
              <th className="text-center px-3 py-2 font-medium text-gray-500">Data Kosong</th>
              <th className="text-center px-3 py-2 font-medium text-gray-500">Format Salah</th>
              <th className="text-left px-3 py-2 font-medium text-gray-500">Status</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-gray-50">
            {validations.map((v) => (
              <tr key={v.column} className="hover:bg-gray-50 transition-colors">
                <td className="px-3 py-2.5 font-medium text-gray-700">{v.column}</td>
                <td className="px-3 py-2.5 text-center">
                  {v.emptyCount > 0 ? (
                    <span className="inline-flex items-center gap-1 text-amber-600 font-medium">
                      <AlertTriangle className="w-3.5 h-3.5" />
                      {v.emptyCount}
                    </span>
                  ) : (
                    <span className="text-gray-400">0</span>
                  )}
                </td>
                <td className="px-3 py-2.5 text-center">
                  {v.invalidCount > 0 ? (
                    <span className="inline-flex items-center gap-1 text-red-600 font-medium">
                      <XCircle className="w-3.5 h-3.5" />
                      {v.invalidCount}
                    </span>
                  ) : (
                    <span className="text-gray-400">0</span>
                  )}
                </td>
                <td className="px-3 py-2.5">
                  <StatusBadge status={v.status} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {hasIssues && (
        <div className="px-4 py-2.5 bg-amber-50 border-t border-amber-100">
          <p className="text-xs text-amber-700 flex items-center gap-1.5">
            <AlertTriangle className="w-3.5 h-3.5 flex-shrink-0" />
            Data dengan masalah akan tetap diproses. Periksa kembali jika diperlukan.
          </p>
        </div>
      )}
    </div>
  );
}

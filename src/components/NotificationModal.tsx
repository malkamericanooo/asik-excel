import { AlertTriangle, X, Upload, Download } from 'lucide-react';

interface NotificationModalProps {
  visible: boolean;
  title?: string;
  message: string;
  detail?: string;
  variant?: 'warning' | 'error' | 'info';
  confirmLabel?: string;
  cancelLabel?: string;
  onConfirm?: () => void;
  onCancel?: () => void;
  onClose: () => void;
}

export function NotificationModal({
  visible,
  title,
  message,
  detail,
  variant = 'warning',
  confirmLabel = 'Lanjutkan',
  cancelLabel = 'Batal',
  onConfirm,
  onCancel,
  onClose,
}: NotificationModalProps) {
  if (!visible) return null;

  const colorMap = {
    warning: {
      bg: 'bg-amber-50',
      border: 'border-amber-200',
      icon: 'text-amber-500',
      iconBg: 'bg-amber-100',
      btn: 'bg-amber-600 hover:bg-amber-700',
      title: 'text-amber-800',
    },
    error: {
      bg: 'bg-red-50',
      border: 'border-red-200',
      icon: 'text-red-500',
      iconBg: 'bg-red-100',
      btn: 'bg-red-600 hover:bg-red-700',
      title: 'text-red-800',
    },
    info: {
      bg: 'bg-blue-50',
      border: 'border-blue-200',
      icon: 'text-blue-500',
      iconBg: 'bg-blue-100',
      btn: 'bg-blue-600 hover:bg-blue-700',
      title: 'text-blue-800',
    },
  };

  const c = colorMap[variant];

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4">
      <div
        className={`w-full max-w-md rounded-2xl border ${c.border} ${c.bg} shadow-2xl animate-in fade-in zoom-in-95 duration-200`}
      >
        <div className="flex items-start justify-between p-5 pb-3">
          <div className="flex items-start gap-3">
            <div className={`p-2 rounded-xl ${c.iconBg}`}>
              {variant === 'error' ? (
                <X className={`w-5 h-5 ${c.icon}`} />
              ) : (
                <AlertTriangle className={`w-5 h-5 ${c.icon}`} />
              )}
            </div>
            <div>
              <h3 className={`font-semibold text-sm ${c.title}`}>
                {title ?? (variant === 'warning' ? 'Perhatian' : variant === 'error' ? 'Error' : 'Informasi')}
              </h3>
              <p className="text-sm text-gray-600 mt-1 max-w-sm">{message}</p>
              {detail && (
                <p className="text-xs text-gray-500 mt-2 bg-white/60 rounded-lg p-2 border border-gray-100">
                  {detail}
                </p>
              )}
            </div>
          </div>
          <button onClick={onClose} className="p-1 rounded-lg hover:bg-white/60 transition-colors">
            <X className="w-4 h-4 text-gray-400" />
          </button>
        </div>

        <div className="flex gap-2 px-5 pb-5 pt-2">
          {onCancel && (
            <button
              onClick={onCancel}
              className="flex-1 py-2.5 rounded-xl text-sm font-medium border border-gray-200 text-gray-600 hover:bg-white active:scale-95 transition-all"
            >
              {cancelLabel}
            </button>
          )}
          {onConfirm && (
            <button
              onClick={onConfirm}
              className={`flex-1 py-2.5 rounded-xl text-sm font-semibold text-white active:scale-95 transition-all flex items-center justify-center gap-2 ${c.btn}`}
            >
              {variant === 'warning' ? <Upload className="w-4 h-4" /> : <Download className="w-4 h-4" />}
              {confirmLabel}
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

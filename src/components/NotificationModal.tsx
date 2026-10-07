import { X, Upload, Download } from 'lucide-react';

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

  const tanda = variant === 'error' ? 'galat' : variant === 'warning' ? 'peringatan' : 'catatan';

  return (
    <div className="lapis" role="dialog" aria-modal="true">
      <div className="dialog">
        <div className="dialog-kepala">
          <div>
            <p className="kop-instansi">{variant === 'warning' ? 'Perhatian' : variant === 'error' ? 'Error' : 'Informasi'}</p>
            <h3>{title ?? (variant === 'warning' ? 'Perhatian' : variant === 'error' ? 'Error' : 'Informasi')}</h3>
          </div>
          <button onClick={onClose} className="tutup" aria-label="Tutup">
            <X className="w-5 h-5" />
          </button>
        </div>
        <p className={tanda}>{message}</p>
        {detail && <p className="catatan">{detail}</p>}

        {(onCancel || onConfirm) && (
          <div className="dialog-aksi">
            {onCancel && (
              <button onClick={onCancel} className="tombol tombol--garis">
                {cancelLabel}
              </button>
            )}
            {onConfirm && (
              <button onClick={onConfirm} className="tombol tombol--utama">
                {variant === 'warning' ? <Upload className="w-4 h-4" /> : <Download className="w-4 h-4" />}
                {confirmLabel}
              </button>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

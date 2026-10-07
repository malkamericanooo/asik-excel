import { useCallback } from 'react';
import { useDropzone } from 'react-dropzone';
import { Upload, FileSpreadsheet, X, Loader2 } from 'lucide-react';

interface FileDropzoneProps {
  onFilesAccepted: (files: File[]) => void;
  disabled?: boolean;
  isProcessing?: boolean;
  accept?: Record<string, string[]>;
  multiple?: boolean;
  label?: string;
  hint?: string;
}

function formatFileSize(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function FileDropzone({
  onFilesAccepted,
  disabled = false,
  isProcessing = false,
  accept = { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'], 'application/vnd.ms-excel': ['.xls'] },
  multiple = false,
  label = 'Upload File',
  hint,
}: FileDropzoneProps) {
  const handleDrop = useCallback(
    (acceptedFiles: File[]) => {
      if (acceptedFiles.length > 0) {
        onFilesAccepted(acceptedFiles);
      }
    },
    [onFilesAccepted],
  );

  const { getRootProps, getInputProps, isDragActive, isDragReject, acceptedFiles } = useDropzone({
    onDrop: handleDrop,
    accept,
    multiple,
    disabled: disabled || isProcessing,
    maxSize: 50 * 1024 * 1024, // 50MB
  });

  if (disabled) {
    return (
      <div className="unggah unggah--mati">
        <p className="unggah-judul">Konfirmasi periode terlebih dahulu</p>
        <p className="unggah-sub">Pilih bulan dan tahun di langkah 1, lalu tekan Konfirmasi.</p>
      </div>
    );
  }

  return (
    <div className="space-y-2">
      <div
        {...getRootProps()}
        className={[
          'unggah',
          isDragActive && !isDragReject ? 'unggah--seret' : isDragReject ? 'unggah--tolak' : '',
          isProcessing ? 'opacity-70 cursor-wait' : '',
        ].join(' ')}
      >
        <input {...getInputProps()} />
        <div className="flex items-start gap-4">
          {isProcessing ? (
            <Loader2 className="w-7 h-7 mt-1 text-blue-600 animate-spin flex-shrink-0" />
          ) : isDragReject ? (
            <X className="w-7 h-7 mt-1 text-red-600 flex-shrink-0" />
          ) : (
            <Upload className="w-7 h-7 mt-1 text-blue-600 flex-shrink-0" />
          )}
          <div>
            <p className="unggah-judul">
              {isProcessing
                ? 'Memproses file...'
                : isDragActive
                  ? isDragReject ? 'Format file tidak didukung' : 'Lepaskan file di sini...'
                  : label}
            </p>
            {hint && !isProcessing && !isDragActive && <p className="unggah-sub">{hint}</p>}
          </div>
        </div>
      </div>

      {acceptedFiles.length > 0 && !isProcessing && (
        <ul className="baris-daftar">
          {acceptedFiles.map((file) => (
            <li key={file.name}>
              <FileSpreadsheet className="w-4 h-4 text-blue-600 flex-shrink-0" />
              <span className="text-sm truncate flex-1">{file.name}</span>
              <span className="kecil angka flex-shrink-0">{formatFileSize(file.size)}</span>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

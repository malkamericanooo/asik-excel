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
      <div className="w-full py-6 rounded-xl border-2 border-dashed border-gray-200 bg-gray-50 flex flex-col items-center justify-center gap-2 cursor-not-allowed">
        <Upload className="w-8 h-8 text-gray-300" />
        <p className="text-sm text-gray-400">Konfirmasi periode terlebih dahulu</p>
      </div>
    );
  }

  return (
    <div className="space-y-2">
      <div
        {...getRootProps()}
        className={[
          'w-full py-6 px-4 rounded-xl border-2 border-dashed transition-all duration-200 flex flex-col items-center justify-center gap-2 cursor-pointer',
          isDragActive && !isDragReject
            ? 'border-blue-400 bg-blue-50 scale-[1.01]'
            : isDragReject
              ? 'border-red-400 bg-red-50'
              : 'border-gray-200 hover:border-blue-300 hover:bg-gray-50',
          isProcessing ? 'opacity-70 cursor-wait' : '',
        ].join(' ')}
      >
        <input {...getInputProps()} />

        {isProcessing ? (
          <Loader2 className="w-8 h-8 text-blue-400 animate-spin" />
        ) : isDragReject ? (
          <X className="w-8 h-8 text-red-400" />
        ) : isDragActive ? (
          <Upload className="w-8 h-8 text-blue-500" />
        ) : (
          <FileSpreadsheet className="w-8 h-8 text-gray-400" />
        )}

        {isProcessing ? (
          <p className="text-sm text-blue-600 font-medium">Memproses file...</p>
        ) : isDragActive ? (
          <p className="text-sm text-blue-600 font-medium">
            {isDragReject ? 'Format file tidak didukung' : 'Lepaskan file di sini...'}
          </p>
        ) : (
          <>
            <p className="text-sm font-medium text-gray-600">
              {label}
            </p>
            {hint && <p className="text-xs text-gray-400">{hint}</p>}
          </>
        )}
      </div>

      {acceptedFiles.length > 0 && !isProcessing && (
        <div className="space-y-1.5">
          {acceptedFiles.map((file) => (
            <div
              key={file.name}
              className="flex items-center gap-2 px-3 py-2 bg-blue-50 rounded-lg border border-blue-100"
            >
              <FileSpreadsheet className="w-4 h-4 text-blue-500 flex-shrink-0" />
              <span className="text-sm text-gray-700 truncate flex-1">{file.name}</span>
              <span className="text-xs text-gray-400 flex-shrink-0">
                {formatFileSize(file.size)}
              </span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

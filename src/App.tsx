import { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import { Download, RotateCcw, Loader2 } from 'lucide-react';
import type { MasterData, UploadLogEntry, ProcessResult, VaccineKey, ChildRecord } from './types';
import { ALL_SHEETS } from './types';
import { createEmptyMasterData, parseAndMergeAsikFile } from './utils/asikParser';
import { buildMasterExcel, getUploadedVaccines } from './utils/masterExcel';
import { sortChildrenGroups } from './utils/sortChildren';
import { loadDefaultTemplate } from './utils/templateLoader';
import { downloadBlob } from './utils/downloadFile';
import { BULAN_INDONESIA } from './utils/dateUtils';
import { VACCINE_DISPLAY_NAMES, VACCINE_ORDER } from './utils/vaccineMapping';
import { FileDropzone } from './components/FileDropzone';
import { PreviewTable } from './components/PreviewTable';
import { NotificationModal } from './components/NotificationModal';
import { StepIndicator } from './components/StepIndicator';

const CURRENT_YEAR = new Date().getFullYear();
const YEARS = Array.from({ length: 5 }, (_, i) => CURRENT_YEAR - 2 + i);

const VACCINE_GROUPS: { label: string; keys: VaccineKey[] }[] = [
  { label: 'Bayi Baru Lahir', keys: ['HB0_24JAM', 'HB0_7HARI', 'BCG'] },
  { label: 'Usia 1–3 Bulan', keys: ['POLIO_1', 'DPT_1', 'POLIO_2', 'PCV_1', 'ROTA_1'] },
  { label: 'Usia 3–4 Bulan', keys: ['DPT_2', 'POLIO_3', 'PCV_2', 'ROTA_2'] },
  { label: 'Usia 4–9 Bulan', keys: ['DPT_3', 'POLIO_4', 'IPV_1', 'ROTA_3', 'MR_1'] },
  { label: 'Baduta (≥ 9 bln)', keys: ['IPV_2', 'PCV_3', 'DPT_4', 'BOOSTER_MR'] },
];

const WORKFLOW_STEPS = [
  { number: 1, label: 'Periode', description: 'Bulan & Tahun' },
  { number: 2, label: 'Upload', description: 'File Vaksin' },
  { number: 3, label: 'Export', description: 'Master Excel' },
];

function App() {
  const [month, setMonth] = useState<number>(new Date().getMonth() + 1);
  const [year, setYear] = useState<number>(CURRENT_YEAR);
  const [periodLocked, setPeriodLocked] = useState(false);
  const [masterData, setMasterData] = useState<MasterData>(createEmptyMasterData());
  const [logs, setLogs] = useState<UploadLogEntry[]>([]);
  const [selectedVaccine, setSelectedVaccine] = useState<string>(VACCINE_ORDER[0]);
  const [isProcessing, setIsProcessing] = useState(false);
  const [templateBuffer, setTemplateBuffer] = useState<ArrayBuffer | null>(null);
  const [templateName, setTemplateName] = useState('Template default');
  const [templateError, setTemplateError] = useState<string | null>(null);
  const [isExporting, setIsExporting] = useState(false);
  const [dataCount, setDataCount] = useState<Record<string, number>>(
    Object.fromEntries(ALL_SHEETS.map((s) => [s, 0])),
  );
  const [showValidationModal, setShowValidationModal] = useState(false);
  const [validationMessage, setValidationMessage] = useState('');
  const [pendingExport, setPendingExport] = useState(false);

  const templateFileRef = useRef<HTMLInputElement>(null);
  const uploadedVaccines = useMemo(() => getUploadedVaccines(masterData), [masterData]);
  const totalChildren = ALL_SHEETS.reduce((sum, s) => sum + (dataCount[s] ?? 0), 0);
  const uploadedCount = uploadedVaccines.size;
  const totalVaccines = VACCINE_ORDER.length;
  const canDownload = totalChildren > 0 && templateBuffer !== null && !isExporting;

  // Pratinjau memakai urutan yang sama persis dengan file master hasil export.
  const allChildren = useMemo(
    () => sortChildrenGroups(ALL_SHEETS.map((sheet) => masterData[sheet])),
    [masterData],
  );

  /** Derive current workflow step from state */
  const currentStep = useMemo(() => {
    if (!periodLocked) return 1;
    if (totalChildren === 0) return 2;
    return 3;
  }, [periodLocked, totalChildren]);

  useEffect(() => {
    loadDefaultTemplate()
      .then((buf) => {
        setTemplateBuffer(buf);
        setTemplateError(null);
      })
      .catch((err) => {
        console.error('Gagal memuat template:', err);
        setTemplateError(
          err instanceof Error ? err.message : 'Gagal memuat template Master default.',
        );
      });
  }, []);

  const handleTemplateUpload = useCallback(async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const buffer = await file.arrayBuffer();
      setTemplateBuffer(buffer);
      setTemplateName(file.name);
      setTemplateError(null);
    } catch (err) {
      alert(`Gagal memuat template: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      e.target.value = '';
    }
  }, []);

  const processVaccineFile = useCallback(
    async (file: File, currentMaster: MasterData): Promise<{ master: MasterData; log: UploadLogEntry }> => {
      const buffer = await file.arrayBuffer();
      const result: ProcessResult = { added: 0, updated: 0, moved: 0, skipped: 0, logs: [] };
      const newMaster = JSON.parse(JSON.stringify(currentMaster)) as MasterData;
      parseAndMergeAsikFile(buffer, newMaster, result, {
        selectedVaccine: selectedVaccine as VaccineKey,
      });
      const total = result.added + result.updated + result.moved;
      return {
        master: newMaster,
        log: {
          id: `${Date.now()}-${file.name}`,
          fileName: file.name,
          antigen:
            VACCINE_DISPLAY_NAMES[selectedVaccine as keyof typeof VACCINE_DISPLAY_NAMES] ??
            selectedVaccine,
          processedAt: new Date().toLocaleTimeString('id-ID'),
          dataCount: total,
          status: result.skipped > 0 && total === 0 ? 'error' : 'success',
          message: [
            `+${result.added} baru`,
            result.updated > 0 ? `~${result.updated} diperbarui` : '',
            result.moved > 0 ? `↗${result.moved} dipindah` : '',
            result.skipped > 0 ? `⊘${result.skipped} dilewati` : '',
            ...result.logs.slice(0, 3),
          ]
            .filter(Boolean)
            .join(' · '),
        },
      };
    },
    [selectedVaccine],
  );

  /** Handle files dropped/selected from FileDropzone */
  const handleFilesAccepted = useCallback(
    async (files: File[]) => {
      if (!periodLocked) {
        setValidationMessage('Harap konfirmasi Bulan dan Tahun terlebih dahulu sebelum upload!');
        setShowValidationModal(true);
        return;
      }
      if (!files.length) return;

      setIsProcessing(true);
      let currentMaster = masterData;
      const newLogs: UploadLogEntry[] = [];

      try {
        for (const file of files) {
          try {
            const { master, log } = await processVaccineFile(file, currentMaster);
            currentMaster = master;
            newLogs.push(log);
          } catch (err) {
            newLogs.push({
              id: `${Date.now()}-${file.name}`,
              fileName: file.name,
              antigen: selectedVaccine,
              processedAt: new Date().toLocaleTimeString('id-ID'),
              dataCount: 0,
              status: 'error',
              message: `Error: ${err instanceof Error ? err.message : String(err)}`,
            });
          }
        }

        setMasterData(currentMaster);
        const counts: Record<string, number> = {};
        for (const s of ALL_SHEETS) counts[s] = currentMaster[s].length;
        setDataCount(counts);
        setLogs((prev) => [...newLogs.reverse(), ...prev]);
      } finally {
        setIsProcessing(false);
      }
    },
    [masterData, periodLocked, processVaccineFile, selectedVaccine],
  );

  const handleDownload = useCallback(async () => {
    if (!templateBuffer) {
      setValidationMessage(
        templateError ?? 'Template Master belum siap. Upload template manual, lalu coba lagi.',
      );
      setShowValidationModal(true);
      return;
    }
    if (totalChildren === 0) {
      setValidationMessage('Belum ada data. Upload file vaksin ASIK terlebih dahulu.');
      setShowValidationModal(true);
      return;
    }

    // Only warn if there are actual errors (not just parsing warnings)
    const totalErrors = logs.filter((l) => l.status === 'error').length;
    if (totalErrors > 0 && !pendingExport) {
      setValidationMessage(
        `${totalErrors} file upload gagal diproses. Data yang berhasil akan tetap diexport. Lanjutkan?`,
      );
      setPendingExport(true);
      setShowValidationModal(true);
      return;
    }

    setIsExporting(true);
    setPendingExport(false);
    try {
      const blob = await buildMasterExcel(masterData, month, year, templateBuffer);
      const filename = `Master_Imunisasi_${BULAN_INDONESIA[month]}_${year}.xlsx`;
      downloadBlob(blob, filename);
    } catch (err) {
      setValidationMessage(`Gagal membuat file: ${err instanceof Error ? err.message : String(err)}`);
      setShowValidationModal(true);
    } finally {
      setIsExporting(false);
    }
  }, [templateBuffer, templateError, totalChildren, logs, pendingExport, masterData, month, year]);

  const handleReset = useCallback(() => {
    if (!confirm('Reset semua data? Semua data yang sudah diupload akan hilang.')) return;
    setMasterData(createEmptyMasterData());
    setLogs([]);
    setPeriodLocked(false);
    setPendingExport(false);
    setDataCount(Object.fromEntries(ALL_SHEETS.map((s) => [s, 0])));
  }, []);

  const namaSheet = (sheet: string) =>
    sheet === 'LUAR WILAYAH' ? 'Luar Wil.' : sheet === 'Kejar' ? 'Kejar' : sheet.charAt(0) + sheet.slice(1).toLowerCase();

  return (
    <div className="halaman">
      <header className="kop">
        <div>
          <p className="kop-instansi">Puskesmas Mabu'un · ASIK Excel</p>
          <h1>Imunisasi Master Merger</h1>
          <p className="kop-lead">
            Gabungkan file ASIK per vaksin menjadi satu Master Excel laporan bulanan: pilih periode, masukkan file
            vaksin, lalu unduh.
          </p>
        </div>
        <p className="kop-privasi">
          File diproses di browser ini saja. Data warga tidak dikirim ke server.
        </p>
      </header>

      <StepIndicator steps={WORKFLOW_STEPS} currentStep={currentStep} />

      {/* Ketentuan & kebijakan data: selalu terlihat sebelum memakai alat */}
      <div role="note" className="peringatan">
        <strong>Sebelum memakai:</strong> alat ini khusus petugas kesehatan dan kader yang berwenang. Kami (pengembang){' '}
        <b>tidak menyimpan, tidak melihat, dan tidak memakai data warga</b>: file diproses di browser Anda dan tidak dikirim
        ke server. Penanggung jawab data adalah <b>instansi kesehatan</b> pengguna. Hasil otomatis bisa keliru, wajib diperiksa.{' '}
        <a href="/kebijakan.html">Baca ketentuan &amp; kebijakan data</a>
      </div>

      {/* Step 1: Period */}
      <section className="bagian">
        <div className="bagian-kepala">
          <h2><span className="bagian-no">1</span>Periode Laporan</h2>
          {periodLocked && <span className="tanda tanda--ok">✓ {BULAN_INDONESIA[month]} {year}</span>}
        </div>
        <div className="bagian-isi flex flex-wrap gap-3 items-end">
          <div className="flex-1 min-w-[140px]">
            <label className="label" htmlFor="bulan">Bulan</label>
            <select id="bulan" value={month} onChange={(e) => setMonth(Number(e.target.value))} disabled={periodLocked} className="isian">
              {BULAN_INDONESIA.slice(1).map((nama, idx) => (
                <option key={idx + 1} value={idx + 1}>{nama}</option>
              ))}
            </select>
          </div>
          <div className="flex-1 min-w-[110px]">
            <label className="label" htmlFor="tahun">Tahun</label>
            <select id="tahun" value={year} onChange={(e) => setYear(Number(e.target.value))} disabled={periodLocked} className="isian">
              {YEARS.map((y) => (
                <option key={y} value={y}>{y}</option>
              ))}
            </select>
          </div>
          {!periodLocked ? (
            <button onClick={() => setPeriodLocked(true)} className="tombol tombol--utama">Konfirmasi</button>
          ) : (
            <button onClick={() => setPeriodLocked(false)} className="tombol tombol--garis">Ubah</button>
          )}
        </div>
      </section>

      {/* Step 2: Upload */}
      <section className="bagian">
        <div className="bagian-kepala">
          <h2><span className="bagian-no">2</span>Upload File Vaksin ASIK</h2>
          {uploadedCount > 0 && <span className="bagian-info angka">{uploadedCount}/{totalVaccines} vaksin</span>}
        </div>
        <div className="bagian-isi space-y-4">
          <div>
            <label className="label" htmlFor="vaksin">
              Jenis Vaksin <small>(info: otomatis terdeteksi dari Nama Antigen)</small>
            </label>
            <select id="vaksin" value={selectedVaccine} onChange={(e) => setSelectedVaccine(e.target.value)} className="isian">
              {VACCINE_ORDER.map((vk) => (
                <option key={vk} value={vk}>
                  {uploadedVaccines.has(vk) ? '✓ ' : ''}{VACCINE_DISPLAY_NAMES[vk]}
                </option>
              ))}
            </select>
          </div>

          <FileDropzone
            onFilesAccepted={handleFilesAccepted}
            disabled={!periodLocked}
            isProcessing={isProcessing}
            multiple={true}
            label="Seret file vaksin ASIK ke sini, atau klik untuk memilih"
            hint="Format .xlsx / .xls, bisa beberapa file sekaligus"
          />

          <div className="flex flex-wrap items-baseline gap-x-3 gap-y-1 text-sm">
            <input ref={templateFileRef} type="file" accept=".xlsx,.xls" onChange={handleTemplateUpload} className="hidden" />
            <button onClick={() => templateFileRef.current?.click()} className="tautan">
              {templateError && !templateBuffer
                ? '⚠ Upload template Master (default gagal dimuat)'
                : `Ganti template (${templateName})`}
            </button>
            <span className="kecil ml-auto">
              {templateBuffer ? '✓ Template siap' : templateError ? '✕ Gagal muat' : 'Memuat default...'}
            </span>
          </div>
        </div>
      </section>

      {/* Data Preview (after upload) */}
      {totalChildren > 0 && (
        <PreviewTable children={allChildren} maxRows={5} title="Pratinjau Data Anak" />
      )}

      {/* Upload Log */}
      {logs.length > 0 && (
        <section className="bagian">
          <div className="bagian-kepala">
            <h2>Log Upload</h2>
            <span className="bagian-info angka">{logs.length} file</span>
          </div>
          <ul className="baris-daftar max-h-72 overflow-y-auto">
            {logs.map((log) => (
              <li key={log.id}>
                <span
                  className={`tanda ${
                    log.status === 'success' ? 'tanda--ok' : log.status === 'error' ? 'tanda--galat' : 'tanda--awas'
                  }`}
                >
                  {log.status === 'success' ? '✓' : log.status === 'error' ? '✕' : '!'}
                </span>
                <div className="flex-1 min-w-0">
                  <div className="flex items-baseline gap-2">
                    <span className="font-medium truncate">{log.fileName}</span>
                    <span className="kecil angka flex-shrink-0">{log.processedAt}</span>
                  </div>
                  <div className="text-sm text-gray-600">
                    <span className="font-medium text-blue-600">{log.antigen}</span>
                    {log.dataCount > 0 && <span className="ml-1">· {log.dataCount} data</span>}
                  </div>
                  {log.message && <p className="kecil mt-0.5 break-words">{log.message}</p>}
                </div>
              </li>
            ))}
          </ul>
        </section>
      )}

      {/* Vaccine Completion Tracker */}
      {totalChildren > 0 && (
        <section className="bagian">
          <div className="bagian-kepala">
            <h2>Kelengkapan Vaksin</h2>
            <span className={`tanda ${uploadedCount === totalVaccines ? 'tanda--ok' : 'tanda--awas'}`}>
              {uploadedCount}/{totalVaccines} terupload
            </span>
          </div>
          <div className="bagian-isi space-y-4">
            {VACCINE_GROUPS.map((group) => (
              <div key={group.label}>
                <p className="label">{group.label}</p>
                <div className="flex flex-wrap gap-2">
                  {group.keys.map((vk) => {
                    const done = uploadedVaccines.has(vk);
                    return (
                      <span key={vk} className={`tanda ${done ? 'tanda--ok' : 'tanda--pelan'}`}>
                        {done ? '✓' : '○'} {VACCINE_DISPLAY_NAMES[vk]}
                      </span>
                    );
                  })}
                </div>
              </div>
            ))}
            {uploadedCount < totalVaccines && (
              <p className="peringatan">
                ⚠ {totalVaccines - uploadedCount} vaksin belum ada datanya — tetap bisa diexport,
                kolom tersebut akan kosong.
              </p>
            )}
          </div>
        </section>
      )}

      {/* Data Summary */}
      {totalChildren > 0 && (
        <section className="bagian">
          <div className="bagian-kepala">
            <h2><span className="bagian-no">3</span>Sebaran Data</h2>
            <span className="bagian-info">{BULAN_INDONESIA[month]} {year}</span>
          </div>
          <div className="bagian-isi flex flex-wrap items-end gap-x-10 gap-y-5">
            <div>
              <span className="angka angka--besar">{totalChildren}</span>
              <span className="kecil">anak di file master</span>
            </div>
            <dl className="rinci-daftar flex-1 min-w-[260px]">
              {ALL_SHEETS.map((sheet) => (
                <div key={sheet} className="rinci">
                  <dt>{namaSheet(sheet)}</dt>
                  <dd>{dataCount[sheet] ?? 0}</dd>
                </div>
              ))}
            </dl>
          </div>
        </section>
      )}

      {/* Download & Reset */}
      {totalChildren > 0 && (
        <section className="mt-8 space-y-3">
          <button onClick={handleDownload} disabled={!canDownload} className="tombol tombol--utama tombol--lebar">
            {isExporting ? (
              <>
                <Loader2 className="w-5 h-5 animate-spin" />
                Menyiapkan file...
              </>
            ) : (
              <>
                <Download className="w-5 h-5" />
                Download Master Excel
                <span className="font-normal text-sm opacity-80 hidden sm:inline">
                  · {totalChildren} anak · {uploadedCount}/{totalVaccines} vaksin
                </span>
              </>
            )}
          </button>

          {!templateBuffer && (
            <p className="peringatan">⚠ Template belum tersedia — upload template Master untuk bisa download.</p>
          )}

          <button onClick={handleReset} className="tombol tombol--garis tombol--bahaya w-full">
            <RotateCcw className="w-4 h-4" />
            Reset Semua Data
          </button>
        </section>
      )}

      <footer className="kaki">
        <p>Data diproses sepenuhnya di browser — tidak ada data yang dikirim ke server.</p>
        <p>Deduplikasi berdasarkan Nama Anak + Tanggal Lahir + Nama Orang Tua.</p>
        <p>
          <a href="/kebijakan.html">Ketentuan &amp; kebijakan data</a> · Penanggung jawab data: instansi kesehatan pengguna
        </p>
      </footer>

      {/* Notification Modal */}
      <NotificationModal
        visible={showValidationModal}
        variant={pendingExport ? 'warning' : totalChildren > 0 ? 'warning' : 'info'}
        title={pendingExport ? 'Data Perlu Dicek' : totalChildren > 0 ? 'Export Data' : 'Informasi'}
        message={validationMessage}
        confirmLabel={pendingExport ? 'Lanjut Export' : 'OK'}
        cancelLabel={pendingExport ? 'Batal' : undefined}
        onConfirm={() => {
          setShowValidationModal(false);
          if (pendingExport) {
            setPendingExport(false);
            handleDownload();
          }
        }}
        onCancel={pendingExport ? () => { setShowValidationModal(false); setPendingExport(false); } : undefined}
        onClose={() => { setShowValidationModal(false); setPendingExport(false); }}
      />
    </div>
  );
}

export default App;

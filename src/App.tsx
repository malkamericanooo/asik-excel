import { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import type { MasterData, UploadLogEntry, ProcessResult, VaccineKey } from './types';
import { ALL_SHEETS } from './types';
import { createEmptyMasterData, parseAndMergeAsikFile } from './utils/asikParser';
import { buildMasterExcel, getUploadedVaccines } from './utils/masterExcel';
import { loadDefaultTemplate } from './utils/templateLoader';
import { downloadBlob } from './utils/downloadFile';
import { BULAN_INDONESIA } from './utils/dateUtils';
import { VACCINE_DISPLAY_NAMES, VACCINE_ORDER } from './utils/vaccineMapping';

const CURRENT_YEAR = new Date().getFullYear();
const YEARS = Array.from({ length: 5 }, (_, i) => CURRENT_YEAR - 2 + i);

const VACCINE_GROUPS: { label: string; keys: VaccineKey[] }[] = [
  { label: 'Bayi Baru Lahir', keys: ['HB0_24JAM', 'HB0_7HARI', 'BCG'] },
  { label: 'Usia 1–3 Bulan', keys: ['POLIO_1', 'DPT_1', 'POLIO_2', 'PCV_1', 'ROTA_1'] },
  { label: 'Usia 3–4 Bulan', keys: ['DPT_2', 'POLIO_3', 'PCV_2', 'ROTA_2'] },
  { label: 'Usia 4–9 Bulan', keys: ['DPT_3', 'POLIO_4', 'IPV_1', 'ROTA_3', 'MR_1'] },
  { label: 'Baduta (≥ 9 bln)', keys: ['IPV_2', 'PCV_3', 'DPT_4', 'BOOSTER_MR'] },
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
    Object.fromEntries(ALL_SHEETS.map((s) => [s, 0]))
  );
  const vaccineFileRef = useRef<HTMLInputElement>(null);
  const templateFileRef = useRef<HTMLInputElement>(null);
  const uploadedVaccines = useMemo(() => getUploadedVaccines(masterData), [masterData]);
  const totalChildren = ALL_SHEETS.reduce((sum, s) => sum + (dataCount[s] ?? 0), 0);
  const uploadedCount = uploadedVaccines.size;
  const totalVaccines = VACCINE_ORDER.length;
  const canDownload = totalChildren > 0 && templateBuffer !== null && !isExporting;

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

  const handleVaccineUpload = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const files = e.target.files;
      if (!files?.length) return;
      if (!periodLocked) {
        alert('Harap konfirmasi Bulan dan Tahun terlebih dahulu sebelum upload!');
        e.target.value = '';
        return;
      }

      setIsProcessing(true);
      let currentMaster = masterData;
      const newLogs: UploadLogEntry[] = [];

      try {
        for (const file of Array.from(files)) {
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
        e.target.value = '';
      }
    },
    [masterData, periodLocked, processVaccineFile, selectedVaccine],
  );

  const handleDownload = async () => {
    if (!templateBuffer) {
      alert(
        templateError ??
          'Template Master belum siap. Upload template manual di langkah 2, lalu coba lagi.',
      );
      return;
    }
    if (totalChildren === 0) {
      alert('Belum ada data. Upload file vaksin ASIK terlebih dahulu.');
      return;
    }

    setIsExporting(true);
    try {
      const blob = buildMasterExcel(masterData, month, year, templateBuffer);
      const filename = `Master_Imunisasi_${BULAN_INDONESIA[month]}_${year}.xlsx`;
      downloadBlob(blob, filename);
    } catch (err) {
      alert(`Gagal membuat file: ${err instanceof Error ? err.message : String(err)}`);
    } finally {
      setIsExporting(false);
    }
  };

  const handleReset = () => {
    if (!confirm('Reset semua data? Semua data yang sudah diupload akan hilang.')) return;
    setMasterData(createEmptyMasterData()); setLogs([]); setPeriodLocked(false);
    setDataCount(Object.fromEntries(ALL_SHEETS.map((s) => [s, 0])));
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-white to-emerald-50">
      <header className="bg-white border-b border-blue-100 shadow-sm">
        <div className="max-w-2xl mx-auto px-4 py-4 flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-blue-600 flex items-center justify-center flex-shrink-0">
            <svg className="w-6 h-6 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
            </svg>
          </div>
          <div>
            <h1 className="text-lg font-bold text-gray-900 leading-tight">Imunisasi Master Merger</h1>
            <p className="text-xs text-gray-500">Puskesmas Mabuun — Gabungkan data ASIK ke Master Excel</p>
          </div>
        </div>
      </header>

      <main className="max-w-2xl mx-auto px-4 py-6 space-y-5">
        {/* Step 1: Period */}
        <section className="bg-white rounded-2xl border border-blue-100 shadow-sm overflow-hidden">
          <div className="bg-blue-600 px-4 py-3 flex items-center gap-2">
            <span className="w-6 h-6 rounded-full bg-white text-blue-600 text-xs font-bold flex items-center justify-center">1</span>
            <h2 className="text-white font-semibold text-sm">Pilih Periode Laporan</h2>
            {periodLocked && <span className="ml-auto bg-emerald-500 text-white text-xs px-2 py-0.5 rounded-full font-medium">✓ Terkunci</span>}
          </div>
          <div className="p-4">
            <div className="flex gap-3 items-end">
              <div className="flex-1">
                <label className="block text-xs font-medium text-gray-600 mb-1">Bulan</label>
                <select value={month} onChange={(e) => setMonth(Number(e.target.value))} disabled={periodLocked}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 disabled:bg-gray-50 disabled:text-gray-500">
                  {BULAN_INDONESIA.slice(1).map((nama, idx) => <option key={idx + 1} value={idx + 1}>{nama}</option>)}
                </select>
              </div>
              <div className="flex-1">
                <label className="block text-xs font-medium text-gray-600 mb-1">Tahun</label>
                <select value={year} onChange={(e) => setYear(Number(e.target.value))} disabled={periodLocked}
                  className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 disabled:bg-gray-50 disabled:text-gray-500">
                  {YEARS.map((y) => <option key={y} value={y}>{y}</option>)}
                </select>
              </div>
              <div>
                {!periodLocked
                  ? <button onClick={() => setPeriodLocked(true)} className="px-4 py-2 bg-blue-600 text-white text-sm font-medium rounded-lg hover:bg-blue-700 active:scale-95 transition-all">Konfirmasi</button>
                  : <button onClick={() => setPeriodLocked(false)} className="px-4 py-2 bg-gray-100 text-gray-600 text-sm font-medium rounded-lg hover:bg-gray-200 active:scale-95 transition-all">Ubah</button>
                }
              </div>
            </div>
            {periodLocked
              ? <p className="mt-2 text-xs text-emerald-600 font-medium">✓ Periode: {BULAN_INDONESIA[month]} {year}</p>
              : <p className="mt-2 text-xs text-amber-600">⚠ Konfirmasi periode sebelum mulai upload file imunisasi.</p>
            }
          </div>
        </section>

        {/* Step 2: Template (opsional) */}
        <section className="bg-white rounded-2xl border border-blue-100 shadow-sm overflow-hidden">
          <div className="bg-blue-600 px-4 py-3 flex items-center gap-2">
            <span className="w-6 h-6 rounded-full bg-white text-blue-600 text-xs font-bold flex items-center justify-center">2</span>
            <h2 className="text-white font-semibold text-sm">Master Template</h2>
            <span className="ml-auto bg-white/20 text-white text-xs px-2 py-0.5 rounded-full font-medium truncate max-w-[140px]">{templateName}</span>
          </div>
          <div className="p-4 space-y-3">
            <p className="text-xs text-gray-500">Upload template Master (opsional). Jika tidak diupload, format default Puskesmas Mabuun akan dipakai.</p>
            {templateError && !templateBuffer && (
              <p className="text-xs text-amber-700 bg-amber-50 rounded-lg px-3 py-2">
                ⚠ {templateError}
              </p>
            )}
            <input ref={templateFileRef} type="file" accept=".xlsx,.xls" onChange={handleTemplateUpload} className="hidden" />
            <button
              onClick={() => templateFileRef.current?.click()}
              className="w-full py-2.5 rounded-xl text-sm font-medium border border-blue-200 text-blue-700 hover:bg-blue-50 active:scale-95 transition-all"
            >
              Upload Master Template
            </button>
          </div>
        </section>

        {/* Step 3: Upload */}
        <section className="bg-white rounded-2xl border border-blue-100 shadow-sm overflow-hidden">
          <div className="bg-blue-600 px-4 py-3 flex items-center gap-2">
            <span className="w-6 h-6 rounded-full bg-white text-blue-600 text-xs font-bold flex items-center justify-center">3</span>
            <h2 className="text-white font-semibold text-sm">Upload File Vaksin ASIK</h2>
            {uploadedCount > 0 && (
              <span className="ml-auto bg-white/20 text-white text-xs px-2 py-0.5 rounded-full font-medium">{uploadedCount}/{totalVaccines} vaksin</span>
            )}
          </div>
          <div className="p-4 space-y-3">
            <div>
              <label className="block text-xs font-medium text-gray-600 mb-1">Jenis Vaksin (info saja — terdeteksi otomatis dari kolom Nama Antigen)</label>
              <select value={selectedVaccine} onChange={(e) => setSelectedVaccine(e.target.value)}
                className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500">
                {VACCINE_ORDER.map((vk) => <option key={vk} value={vk}>{uploadedVaccines.has(vk) ? '✓ ' : ''}{VACCINE_DISPLAY_NAMES[vk]}</option>)}
              </select>
            </div>
            <input ref={vaccineFileRef} type="file" accept=".xlsx,.xls,.csv" multiple onChange={handleVaccineUpload} className="hidden" />
            <button
              onClick={() => { if (!periodLocked) { alert('Harap konfirmasi Bulan dan Tahun terlebih dahulu!'); return; } vaccineFileRef.current?.click(); }}
              disabled={isProcessing}
              className={['w-full py-3 rounded-xl text-sm font-semibold transition-all flex items-center justify-center gap-2', periodLocked ? 'bg-blue-600 text-white hover:bg-blue-700 active:scale-95 shadow-md shadow-blue-200' : 'bg-gray-100 text-gray-400 cursor-not-allowed', isProcessing ? 'opacity-70' : ''].join(' ')}
            >
              {isProcessing ? (
                <><svg className="animate-spin w-4 h-4" fill="none" viewBox="0 0 24 24"><circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" /><path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z" /></svg>Memproses...</>
              ) : (
                <><svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-8l-4-4m0 0L8 8m4-4v12" /></svg>Upload File Vaksin (bisa banyak sekaligus)</>
              )}
            </button>
            <p className="text-xs text-gray-400 text-center">Pilih banyak file sekaligus (Ctrl/Cmd+klik atau Shift). Urutan bebas — duplikat anak otomatis digabung.</p>
          </div>
        </section>

        {/* Vaccine Completion Tracker */}
        {totalChildren > 0 && (
          <section className="bg-white rounded-2xl border border-blue-100 shadow-sm overflow-hidden">
            <div className="px-4 py-3 border-b border-gray-100 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <svg className="w-4 h-4 text-blue-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2M9 5a2 2 0 002 2h2a2 2 0 002-2M9 5a2 2 0 012-2h2a2 2 0 012 2m-6 9l2 2 4-4" />
                </svg>
                <h2 className="text-gray-700 font-semibold text-sm">Kelengkapan Vaksin</h2>
              </div>
              <span className={`text-xs font-semibold px-2.5 py-1 rounded-full ${uploadedCount === totalVaccines ? 'bg-emerald-100 text-emerald-700' : 'bg-amber-100 text-amber-700'}`}>
                {uploadedCount}/{totalVaccines} terupload
              </span>
            </div>
            <div className="p-4 space-y-4">
              {VACCINE_GROUPS.map((group) => (
                <div key={group.label}>
                  <p className="text-xs font-semibold text-gray-400 uppercase tracking-wide mb-2">{group.label}</p>
                  <div className="flex flex-wrap gap-2">
                    {group.keys.map((vk) => {
                      const done = uploadedVaccines.has(vk);
                      return (
                        <span key={vk} className={['inline-flex items-center gap-1 px-2.5 py-1 rounded-lg text-xs font-medium border transition-colors', done ? 'bg-emerald-50 border-emerald-200 text-emerald-700' : 'bg-gray-50 border-gray-200 text-gray-400'].join(' ')}>
                          {done
                            ? <svg className="w-3 h-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2.5} d="M5 13l4 4L19 7" /></svg>
                            : <svg className="w-3 h-3 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><circle cx="12" cy="12" r="8" strokeWidth="2" /></svg>
                          }
                          {VACCINE_DISPLAY_NAMES[vk]}
                        </span>
                      );
                    })}
                  </div>
                </div>
              ))}
              {uploadedCount < totalVaccines && (
                <p className="text-xs text-amber-600 bg-amber-50 rounded-lg px-3 py-2">
                  ⚠ {totalVaccines - uploadedCount} vaksin belum ada datanya — tetap bisa diexport, kolom tersebut akan kosong.
                </p>
              )}
            </div>
          </section>
        )}

        {/* Data Summary */}
        {totalChildren > 0 && (
          <section className="bg-white rounded-2xl border border-emerald-100 shadow-sm overflow-hidden">
            <div className="bg-emerald-600 px-4 py-3 flex items-center gap-2">
              <svg className="w-4 h-4 text-white" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z" />
              </svg>
              <h2 className="text-white font-semibold text-sm">Sebaran Data ({totalChildren} anak)</h2>
            </div>
            <div className="p-4">
              <div className="grid grid-cols-4 gap-2 sm:grid-cols-7">
                {ALL_SHEETS.map((sheet) => (
                  <div key={sheet} className="bg-gray-50 rounded-lg px-2 py-2 text-center">
                    <div className="text-lg font-bold text-blue-600">{dataCount[sheet] ?? 0}</div>
                    <div className="text-xs text-gray-500 leading-tight mt-0.5">
                      {sheet === 'LUAR WILAYAH' ? 'Luar Wil.' : sheet === 'Kejar' ? 'Kejar' : sheet.charAt(0) + sheet.slice(1).toLowerCase()}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </section>
        )}

        {/* Upload Log */}
        {logs.length > 0 && (
          <section className="bg-white rounded-2xl border border-gray-100 shadow-sm overflow-hidden">
            <div className="bg-gray-50 px-4 py-3 border-b border-gray-100 flex items-center justify-between">
              <h2 className="text-gray-700 font-semibold text-sm">Log Upload</h2>
              <span className="text-xs text-gray-400">{logs.length} file</span>
            </div>
            <div className="max-h-64 overflow-y-auto divide-y divide-gray-50">
              {logs.map((log) => (
                <div key={log.id} className={`px-4 py-3 flex gap-3 items-start ${log.status === 'error' ? 'bg-red-50' : log.status === 'warning' ? 'bg-amber-50' : ''}`}>
                  <span className={`mt-0.5 w-5 h-5 rounded-full flex-shrink-0 flex items-center justify-center text-xs font-bold ${log.status === 'success' ? 'bg-emerald-100 text-emerald-600' : log.status === 'error' ? 'bg-red-100 text-red-600' : 'bg-amber-100 text-amber-600'}`}>
                    {log.status === 'success' ? '✓' : log.status === 'error' ? '✕' : '!'}
                  </span>
                  <div className="flex-1 min-w-0">
                    <div className="flex items-baseline gap-2">
                      <span className="text-sm font-medium text-gray-800 truncate">{log.fileName}</span>
                      <span className="text-xs text-gray-400 flex-shrink-0">{log.processedAt}</span>
                    </div>
                    <div className="text-xs text-gray-500 mt-0.5">
                      <span className="font-medium text-blue-600">{log.antigen}</span>
                      {log.dataCount > 0 && <span className="ml-1">· {log.dataCount} data</span>}
                    </div>
                    {log.message && <p className="text-xs text-gray-500 mt-0.5 break-words">{log.message}</p>}
                  </div>
                </div>
              ))}
            </div>
          </section>
        )}

        {/* Download & Reset */}
        <section className="space-y-3">
          <button onClick={handleDownload} disabled={!canDownload}
            className={`w-full py-4 rounded-2xl text-base font-bold transition-all flex items-center justify-center gap-3 ${canDownload ? 'bg-emerald-600 text-white hover:bg-emerald-700 active:scale-95 shadow-lg shadow-emerald-200' : 'bg-gray-100 text-gray-400 cursor-not-allowed'}`}>
            <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4" />
            </svg>
            {isExporting ? 'Menyiapkan file...' : 'Download Master Excel'}
            {totalChildren > 0 && canDownload && <span className="text-emerald-200 text-sm font-normal">· {totalChildren} anak · {uploadedCount}/{totalVaccines} vaksin · {BULAN_INDONESIA[month]} {year}</span>}
          </button>
          {totalChildren > 0 && !templateBuffer && (
            <p className="text-xs text-amber-700 bg-amber-50 rounded-lg px-3 py-2 text-center">
              ⚠ Template belum tersedia — upload template Master di langkah 2 untuk bisa download.
            </p>
          )}
          {totalChildren > 0 && (
            <button onClick={handleReset} className="w-full py-2.5 rounded-xl text-sm text-gray-500 hover:text-red-500 hover:bg-red-50 transition-all border border-gray-100">
              Reset Semua Data
            </button>
          )}
        </section>

        <div className="text-center text-xs text-gray-400 pb-4 space-y-1">
          <p>Data diproses sepenuhnya di browser — tidak ada data yang dikirim ke server.</p>
          <p>Deduplikasi berdasarkan Nama Anak + Tanggal Lahir + Nama Orang Tua.</p>
        </div>
      </main>
    </div>
  );
}

export default App;

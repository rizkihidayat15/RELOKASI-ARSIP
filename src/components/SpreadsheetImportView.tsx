import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  Check,
  CheckCircle2,
  Download,
  FilePlus2,
  FileSpreadsheet,
  Lock,
  Pencil,
  Plus,
  Search,
  Trash2,
  Upload,
  X,
} from 'lucide-react';
import type {
  ImportedSpreadsheet,
  LockedLocationInfo,
} from '../types/archive';
import { buildTemplateBackupExportRows } from '../utils/backupHistory';
import {
  downloadSpreadsheetTemplate,
  downloadTemplateBackupExcel,
  downloadUpdatedTemplateExcel,
  parseSpreadsheetFile,
} from '../utils/excelService';
import { extractSpreadsheetItems } from '../utils/relocationEngine';
import {
  addBulkTemplateItems,
  addTemplateItem,
  createEmptyTemplateSpreadsheet,
  mergeSpreadsheetIntoTemplate,
  parseBulkTemplateInput,
  removeTemplateItem,
  updateTemplateItem,
} from '../utils/templateEditor';

export type TemplateLockFilter = 'all' | 'available' | 'locked';
type AddMode = 'single' | 'bulk';

interface SpreadsheetImportViewProps {
  spreadsheet: ImportedSpreadsheet | null;
  lockedMap: Map<string, LockedLocationInfo>;
  allowUnlistedWithSpreadsheet: boolean;
  initialLockFilter?: TemplateLockFilter;
  onImportSpreadsheet: (sheet: ImportedSpreadsheet) => void;
  onUpdateSpreadsheet: (
    nextSheet: ImportedSpreadsheet,
    actionLabel: string,
    feedbackMessage: string
  ) => void;
  onChangeColumn: (nextColumn: string) => void;
  onClearSpreadsheet: () => void;
  onToggleAllowUnlisted: (allowed: boolean) => void;
  onBackToRack: () => void;
}

export const SpreadsheetImportView: React.FC<SpreadsheetImportViewProps> = ({
  spreadsheet,
  lockedMap,
  allowUnlistedWithSpreadsheet,
  initialLockFilter = 'all',
  onImportSpreadsheet,
  onUpdateSpreadsheet,
  onChangeColumn,
  onClearSpreadsheet,
  onToggleAllowUnlisted,
  onBackToRack,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const mergeFileInputRef = useRef<HTMLInputElement>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [lockFilter, setLockFilter] =
    useState<TemplateLockFilter>(initialLockFilter);
  const [isDragging, setIsDragging] = useState(false);

  // Template Editor state (add single / bulk)
  const [addMode, setAddMode] = useState<AddMode>('single');
  const [newCode, setNewCode] = useState('');
  const [newDesc, setNewDesc] = useState('');
  const [bulkText, setBulkText] = useState('');
  const [editorError, setEditorError] = useState<string | null>(null);
  const [editorSuccess, setEditorSuccess] = useState<string | null>(null);

  // Inline row edit state
  const [editingKey, setEditingKey] = useState<string | null>(null);
  const [editCode, setEditCode] = useState('');
  const [editDesc, setEditDesc] = useState('');
  const [editRowError, setEditRowError] = useState<string | null>(null);

  useEffect(() => {
    setLockFilter(initialLockFilter);
  }, [initialLockFilter]);

  const parsedBulkPreview = useMemo(
    () => parseBulkTemplateInput(bulkText),
    [bulkText]
  );

  const handleFileSelected = async (file: File | null) => {
    if (!file) return;
    setErrorMsg(null);
    setEditorError(null);
    setEditorSuccess(null);
    setIsLoading(true);
    try {
      const parsed = await parseSpreadsheetFile(file);
      onImportSpreadsheet(parsed);
      setEditingKey(null);
    } catch (err) {
      setErrorMsg(
        err instanceof Error
          ? err.message
          : 'Gagal membaca file spreadsheet. Pastikan format .xlsx, .xls, atau .csv valid.'
      );
    } finally {
      setIsLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  };

  const handleMergeFileSelected = async (file: File | null) => {
    if (!file) return;
    setErrorMsg(null);
    setEditorError(null);
    setEditorSuccess(null);
    setIsLoading(true);
    try {
      const parsed = await parseSpreadsheetFile(file);
      const merged = mergeSpreadsheetIntoTemplate(spreadsheet, parsed);
      if (merged.addedCount === 0) {
        setEditorError(
          `Semua ${merged.skippedDuplicates} kode pada "${file.name}" sudah ada di template aktif.`
        );
      } else {
        const msg =
          merged.skippedDuplicates > 0
            ? `${merged.addedCount} kode baru dari "${file.name}" ditambahkan ke template (${merged.skippedDuplicates} duplikat dilewati) & langsung aktif di kata kunci pencarian.`
            : `${merged.addedCount} kode baru dari "${file.name}" ditambahkan ke template & langsung aktif di kata kunci pencarian.`;
        onUpdateSpreadsheet(
          merged.nextSpreadsheet,
          `Gabung Template (${file.name})`,
          msg
        );
        setEditorSuccess(msg);
      }
    } catch (err) {
      setErrorMsg(
        err instanceof Error
          ? err.message
          : 'Gagal menggabungkan file spreadsheet tambahan.'
      );
    } finally {
      setIsLoading(false);
      if (mergeFileInputRef.current) mergeFileInputRef.current.value = '';
    }
  };

  const handleCreateManualTemplate = () => {
    const empty = createEmptyTemplateSpreadsheet();
    onUpdateSpreadsheet(
      empty,
      'Buat Template Manual',
      'Template lokasi simpan lama siap diisi. Tambahkan kode baru di bawah ini.'
    );
    setErrorMsg(null);
    setEditorError(null);
    setEditorSuccess(null);
  };

  const handleAddSingleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setEditorError(null);
    setEditorSuccess(null);

    const result = addTemplateItem(spreadsheet, newCode, newDesc);
    if (result.error || !result.addedItem) {
      setEditorError(result.error ?? 'Gagal menambahkan lokasi simpan lama.');
      return;
    }

    const msg = `Kode "${result.addedItem.oldLocation}" ditambahkan ke template & langsung terupdate di kata kunci pencarian.`;
    onUpdateSpreadsheet(
      result.nextSpreadsheet,
      `Tambah Kode Template (${result.addedItem.oldLocation})`,
      msg
    );
    setNewCode('');
    setNewDesc('');
    setEditorSuccess(msg);
  };

  const handleAddBulkSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setEditorError(null);
    setEditorSuccess(null);

    if (parsedBulkPreview.length === 0) {
      setEditorError(
        'Masukkan minimal satu kode lokasi simpan lama pada kolom tempel daftar.'
      );
      return;
    }

    const result = addBulkTemplateItems(spreadsheet, parsedBulkPreview);
    if (result.addedCount === 0) {
      setEditorError(
        `Tidak ada kode baru yang ditambahkan (${result.duplicateCodes.length} kode sudah terdaftar di template).`
      );
      return;
    }

    const msg =
      result.duplicateCodes.length > 0
        ? `${result.addedCount} kode lokasi lama ditambahkan ke template (${result.duplicateCodes.length} duplikat dilewati) & langsung terupdate di kata kunci pencarian.`
        : `${result.addedCount} kode lokasi lama ditambahkan ke template & langsung terupdate di kata kunci pencarian.`;

    onUpdateSpreadsheet(
      result.nextSpreadsheet,
      `Tambah ${result.addedCount} Kode Template`,
      msg
    );
    setBulkText('');
    setEditorSuccess(msg);
  };

  const startInlineEdit = (
    normalizedKey: string,
    currentCode: string,
    currentDesc?: string
  ) => {
    setEditingKey(normalizedKey);
    setEditCode(currentCode);
    setEditDesc(currentDesc ?? '');
    setEditRowError(null);
  };

  const cancelInlineEdit = () => {
    setEditingKey(null);
    setEditCode('');
    setEditDesc('');
    setEditRowError(null);
  };

  const handleSaveInlineEdit = (normalizedKey: string) => {
    if (!spreadsheet) return;
    setEditRowError(null);
    const isLocked = lockedMap.has(normalizedKey);
    const result = updateTemplateItem(
      spreadsheet,
      normalizedKey,
      editCode,
      editDesc,
      { isLocked }
    );
    if (result.error || !result.updatedItem) {
      setEditRowError(result.error ?? 'Gagal menyimpan perubahan baris.');
      return;
    }
    const msg = `Data template "${result.updatedItem.oldLocation}" diperbarui & langsung terupdate di kata kunci pencarian.`;
    onUpdateSpreadsheet(
      result.nextSpreadsheet,
      `Edit Template (${result.updatedItem.oldLocation})`,
      msg
    );
    setEditingKey(null);
    setEditorSuccess(msg);
  };

  const handleDeleteTemplateRow = (normalizedKey: string) => {
    if (!spreadsheet) return;
    if (lockedMap.has(normalizedKey)) return;
    const result = removeTemplateItem(spreadsheet, normalizedKey);
    if (!result.removedItem) return;
    if (editingKey === normalizedKey) {
      cancelInlineEdit();
    }
    const msg = `Kode "${result.removedItem.oldLocation}" dihapus dari template & kata kunci pencarian.`;
    onUpdateSpreadsheet(
      result.nextSpreadsheet,
      `Hapus Kode Template (${result.removedItem.oldLocation})`,
      msg
    );
    setEditorSuccess(msg);
  };

  const filteredItems = useMemo(() => {
    if (!spreadsheet) return [];
    const q = searchQuery.trim().toLowerCase();
    return spreadsheet.items.filter((item) => {
      const isLocked = lockedMap.has(item.normalizedKey);
      if (lockFilter === 'available' && isLocked) return false;
      if (lockFilter === 'locked' && !isLocked) return false;
      if (!q) return true;
      const locked = lockedMap.get(item.normalizedKey);
      return (
        item.oldLocation.toLowerCase().includes(q) ||
        (item.description?.toLowerCase().includes(q) ?? false) ||
        (locked?.newLocation.toLowerCase().includes(q) ?? false)
      );
    });
  }, [spreadsheet, searchQuery, lockFilter, lockedMap]);

  const lockedCount = useMemo(() => {
    if (!spreadsheet) return 0;
    return spreadsheet.items.filter((item) => lockedMap.has(item.normalizedKey))
      .length;
  }, [spreadsheet, lockedMap]);

  const availableCount = spreadsheet
    ? Math.max(0, spreadsheet.items.length - lockedCount)
    : 0;

  const handleDownloadRemainingTemplate = () => {
    if (!spreadsheet) return;
    const rows = buildTemplateBackupExportRows(spreadsheet, lockedMap, true);
    downloadTemplateBackupExcel(
      rows,
      new Date().toISOString(),
      'TEMPLATE_TERSISA'
    );
  };

  return (
    <div className="space-y-6">
      {/* Upload & File Management Card */}
      <section className="rounded-xl border border-slate-200 bg-white p-5">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h2 className="text-base font-bold tracking-tight text-slate-900">
              Template Spreadsheet Lokasi Simpan Lama
            </h2>
            <p className="mt-0.5 max-w-2xl text-xs leading-relaxed text-slate-500">
              Unggah, edit, atau tambahkan data lokasi simpan lama (satuan,
              tempel banyak, atau gabung file). Setiap perubahan langsung
              terupdate pada kata kunci pencarian modal box rak &amp; Dashboard.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {spreadsheet && (
              <>
                <label
                  htmlFor="spreadsheet-merge-input"
                  data-testid="label-merge-template"
                  className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-[#0F4C81]/30 bg-sky-50 px-3.5 py-2 text-xs font-semibold text-[#0F4C81] transition-colors hover:bg-sky-100 whitespace-nowrap shrink-0"
                >
                  <FilePlus2 className="h-3.5 w-3.5" />
                  Tambah dari File (.xlsx)
                </label>
                <label
                  htmlFor="spreadsheet-file-input"
                  className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap shrink-0"
                >
                  <Upload className="h-3.5 w-3.5" />
                  Ganti File
                </label>
                <button
                  type="button"
                  disabled={spreadsheet.items.length === 0}
                  onClick={() => downloadUpdatedTemplateExcel(spreadsheet)}
                  data-testid="btn-download-updated-template"
                  className="inline-flex items-center gap-1.5 rounded-lg bg-emerald-700 px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-emerald-800 disabled:opacity-45 whitespace-nowrap shrink-0"
                >
                  <Download className="h-3.5 w-3.5" />
                  Unduh Template Terupdate (.xlsx)
                </button>
              </>
            )}
            {!spreadsheet && (
              <button
                type="button"
                onClick={downloadSpreadsheetTemplate}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap shrink-0"
              >
                <Download className="h-3.5 w-3.5" />
                Unduh Template Contoh (.xlsx)
              </button>
            )}
          </div>
        </div>

        <input
          ref={fileInputRef}
          id="spreadsheet-file-input"
          type="file"
          accept=".xlsx,.xls,.csv"
          onChange={(e) => void handleFileSelected(e.target.files?.[0] ?? null)}
          className="sr-only"
        />
        <input
          ref={mergeFileInputRef}
          id="spreadsheet-merge-input"
          type="file"
          accept=".xlsx,.xls,.csv"
          onChange={(e) =>
            void handleMergeFileSelected(e.target.files?.[0] ?? null)
          }
          className="sr-only"
        />

        {/* Dropzone shown in full only when no spreadsheet is loaded yet */}
        {!spreadsheet && (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragging(true);
            }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={(e) => {
              e.preventDefault();
              setIsDragging(false);
              const dropped = e.dataTransfer.files?.[0] ?? null;
              void handleFileSelected(dropped);
            }}
            className={`mt-5 flex flex-col items-center justify-center rounded-xl border border-dashed px-6 py-10 text-center transition-colors ${
              isDragging
                ? 'border-[#0F4C81] bg-sky-50/70'
                : 'border-slate-300 bg-slate-50/60'
            }`}
          >
            <div className="flex h-11 w-11 items-center justify-center rounded-lg border border-slate-200 bg-white text-[#0F4C81]">
              <Upload className="h-5 w-5" />
            </div>
            <p className="mt-3 text-sm font-semibold text-slate-900">
              {isLoading
                ? 'Membaca file spreadsheet...'
                : 'Pilih atau seret file spreadsheet template ke sini'}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              Mendukung format .xlsx, .xls, dan .csv — atau buat daftar template
              baru secara manual.
            </p>

            <div className="mt-4 flex flex-wrap items-center justify-center gap-2.5">
              <label
                htmlFor="spreadsheet-file-input"
                className="inline-flex cursor-pointer items-center gap-2 rounded-lg bg-[#0F4C81] px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] whitespace-nowrap"
              >
                <FileSpreadsheet className="h-4 w-4" />
                Pilih File Spreadsheet
              </label>
              <button
                type="button"
                data-testid="btn-create-manual-template"
                onClick={handleCreateManualTemplate}
                className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-4 py-2 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap"
              >
                <Plus className="h-4 w-4 text-[#0F4C81]" />
                Buat &amp; Isi Template Manual
              </button>
            </div>
          </div>
        )}

        {errorMsg && (
          <div
            role="alert"
            className="mt-4 flex items-start gap-2.5 rounded-lg border border-red-200 bg-red-50 px-4 py-3 text-xs text-red-800"
          >
            <AlertCircle className="mt-0.5 h-4 w-4 shrink-0 text-red-600" />
            <span>{errorMsg}</span>
          </div>
        )}
      </section>

      {/* Active Imported / Editable Template Section */}
      {spreadsheet && (
        <section className="rounded-xl border border-slate-200 bg-white">
          <div className="border-b border-slate-200 px-6 py-4">
            <div className="flex flex-wrap items-start justify-between gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <CheckCircle2 className="h-4 w-4 text-emerald-600" />
                  <h3 className="text-sm font-bold text-slate-900">
                    {spreadsheet.fileName}
                  </h3>
                </div>
                <p className="mt-1 text-xs text-slate-500 tabular-nums">
                  Sheet: {spreadsheet.sheetName} ·{' '}
                  <strong
                    data-testid="template-total-count"
                    className="font-semibold text-slate-800"
                  >
                    {spreadsheet.items.length}
                  </strong>{' '}
                  lokasi unik ·{' '}
                  <span className="font-semibold text-amber-800">
                    {lockedCount} terkunci (sudah digunakan)
                  </span>{' '}
                  ·{' '}
                  <span className="font-semibold text-emerald-700">
                    {availableCount} tersedia
                  </span>
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  type="button"
                  onClick={onBackToRack}
                  className="rounded-lg bg-[#0F4C81] px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] whitespace-nowrap shrink-0"
                >
                  Lanjut ke Visualisasi Rak
                </button>
                <button
                  type="button"
                  disabled={availableCount === 0}
                  onClick={handleDownloadRemainingTemplate}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-40 whitespace-nowrap shrink-0"
                >
                  <Download className="h-3.5 w-3.5" />
                  Unduh Tersisa ({availableCount})
                </button>
                <button
                  type="button"
                  onClick={onClearSpreadsheet}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-white px-3 py-2 text-xs font-semibold text-red-700 transition-colors hover:bg-red-50 whitespace-nowrap shrink-0"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                  Hapus Sumber Data
                </button>
              </div>
            </div>

            {/* Inline Template Editor: Add Single or Bulk Old Locations */}
            <div
              data-testid="template-editor-panel"
              className="mt-4 rounded-xl border border-slate-200 bg-slate-50/70 p-4"
            >
              <div className="flex flex-wrap items-center justify-between gap-3">
                <div>
                  <h4 className="text-xs font-bold text-slate-900">
                    Tambah Data Lokasi Simpan Lama ke Template
                  </h4>
                  <p className="mt-0.5 text-[11px] text-slate-500">
                    Kode &amp; keterangan yang ditambahkan langsung muncul pada
                    pencarian kata kunci di modal klik box rak.
                  </p>
                </div>

                <div className="inline-flex items-center gap-1 rounded-lg border border-slate-200 bg-white p-0.5">
                  <button
                    type="button"
                    data-testid="tab-add-single"
                    onClick={() => {
                      setAddMode('single');
                      setEditorError(null);
                    }}
                    className={`rounded-md px-2.5 py-1 text-xs font-semibold transition-colors whitespace-nowrap ${
                      addMode === 'single'
                        ? 'bg-[#0F4C81] text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Tambah Satuan
                  </button>
                  <button
                    type="button"
                    data-testid="tab-add-bulk"
                    onClick={() => {
                      setAddMode('bulk');
                      setEditorError(null);
                    }}
                    className={`rounded-md px-2.5 py-1 text-xs font-semibold transition-colors whitespace-nowrap ${
                      addMode === 'bulk'
                        ? 'bg-[#0F4C81] text-white'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    Tambah Banyak (Tempel)
                  </button>
                </div>
              </div>

              {addMode === 'single' ? (
                <form
                  onSubmit={handleAddSingleSubmit}
                  className="mt-3 grid grid-cols-1 gap-2.5 sm:grid-cols-12"
                >
                  <div className="sm:col-span-4">
                    <label
                      htmlFor="template-add-code"
                      className="sr-only"
                    >
                      Kode Lokasi Simpan Lama Baru
                    </label>
                    <input
                      id="template-add-code"
                      data-testid="template-add-code-input"
                      type="text"
                      value={newCode}
                      onChange={(e) => {
                        setNewCode(e.target.value);
                        setEditorError(null);
                      }}
                      placeholder="Kode lokasi lama (mis. 401)"
                      className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-xs font-semibold text-slate-900 placeholder:font-sans placeholder:font-normal placeholder:text-slate-400 focus:border-[#0F4C81] focus:outline-none"
                    />
                  </div>
                  <div className="sm:col-span-5">
                    <label
                      htmlFor="template-add-desc"
                      className="sr-only"
                    >
                      Keterangan atau Kata Kunci (opsional)
                    </label>
                    <input
                      id="template-add-desc"
                      data-testid="template-add-desc-input"
                      type="text"
                      value={newDesc}
                      onChange={(e) => setNewDesc(e.target.value)}
                      placeholder="Keterangan / kata kunci opsional (mis. Arsip Pajak 2025)"
                      className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:outline-none"
                    />
                  </div>
                  <div className="sm:col-span-3">
                    <button
                      type="submit"
                      data-testid="template-add-submit-btn"
                      className="inline-flex w-full items-center justify-center gap-1.5 rounded-lg bg-[#0F4C81] px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] whitespace-nowrap"
                    >
                      <Plus className="h-3.5 w-3.5" />
                      Tambah ke Template
                    </button>
                  </div>
                </form>
              ) : (
                <form onSubmit={handleAddBulkSubmit} className="mt-3 space-y-2.5">
                  <textarea
                    data-testid="template-bulk-input"
                    rows={3}
                    value={bulkText}
                    onChange={(e) => {
                      setBulkText(e.target.value);
                      setEditorError(null);
                    }}
                    placeholder={
                      'Tempel daftar kode lokasi lama di sini (1 per baris, pisahkan koma, atau salin langsung dari Excel):\n401\tArsip Keuangan 2025\n402 | Arsip Audit\n403, 404'
                    }
                    className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-xs text-slate-900 placeholder:font-sans placeholder:text-slate-400 focus:border-[#0F4C81] focus:outline-none"
                  />
                  <div className="flex flex-wrap items-center justify-between gap-2">
                    <span className="text-[11px] text-slate-500 tabular-nums">
                      {parsedBulkPreview.length > 0
                        ? `${parsedBulkPreview.length} entri terbaca siap ditambahkan (duplikat dilewati otomatis).`
                        : 'Mendukung format baris baru, koma, tanda |, atau salin-tempel 2 kolom dari Excel.'}
                    </span>
                    <button
                      type="submit"
                      data-testid="template-bulk-submit-btn"
                      disabled={parsedBulkPreview.length === 0}
                      className="inline-flex items-center gap-1.5 rounded-lg bg-[#0F4C81] px-3.5 py-1.5 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap"
                    >
                      <Plus className="h-3.5 w-3.5" />
                      Tambahkan {parsedBulkPreview.length || ''} Kode
                    </button>
                  </div>
                </form>
              )}

              {editorError && (
                <div
                  role="alert"
                  data-testid="template-editor-error"
                  className="mt-2.5 flex items-center gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-800"
                >
                  <AlertCircle className="h-3.5 w-3.5 shrink-0 text-red-600" />
                  <span>{editorError}</span>
                </div>
              )}

              {editorSuccess && !editorError && (
                <div
                  role="status"
                  data-testid="template-editor-success"
                  className="mt-2.5 flex items-center justify-between gap-2 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs text-emerald-900"
                >
                  <div className="flex items-center gap-2">
                    <CheckCircle2 className="h-3.5 w-3.5 shrink-0 text-emerald-600" />
                    <span>{editorSuccess}</span>
                  </div>
                  <button
                    type="button"
                    onClick={() => setEditorSuccess(null)}
                    aria-label="Tutup pemberitahuan"
                    className="rounded p-0.5 text-emerald-700 hover:bg-emerald-100"
                  >
                    <X className="h-3.5 w-3.5" />
                  </button>
                </div>
              )}
            </div>

            {/* Column Picker & Validation Option */}
            <div className="mt-4 flex flex-wrap items-center justify-between gap-4 border-t border-slate-100 pt-4">
              <div className="flex flex-wrap items-center gap-3">
                <label
                  htmlFor="spreadsheet-col-select"
                  className="text-xs font-semibold text-slate-700"
                >
                  Kolom Lokasi Simpan Lama:
                </label>
                <select
                  id="spreadsheet-col-select"
                  value={spreadsheet.selectedColumn}
                  onChange={(e) => onChangeColumn(e.target.value)}
                  className="rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-medium text-slate-900 focus:border-[#0F4C81] focus:outline-none"
                >
                  {spreadsheet.columns.map((col) => {
                    const count = extractSpreadsheetItems(
                      spreadsheet.rawRows,
                      col
                    ).length;
                    return (
                      <option key={col} value={col}>
                        {col} ({count} nilai)
                      </option>
                    );
                  })}
                </select>
              </div>

              <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-600 select-none">
                <input
                  type="checkbox"
                  checked={allowUnlistedWithSpreadsheet}
                  onChange={(e) => onToggleAllowUnlisted(e.target.checked)}
                  className="h-3.5 w-3.5 rounded border-slate-300 text-[#0F4C81] focus:ring-[#0F4C81]"
                />
                <span>Izinkan input kode di luar spreadsheet</span>
              </label>
            </div>

            {/* Filter Tabs + Search within Imported Spreadsheet */}
            <div className="mt-4 flex flex-wrap items-center justify-between gap-3">
              <div className="inline-flex items-center gap-1 rounded-lg bg-slate-100 p-1">
                {(
                  [
                    { id: 'all', label: `Semua (${spreadsheet.items.length})` },
                    { id: 'available', label: `Tersedia (${availableCount})` },
                    { id: 'locked', label: `Terkunci (${lockedCount})` },
                  ] as const
                ).map((opt) => (
                  <button
                    key={opt.id}
                    type="button"
                    onClick={() => setLockFilter(opt.id)}
                    className={`rounded-md px-3 py-1 text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                      lockFilter === opt.id
                        ? 'bg-white text-slate-900 shadow-xs'
                        : 'text-slate-600 hover:text-slate-900'
                    }`}
                  >
                    {opt.label}
                  </button>
                ))}
              </div>

              <div className="relative min-w-[220px] flex-1 sm:max-w-xs">
                <Search
                  aria-hidden="true"
                  className="pointer-events-none absolute top-1/2 left-3 h-3.5 w-3.5 -translate-y-1/2 text-slate-400"
                />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  aria-label="Cari lokasi simpan lama pada spreadsheet"
                  placeholder="Cari kode lama, keterangan, atau kode baru..."
                  className="w-full rounded-lg border border-slate-200 bg-slate-50 py-1.5 pr-3 pl-8 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:bg-white focus:outline-none"
                />
              </div>
            </div>

            {editRowError && (
              <div
                role="alert"
                className="mt-3 flex items-center gap-2 rounded-lg border border-red-200 bg-red-50 px-3 py-2 text-xs text-red-800"
              >
                <AlertCircle className="h-3.5 w-3.5 shrink-0 text-red-600" />
                <span>{editRowError}</span>
              </div>
            )}
          </div>

          <div className="max-h-[440px] overflow-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead className="sticky top-0 z-10 border-b border-slate-200 bg-slate-50 text-[11px] font-bold tracking-wide text-slate-600">
                <tr>
                  <th scope="col" className="w-16 px-4 py-2.5 text-right">
                    Baris
                  </th>
                  <th scope="col" className="px-4 py-2.5">
                    LOKASI SIMPAN LAMA
                  </th>
                  <th scope="col" className="px-4 py-2.5">
                    KETERANGAN / KATA KUNCI
                  </th>
                  <th scope="col" className="px-4 py-2.5">
                    STATUS PENEMPATAN
                  </th>
                  <th scope="col" className="w-28 px-4 py-2.5 text-right">
                    AKSI
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredItems.length === 0 ? (
                  <tr>
                    <td
                      colSpan={5}
                      className="px-4 py-8 text-center text-xs text-slate-500"
                    >
                      {spreadsheet.items.length === 0
                        ? 'Template masih kosong. Gunakan formulir di atas untuk menambahkan kode lokasi simpan lama.'
                        : `Data lokasi simpan lama "${searchQuery}" tidak tersedia pada template ini.`}
                    </td>
                  </tr>
                ) : (
                  filteredItems.map((item) => {
                    const locked = lockedMap.get(item.normalizedKey);
                    const isEditing = editingKey === item.normalizedKey;

                    return (
                      <tr
                        key={`${item.normalizedKey}-${item.rowNumber}`}
                        data-testid={`template-row-${item.oldLocation}`}
                        className={
                          isEditing
                            ? 'bg-sky-50/70'
                            : locked
                              ? 'bg-amber-50/35 hover:bg-amber-50/60'
                              : 'hover:bg-slate-50/80'
                        }
                      >
                        <td className="px-4 py-2.5 text-right font-mono text-slate-400 tabular-nums">
                          {item.rowNumber}
                        </td>

                        <td className="px-4 py-2.5 font-mono font-semibold tabular-nums">
                          {isEditing ? (
                            <input
                              type="text"
                              data-testid="template-edit-code-input"
                              value={editCode}
                              disabled={Boolean(locked)}
                              title={
                                locked
                                  ? 'Kode sudah terkunci di hasil simulasi; hanya keterangan yang bisa diedit'
                                  : 'Edit kode lokasi simpan lama'
                              }
                              onChange={(e) => {
                                setEditCode(e.target.value);
                                setEditRowError(null);
                              }}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  handleSaveInlineEdit(item.normalizedKey);
                                } else if (e.key === 'Escape') {
                                  cancelInlineEdit();
                                }
                              }}
                              className="w-36 rounded-md border border-[#0F4C81] bg-white px-2 py-1 font-mono text-xs font-semibold text-slate-900 focus:outline-none disabled:cursor-not-allowed disabled:border-slate-200 disabled:bg-slate-100 disabled:text-slate-500"
                            />
                          ) : (
                            <span
                              className={
                                locked
                                  ? 'text-slate-400 line-through'
                                  : 'text-slate-900'
                              }
                            >
                              {item.oldLocation}
                            </span>
                          )}
                        </td>

                        <td className="px-4 py-2.5 text-slate-600">
                          {isEditing ? (
                            <input
                              type="text"
                              data-testid="template-edit-desc-input"
                              value={editDesc}
                              onChange={(e) => setEditDesc(e.target.value)}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') {
                                  e.preventDefault();
                                  handleSaveInlineEdit(item.normalizedKey);
                                } else if (e.key === 'Escape') {
                                  cancelInlineEdit();
                                }
                              }}
                              placeholder="Keterangan / kata kunci..."
                              className="w-full min-w-[160px] rounded-md border border-[#0F4C81] bg-white px-2 py-1 text-xs text-slate-900 focus:outline-none"
                            />
                          ) : (
                            item.description || '—'
                          )}
                        </td>

                        <td className="px-4 py-2.5">
                          {locked ? (
                            <span className="inline-flex items-center gap-1.5 font-mono font-semibold text-amber-800 tabular-nums">
                              <Lock className="h-3.5 w-3.5 shrink-0 text-amber-700" />
                              {locked.source === 'draft'
                                ? `Terkunci · Draft Rak → ${locked.newLocation}`
                                : `Terkunci · Sudah Tersimulasi → ${locked.newLocation}`}
                            </span>
                          ) : (
                            <span className="font-medium text-emerald-700">
                              Tersedia
                            </span>
                          )}
                        </td>

                        <td className="px-4 py-2.5 text-right">
                          {isEditing ? (
                            <div className="inline-flex items-center justify-end gap-1">
                              <button
                                type="button"
                                data-testid="template-edit-save-btn"
                                onClick={() =>
                                  handleSaveInlineEdit(item.normalizedKey)
                                }
                                title="Simpan perubahan"
                                aria-label={`Simpan perubahan ${item.oldLocation}`}
                                className="inline-flex items-center gap-1 rounded-md bg-emerald-700 px-2 py-1 text-[11px] font-semibold text-white hover:bg-emerald-800"
                              >
                                <Check className="h-3 w-3" />
                                Simpan
                              </button>
                              <button
                                type="button"
                                onClick={cancelInlineEdit}
                                title="Batalkan edit"
                                aria-label="Batalkan edit"
                                className="rounded-md border border-slate-300 bg-white p-1 text-slate-600 hover:bg-slate-50"
                              >
                                <X className="h-3 w-3" />
                              </button>
                            </div>
                          ) : (
                            <div className="inline-flex items-center justify-end gap-1">
                              <button
                                type="button"
                                data-testid={`template-edit-btn-${item.oldLocation}`}
                                onClick={() =>
                                  startInlineEdit(
                                    item.normalizedKey,
                                    item.oldLocation,
                                    item.description
                                  )
                                }
                                title={
                                  locked
                                    ? 'Edit keterangan baris ini'
                                    : `Edit kode & keterangan ${item.oldLocation}`
                                }
                                aria-label={`Edit ${item.oldLocation}`}
                                className="rounded-md p-1.5 text-slate-500 transition-colors hover:bg-sky-50 hover:text-[#0F4C81]"
                              >
                                <Pencil className="h-3.5 w-3.5" />
                              </button>
                              <button
                                type="button"
                                data-testid={`template-delete-btn-${item.oldLocation}`}
                                disabled={Boolean(locked)}
                                onClick={() =>
                                  handleDeleteTemplateRow(item.normalizedKey)
                                }
                                title={
                                  locked
                                    ? 'Tidak dapat menghapus kode yang sedang terkunci di hasil simulasi'
                                    : `Hapus ${item.oldLocation} dari template`
                                }
                                aria-label={`Hapus ${item.oldLocation}`}
                                className="rounded-md p-1.5 text-slate-400 transition-colors hover:bg-red-50 hover:text-red-600 disabled:cursor-not-allowed disabled:opacity-35"
                              >
                                <Trash2 className="h-3.5 w-3.5" />
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </div>
  );
};

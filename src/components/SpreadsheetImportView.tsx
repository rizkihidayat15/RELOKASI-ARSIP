import React, { useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Search,
  Trash2,
  Upload,
} from 'lucide-react';
import type { ImportedSpreadsheet, PlacementRecord } from '../types/archive';
import {
  downloadSpreadsheetTemplate,
  parseSpreadsheetFile,
} from '../utils/excelService';
import {
  extractSpreadsheetItems,
  normalizeLocationCode,
} from '../utils/relocationEngine';

interface SpreadsheetImportViewProps {
  spreadsheet: ImportedSpreadsheet | null;
  placements: PlacementRecord[];
  allowUnlistedWithSpreadsheet: boolean;
  onImportSpreadsheet: (sheet: ImportedSpreadsheet) => void;
  onChangeColumn: (nextColumn: string) => void;
  onClearSpreadsheet: () => void;
  onToggleAllowUnlisted: (allowed: boolean) => void;
  onBackToRack: () => void;
}

export const SpreadsheetImportView: React.FC<SpreadsheetImportViewProps> = ({
  spreadsheet,
  placements,
  allowUnlistedWithSpreadsheet,
  onImportSpreadsheet,
  onChangeColumn,
  onClearSpreadsheet,
  onToggleAllowUnlisted,
  onBackToRack,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isLoading, setIsLoading] = useState(false);
  const [errorMsg, setErrorMsg] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState('');
  const [isDragging, setIsDragging] = useState(false);

  const placedMap = useMemo(() => {
    const map = new Map<string, PlacementRecord>();
    for (const p of placements) {
      map.set(normalizeLocationCode(p.oldLocation), p);
    }
    return map;
  }, [placements]);

  const handleFileSelected = async (file: File | null) => {
    if (!file) return;
    setErrorMsg(null);
    setIsLoading(true);
    try {
      const parsed = await parseSpreadsheetFile(file);
      onImportSpreadsheet(parsed);
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

  const filteredItems = useMemo(() => {
    if (!spreadsheet) return [];
    const q = searchQuery.trim().toLowerCase();
    if (!q) return spreadsheet.items;
    return spreadsheet.items.filter(
      (item) =>
        item.oldLocation.toLowerCase().includes(q) ||
        (item.description?.toLowerCase().includes(q) ?? false)
    );
  }, [spreadsheet, searchQuery]);

  const movedCount = useMemo(() => {
    if (!spreadsheet) return 0;
    return spreadsheet.items.filter((item) => placedMap.has(item.normalizedKey))
      .length;
  }, [spreadsheet, placedMap]);

  return (
    <div className="space-y-6">
      {/* Upload Card */}
      <section className="rounded-xl border border-slate-200 bg-white p-5">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h2 className="text-base font-bold tracking-tight text-slate-900">
              Impor Spreadsheet Lokasi Simpan Lama
            </h2>
            <p className="mt-0.5 max-w-2xl text-xs leading-relaxed text-slate-500">
              Unggah file Excel (.xlsx, .xls) atau CSV berisi daftar lokasi simpan
              lama aktual sebagai sumber pencarian dan validasi pada modal klik box rak.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {spreadsheet && (
              <label
                htmlFor="spreadsheet-file-input"
                className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg bg-[#0F4C81] px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] whitespace-nowrap shrink-0"
              >
                <Upload className="h-3.5 w-3.5" />
                Ganti File Spreadsheet
              </label>
            )}
            <button
              type="button"
              onClick={downloadSpreadsheetTemplate}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap shrink-0"
            >
              <Download className="h-3.5 w-3.5" />
              Unduh Template Contoh (.xlsx)
            </button>
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
                : 'Pilih atau seret file spreadsheet ke sini'}
            </p>
            <p className="mt-1 text-xs text-slate-500">
              Mendukung format .xlsx, .xls, dan .csv
            </p>

            <label
              htmlFor="spreadsheet-file-input"
              className="mt-4 inline-flex cursor-pointer items-center gap-2 rounded-lg bg-[#0F4C81] px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] whitespace-nowrap"
            >
              <FileSpreadsheet className="h-4 w-4" />
              Pilih File Spreadsheet
            </label>
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

      {/* Active Imported Spreadsheet Preview */}
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
                  Sheet: {spreadsheet.sheetName} · {spreadsheet.items.length}{' '}
                  lokasi unik · {movedCount} sudah dipindahkan ·{' '}
                  {spreadsheet.items.length - movedCount} belum dipindahkan
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
                  onClick={onClearSpreadsheet}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-white px-3 py-2 text-xs font-semibold text-red-700 transition-colors hover:bg-red-50 whitespace-nowrap shrink-0"
                >
                  <Trash2 className="h-3.5 w-3.5" />
                  Hapus Sumber Data
                </button>
              </div>
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

            {/* Search within Imported Spreadsheet */}
            <div className="relative mt-4 max-w-sm">
              <Search
                aria-hidden="true"
                className="pointer-events-none absolute top-1/2 left-3 h-3.5 w-3.5 -translate-y-1/2 text-slate-400"
              />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                aria-label="Cari lokasi simpan lama pada spreadsheet"
                placeholder="Cari kode lokasi pada spreadsheet..."
                className="w-full rounded-lg border border-slate-200 bg-slate-50 py-1.5 pr-3 pl-8 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:bg-white focus:outline-none"
              />
            </div>
          </div>

          <div className="max-h-[420px] overflow-auto">
            <table className="w-full border-collapse text-left text-xs">
              <thead className="sticky top-0 border-b border-slate-200 bg-slate-50 text-[11px] font-bold tracking-wide text-slate-600">
                <tr>
                  <th scope="col" className="w-16 px-4 py-2.5 text-right">
                    Baris
                  </th>
                  <th scope="col" className="px-4 py-2.5">
                    LOKASI SIMPAN LAMA
                  </th>
                  <th scope="col" className="px-4 py-2.5">
                    KETERANGAN
                  </th>
                  <th scope="col" className="px-4 py-2.5">
                    STATUS PENEMPATAN
                  </th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-100">
                {filteredItems.length === 0 ? (
                  <tr>
                    <td
                      colSpan={4}
                      className="px-4 py-8 text-center text-xs text-slate-500"
                    >
                      Data lokasi simpan lama "{searchQuery}" tidak tersedia pada
                      spreadsheet ini.
                    </td>
                  </tr>
                ) : (
                  filteredItems.map((item) => {
                    const placed = placedMap.get(item.normalizedKey);
                    return (
                      <tr
                        key={`${item.normalizedKey}-${item.rowNumber}`}
                        className="hover:bg-slate-50/80"
                      >
                        <td className="px-4 py-2.5 text-right font-mono text-slate-400 tabular-nums">
                          {item.rowNumber}
                        </td>
                        <td className="px-4 py-2.5 font-mono font-semibold text-slate-900 tabular-nums">
                          {item.oldLocation}
                        </td>
                        <td className="px-4 py-2.5 text-slate-600">
                          {item.description || '—'}
                        </td>
                        <td className="px-4 py-2.5">
                          {placed ? (
                            <span className="font-mono font-semibold text-emerald-700 tabular-nums">
                              Terpasang di {placed.newLocation}
                            </span>
                          ) : (
                            <span className="text-slate-500">
                              Belum dipindahkan
                            </span>
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

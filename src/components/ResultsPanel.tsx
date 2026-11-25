import React, { useMemo, useState } from 'react';
import {
  CheckCircle2,
  Download,
  FileSpreadsheet,
  Info,
  MapPin,
  Search,
  Trash2,
} from 'lucide-react';
import type { BoxCoordinate, PlacementRecord } from '../types/archive';
import {
  EXPORT_EXCEL_FILENAME,
  type GeneratedExcelArtifact,
} from '../utils/excelService';

interface ResultsPanelProps {
  savedRecords: PlacementRecord[];
  unsavedDraftCount: number;
  includeExtraColumns: boolean;
  generatedArtifact: GeneratedExcelArtifact | null;
  isArtifactStale: boolean;
  fullWidth?: boolean;
  onToggleExtraColumns: (checked: boolean) => void;
  onRemoveSavedRecord: (placementId: string) => void;
  onRequestClearSavedExcel: () => void;
  onJumpToBox: (coord: BoxCoordinate) => void;
  onGenerateExcel: () => void;
  onDownloadExcel: () => void;
}

export const ResultsPanel: React.FC<ResultsPanelProps> = ({
  savedRecords,
  unsavedDraftCount,
  includeExtraColumns,
  generatedArtifact,
  isArtifactStale,
  fullWidth = false,
  onToggleExtraColumns,
  onRemoveSavedRecord,
  onRequestClearSavedExcel,
  onJumpToBox,
  onGenerateExcel,
  onDownloadExcel,
}) => {
  const [tableSearch, setTableSearch] = useState('');

  const orderedRecords = useMemo(
    () => [...savedRecords].sort((a, b) => a.sequence - b.sequence),
    [savedRecords]
  );

  const filteredRecords = useMemo(() => {
    const q = tableSearch.trim().toLowerCase();
    if (!q) return orderedRecords;
    return orderedRecords.filter(
      (p) =>
        p.oldLocation.toLowerCase().includes(q) ||
        p.newLocation.toLowerCase().includes(q)
    );
  }, [orderedRecords, tableSearch]);

  const hasData = savedRecords.length > 0;

  return (
    <section
      aria-label="Tabel Hasil Simpan Excel"
      className="flex flex-col rounded-xl border border-slate-200 bg-white"
    >
      {/* Header */}
      <div className="border-b border-slate-200 px-5 py-4">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <div>
            <h2 className="text-base font-bold tracking-tight text-slate-900">
              Hasil Simpan Excel
            </h2>
            <p className="mt-0.5 text-xs text-slate-500 tabular-nums">
              Akumulasi antar siklus rak · {savedRecords.length} baris tersimpan
            </p>
          </div>
          {hasData ? (
            <button
              type="button"
              onClick={onRequestClearSavedExcel}
              className="text-xs font-medium text-red-600 hover:text-red-700 hover:underline whitespace-nowrap shrink-0"
            >
              Kosongkan Excel
            </button>
          ) : (
            <span className="font-mono text-xs text-slate-500">
              {includeExtraColumns ? '6 kolom' : '2 kolom utama'}
            </span>
          )}
        </div>

        {/* Unsaved Draft Reminder */}
        {unsavedDraftCount > 0 && (
          <div
            data-testid="unsaved-draft-notice"
            className="mt-3 flex items-start gap-2 rounded-lg border border-sky-200 bg-sky-50 px-3 py-2 text-xs text-sky-900"
          >
            <Info className="mt-0.5 h-3.5 w-3.5 shrink-0 text-[#0F4C81]" />
            <span>
              <strong className="font-semibold tabular-nums">
                {unsavedDraftCount} box
              </strong>{' '}
              pada simulasi rak aktif belum masuk Excel. Klik tombol{' '}
              <strong className="font-semibold">Simpan</strong> untuk
              menambahkannya ke daftar ini.
            </span>
          </div>
        )}

        {/* Export Action Bar */}
        <div className="mt-3.5 flex flex-wrap items-center gap-2">
          <button
            type="button"
            disabled={!hasData}
            onClick={onGenerateExcel}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3.5 py-2 text-xs font-semibold text-slate-800 transition-colors hover:bg-slate-50 disabled:opacity-45 whitespace-nowrap shrink-0"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-[#0F4C81]" />
            Generate Excel
          </button>
          <button
            type="button"
            disabled={!hasData}
            onClick={onDownloadExcel}
            className="inline-flex items-center gap-1.5 rounded-lg bg-[#0F4C81] px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap shrink-0"
          >
            <Download className="h-3.5 w-3.5" />
            Download Excel
          </button>
        </div>

        {/* Optional Extra Columns Toggle */}
        <label className="mt-3 flex cursor-pointer items-center gap-2 text-xs text-slate-600 select-none">
          <input
            type="checkbox"
            checked={includeExtraColumns}
            onChange={(e) => onToggleExtraColumns(e.target.checked)}
            className="h-3.5 w-3.5 rounded border-slate-300 text-[#0F4C81] focus:ring-[#0F4C81]"
          />
          <span>Sertakan kolom tambahan (Rak, Row, Box, Waktu)</span>
        </label>

        {/* Generated Excel Readiness Banner */}
        {generatedArtifact && hasData && (
          <div
            data-testid="excel-generated-banner"
            className={`mt-3 rounded-lg border px-3.5 py-2.5 text-xs ${
              isArtifactStale
                ? 'border-amber-200 bg-amber-50 text-amber-900'
                : 'border-emerald-200 bg-emerald-50 text-emerald-900'
            }`}
          >
            <div className="flex items-start gap-2">
              <CheckCircle2
                className={`mt-0.5 h-4 w-4 shrink-0 ${
                  isArtifactStale ? 'text-amber-600' : 'text-emerald-600'
                }`}
              />
              <div className="min-w-0">
                <p className="font-semibold">
                  {isArtifactStale
                    ? 'Data berubah sejak terakhir Generate Excel'
                    : 'File Excel siap diunduh'}
                </p>
                <p className="mt-0.5 truncate font-mono text-[11px]">
                  {EXPORT_EXCEL_FILENAME}
                </p>
                <p className="mt-0.5 text-[11px] opacity-85 tabular-nums">
                  {generatedArtifact.rowCount} baris · Kolom:{' '}
                  {generatedArtifact.columnNames.join(', ')} ·{' '}
                  {(generatedArtifact.byteSize / 1024).toFixed(1)} KB
                </p>
              </div>
            </div>
          </div>
        )}

        {/* Table Search */}
        {hasData && (
          <div className="relative mt-3">
            <Search
              aria-hidden="true"
              className="pointer-events-none absolute top-1/2 left-3 h-3.5 w-3.5 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              value={tableSearch}
              onChange={(e) => setTableSearch(e.target.value)}
              aria-label="Filter tabel hasil simpan Excel"
              placeholder="Filter FORMAT LAMA atau FORMAT BARU..."
              className="w-full rounded-lg border border-slate-200 bg-slate-50 py-1.5 pr-3 pl-8 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:bg-white focus:outline-none"
            />
          </div>
        )}
      </div>

      {/* Table Body / Empty State */}
      {!hasData ? (
        <div className="flex flex-1 flex-col items-center justify-center px-6 py-12 text-center">
          <div className="flex h-10 w-10 items-center justify-center rounded-lg border border-slate-200 bg-slate-50 text-slate-400">
            <FileSpreadsheet className="h-5 w-5" />
          </div>
          <p className="mt-3 text-sm font-semibold text-slate-800">
            Belum ada data tersimpan di Excel
          </p>
          <p className="mt-1 max-w-xs text-xs leading-relaxed text-slate-500">
            Isi box pada simulasi rak di samping, lalu klik tombol{' '}
            <span className="font-semibold text-slate-700">Simpan</span> agar
            hasil pemindahan masuk ke daftar Excel ini.
          </p>
        </div>
      ) : (
        <div
          className={`overflow-x-auto ${
            fullWidth ? '' : 'max-h-[540px] overflow-y-auto'
          }`}
        >
          <table className="w-full border-collapse text-left text-xs">
            <thead className="sticky top-0 z-10 border-b border-slate-200 bg-slate-50 text-[11px] font-bold tracking-wide text-slate-600">
              <tr>
                <th scope="col" className="w-10 px-3 py-2.5 text-right">
                  No
                </th>
                <th scope="col" className="px-3 py-2.5">
                  FORMAT LAMA
                </th>
                <th scope="col" className="px-3 py-2.5">
                  FORMAT BARU
                </th>
                {includeExtraColumns && (
                  <>
                    <th scope="col" className="px-3 py-2.5 text-right">
                      RAK
                    </th>
                    <th scope="col" className="px-3 py-2.5 text-right">
                      ROW
                    </th>
                    <th scope="col" className="px-3 py-2.5 text-right">
                      BOX
                    </th>
                  </>
                )}
                <th scope="col" className="w-20 px-3 py-2.5 text-right">
                  Aksi
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {filteredRecords.length === 0 ? (
                <tr>
                  <td
                    colSpan={includeExtraColumns ? 7 : 4}
                    className="px-4 py-8 text-center text-xs text-slate-500"
                  >
                    Tidak ada baris yang cocok dengan pencarian "{tableSearch}".
                  </td>
                </tr>
              ) : (
                filteredRecords.map((item, idx) => (
                  <tr
                    key={item.id}
                    data-testid={`excel-row-${idx + 1}`}
                    className="transition-colors hover:bg-slate-50/90"
                  >
                    <td className="px-3 py-2.5 text-right font-mono text-slate-400 tabular-nums">
                      {idx + 1}
                    </td>
                    <td className="px-3 py-2.5 font-mono font-semibold text-slate-900 tabular-nums">
                      {item.oldLocation}
                    </td>
                    <td className="px-3 py-2.5 font-mono font-semibold text-[#0F4C81] tabular-nums">
                      {item.newLocation}
                    </td>
                    {includeExtraColumns && (
                      <>
                        <td className="px-3 py-2.5 text-right font-mono text-slate-600 tabular-nums">
                          {item.rack}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono text-slate-600 tabular-nums">
                          {item.row}
                        </td>
                        <td className="px-3 py-2.5 text-right font-mono text-slate-600 tabular-nums">
                          {item.box}
                        </td>
                      </>
                    )}
                    <td className="px-3 py-2.5 text-right">
                      <div className="inline-flex items-center justify-end gap-1">
                        <button
                          type="button"
                          onClick={() =>
                            onJumpToBox({
                              locationCode: item.locationCode,
                              rack: item.rack,
                              row: item.row,
                              box: item.box,
                            })
                          }
                          aria-label={`Lihat lokasi ${item.oldLocation} di ${item.newLocation}`}
                          title="Tampilkan rak lokasi ini"
                          className="rounded p-1 text-slate-400 transition-colors hover:bg-slate-100 hover:text-[#0F4C81]"
                        >
                          <MapPin className="h-3.5 w-3.5" />
                        </button>
                        <button
                          type="button"
                          onClick={() => onRemoveSavedRecord(item.id)}
                          aria-label={`Hapus ${item.oldLocation} (${item.newLocation}) dari Excel`}
                          title="Hapus dari daftar Excel"
                          className="rounded p-1 text-slate-400 transition-colors hover:bg-red-50 hover:text-red-600"
                        >
                          <Trash2 className="h-3.5 w-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      )}

      {/* Footer note on Excel output */}
      <div className="border-t border-slate-100 bg-slate-50/70 px-5 py-2.5 text-[11px] text-slate-500">
        Nama file ekspor:{' '}
        <span className="font-mono font-medium text-slate-700">
          {EXPORT_EXCEL_FILENAME}
        </span>
      </div>
    </section>
  );
};

import React, { useRef } from 'react';
import { FileSpreadsheet, RefreshCw, X } from 'lucide-react';
import type { SimulationMetrics } from '../types/archive';
import type { ConnectedGoogleSheet } from '../utils/googleSheetsSync';

export type TemplateFilterPreset = 'all' | 'available' | 'locked';

interface MetricsBarProps {
  metrics: SimulationMetrics;
  activeLocationCode: string;
  activeRack: number;
  spreadsheetFileName?: string;
  connectedSheet: ConnectedGoogleSheet | null;
  isGoogleAuthenticated: boolean;
  isRefreshingDashboard: boolean;
  externalResultSourceName?: string | null;
  externalResultSourceType?: 'google_sheets' | 'local_file' | null;
  onRefreshFromGoogleSheet: () => void;
  onImportExternalResultsFile: (file: File) => void;
  onClearExternalResults: () => void;
  onOpenTemplateWithFilter: (filter: TemplateFilterPreset) => void;
  onOpenResultsTab: () => void;
}

export const MetricsBar: React.FC<MetricsBarProps> = ({
  metrics,
  activeLocationCode,
  activeRack,
  spreadsheetFileName,
  connectedSheet,
  isGoogleAuthenticated,
  isRefreshingDashboard,
  externalResultSourceName,
  externalResultSourceType,
  onRefreshFromGoogleSheet,
  onImportExternalResultsFile,
  onClearExternalResults,
  onOpenTemplateWithFilter,
  onOpenResultsTab,
}) => {
  const resultFileInputRef = useRef<HTMLInputElement>(null);
  const hasTemplate = metrics.templateTotal !== null;

  const handleResultFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onImportExternalResultsFile(file);
    }
    if (resultFileInputRef.current) {
      resultFileInputRef.current.value = '';
    }
  };

  return (
    <section
      aria-label="Dashboard Progres Relokasi"
      data-testid="dashboard-metrics"
      className="rounded-xl border border-slate-200 bg-white px-5 py-4"
    >
      {/* Dashboard Header: Template & Spreadsheet Hasil Sync Context */}
      <div className="flex flex-wrap items-start justify-between gap-3 border-b border-slate-100 pb-3">
        <div className="min-w-0">
          <h1 className="text-sm font-bold tracking-tight text-slate-900">
            Dashboard Progres Relokasi
          </h1>
          <div className="mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-0.5 text-xs text-slate-500">
            {spreadsheetFileName ? (
              <span>
                Template:{' '}
                <button
                  type="button"
                  onClick={() => onOpenTemplateWithFilter('all')}
                  className="font-semibold text-slate-700 hover:text-[#0F4C81] hover:underline"
                >
                  {spreadsheetFileName}
                </button>
              </span>
            ) : (
              <span>
                Template lokasi lama belum dimuat ·{' '}
                <button
                  type="button"
                  onClick={() => onOpenTemplateWithFilter('all')}
                  className="font-semibold text-[#0F4C81] hover:underline"
                >
                  Impor Template
                </button>
              </span>
            )}

            <span aria-hidden="true" className="text-slate-300">
              ·
            </span>

            {connectedSheet ? (
              <span className="tabular-nums">
                Spreadsheet Hasil:{' '}
                <button
                  type="button"
                  onClick={onOpenResultsTab}
                  className={`font-semibold hover:underline ${
                    isGoogleAuthenticated
                      ? 'text-emerald-700'
                      : 'text-amber-700'
                  }`}
                >
                  {connectedSheet.title} ({connectedSheet.sheetName})
                </button>
                {!isGoogleAuthenticated
                  ? ' · Sesi Google belum aktif'
                  : metrics.externalResultsCount > 0
                    ? ` · ${metrics.externalResultsCount} baris terbaca`
                    : ''}
              </span>
            ) : externalResultSourceType === 'local_file' &&
              externalResultSourceName ? (
              <span className="inline-flex items-center gap-1.5 tabular-nums">
                <span>
                  File Hasil:{' '}
                  <strong className="font-semibold text-emerald-700">
                    {externalResultSourceName}
                  </strong>{' '}
                  ({metrics.externalResultsCount} baris)
                </span>
                <button
                  type="button"
                  onClick={onClearExternalResults}
                  aria-label="Lepas file hasil relokasi"
                  title="Lepas file hasil relokasi"
                  className="rounded p-0.5 text-slate-400 hover:bg-red-50 hover:text-red-600"
                >
                  <X className="h-3 w-3" />
                </button>
              </span>
            ) : (
              <span className="tabular-nums">
                Hasil: Daftar Simpan Lokal ({metrics.savedExcelCount} baris) ·{' '}
                <button
                  type="button"
                  onClick={onOpenResultsTab}
                  className="font-semibold text-[#0F4C81] hover:underline"
                >
                  Hubungkan Google Sheets
                </button>
              </span>
            )}
          </div>
        </div>

        {/* Sync Actions */}
        <div className="flex flex-wrap items-center gap-2 shrink-0">
          {connectedSheet && isGoogleAuthenticated && (
            <button
              type="button"
              disabled={isRefreshingDashboard}
              onClick={onRefreshFromGoogleSheet}
              data-testid="btn-sync-dashboard"
              className="inline-flex items-center gap-1.5 rounded-lg border border-emerald-200 bg-emerald-50 px-3 py-1.5 text-xs font-semibold text-emerald-800 transition-colors hover:bg-emerald-100 disabled:opacity-50 whitespace-nowrap shrink-0"
            >
              <RefreshCw
                className={`h-3.5 w-3.5 ${isRefreshingDashboard ? 'animate-spin' : ''}`}
              />
              {isRefreshingDashboard
                ? 'Membaca Sheets...'
                : 'Sinkronkan Dashboard'}
            </button>
          )}

          <input
            ref={resultFileInputRef}
            id="dashboard-result-file-input"
            type="file"
            accept=".xlsx,.xls,.csv"
            onChange={handleResultFileChange}
            className="sr-only"
          />
          <label
            htmlFor="dashboard-result-file-input"
            data-testid="label-import-result-file"
            className="inline-flex cursor-pointer items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap shrink-0"
          >
            <FileSpreadsheet className="h-3.5 w-3.5 text-[#0F4C81]" />
            Baca File Hasil (.xlsx)
          </label>
        </div>
      </div>

      {/* Primary KPI Strip */}
      {hasTemplate ? (
        <dl className="mt-3.5 grid grid-cols-2 gap-4 sm:grid-cols-4 lg:gap-0 lg:divide-x lg:divide-slate-200">
          <div className="lg:pr-6">
            <dt className="text-xs text-slate-500">Total Kode Template</dt>
            <dd
              data-testid="metric-template-total"
              className="mt-0.5 font-mono text-2xl font-bold text-slate-900 tabular-nums"
            >
              {metrics.templateTotal}
            </dd>
            <p className="mt-0.5 text-[11px] text-slate-500">
              kode lokasi simpan lama
            </p>
          </div>

          <div className="lg:px-6">
            <button
              type="button"
              onClick={() => onOpenTemplateWithFilter('locked')}
              className="group w-full text-left focus:outline-none"
            >
              <dt className="text-xs text-slate-500 group-hover:text-emerald-700">
                Sudah Tersimulasi (Tabel Hasil)
              </dt>
              <dd className="mt-0.5 flex items-baseline gap-2">
                <span
                  data-testid="metric-template-registered"
                  className="font-mono text-2xl font-bold text-emerald-700 tabular-nums"
                >
                  {metrics.templateRegisteredCount}
                </span>
                {(metrics.templateDraftOnlyCount ?? 0) > 0 && (
                  <span
                    data-testid="metric-template-draft"
                    className="font-mono text-xs font-semibold text-[#0F4C81] tabular-nums"
                  >
                    +{metrics.templateDraftOnlyCount} draft
                  </span>
                )}
              </dd>
              <p className="mt-0.5 text-[11px] text-slate-500">
                terdaftar di tabel hasil
              </p>
            </button>
          </div>

          <div className="lg:px-6">
            <button
              type="button"
              onClick={() => onOpenTemplateWithFilter('available')}
              className="group w-full text-left focus:outline-none"
            >
              <dt className="text-xs text-slate-500 group-hover:text-amber-700">
                Belum Tersimulasi
              </dt>
              <dd
                data-testid="metric-template-remaining"
                className={`mt-0.5 font-mono text-2xl font-bold tabular-nums ${
                  (metrics.templateRemainingCount ?? 0) === 0
                    ? 'text-emerald-700'
                    : 'text-amber-700'
                }`}
              >
                {metrics.templateRemainingCount}
              </dd>
              <p className="mt-0.5 text-[11px] text-slate-500">
                {(metrics.templateRemainingCount ?? 0) === 0
                  ? 'seluruh kode sudah tersimulasi'
                  : 'belum terdaftar di tabel hasil'}
              </p>
            </button>
          </div>

          <div className="lg:pl-6">
            <dt className="text-xs text-slate-500">Progres Relokasi</dt>
            <dd className="mt-0.5 flex items-baseline gap-2">
              <span
                data-testid="metric-progress-percent"
                className="font-mono text-2xl font-bold text-[#0F4C81] tabular-nums"
              >
                {metrics.progressPercent}%
              </span>
              <span className="text-xs text-slate-500 tabular-nums">
                {metrics.templateSimulatedCount} / {metrics.templateTotal} kode
              </span>
            </dd>
            <p className="mt-0.5 text-[11px] text-slate-500 tabular-nums">
              {metrics.registeredProgressPercent}% tersimpan di hasil
            </p>
          </div>
        </dl>
      ) : (
        <dl className="mt-3.5 grid grid-cols-2 gap-4 sm:grid-cols-4 lg:gap-0 lg:divide-x lg:divide-slate-200">
          <div className="lg:pr-6">
            <dt className="text-xs text-slate-500">
              Box Terisi ({activeLocationCode || 'RC2'} · Rak {activeRack})
            </dt>
            <dd
              data-testid="metric-filled-boxes"
              className="mt-0.5 font-mono text-2xl font-bold text-[#0F4C81] tabular-nums"
            >
              {metrics.rackFilledBoxes}
            </dd>
          </div>
          <div className="lg:px-6">
            <dt className="text-xs text-slate-500">Box Kosong (Rak Aktif)</dt>
            <dd
              data-testid="metric-empty-boxes"
              className="mt-0.5 font-mono text-2xl font-bold text-slate-900 tabular-nums"
            >
              {metrics.rackEmptyBoxes}
            </dd>
          </div>
          <div className="lg:px-6">
            <dt className="text-xs text-slate-500">
              Terdaftar di Spreadsheet Hasil
            </dt>
            <dd
              data-testid="metric-saved-excel"
              className="mt-0.5 font-mono text-2xl font-bold text-emerald-700 tabular-nums"
            >
              {metrics.spreadsheetResultsTotal}
            </dd>
          </div>
          <div className="lg:pl-6">
            <dt className="text-xs text-slate-500">Progres Rak Aktif</dt>
            <dd className="mt-0.5 flex items-baseline gap-2">
              <span
                data-testid="metric-progress-percent"
                className="font-mono text-2xl font-bold text-slate-900 tabular-nums"
              >
                {metrics.progressPercent}%
              </span>
              <span className="text-xs text-slate-500 tabular-nums">
                {metrics.progressLabel}
              </span>
            </dd>
          </div>
        </dl>
      )}

      {/* Dual-Segment Linear Progress Track + Secondary Telemetry */}
      <div className="mt-3.5">
        <div
          role="progressbar"
          aria-label="Progres relokasi kode template terhadap spreadsheet hasil"
          aria-valuenow={metrics.progressPercent}
          aria-valuemin={0}
          aria-valuemax={100}
          className="flex h-2 w-full overflow-hidden rounded-full bg-slate-100"
        >
          {metrics.registeredProgressPercent > 0 && (
            <div
              className="h-full bg-emerald-600 transition-all duration-200"
              style={{ width: `${metrics.registeredProgressPercent}%` }}
            />
          )}
          {metrics.draftProgressPercent > 0 && (
            <div
              className="h-full bg-[#0F4C81] transition-all duration-200"
              style={{ width: `${metrics.draftProgressPercent}%` }}
            />
          )}
        </div>

        <div className="mt-2.5 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          {hasTemplate ? (
            <>
              <span className="tabular-nums">
                Rak Aktif ({activeLocationCode || 'RC2'} · Rak {activeRack}):{' '}
                <strong
                  data-testid="metric-filled-boxes"
                  className="font-semibold text-slate-700"
                >
                  {metrics.rackFilledBoxes}
                </strong>{' '}
                terisi ·{' '}
                <strong
                  data-testid="metric-empty-boxes"
                  className="font-semibold text-slate-700"
                >
                  {metrics.rackEmptyBoxes}
                </strong>{' '}
                kosong · Total baris Hasil:{' '}
                <strong
                  data-testid="metric-saved-excel"
                  className="font-semibold text-slate-700"
                >
                  {metrics.spreadsheetResultsTotal}
                </strong>
                {metrics.unmatchedResultCount > 0
                  ? ` (${metrics.unmatchedResultCount} di luar template)`
                  : ''}
              </span>

              <div className="flex flex-wrap items-center gap-4 text-[11px] tabular-nums">
                <span className="inline-flex items-center gap-1.5 text-slate-600">
                  <span className="h-2 w-2 rounded-xs bg-emerald-600" />
                  Sudah Tersimulasi ({metrics.templateRegisteredCount})
                </span>
                {(metrics.templateDraftOnlyCount ?? 0) > 0 && (
                  <span className="inline-flex items-center gap-1.5 text-slate-600">
                    <span className="h-2 w-2 rounded-xs bg-[#0F4C81]" />
                    Draft Rak ({metrics.templateDraftOnlyCount})
                  </span>
                )}
                <span className="inline-flex items-center gap-1.5 text-slate-600">
                  <span className="h-2 w-2 rounded-xs bg-slate-300" />
                  Belum Tersimulasi ({metrics.templateRemainingCount})
                </span>
              </div>
            </>
          ) : (
            <>
              <span className="tabular-nums">
                Kapasitas 1 rak: {metrics.rackTotalBoxes} box
                {metrics.unsavedDraftCount > 0
                  ? ` · ${metrics.unsavedDraftCount} box belum disimpan ke Hasil`
                  : ''}
              </span>
              <span className="text-[11px] text-slate-500">
                Impor template lokasi lama agar dashboard melacak jumlah kode
                yang sudah &amp; belum tersimulasi.
              </span>
            </>
          )}
        </div>
      </div>
    </section>
  );
};

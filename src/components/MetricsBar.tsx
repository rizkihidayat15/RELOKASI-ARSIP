import React from 'react';
import type { SimulationMetrics } from '../types/archive';

interface MetricsBarProps {
  metrics: SimulationMetrics;
  activeLocationCode: string;
  activeRack: number;
  spreadsheetFileName?: string;
}

export const MetricsBar: React.FC<MetricsBarProps> = ({
  metrics,
  activeLocationCode,
  activeRack,
  spreadsheetFileName,
}) => {
  return (
    <section
      aria-label="Ringkasan dan Progres Simulasi"
      className="rounded-xl border border-slate-200 bg-white px-5 py-4"
    >
      <dl className="grid grid-cols-2 gap-4 sm:grid-cols-4 lg:flex lg:flex-wrap lg:items-center lg:gap-0 lg:divide-x lg:divide-slate-200">
        <div className="lg:pr-6">
          <dt className="text-xs text-slate-500">
            Box Terisi ({activeLocationCode || 'RC2'} · Rak {activeRack})
          </dt>
          <dd
            data-testid="metric-filled-boxes"
            className="mt-0.5 font-mono text-xl font-bold text-[#0F4C81] tabular-nums"
          >
            {metrics.rackFilledBoxes}
          </dd>
        </div>
        <div className="lg:px-6">
          <dt className="text-xs text-slate-500">Box Kosong (Rak Aktif)</dt>
          <dd
            data-testid="metric-empty-boxes"
            className="mt-0.5 font-mono text-xl font-bold text-slate-900 tabular-nums"
          >
            {metrics.rackEmptyBoxes}
          </dd>
        </div>
        <div className="lg:px-6">
          <dt className="text-xs text-slate-500">Tersimpan di Excel</dt>
          <dd
            data-testid="metric-saved-excel"
            className="mt-0.5 font-mono text-xl font-bold text-emerald-700 tabular-nums"
          >
            {metrics.savedExcelCount}
          </dd>
        </div>
        <div className="lg:pl-6">
          <dt className="text-xs text-slate-500">Progres Simulasi</dt>
          <dd className="mt-0.5 flex items-baseline gap-2">
            <span
              data-testid="metric-progress-percent"
              className="font-mono text-xl font-bold text-slate-900 tabular-nums"
            >
              {metrics.progressPercent}%
            </span>
            <span className="text-xs text-slate-500 tabular-nums">
              {metrics.progressLabel}
            </span>
          </dd>
        </div>
      </dl>

      {/* Linear Progress Bar + Source Info */}
      <div className="mt-3.5">
        <div
          role="progressbar"
          aria-label="Progres simulasi pemindahan lokasi"
          aria-valuenow={metrics.progressPercent}
          aria-valuemin={0}
          aria-valuemax={100}
          className="h-1.5 w-full overflow-hidden rounded-full bg-slate-100"
        >
          <div
            className="h-full rounded-full bg-[#0F4C81] transition-all duration-200"
            style={{ width: `${metrics.progressPercent}%` }}
          />
        </div>
        <div className="mt-2 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          <span className="tabular-nums">
            Kapasitas 1 rak: {metrics.rackTotalBoxes} box
            {metrics.unsavedDraftCount > 0
              ? ` · ${metrics.unsavedDraftCount} box belum disimpan ke Excel`
              : ''}
            {spreadsheetFileName
              ? ` · Sumber spreadsheet: ${spreadsheetFileName}`
              : ''}
          </span>
          {metrics.spreadsheetRemaining !== null && (
            <span className="tabular-nums">
              Sisa belum dipindahkan: {metrics.spreadsheetRemaining} lokasi
            </span>
          )}
        </div>
      </div>
    </section>
  );
};

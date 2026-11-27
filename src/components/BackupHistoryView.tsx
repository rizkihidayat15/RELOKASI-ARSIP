import React, { useMemo, useState } from 'react';
import {
  Archive,
  Clock,
  Download,
  ExternalLink,
  FileSpreadsheet,
  History,
  RotateCcw,
  Save,
  Search,
  Trash2,
} from 'lucide-react';
import type {
  BackupHistoryCategory,
  BackupHistoryEntry,
  LockedLocationInfo,
} from '../types/archive';
import {
  buildTemplateBackupExportRows,
  filterBackupHistory,
} from '../utils/backupHistory';
import {
  downloadResultBackupExcel,
  downloadTemplateBackupExcel,
} from '../utils/excelService';

interface BackupHistoryViewProps {
  entries: BackupHistoryEntry[];
  lockedMap: Map<string, LockedLocationInfo>;
  canCreateManualSnapshot: boolean;
  onCreateManualSnapshot: () => void;
  onRequestRestoreEntry: (entry: BackupHistoryEntry) => void;
  onRemoveEntry: (entryId: string) => void;
  onRequestClearAllHistory: () => void;
}

export const BackupHistoryView: React.FC<BackupHistoryViewProps> = ({
  entries,
  lockedMap,
  canCreateManualSnapshot,
  onCreateManualSnapshot,
  onRequestRestoreEntry,
  onRemoveEntry,
  onRequestClearAllHistory,
}) => {
  const [categoryFilter, setCategoryFilter] = useState<
    'all' | BackupHistoryCategory
  >('all');
  const [searchQuery, setSearchQuery] = useState('');

  const filteredEntries = useMemo(
    () => filterBackupHistory(entries, categoryFilter, searchQuery),
    [entries, categoryFilter, searchQuery]
  );

  const templateCount = useMemo(
    () => entries.filter((e) => e.category === 'template').length,
    [entries]
  );
  const resultCount = useMemo(
    () => entries.filter((e) => e.category === 'result').length,
    [entries]
  );

  const handleDownloadEntry = (entry: BackupHistoryEntry) => {
    if (entry.category === 'template' && entry.templateSnapshot) {
      const rows = buildTemplateBackupExportRows(
        entry.templateSnapshot,
        lockedMap,
        false
      );
      downloadTemplateBackupExcel(rows, entry.createdAt, 'TEMPLATE');
    } else if (entry.category === 'result' && entry.recordsSnapshot) {
      downloadResultBackupExcel(
        entry.recordsSnapshot,
        Boolean(entry.includeExtraColumns),
        entry.createdAt
      );
    }
  };

  return (
    <section
      aria-label="Riwayat dan Cadangan Update"
      className="rounded-xl border border-slate-200 bg-white"
    >
      {/* Header */}
      <div className="border-b border-slate-200 px-6 py-5">
        <div className="flex flex-wrap items-start justify-between gap-4">
          <div>
            <h2 className="text-base font-bold tracking-tight text-slate-900">
              Riwayat & Cadangan Update
            </h2>
            <p className="mt-0.5 max-w-2xl text-xs leading-relaxed text-slate-500">
              Setiap pembaruan file template lokasi simpan lama maupun hasil
              simpan, generate Excel, dan sinkronisasi Google Sheets otomatis
              dicatat di sini sebagai cadangan yang dapat diunduh atau dipulihkan
              kapan saja.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              disabled={!canCreateManualSnapshot}
              onClick={onCreateManualSnapshot}
              data-testid="btn-create-manual-backup"
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#0F4C81] px-3.5 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap shrink-0"
            >
              <Save className="h-3.5 w-3.5" />
              Buat Cadangan Sekarang
            </button>
            {entries.length > 0 && (
              <button
                type="button"
                onClick={onRequestClearAllHistory}
                className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-white px-3 py-2 text-xs font-semibold text-red-700 transition-colors hover:bg-red-50 whitespace-nowrap shrink-0"
              >
                <Trash2 className="h-3.5 w-3.5" />
                Bersihkan Riwayat
              </button>
            )}
          </div>
        </div>

        {/* Filter Tabs + Search */}
        <div className="mt-4 flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 pt-4">
          <div className="inline-flex items-center gap-1 rounded-lg bg-slate-100 p-1">
            {(
              [
                { id: 'all', label: `Semua (${entries.length})` },
                { id: 'template', label: `Template (${templateCount})` },
                { id: 'result', label: `Hasil Generate (${resultCount})` },
              ] as const
            ).map((opt) => (
              <button
                key={opt.id}
                type="button"
                onClick={() => setCategoryFilter(opt.id)}
                className={`rounded-md px-3 py-1 text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                  categoryFilter === opt.id
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
              aria-label="Cari riwayat cadangan"
              placeholder="Cari nama file, aksi, atau kode..."
              className="w-full rounded-lg border border-slate-200 bg-slate-50 py-1.5 pr-3 pl-8 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:bg-white focus:outline-none"
            />
          </div>
        </div>
      </div>

      {/* History List / Empty State */}
      {entries.length === 0 ? (
        <div className="flex flex-col items-center justify-center px-6 py-14 text-center">
          <div className="flex h-11 w-11 items-center justify-center rounded-lg border border-slate-200 bg-slate-50 text-slate-400">
            <History className="h-5 w-5" />
          </div>
          <p className="mt-3 text-sm font-semibold text-slate-800">
            Belum ada riwayat cadangan tercatat
          </p>
          <p className="mt-1 max-w-md text-xs leading-relaxed text-slate-500">
            Riwayat akan otomatis tersimpan setiap kali Anda mengimpor atau
            memperbarui template lokasi simpan lama, menyimpan hasil simulasi,
            men-generate Excel, atau memperbarui Google Sheets.
          </p>
        </div>
      ) : filteredEntries.length === 0 ? (
        <div className="px-6 py-10 text-center text-xs text-slate-500">
          Tidak ada entri riwayat yang cocok dengan filter atau kata kunci "
          {searchQuery}".
        </div>
      ) : (
        <ul
          data-testid="backup-history-list"
          className="divide-y divide-slate-100"
        >
          {filteredEntries.map((entry) => {
            const isTemplate = entry.category === 'template';
            const canDownload = isTemplate
              ? Boolean(entry.templateSnapshot)
              : Boolean(
                  entry.recordsSnapshot && entry.recordsSnapshot.length > 0
                );
            const canRestore = isTemplate
              ? Boolean(entry.templateSnapshot)
              : Boolean(entry.recordsSnapshot);

            return (
              <li
                key={entry.id}
                data-testid={`history-entry-${entry.category}`}
                className="flex flex-col justify-between gap-3 px-6 py-4 transition-colors hover:bg-slate-50/70 sm:flex-row sm:items-center"
              >
                <div className="flex min-w-0 items-start gap-3">
                  <div
                    className={`mt-0.5 flex h-9 w-9 shrink-0 items-center justify-center rounded-lg border ${
                      isTemplate
                        ? 'border-sky-200 bg-sky-50 text-[#0F4C81]'
                        : 'border-emerald-200 bg-emerald-50 text-emerald-700'
                    }`}
                  >
                    {isTemplate ? (
                      <FileSpreadsheet className="h-4 w-4" />
                    ) : (
                      <Archive className="h-4 w-4" />
                    )}
                  </div>

                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-x-2 gap-y-0.5">
                      <span className="text-sm font-bold text-slate-900">
                        {entry.actionLabel}
                      </span>
                      <span
                        className={`text-[11px] font-semibold ${
                          isTemplate ? 'text-[#0F4C81]' : 'text-emerald-700'
                        }`}
                      >
                        ·{' '}
                        {isTemplate
                          ? 'Template Lokasi Lama'
                          : 'Hasil Generate / Simpan'}
                      </span>
                    </div>

                    <p className="mt-0.5 text-xs text-slate-600 tabular-nums">
                      {entry.summary}
                    </p>

                    <div className="mt-1.5 flex flex-wrap items-center gap-x-4 gap-y-1 text-[11px] text-slate-500">
                      <span className="inline-flex items-center gap-1 tabular-nums">
                        <Clock className="h-3 w-3" />
                        {new Date(entry.createdAt).toLocaleString('id-ID', {
                          day: '2-digit',
                          month: 'short',
                          year: 'numeric',
                          hour: '2-digit',
                          minute: '2-digit',
                        })}
                      </span>
                      {entry.fileName && (
                        <span className="font-mono text-slate-600">
                          {entry.fileName}
                        </span>
                      )}
                      {entry.targetSheetTitle && entry.targetSheetUrl && (
                        <a
                          href={entry.targetSheetUrl}
                          target="_blank"
                          rel="noopener noreferrer"
                          className="inline-flex items-center gap-1 font-medium text-[#0F4C81] hover:underline"
                        >
                          <span>
                            {entry.targetSheetTitle}
                            {entry.targetSheetTab
                              ? ` (${entry.targetSheetTab})`
                              : ''}
                          </span>
                          <ExternalLink className="h-3 w-3" />
                        </a>
                      )}
                    </div>
                  </div>
                </div>

                {/* Per-entry Backup Actions */}
                <div className="flex flex-wrap items-center gap-2 shrink-0">
                  <button
                    type="button"
                    disabled={!canDownload}
                    onClick={() => handleDownloadEntry(entry)}
                    className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-40 whitespace-nowrap"
                  >
                    <Download className="h-3.5 w-3.5" />
                    Unduh Cadangan
                  </button>
                  <button
                    type="button"
                    disabled={!canRestore}
                    onClick={() => onRequestRestoreEntry(entry)}
                    className="inline-flex items-center gap-1.5 rounded-lg border border-[#0F4C81]/30 bg-sky-50/70 px-3 py-1.5 text-xs font-semibold text-[#0F4C81] transition-colors hover:bg-sky-100 disabled:opacity-40 whitespace-nowrap"
                  >
                    <RotateCcw className="h-3.5 w-3.5" />
                    Pulihkan
                  </button>
                  <button
                    type="button"
                    onClick={() => onRemoveEntry(entry.id)}
                    aria-label={`Hapus riwayat ${entry.actionLabel}`}
                    title="Hapus entri riwayat"
                    className="rounded-lg p-1.5 text-slate-400 transition-colors hover:bg-red-50 hover:text-red-600"
                  >
                    <Trash2 className="h-3.5 w-3.5" />
                  </button>
                </div>
              </li>
            );
          })}
        </ul>
      )}
    </section>
  );
};

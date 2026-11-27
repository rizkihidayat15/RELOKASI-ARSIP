import type {
  BackupHistoryCategory,
  BackupHistoryEntry,
  ImportedSpreadsheet,
  LockedLocationInfo,
  PlacementRecord,
} from '../types/archive';

export const BACKUP_HISTORY_STORAGE_KEY = 'arsip_relokasi_backup_history_v1';
export const MAX_BACKUP_HISTORY = 25;

export function createTemplateBackupEntry(params: {
  actionLabel: string;
  spreadsheet: ImportedSpreadsheet;
  lockedCount: number;
  timestamp?: string;
}): BackupHistoryEntry {
  const createdAt = params.timestamp ?? new Date().toISOString();
  const total = params.spreadsheet.items.length;
  const available = Math.max(0, total - params.lockedCount);

  return {
    id: `tpl-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    category: 'template',
    actionLabel: params.actionLabel,
    createdAt,
    fileName: params.spreadsheet.fileName,
    itemCount: total,
    lockedCount: params.lockedCount,
    summary: `${total} kode lokasi (${params.lockedCount} terkunci, ${available} tersedia) · Kolom: ${params.spreadsheet.selectedColumn}`,
    templateSnapshot: {
      ...params.spreadsheet,
      columns: [...params.spreadsheet.columns],
      items: params.spreadsheet.items.map((item) => ({ ...item })),
      // Cap rawRows inside backup snapshot to keep localStorage fast while keeping items 100% intact
      rawRows: params.spreadsheet.rawRows.slice(0, 5000).map((r) => ({ ...r })),
    },
  };
}

export function createResultBackupEntry(params: {
  actionLabel: string;
  records: PlacementRecord[];
  includeExtraColumns: boolean;
  fileName?: string;
  targetSheetTitle?: string;
  targetSheetTab?: string;
  targetSheetUrl?: string;
  summaryOverride?: string;
  timestamp?: string;
}): BackupHistoryEntry {
  const createdAt = params.timestamp ?? new Date().toISOString();
  const ordered = [...params.records]
    .sort((a, b) => a.sequence - b.sequence)
    .map((r) => ({ ...r }));

  return {
    id: `res-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    category: 'result',
    actionLabel: params.actionLabel,
    createdAt,
    fileName: params.fileName,
    itemCount: ordered.length,
    includeExtraColumns: params.includeExtraColumns,
    targetSheetTitle: params.targetSheetTitle,
    targetSheetTab: params.targetSheetTab,
    targetSheetUrl: params.targetSheetUrl,
    summary:
      params.summaryOverride ??
      `${ordered.length} baris hasil relokasi (${params.includeExtraColumns ? '6 kolom' : '2 kolom utama'})`,
    recordsSnapshot: ordered,
  };
}

export function appendBackupHistoryEntry(
  entries: BackupHistoryEntry[],
  nextEntry: BackupHistoryEntry,
  maxEntries = MAX_BACKUP_HISTORY
): BackupHistoryEntry[] {
  return [nextEntry, ...entries].slice(0, Math.max(1, maxEntries));
}

export function removeBackupHistoryEntry(
  entries: BackupHistoryEntry[],
  entryId: string
): BackupHistoryEntry[] {
  return entries.filter((entry) => entry.id !== entryId);
}

export function filterBackupHistory(
  entries: BackupHistoryEntry[],
  category: 'all' | BackupHistoryCategory,
  query: string
): BackupHistoryEntry[] {
  const q = query.trim().toLowerCase();
  return entries.filter((entry) => {
    if (category !== 'all' && entry.category !== category) return false;
    if (!q) return true;
    return (
      entry.actionLabel.toLowerCase().includes(q) ||
      entry.summary.toLowerCase().includes(q) ||
      (entry.fileName?.toLowerCase().includes(q) ?? false) ||
      (entry.targetSheetTitle?.toLowerCase().includes(q) ?? false) ||
      (entry.recordsSnapshot?.some(
        (r) =>
          r.oldLocation.toLowerCase().includes(q) ||
          r.newLocation.toLowerCase().includes(q)
      ) ??
        false)
    );
  });
}

export function loadBackupHistory(): BackupHistoryEntry[] {
  try {
    const raw = window.localStorage.getItem(BACKUP_HISTORY_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed as BackupHistoryEntry[];
  } catch {
    return [];
  }
}

export function saveBackupHistory(entries: BackupHistoryEntry[]): void {
  try {
    window.localStorage.setItem(
      BACKUP_HISTORY_STORAGE_KEY,
      JSON.stringify(entries)
    );
  } catch {
    // If localStorage quota is tight, retry keeping the 8 most recent entries without rawRows
    try {
      const compacted = entries.slice(0, 8).map((entry) =>
        entry.templateSnapshot
          ? {
              ...entry,
              templateSnapshot: {
                ...entry.templateSnapshot,
                rawRows: [],
              },
            }
          : entry
      );
      window.localStorage.setItem(
        BACKUP_HISTORY_STORAGE_KEY,
        JSON.stringify(compacted)
      );
    } catch {
      // Ignore storage quota failure safely
    }
  }
}

export function buildTemplateBackupExportRows(
  spreadsheet: ImportedSpreadsheet,
  lockedMap: Map<string, LockedLocationInfo>,
  onlyAvailable = false
): Record<string, string>[] {
  const rows: Record<string, string>[] = [];
  for (const item of spreadsheet.items) {
    const locked = lockedMap.get(item.normalizedKey);
    if (onlyAvailable && locked) continue;
    rows.push({
      'LOKASI SIMPAN LAMA': item.oldLocation,
      KETERANGAN: item.description ?? '',
      STATUS: locked ? 'TERKUNCI (SUDAH DIGUNAKAN)' : 'TERSEDIA',
      'LOKASI SIMPAN BARU': locked ? locked.newLocation : '',
    });
  }
  return rows;
}

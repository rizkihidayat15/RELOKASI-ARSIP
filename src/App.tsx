import { useEffect, useMemo, useRef, useState } from 'react';
import { Download, RefreshCw } from 'lucide-react';
import { BackupHistoryView } from './components/BackupHistoryView';
import { BoxInputModal } from './components/BoxInputModal';
import { ConfirmDialog } from './components/ConfirmDialog';
import {
  MetricsBar,
  type TemplateFilterPreset,
} from './components/MetricsBar';
import { RackVisualizer } from './components/RackVisualizer';
import {
  ResultsPanel,
  type GoogleSheetsPanelProps,
} from './components/ResultsPanel';
import { SpreadsheetImportView } from './components/SpreadsheetImportView';
import type {
  BackupHistoryEntry,
  BoxCoordinate,
  ImportedSpreadsheet,
  LockedLocationInfo,
  PlacementRecord,
  RackConfig,
  SimulationState,
} from './types/archive';
import {
  appendBackupHistoryEntry,
  createResultBackupEntry,
  createTemplateBackupEntry,
  loadBackupHistory,
  removeBackupHistoryEntry,
  saveBackupHistory,
} from './utils/backupHistory';
import {
  downloadExcelArtifact,
  generateRelocationExcelArtifact,
  parseResultSpreadsheetFile,
  type GeneratedExcelArtifact,
} from './utils/excelService';
import {
  clearWorkspaceAccessToken,
  getWorkspaceAccessToken,
  initGoogleAuth,
  logoutGoogleWorkspace,
  signInWithGoogleWorkspace,
} from './utils/googleAuth';
import {
  buildSafeSpreadsheetUrl,
  computePendingGoogleSheetSyncCount,
  createRelocationSpreadsheet,
  extractSpreadsheetGid,
  extractSpreadsheetId,
  fetchSpreadsheetMetadata,
  listDriveSpreadsheets,
  readSheetRows,
  syncRecordsToGoogleSheet,
  WorkspaceAuthExpiredError,
  type ConnectedGoogleSheet,
  type DriveSpreadsheetItem,
} from './utils/googleSheetsSync';
import {
  applyPlacement,
  calculateSimulationMetrics,
  clearSavedExcelRecords,
  commitDraftToExcel,
  createInitialSimulationState,
  extractSpreadsheetItems,
  getCombinedResultRecords,
  getEffectiveRackPlacements,
  getLockedLocationsMap,
  parseExternalSheetLocks,
  removePlacementByCoordinate,
  removeSavedRecordById,
  resetActiveRackSimulation,
  sanitizeLocationCode,
  undoLastAction,
  updateRackConfig,
} from './utils/relocationEngine';
import { addTemplateItem } from './utils/templateEditor';

type ActiveTab = 'rack' | 'results' | 'import' | 'history';

const STORAGE_KEY = 'arsip_relokasi_simulasi_v2';
const CONNECTED_SHEET_STORAGE_KEY = 'arsip_relokasi_google_sheet_v1';

function loadConnectedGoogleSheet(): ConnectedGoogleSheet | null {
  try {
    const raw = window.localStorage.getItem(CONNECTED_SHEET_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as Partial<ConnectedGoogleSheet>;
    if (!parsed.spreadsheetId) return null;
    const validId = extractSpreadsheetId(parsed.spreadsheetId);
    if (!validId) return null;
    const preferredGid =
      typeof parsed.preferredGid === 'number' &&
      Number.isFinite(parsed.preferredGid)
        ? parsed.preferredGid
        : undefined;
    return {
      spreadsheetId: validId,
      spreadsheetUrl: buildSafeSpreadsheetUrl(validId, preferredGid),
      title: parsed.title || 'Google Spreadsheet Terhubung',
      sheetName: parsed.sheetName || 'Sheet1',
      availableSheets:
        Array.isArray(parsed.availableSheets) &&
        parsed.availableSheets.length > 0
          ? parsed.availableSheets
          : [parsed.sheetName || 'Sheet1'],
      preferredGid,
      lastSyncedAt: parsed.lastSyncedAt,
      lastSyncSummary: parsed.lastSyncSummary,
    };
  } catch {
    return null;
  }
}

function normalizeStoredRecords(
  records: unknown,
  defaultLoc: string
): PlacementRecord[] {
  if (!Array.isArray(records)) return [];
  return records.map((p: PlacementRecord) => ({
    ...p,
    locationCode: p.locationCode || p.newLocation?.split('.')[0] || defaultLoc,
  }));
}

function loadSavedState(): SimulationState {
  try {
    const raw = window.localStorage.getItem(STORAGE_KEY);
    if (!raw) return createInitialSimulationState();
    const parsed = JSON.parse(raw) as Partial<SimulationState>;
    const base = createInitialSimulationState(parsed.config);
    const defaultLoc = base.config.prefix || 'RC2';
    return {
      ...base,
      draftPlacements: normalizeStoredRecords(
        parsed.draftPlacements,
        defaultLoc
      ),
      savedRecords: normalizeStoredRecords(parsed.savedRecords, defaultLoc),
      nextSequence:
        typeof parsed.nextSequence === 'number' && parsed.nextSequence > 0
          ? parsed.nextSequence
          : 1,
      spreadsheet: parsed.spreadsheet ?? null,
      externalLockedLocations: Array.isArray(parsed.externalLockedLocations)
        ? (parsed.externalLockedLocations as LockedLocationInfo[])
        : [],
      externalResultSourceName:
        typeof parsed.externalResultSourceName === 'string'
          ? parsed.externalResultSourceName
          : null,
      externalResultSourceType:
        parsed.externalResultSourceType === 'google_sheets' ||
        parsed.externalResultSourceType === 'local_file'
          ? parsed.externalResultSourceType
          : null,
      allowUnlistedWithSpreadsheet: Boolean(parsed.allowUnlistedWithSpreadsheet),
      includeExtraExcelColumns: Boolean(parsed.includeExtraExcelColumns),
      history: [],
    };
  } catch {
    return createInitialSimulationState();
  }
}

export default function App() {
  const [state, setState] = useState<SimulationState>(() => loadSavedState());
  const [activeTab, setActiveTab] = useState<ActiveTab>('rack');
  const [activeLocationCode, setActiveLocationCode] = useState<string>(
    () => state.config.prefix || 'RC2'
  );
  const [activeRack, setActiveRack] = useState<number>(1);
  const [startBoxNumber, setStartBoxNumber] = useState<number>(1);
  const [continuousBoxNumbering, setContinuousBoxNumbering] =
    useState<boolean>(true);
  const [selectedCoordinate, setSelectedCoordinate] =
    useState<BoxCoordinate | null>(null);
  const [highlightedCoordinate, setHighlightedCoordinate] =
    useState<BoxCoordinate | null>(null);
  const [showClearExcelDialog, setShowClearExcelDialog] = useState(false);
  const [generatedArtifact, setGeneratedArtifact] =
    useState<GeneratedExcelArtifact | null>(null);
  const [artifactSignature, setArtifactSignature] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);

  // Google Workspace Auth & Same-Link Google Sheets Sync State
  const [isGoogleAuthenticated, setIsGoogleAuthenticated] = useState(false);
  const [googleUserEmail, setGoogleUserEmail] = useState<string | null>(null);
  const [isSigningInGoogle, setIsSigningInGoogle] = useState(false);
  const [connectedSheet, setConnectedSheet] =
    useState<ConnectedGoogleSheet | null>(() => loadConnectedGoogleSheet());
  const [driveFiles, setDriveFiles] = useState<DriveSpreadsheetItem[]>([]);
  const [isLoadingDriveFiles, setIsLoadingDriveFiles] = useState(false);
  const [isConnectingSheet, setIsConnectingSheet] = useState(false);
  const [isSyncingSheet, setIsSyncingSheet] = useState(false);
  const [isRefreshingDashboard, setIsRefreshingDashboard] = useState(false);
  const [sheetSyncError, setSheetSyncError] = useState<string | null>(null);
  const [showSyncConfirmDialog, setShowSyncConfirmDialog] = useState(false);
  const [templateFilterPreset, setTemplateFilterPreset] =
    useState<TemplateFilterPreset>('all');

  // Backup History State (Template & Generated Result snapshots)
  const [backupHistory, setBackupHistory] = useState<BackupHistoryEntry[]>(() =>
    loadBackupHistory()
  );
  const [pendingRestoreEntry, setPendingRestoreEntry] =
    useState<BackupHistoryEntry | null>(null);
  const [showClearHistoryDialog, setShowClearHistoryDialog] = useState(false);

  const metrics = useMemo(
    () => calculateSimulationMetrics(state, activeLocationCode, activeRack),
    [state, activeLocationCode, activeRack]
  );

  const lockedMap = useMemo(() => getLockedLocationsMap(state), [state]);

  const combinedResultRecords = useMemo(
    () => getCombinedResultRecords(state),
    [state]
  );

  const effectiveRackPlacements = useMemo(
    () => getEffectiveRackPlacements(state),
    [state]
  );

  const pendingSheetSync = useMemo(
    () =>
      computePendingGoogleSheetSyncCount(
        combinedResultRecords,
        state.externalLockedLocations,
        Boolean(connectedSheet)
      ),
    [combinedResultRecords, state.externalLockedLocations, connectedSheet]
  );

  const isSyncingSheetRef = useRef(false);
  const isRefreshingDashboardRef = useRef(false);
  isSyncingSheetRef.current = isSyncingSheet;
  isRefreshingDashboardRef.current = isRefreshingDashboard;

  const currentDataSignature = useMemo(() => {
    const rowsSig = combinedResultRecords
      .map((p) => `${p.sequence}:${p.oldLocation}:${p.newLocation}`)
      .join('|');
    return `${state.includeExtraExcelColumns ? 'extra' : 'basic'}::${rowsSig}`;
  }, [combinedResultRecords, state.includeExtraExcelColumns]);

  const isArtifactStale =
    generatedArtifact !== null && artifactSignature !== currentDataSignature;

  useEffect(() => {
    try {
      const serializable = {
        config: state.config,
        draftPlacements: state.draftPlacements,
        savedRecords: state.savedRecords,
        nextSequence: state.nextSequence,
        spreadsheet: state.spreadsheet,
        externalLockedLocations: state.externalLockedLocations,
        externalResultSourceName: state.externalResultSourceName ?? null,
        externalResultSourceType: state.externalResultSourceType ?? null,
        allowUnlistedWithSpreadsheet: state.allowUnlistedWithSpreadsheet,
        includeExtraExcelColumns: state.includeExtraExcelColumns,
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(serializable));
    } catch {
      // Ignore quota errors cleanly
    }
  }, [state]);

  useEffect(() => {
    saveBackupHistory(backupHistory);
  }, [backupHistory]);

  useEffect(() => {
    if (!notice) return;
    const timer = setTimeout(() => setNotice(null), 4500);
    return () => clearTimeout(timer);
  }, [notice]);

  useEffect(() => {
    if (!highlightedCoordinate) return;
    const timer = setTimeout(() => setHighlightedCoordinate(null), 2600);
    return () => clearTimeout(timer);
  }, [highlightedCoordinate]);

  useEffect(() => {
    const unsubscribe = initGoogleAuth(
      (user) => {
        setIsGoogleAuthenticated(true);
        setGoogleUserEmail(user.email);
      },
      (user) => {
        setIsGoogleAuthenticated(false);
        setGoogleUserEmail(user?.email ?? null);
      }
    );
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    try {
      if (connectedSheet) {
        window.localStorage.setItem(
          CONNECTED_SHEET_STORAGE_KEY,
          JSON.stringify(connectedSheet)
        );
      } else {
        window.localStorage.removeItem(CONNECTED_SHEET_STORAGE_KEY);
      }
    } catch {
      // Ignore quota errors cleanly
    }
  }, [connectedSheet]);

  const handleSubmitPlacement = (
    coord: BoxCoordinate,
    oldLocation: string,
    confirmReplace: boolean
  ) => {
    setState((prev) => {
      const { nextState, result } = applyPlacement(prev, coord, oldLocation, {
        confirmReplace,
      });
      if (result.status === 'placed') {
        setHighlightedCoordinate(coord);
        setNotice(
          `Lokasi lama ${result.record.oldLocation} ditempatkan pada ${result.record.newLocation}. Klik Simpan untuk memasukkan ke Excel.`
        );
      }
      return nextState;
    });
  };

  const handleSaveToExcel = () => {
    const { nextState, addedCount } = commitDraftToExcel(state);
    if (addedCount === 0) return;

    setState(nextState);
    const nextCombined = getCombinedResultRecords(nextState);
    const snapshot = createResultBackupEntry({
      actionLabel: `Simpan Siklus Rak ${activeRack}`,
      records: nextCombined,
      includeExtraColumns: nextState.includeExtraExcelColumns,
    });
    setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));

    if (isGoogleAuthenticated && connectedSheet) {
      setNotice(
        `${addedCount} lokasi disimpan & dicadangkan. Konfirmasi untuk langsung update Google Sheets "${connectedSheet.title}".`
      );
      setShowSyncConfirmDialog(true);
    } else if (connectedSheet) {
      setNotice(
        `${addedCount} lokasi disimpan di aplikasi (${nextCombined.length} baris). Sesi Google belum aktif — klik "Sign in with Google" agar perubahan terkirim ke "${connectedSheet.title}".`
      );
    } else {
      setNotice(
        `${addedCount} lokasi dari simulasi rak berhasil disimpan & dicadangkan (total ${nextCombined.length} baris terdaftar).`
      );
    }
  };

  const handleResetActiveRack = () => {
    setState((prev) => resetActiveRackSimulation(prev));
    setNotice(
      `Draft simulasi rak aktif direset. ${combinedResultRecords.length} data yang sudah terdaftar di Tabel Hasil tetap tersimpan.`
    );
  };

  const handleRemoveBoxPlacement = (coord: BoxCoordinate) => {
    setState((prev) => removePlacementByCoordinate(prev, coord));
    setNotice(
      `Isi box ${coord.locationCode} · RAK ${coord.rack} · ROW ${coord.row} · BOX ${coord.box} dihapus dari simulasi & Tabel Hasil.`
    );
  };

  const handleRemoveSavedRecord = (placementId: string) => {
    const target = combinedResultRecords.find((p) => p.id === placementId);
    setState((prev) => removeSavedRecordById(prev, placementId));
    if (target) {
      setNotice(
        `Baris ${target.oldLocation} (${target.newLocation}) dihapus dari Tabel Hasil.`
      );
    }
  };

  const handleConfirmClearSavedExcel = () => {
    setState((prev) => clearSavedExcelRecords(prev));
    setShowClearExcelDialog(false);
    setGeneratedArtifact(null);
    setArtifactSignature(null);
    setNotice(
      'Seluruh data hasil simpan Excel telah dikosongkan. Gunakan Undo jika ingin mengembalikan.'
    );
  };

  const handleUndo = () => {
    setState((prev) => undoLastAction(prev));
    setNotice('Tindakan terakhir dibatalkan (Undo).');
  };

  const handleUpdateConfig = (nextConfig: RackConfig) => {
    setState((prev) => updateRackConfig(prev, nextConfig));
    setNotice(
      `Ukuran 1 rak diperbarui: ${nextConfig.rowsPerRack} row × ${nextConfig.boxesPerRow} box.`
    );
  };

  const handleChangeLocationCode = (value: string) => {
    setActiveLocationCode(value);
    const cleaned = value.trim().replace(/\s+/g, '').toUpperCase();
    if (cleaned) {
      setState((prev) => ({
        ...prev,
        config: {
          ...prev.config,
          prefix: cleaned,
        },
      }));
    }
  };

  const handleImportSpreadsheet = (sheet: ImportedSpreadsheet) => {
    const lockedCount = sheet.items.filter((item) =>
      lockedMap.has(item.normalizedKey)
    ).length;

    setState((prev) => ({
      ...prev,
      spreadsheet: sheet,
    }));

    const snapshot = createTemplateBackupEntry({
      actionLabel: 'Impor Template Lokasi Lama',
      spreadsheet: sheet,
      lockedCount,
    });
    setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));

    setNotice(
      lockedCount > 0
        ? `Berhasil mengimpor ${sheet.items.length} lokasi dari ${sheet.fileName} (${lockedCount} kode sudah digunakan otomatis dikunci) & dicadangkan ke Riwayat.`
        : `Berhasil mengimpor ${sheet.items.length} lokasi simpan lama dari ${sheet.fileName} & dicadangkan ke Riwayat.`
    );
  };

  const handleChangeSpreadsheetColumn = (nextColumn: string) => {
    if (!state.spreadsheet) return;
    const descCol = state.spreadsheet.columns.find(
      (col) =>
        col !== nextColumn &&
        /keterangan|uraian|judul|nama|arsip|unit|deskripsi/i.test(col)
    );
    const nextItems = extractSpreadsheetItems(
      state.spreadsheet.rawRows,
      nextColumn,
      descCol
    );
    const updatedSheet: ImportedSpreadsheet = {
      ...state.spreadsheet,
      selectedColumn: nextColumn,
      items: nextItems,
    };
    const lockedCount = nextItems.filter((item) =>
      lockedMap.has(item.normalizedKey)
    ).length;

    setState((prev) => ({
      ...prev,
      spreadsheet: updatedSheet,
    }));

    const snapshot = createTemplateBackupEntry({
      actionLabel: 'Ubah Kolom Template',
      spreadsheet: updatedSheet,
      lockedCount,
    });
    setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));
    setNotice(
      `Kolom template diperbarui ke "${nextColumn}" (${nextItems.length} kode) & dicadangkan ke Riwayat.`
    );
  };

  const handleUpdateTemplateSpreadsheet = (
    nextSheet: ImportedSpreadsheet,
    actionLabel: string,
    feedbackMessage: string
  ) => {
    const lockedCount = nextSheet.items.filter((item) =>
      lockedMap.has(item.normalizedKey)
    ).length;

    setState((prev) => ({
      ...prev,
      spreadsheet: nextSheet,
    }));

    const snapshot = createTemplateBackupEntry({
      actionLabel,
      spreadsheet: nextSheet,
      lockedCount,
    });
    setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));
    setNotice(feedbackMessage);
  };

  const handleQuickAddToTemplate = (oldLocation: string) => {
    const result = addTemplateItem(state.spreadsheet, oldLocation);
    if (result.error || !result.addedItem) {
      setNotice(result.error ?? 'Gagal menambahkan kode ke template.');
      return;
    }
    handleUpdateTemplateSpreadsheet(
      result.nextSpreadsheet,
      `Tambah Cepat dari Box (${result.addedItem.oldLocation})`,
      `Kode "${result.addedItem.oldLocation}" ditambahkan ke template & langsung tersedia pada kata kunci pencarian.`
    );
  };

  const handleClearSpreadsheet = () => {
    setState((prev) => ({
      ...prev,
      spreadsheet: null,
    }));
    setNotice(
      'Sumber data spreadsheet dihapus. Aplikasi kembali ke mode input langsung.'
    );
  };

  const handleJumpToBox = (coord: BoxCoordinate) => {
    setActiveTab('rack');
    setActiveLocationCode(coord.locationCode);
    setActiveRack(coord.rack);
    const minVisible = startBoxNumber;
    const maxVisible = continuousBoxNumbering
      ? startBoxNumber +
        state.config.rowsPerRack * state.config.boxesPerRow -
        1
      : startBoxNumber + state.config.boxesPerRow - 1;
    if (coord.box < minVisible || coord.box > maxVisible) {
      const nextStart = continuousBoxNumbering
        ? Math.max(1, coord.box - (coord.row - 1) * state.config.boxesPerRow)
        : Math.max(1, coord.box);
      setStartBoxNumber(nextStart);
    }
    setHighlightedCoordinate(coord);
    setSelectedCoordinate(coord);
  };

  const handleGenerateExcel = () => {
    if (combinedResultRecords.length === 0) return;
    const artifact = generateRelocationExcelArtifact(
      combinedResultRecords,
      state.includeExtraExcelColumns
    );
    setGeneratedArtifact(artifact);
    setArtifactSignature(currentDataSignature);

    const snapshot = createResultBackupEntry({
      actionLabel: 'Generate Excel (.xlsx)',
      records: combinedResultRecords,
      includeExtraColumns: state.includeExtraExcelColumns,
      fileName: artifact.fileName,
    });
    setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));

    setNotice(
      `File ${artifact.fileName} berhasil dibuat (${artifact.rowCount} baris) & dicadangkan ke Riwayat.`
    );
  };

  const handleDownloadExcel = () => {
    if (combinedResultRecords.length === 0) return;
    let artifact = generatedArtifact;
    if (!artifact || isArtifactStale) {
      artifact = generateRelocationExcelArtifact(
        combinedResultRecords,
        state.includeExtraExcelColumns
      );
      setGeneratedArtifact(artifact);
      setArtifactSignature(currentDataSignature);

      const snapshot = createResultBackupEntry({
        actionLabel: 'Download Excel (.xlsx)',
        records: combinedResultRecords,
        includeExtraColumns: state.includeExtraExcelColumns,
        fileName: artifact.fileName,
      });
      setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));
    }
    downloadExcelArtifact(artifact);
    setNotice(`Mengunduh ${artifact.fileName}...`);
  };

  const handleCreateManualBackupSnapshot = () => {
    const newEntries: BackupHistoryEntry[] = [];
    if (state.spreadsheet) {
      const lockedCount = state.spreadsheet.items.filter((item) =>
        lockedMap.has(item.normalizedKey)
      ).length;
      newEntries.push(
        createTemplateBackupEntry({
          actionLabel: 'Cadangan Manual Template',
          spreadsheet: state.spreadsheet,
          lockedCount,
        })
      );
    }
    if (combinedResultRecords.length > 0) {
      newEntries.push(
        createResultBackupEntry({
          actionLabel: 'Cadangan Manual Hasil Simulasi',
          records: combinedResultRecords,
          includeExtraColumns: state.includeExtraExcelColumns,
          fileName: 'HASIL_LOKASI_SIMPAN_LAMA_DAN_BARU.xlsx',
        })
      );
    }
    if (newEntries.length === 0) return;
    setBackupHistory((prev) =>
      newEntries.reduce(
        (acc, entry) => appendBackupHistoryEntry(acc, entry),
        prev
      )
    );
    setNotice(
      `${newEntries.length} cadangan manual berhasil disimpan ke Riwayat.`
    );
  };

  const handleConfirmRestoreBackupEntry = () => {
    if (!pendingRestoreEntry) return;

    if (
      pendingRestoreEntry.category === 'template' &&
      pendingRestoreEntry.templateSnapshot
    ) {
      const restoredSheet = pendingRestoreEntry.templateSnapshot;
      setState((prev) => ({
        ...prev,
        spreadsheet: {
          ...restoredSheet,
          columns: [...restoredSheet.columns],
          items: restoredSheet.items.map((item) => ({ ...item })),
          rawRows: restoredSheet.rawRows.map((row) => ({ ...row })),
        },
      }));
      setNotice(
        `Template "${restoredSheet.fileName}" (${restoredSheet.items.length} lokasi) berhasil dipulihkan dari cadangan.`
      );
    } else if (
      pendingRestoreEntry.category === 'result' &&
      pendingRestoreEntry.recordsSnapshot
    ) {
      const restoredRecords = pendingRestoreEntry.recordsSnapshot.map((r) => ({
        ...r,
      }));
      const maxSeq = restoredRecords.reduce(
        (max, r) => Math.max(max, r.sequence),
        0
      );
      setState((prev) => ({
        ...prev,
        savedRecords: restoredRecords,
        nextSequence: Math.max(prev.nextSequence, maxSeq + 1),
        includeExtraExcelColumns:
          pendingRestoreEntry.includeExtraColumns ??
          prev.includeExtraExcelColumns,
        history: [
          ...prev.history,
          {
            draftPlacements: prev.draftPlacements.map((p) => ({ ...p })),
            savedRecords: prev.savedRecords.map((p) => ({ ...p })),
            nextSequence: prev.nextSequence,
          },
        ].slice(-40),
      }));
      setGeneratedArtifact(null);
      setArtifactSignature(null);
      setNotice(
        `${restoredRecords.length} baris hasil simulasi berhasil dipulihkan dari cadangan (dapat dibatalkan via Undo).`
      );
    }
    setPendingRestoreEntry(null);
  };

  const handleRemoveBackupEntry = (entryId: string) => {
    setBackupHistory((prev) => removeBackupHistoryEntry(prev, entryId));
    setNotice('Satu entri riwayat cadangan dihapus.');
  };

  const handleConfirmClearAllHistory = () => {
    setBackupHistory([]);
    setShowClearHistoryDialog(false);
    setNotice('Seluruh riwayat cadangan telah dibersihkan.');
  };

  const handleWorkspaceError = (err: unknown, fallbackMessage: string) => {
    if (err instanceof WorkspaceAuthExpiredError) {
      clearWorkspaceAccessToken();
      setIsGoogleAuthenticated(false);
      setSheetSyncError(err.message);
      setNotice(err.message);
      return;
    }
    const msg = err instanceof Error ? err.message : fallbackMessage;
    setSheetSyncError(msg);
    setNotice(msg);
  };

  const refreshExternalLocksFromSheet = async (
    sheet: ConnectedGoogleSheet,
    token: string,
    options?: { silent?: boolean }
  ): Promise<number | null> => {
    try {
      const rows = await readSheetRows(
        sheet.spreadsheetId,
        sheet.sheetName,
        token
      );
      const parsedLocks = parseExternalSheetLocks(rows, 'google_sheets');
      setState((prev) => ({
        ...prev,
        externalLockedLocations: parsedLocks,
        externalResultSourceName: sheet.title,
        externalResultSourceType: 'google_sheets',
      }));
      return parsedLocks.length;
    } catch (err) {
      if (err instanceof WorkspaceAuthExpiredError || !options?.silent) {
        handleWorkspaceError(
          err,
          'Gagal membaca baris dari Google Spreadsheet tujuan.'
        );
      }
      return null;
    }
  };

  // Automatic read-only pull from Google Sheets when authenticated (on focus, tab return, and every 30s)
  useEffect(() => {
    if (!isGoogleAuthenticated || !connectedSheet) return;
    let isPulling = false;
    let isCancelled = false;

    const pullLatestFromSheet = async () => {
      if (
        isPulling ||
        isCancelled ||
        isSyncingSheetRef.current ||
        isRefreshingDashboardRef.current
      ) {
        return;
      }
      const token = getWorkspaceAccessToken();
      if (!token) return;
      isPulling = true;
      try {
        const rows = await readSheetRows(
          connectedSheet.spreadsheetId,
          connectedSheet.sheetName,
          token
        );
        if (isCancelled) return;
        const parsedLocks = parseExternalSheetLocks(rows, 'google_sheets');
        setState((prev) => ({
          ...prev,
          externalLockedLocations: parsedLocks,
          externalResultSourceName: connectedSheet.title,
          externalResultSourceType: 'google_sheets',
        }));
      } catch (err) {
        if (!isCancelled && err instanceof WorkspaceAuthExpiredError) {
          clearWorkspaceAccessToken();
          setIsGoogleAuthenticated(false);
          setSheetSyncError(err.message);
        }
      } finally {
        isPulling = false;
      }
    };

    const handleFocus = () => {
      void pullLatestFromSheet();
    };
    const handleVisibilityChange = () => {
      if (document.visibilityState === 'visible') {
        void pullLatestFromSheet();
      }
    };

    window.addEventListener('focus', handleFocus);
    document.addEventListener('visibilitychange', handleVisibilityChange);
    const intervalId = window.setInterval(() => {
      if (document.visibilityState === 'visible') {
        void pullLatestFromSheet();
      }
    }, 30_000);

    return () => {
      isCancelled = true;
      window.removeEventListener('focus', handleFocus);
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      window.clearInterval(intervalId);
    };
  }, [
    isGoogleAuthenticated,
    connectedSheet?.spreadsheetId,
    connectedSheet?.sheetName,
    connectedSheet?.title,
  ]);

  const handleRefreshDashboardFromGoogleSheet = async () => {
    const token = getWorkspaceAccessToken();
    if (!token || !connectedSheet) {
      setIsGoogleAuthenticated(false);
      setNotice(
        'Silakan masuk dengan akun Google terlebih dahulu untuk menyinkronkan Dashboard dengan Google Sheets.'
      );
      return;
    }
    setIsRefreshingDashboard(true);
    setSheetSyncError(null);
    try {
      const rows = await readSheetRows(
        connectedSheet.spreadsheetId,
        connectedSheet.sheetName,
        token
      );
      const parsedLocks = parseExternalSheetLocks(rows, 'google_sheets');
      const nextState: SimulationState = {
        ...state,
        externalLockedLocations: parsedLocks,
        externalResultSourceName: connectedSheet.title,
        externalResultSourceType: 'google_sheets',
      };
      setState(nextState);
      const refreshedMetrics = calculateSimulationMetrics(
        nextState,
        activeLocationCode,
        activeRack
      );
      if (refreshedMetrics.templateTotal !== null) {
        setNotice(
          `Dashboard disinkronkan dengan "${connectedSheet.title}": ${refreshedMetrics.templateRegisteredCount} dari ${refreshedMetrics.templateTotal} kode template sudah terdaftar, ${refreshedMetrics.templateRemainingCount} belum tersimulasi.`
        );
      } else {
        setNotice(
          `Dashboard membaca ${parsedLocks.length} baris hasil relokasi dari "${connectedSheet.title}" (${connectedSheet.sheetName}).`
        );
      }
    } catch (err) {
      handleWorkspaceError(
        err,
        'Gagal membaca data terbaru dari Google Spreadsheet tujuan.'
      );
    } finally {
      setIsRefreshingDashboard(false);
    }
  };

  const handleImportExternalResultsFile = async (file: File) => {
    try {
      const parsed = await parseResultSpreadsheetFile(file);
      const nextState: SimulationState = {
        ...state,
        externalLockedLocations: parsed.locks,
        externalResultSourceName: parsed.fileName,
        externalResultSourceType: 'local_file',
      };
      setState(nextState);
      const refreshedMetrics = calculateSimulationMetrics(
        nextState,
        activeLocationCode,
        activeRack
      );
      if (refreshedMetrics.templateTotal !== null) {
        setNotice(
          `File hasil "${parsed.fileName}" dibaca (${parsed.locks.length} baris): ${refreshedMetrics.templateRegisteredCount} dari ${refreshedMetrics.templateTotal} kode template sudah terdaftar & terkunci, ${refreshedMetrics.templateRemainingCount} belum tersimulasi.`
        );
      } else {
        setNotice(
          `File hasil "${parsed.fileName}" dibaca (${parsed.locks.length} kode lokasi lama terkunci). Impor template lokasi lama untuk melacak sisa kode belum tersimulasi.`
        );
      }
    } catch (err) {
      setNotice(
        err instanceof Error
          ? err.message
          : 'Gagal membaca file hasil relokasi (.xlsx).'
      );
    }
  };

  const handleClearExternalResults = () => {
    setState((prev) => ({
      ...prev,
      externalLockedLocations: [],
      externalResultSourceName: null,
      externalResultSourceType: null,
    }));
    setNotice('Referensi file hasil relokasi eksternal dilepas dari Dashboard.');
  };

  const handleOpenTemplateWithFilter = (filter: TemplateFilterPreset) => {
    setTemplateFilterPreset(filter);
    setActiveTab('import');
  };

  const handleGoogleSignIn = async () => {
    setIsSigningInGoogle(true);
    setSheetSyncError(null);
    try {
      const { user, accessToken } = await signInWithGoogleWorkspace();
      setIsGoogleAuthenticated(true);
      setGoogleUserEmail(user.email);
      setNotice(`Terhubung dengan akun Google ${user.email ?? ''}.`);

      // Refresh metadata & read existing rows to lock already-used old codes
      if (connectedSheet) {
        try {
          const refreshed = await fetchSpreadsheetMetadata(
            connectedSheet.spreadsheetId,
            accessToken,
            connectedSheet.sheetName,
            connectedSheet.preferredGid
          );
          setConnectedSheet((prev) => ({
            ...refreshed,
            lastSyncedAt: prev?.lastSyncedAt,
            lastSyncSummary: prev?.lastSyncSummary,
          }));
          const lockedCount = await refreshExternalLocksFromSheet(
            refreshed,
            accessToken
          );
          if (lockedCount !== null) {
            if (lockedCount > 0 && state.savedRecords.length > 0) {
              setNotice(
                `Terhubung ke "${refreshed.title}" (tab "${refreshed.sheetName}"). ${lockedCount} baris dibaca dari Sheets — klik "Update Google Sheets" jika ada perubahan lokal yang ingin dikirim.`
              );
            } else if (lockedCount > 0) {
              setNotice(
                `Terhubung ke "${refreshed.title}" (tab "${refreshed.sheetName}"). ${lockedCount} baris dari Google Sheets otomatis disinkronkan.`
              );
            } else if (state.savedRecords.length > 0) {
              setNotice(
                `Terhubung ke "${refreshed.title}" (tab "${refreshed.sheetName}"). Klik "Update Google Sheets" untuk mengirim ${state.savedRecords.length} baris tersimpan ke link ini.`
              );
            }
          }
        } catch (metaErr) {
          handleWorkspaceError(
            metaErr,
            'Akun Google terhubung, tetapi gagal membaca Google Spreadsheet tujuan. Pastikan link valid dan Anda memiliki akses.'
          );
        }
      }
    } catch (err) {
      handleWorkspaceError(
        err,
        'Gagal masuk dengan akun Google. Silakan coba lagi.'
      );
    } finally {
      setIsSigningInGoogle(false);
    }
  };

  const handleGoogleSignOut = async () => {
    await logoutGoogleWorkspace();
    setIsGoogleAuthenticated(false);
    setGoogleUserEmail(null);
    setDriveFiles([]);
    setSheetSyncError(null);
    setNotice('Anda telah keluar dari akun Google.');
  };

  const handleLoadDriveFiles = async () => {
    const token = getWorkspaceAccessToken();
    if (!token) {
      setIsGoogleAuthenticated(false);
      return;
    }
    setIsLoadingDriveFiles(true);
    setSheetSyncError(null);
    try {
      const files = await listDriveSpreadsheets(token);
      setDriveFiles(files);
    } catch (err) {
      handleWorkspaceError(
        err,
        'Gagal memuat daftar Google Spreadsheet dari Google Drive.'
      );
    } finally {
      setIsLoadingDriveFiles(false);
    }
  };

  const handleConnectSheetByUrl = async (urlOrId: string): Promise<boolean> => {
    const spreadsheetId = extractSpreadsheetId(urlOrId);
    if (!spreadsheetId) return false;
    const preferredGid = extractSpreadsheetGid(urlOrId);

    setSheetSyncError(null);
    const token = getWorkspaceAccessToken();

    if (!token) {
      setConnectedSheet({
        spreadsheetId,
        spreadsheetUrl: buildSafeSpreadsheetUrl(
          spreadsheetId,
          preferredGid ?? undefined
        ),
        title: 'Google Spreadsheet (Menunggu Sign-in)',
        sheetName: 'Sheet1',
        availableSheets: ['Sheet1'],
        preferredGid: preferredGid ?? undefined,
      });
      setNotice(
        'Link Google Spreadsheet tersimpan. Silakan klik "Sign in with Google" untuk mulai sinkronisasi.'
      );
      return true;
    }

    setIsConnectingSheet(true);
    try {
      const metadata = await fetchSpreadsheetMetadata(
        spreadsheetId,
        token,
        undefined,
        preferredGid
      );
      setConnectedSheet(metadata);
      const extLocked = await refreshExternalLocksFromSheet(metadata, token);
      if (extLocked !== null) {
        setNotice(
          extLocked > 0
            ? `Terhubung ke Google Sheets "${metadata.title}" (tab "${metadata.sheetName}") · ${extLocked} baris dari Sheets otomatis disinkronkan.`
            : `Terhubung ke Google Sheets "${metadata.title}" (tab: "${metadata.sheetName}").`
        );
      }
      return true;
    } catch (err) {
      handleWorkspaceError(
        err,
        'Gagal membaca Google Spreadsheet. Pastikan link benar dan Anda memiliki akses edit.'
      );
      return false;
    } finally {
      setIsConnectingSheet(false);
    }
  };

  const handleCreateNewGoogleSheet = async () => {
    const token = getWorkspaceAccessToken();
    if (!token) {
      setIsGoogleAuthenticated(false);
      return;
    }
    setIsConnectingSheet(true);
    setSheetSyncError(null);
    try {
      const created = await createRelocationSpreadsheet(token);
      setConnectedSheet(created);
      setState((prev) => ({
        ...prev,
        externalLockedLocations: [],
        externalResultSourceName: created.title,
        externalResultSourceType: 'google_sheets',
      }));
      setNotice(
        `Spreadsheet baru "${created.title}" dibuat di Google Drive dan siap menerima update baris.`
      );
    } catch (err) {
      handleWorkspaceError(
        err,
        'Gagal membuat Google Spreadsheet baru di Google Drive.'
      );
    } finally {
      setIsConnectingSheet(false);
    }
  };

  const handleSelectSheetTab = (sheetName: string) => {
    setConnectedSheet((prev) => (prev ? { ...prev, sheetName } : prev));
    const token = getWorkspaceAccessToken();
    if (token && connectedSheet) {
      void refreshExternalLocksFromSheet(
        { ...connectedSheet, sheetName },
        token
      );
    }
  };

  const handleDisconnectSheet = () => {
    setConnectedSheet(null);
    setState((prev) => ({
      ...prev,
      externalLockedLocations: [],
      externalResultSourceName: null,
      externalResultSourceType: null,
    }));
    setSheetSyncError(null);
    setNotice('Koneksi link Google Spreadsheet dilepas.');
  };

  const handleRequestSyncToGoogleSheet = () => {
    if (combinedResultRecords.length === 0) return;
    if (!isGoogleAuthenticated || !getWorkspaceAccessToken()) {
      setSheetSyncError(
        'Silakan masuk dengan akun Google terlebih dahulu sebelum menyinkronkan ke Google Sheets.'
      );
      return;
    }
    if (!connectedSheet) {
      setSheetSyncError(
        'Hubungkan link Google Spreadsheet tujuan terlebih dahulu.'
      );
      return;
    }
    setSheetSyncError(null);
    setShowSyncConfirmDialog(true);
  };

  const handleConfirmSyncToGoogleSheet = async () => {
    const token = getWorkspaceAccessToken();
    if (!token || !connectedSheet) {
      setShowSyncConfirmDialog(false);
      setIsGoogleAuthenticated(false);
      return;
    }

    setIsSyncingSheet(true);
    setSheetSyncError(null);
    try {
      const result = await syncRecordsToGoogleSheet(
        connectedSheet,
        combinedResultRecords,
        state.includeExtraExcelColumns,
        token
      );
      const syncedAt = new Date().toISOString();
      setConnectedSheet((prev) =>
        prev
          ? {
              ...prev,
              lastSyncedAt: syncedAt,
              lastSyncSummary: {
                updatedCount: result.updatedCount,
                appendedCount: result.appendedCount,
                unchangedCount: result.unchangedCount,
                totalSyncedRows: result.totalSheetRows,
              },
            }
          : prev
      );

      // Re-read sheet rows to keep external locked locations in sync
      await refreshExternalLocksFromSheet(connectedSheet, token);

      const snapshot = createResultBackupEntry({
        actionLabel: 'Update Google Sheets',
        records: combinedResultRecords,
        includeExtraColumns: state.includeExtraExcelColumns,
        targetSheetTitle: connectedSheet.title,
        targetSheetTab: connectedSheet.sheetName,
        targetSheetUrl: connectedSheet.spreadsheetUrl,
        summaryOverride: `${combinedResultRecords.length} baris disinkronkan (+${result.appendedCount} baru, ${result.updatedCount} diperbarui, ${result.unchangedCount} tetap)`,
        timestamp: syncedAt,
      });
      setBackupHistory((prev) => appendBackupHistoryEntry(prev, snapshot));

      setShowSyncConfirmDialog(false);
      if (result.appendedCount === 0 && result.updatedCount === 0) {
        setNotice(
          `Google Sheets "${connectedSheet.title}" sudah sinkron (${result.unchangedCount} baris sama) & dicadangkan ke Riwayat.`
        );
      } else {
        setNotice(
          `Berhasil sinkron ke "${connectedSheet.title}" (${connectedSheet.sheetName}): ${result.appendedCount} baris baru, ${result.updatedCount} diperbarui & dicadangkan ke Riwayat.`
        );
      }
    } catch (err) {
      setShowSyncConfirmDialog(false);
      handleWorkspaceError(
        err,
        'Gagal memperbarui baris pada Google Spreadsheet.'
      );
    } finally {
      setIsSyncingSheet(false);
    }
  };

  const googleSheetsPanelProps: GoogleSheetsPanelProps = {
    isAuthenticated: isGoogleAuthenticated,
    userEmail: googleUserEmail,
    isSigningIn: isSigningInGoogle,
    connectedSheet,
    driveFiles,
    isLoadingDriveFiles,
    isConnectingSheet,
    isSyncing: isSyncingSheet,
    isRefreshingSheet: isRefreshingDashboard,
    syncError: sheetSyncError,
    pendingSyncCount: pendingSheetSync.totalPendingCount,
    pendingNewCount: pendingSheetSync.pendingNewCount,
    pendingUpdateCount: pendingSheetSync.pendingUpdateCount,
    onSignIn: handleGoogleSignIn,
    onSignOut: handleGoogleSignOut,
    onConnectByUrl: handleConnectSheetByUrl,
    onLoadDriveFiles: handleLoadDriveFiles,
    onCreateNewSheet: handleCreateNewGoogleSheet,
    onSelectSheetTab: handleSelectSheetTab,
    onDisconnectSheet: handleDisconnectSheet,
    onRefreshFromSheet: () => void handleRefreshDashboardFromGoogleSheet(),
    onRequestSyncToGoogleSheet: handleRequestSyncToGoogleSheet,
  };

  const canDirectSyncHeader = Boolean(isGoogleAuthenticated && connectedSheet);

  return (
    <div className="min-h-screen bg-[#F8FAFC] text-[#0F172A]">
      {/* Top Bar Contract: 3 zones separated cleanly */}
      <header className="sticky top-0 z-30 border-b border-slate-200 bg-white/95 backdrop-blur-xs">
        <div className="mx-auto flex max-w-[1440px] items-center justify-between gap-6 px-4 py-3 sm:px-6 lg:gap-8">
          {/* Zone 1: Single-line Brand Wordmark */}
          <a
            href="#top"
            onClick={(e) => {
              e.preventDefault();
              setActiveTab('rack');
            }}
            className="text-base font-bold tracking-tight text-slate-900 whitespace-nowrap shrink-0"
          >
            Relokasi Arsip
          </a>

          {/* Zone 2: Clean single-line navigation links (desktop) */}
          <nav
            aria-label="Navigasi Utama"
            className="hidden items-center gap-7 text-sm font-medium md:flex"
          >
            <button
              type="button"
              onClick={() => setActiveTab('rack')}
              className={`border-b-2 py-1 transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'rack'
                  ? 'border-[#0F4C81] text-[#0F4C81]'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Simulasi Rak
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('results')}
              className={`border-b-2 py-1 transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'results'
                  ? 'border-[#0F4C81] text-[#0F4C81]'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Hasil Excel
            </button>
            <button
              type="button"
              onClick={() => setActiveTab('import')}
              className={`border-b-2 py-1 transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'import'
                  ? 'border-[#0F4C81] text-[#0F4C81]'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Impor Spreadsheet
            </button>
            <button
              type="button"
              data-testid="nav-tab-history"
              onClick={() => setActiveTab('history')}
              className={`border-b-2 py-1 transition-colors whitespace-nowrap shrink-0 ${
                activeTab === 'history'
                  ? 'border-[#0F4C81] text-[#0F4C81]'
                  : 'border-transparent text-slate-600 hover:text-slate-900'
              }`}
            >
              Riwayat
            </button>
          </nav>

          {/* Zone 3: Single Primary Action */}
          <div className="shrink-0">
            {canDirectSyncHeader ? (
              <button
                type="button"
                disabled={combinedResultRecords.length === 0 || isSyncingSheet}
                onClick={handleRequestSyncToGoogleSheet}
                className="inline-flex items-center gap-2 rounded-lg bg-emerald-700 px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-emerald-800 disabled:opacity-45 whitespace-nowrap shrink-0"
              >
                <RefreshCw
                  className={`h-3.5 w-3.5 ${isSyncingSheet ? 'animate-spin' : ''}`}
                />
                {pendingSheetSync.totalPendingCount > 0
                  ? `Update Google Sheets (${pendingSheetSync.totalPendingCount})`
                  : 'Update Google Sheets'}
              </button>
            ) : (
              <button
                type="button"
                disabled={combinedResultRecords.length === 0}
                onClick={handleDownloadExcel}
                className="inline-flex items-center gap-2 rounded-lg bg-[#0F4C81] px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap shrink-0"
              >
                <Download className="h-3.5 w-3.5" />
                Download Excel
              </button>
            )}
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main id="top" className="mx-auto max-w-[1440px] space-y-5 px-4 py-5 sm:px-6">
        {/* Mobile Navigation Segmented Tabs */}
        <nav
          aria-label="Navigasi Mobile"
          className="grid grid-cols-4 gap-1 rounded-xl border border-slate-200 bg-white p-1 md:hidden"
        >
          <button
            type="button"
            onClick={() => setActiveTab('rack')}
            className={`rounded-lg px-2 py-2 text-xs font-semibold transition-colors whitespace-nowrap ${
              activeTab === 'rack'
                ? 'bg-[#0F4C81] text-white'
                : 'text-slate-600 hover:bg-slate-50'
            }`}
          >
            Simulasi
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('results')}
            className={`rounded-lg px-2 py-2 text-xs font-semibold transition-colors whitespace-nowrap ${
              activeTab === 'results'
                ? 'bg-[#0F4C81] text-white'
                : 'text-slate-600 hover:bg-slate-50'
            }`}
          >
            Hasil
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('import')}
            className={`rounded-lg px-2 py-2 text-xs font-semibold transition-colors whitespace-nowrap ${
              activeTab === 'import'
                ? 'bg-[#0F4C81] text-white'
                : 'text-slate-600 hover:bg-slate-50'
            }`}
          >
            Template
          </button>
          <button
            type="button"
            onClick={() => setActiveTab('history')}
            className={`rounded-lg px-2 py-2 text-xs font-semibold transition-colors whitespace-nowrap ${
              activeTab === 'history'
                ? 'bg-[#0F4C81] text-white'
                : 'text-slate-600 hover:bg-slate-50'
            }`}
          >
            Riwayat
          </button>
        </nav>

        {/* Synchronized Template & Spreadsheet Hasil Dashboard */}
        <MetricsBar
          metrics={metrics}
          activeLocationCode={sanitizeLocationCode(activeLocationCode)}
          activeRack={activeRack}
          spreadsheetFileName={state.spreadsheet?.fileName}
          connectedSheet={connectedSheet}
          isGoogleAuthenticated={isGoogleAuthenticated}
          isRefreshingDashboard={isRefreshingDashboard}
          externalResultSourceName={state.externalResultSourceName}
          externalResultSourceType={state.externalResultSourceType}
          onRefreshFromGoogleSheet={() =>
            void handleRefreshDashboardFromGoogleSheet()
          }
          onImportExternalResultsFile={(file) =>
            void handleImportExternalResultsFile(file)
          }
          onClearExternalResults={handleClearExternalResults}
          onOpenTemplateWithFilter={handleOpenTemplateWithFilter}
          onOpenResultsTab={() => setActiveTab('results')}
        />

        {/* Tab Views */}
        {activeTab === 'rack' && (
          <div className="grid grid-cols-1 items-start gap-5 xl:grid-cols-12">
            <div className="xl:col-span-8">
              <RackVisualizer
                config={state.config}
                draftPlacements={effectiveRackPlacements}
                activeLocationCode={activeLocationCode}
                activeRack={activeRack}
                startBoxNumber={startBoxNumber}
                continuousBoxNumbering={continuousBoxNumbering}
                unsavedCount={metrics.unsavedDraftCount}
                canUndo={state.history.length > 0}
                canResetRack={state.draftPlacements.length > 0}
                highlightedCoordinate={highlightedCoordinate}
                onChangeLocationCode={handleChangeLocationCode}
                onSelectRack={setActiveRack}
                onChangeStartBoxNumber={setStartBoxNumber}
                onToggleContinuousBoxNumbering={setContinuousBoxNumbering}
                onClickBox={(coord) => setSelectedCoordinate(coord)}
                onSaveToExcel={handleSaveToExcel}
                onResetRack={handleResetActiveRack}
                onUndo={handleUndo}
                onUpdateConfig={handleUpdateConfig}
              />
            </div>

            <div className="xl:sticky xl:top-[76px] xl:col-span-4">
              <ResultsPanel
                savedRecords={combinedResultRecords}
                unsavedDraftCount={metrics.unsavedDraftCount}
                includeExtraColumns={state.includeExtraExcelColumns}
                generatedArtifact={generatedArtifact}
                isArtifactStale={isArtifactStale}
                googleSheets={googleSheetsPanelProps}
                onToggleExtraColumns={(checked) =>
                  setState((prev) => ({
                    ...prev,
                    includeExtraExcelColumns: checked,
                  }))
                }
                onRemoveSavedRecord={handleRemoveSavedRecord}
                onRequestClearSavedExcel={() => setShowClearExcelDialog(true)}
                onJumpToBox={handleJumpToBox}
                onGenerateExcel={handleGenerateExcel}
                onDownloadExcel={handleDownloadExcel}
              />
            </div>
          </div>
        )}

        {activeTab === 'results' && (
          <ResultsPanel
            savedRecords={combinedResultRecords}
            unsavedDraftCount={metrics.unsavedDraftCount}
            includeExtraColumns={state.includeExtraExcelColumns}
            generatedArtifact={generatedArtifact}
            isArtifactStale={isArtifactStale}
            googleSheets={googleSheetsPanelProps}
            fullWidth
            onToggleExtraColumns={(checked) =>
              setState((prev) => ({
                ...prev,
                includeExtraExcelColumns: checked,
              }))
            }
            onRemoveSavedRecord={handleRemoveSavedRecord}
            onRequestClearSavedExcel={() => setShowClearExcelDialog(true)}
            onJumpToBox={handleJumpToBox}
            onGenerateExcel={handleGenerateExcel}
            onDownloadExcel={handleDownloadExcel}
          />
        )}

        {activeTab === 'import' && (
          <SpreadsheetImportView
            spreadsheet={state.spreadsheet}
            lockedMap={lockedMap}
            allowUnlistedWithSpreadsheet={state.allowUnlistedWithSpreadsheet}
            initialLockFilter={templateFilterPreset}
            onImportSpreadsheet={handleImportSpreadsheet}
            onUpdateSpreadsheet={handleUpdateTemplateSpreadsheet}
            onChangeColumn={handleChangeSpreadsheetColumn}
            onClearSpreadsheet={handleClearSpreadsheet}
            onToggleAllowUnlisted={(allowed) =>
              setState((prev) => ({
                ...prev,
                allowUnlistedWithSpreadsheet: allowed,
              }))
            }
            onBackToRack={() => setActiveTab('rack')}
          />
        )}

        {activeTab === 'history' && (
          <BackupHistoryView
            entries={backupHistory}
            lockedMap={lockedMap}
            canCreateManualSnapshot={
              Boolean(state.spreadsheet) || combinedResultRecords.length > 0
            }
            onCreateManualSnapshot={handleCreateManualBackupSnapshot}
            onRequestRestoreEntry={(entry) => setPendingRestoreEntry(entry)}
            onRemoveEntry={handleRemoveBackupEntry}
            onRequestClearAllHistory={() => setShowClearHistoryDialog(true)}
          />
        )}
      </main>

      {/* Interactive Box Input Modal */}
      <BoxInputModal
        coordinate={selectedCoordinate}
        state={state}
        onClose={() => setSelectedCoordinate(null)}
        onSubmitPlacement={handleSubmitPlacement}
        onRemovePlacement={handleRemoveBoxPlacement}
        onJumpToBox={handleJumpToBox}
        onOpenSpreadsheetTab={() => setActiveTab('import')}
        onQuickAddToTemplate={handleQuickAddToTemplate}
      />

      {/* Explicit Confirmation Dialog ONLY for clearing Saved Excel Records */}
      <ConfirmDialog
        open={showClearExcelDialog}
        title="Kosongkan Semua Data Hasil Simpan?"
        description={`Tindakan ini akan menghapus ${combinedResultRecords.length} baris dari daftar hasil lokal di aplikasi (data yang sudah tersinkron di Google Sheets tidak ikut terhapus). Anda masih dapat membatalkannya menggunakan tombol Undo.`}
        confirmLabel="Ya, Kosongkan Daftar"
        onConfirm={handleConfirmClearSavedExcel}
        onCancel={() => setShowClearExcelDialog(false)}
      />

      {/* Mandatory Explicit Confirmation Dialog Before Updating Google Sheets */}
      <ConfirmDialog
        open={showSyncConfirmDialog && connectedSheet !== null}
        variant="primary"
        busy={isSyncingSheet}
        title="Update Baris ke Google Sheets?"
        description={
          connectedSheet
            ? pendingSheetSync.totalPendingCount > 0
              ? `Sebanyak ${pendingSheetSync.totalPendingCount} perubahan (+${pendingSheetSync.pendingNewCount} baris baru, ${pendingSheetSync.pendingUpdateCount} diperbarui) dari total ${combinedResultRecords.length} baris akan disinkronkan ke spreadsheet "${connectedSheet.title}" (tab "${connectedSheet.sheetName}"). Baris dengan FORMAT LAMA yang sudah ada akan diperbarui di tempat, dan baris baru akan ditambahkan di bawahnya pada link yang sama.`
              : `Sebanyak ${combinedResultRecords.length} baris akan diperiksa & disinkronkan ke spreadsheet "${connectedSheet.title}" (tab "${connectedSheet.sheetName}"). Baris dengan FORMAT LAMA yang sudah ada akan diperbarui di tempat, dan baris baru akan ditambahkan di bawahnya pada link yang sama.`
            : ''
        }
        confirmLabel="Ya, Update Google Sheets"
        onConfirm={handleConfirmSyncToGoogleSheet}
        onCancel={() => setShowSyncConfirmDialog(false)}
      />

      {/* Explicit Confirmation Dialog for Restoring a Backup Snapshot */}
      <ConfirmDialog
        open={pendingRestoreEntry !== null}
        variant="primary"
        title="Pulihkan Cadangan Ini?"
        description={
          pendingRestoreEntry
            ? pendingRestoreEntry.category === 'template'
              ? `Template lokasi simpan lama aktif saat ini akan diganti dengan cadangan "${pendingRestoreEntry.fileName ?? pendingRestoreEntry.actionLabel}" (${pendingRestoreEntry.itemCount} kode).`
              : `Daftar hasil simpan Excel saat ini akan dipulihkan ke versi "${pendingRestoreEntry.actionLabel}" (${pendingRestoreEntry.itemCount} baris). Anda tetap dapat membatalkannya menggunakan tombol Undo.`
            : ''
        }
        confirmLabel="Ya, Pulihkan Cadangan"
        onConfirm={handleConfirmRestoreBackupEntry}
        onCancel={() => setPendingRestoreEntry(null)}
      />

      {/* Explicit Confirmation Dialog for Clearing All Backup History */}
      <ConfirmDialog
        open={showClearHistoryDialog}
        title="Bersihkan Semua Riwayat Cadangan?"
        description={`Tindakan ini akan menghapus ${backupHistory.length} entri cadangan dari riwayat lokal browser. Data simulasi aktif dan file yang sudah diunduh atau disinkron ke Google Sheets tidak akan terpengaruh.`}
        confirmLabel="Ya, Bersihkan Riwayat"
        onConfirm={handleConfirmClearAllHistory}
        onCancel={() => setShowClearHistoryDialog(false)}
      />

      {/* Non-shifting Floating Feedback Toast */}
      {notice && (
        <div
          role="status"
          aria-live="polite"
          className="fixed right-5 bottom-5 z-40 flex max-w-md items-center gap-3 rounded-lg border border-slate-800 bg-slate-900 px-4 py-2.5 text-xs font-medium text-white shadow-lg"
        >
          <span>{notice}</span>
          <button
            type="button"
            onClick={() => setNotice(null)}
            className="text-slate-300 hover:text-white whitespace-nowrap shrink-0"
          >
            Tutup
          </button>
        </div>
      )}
    </div>
  );
}

import { useEffect, useMemo, useState } from 'react';
import { Download } from 'lucide-react';
import { BoxInputModal } from './components/BoxInputModal';
import { ConfirmDialog } from './components/ConfirmDialog';
import { MetricsBar } from './components/MetricsBar';
import { RackVisualizer } from './components/RackVisualizer';
import { ResultsPanel } from './components/ResultsPanel';
import { SpreadsheetImportView } from './components/SpreadsheetImportView';
import type {
  BoxCoordinate,
  ImportedSpreadsheet,
  PlacementRecord,
  RackConfig,
  SimulationState,
} from './types/archive';
import {
  downloadExcelArtifact,
  generateRelocationExcelArtifact,
  type GeneratedExcelArtifact,
} from './utils/excelService';
import {
  applyPlacement,
  calculateSimulationMetrics,
  clearSavedExcelRecords,
  commitDraftToExcel,
  createInitialSimulationState,
  extractSpreadsheetItems,
  getAllTrackedPlacements,
  removeDraftByCoordinate,
  removeSavedRecordById,
  resetActiveRackSimulation,
  sanitizeLocationCode,
  undoLastAction,
  updateRackConfig,
} from './utils/relocationEngine';

type ActiveTab = 'rack' | 'results' | 'import';

const STORAGE_KEY = 'arsip_relokasi_simulasi_v2';

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

  const metrics = useMemo(
    () => calculateSimulationMetrics(state, activeLocationCode, activeRack),
    [state, activeLocationCode, activeRack]
  );

  const allTrackedPlacements = useMemo(
    () => getAllTrackedPlacements(state),
    [state]
  );

  const currentDataSignature = useMemo(() => {
    const rowsSig = state.savedRecords
      .map((p) => `${p.sequence}:${p.oldLocation}:${p.newLocation}`)
      .join('|');
    return `${state.includeExtraExcelColumns ? 'extra' : 'basic'}::${rowsSig}`;
  }, [state.savedRecords, state.includeExtraExcelColumns]);

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
        allowUnlistedWithSpreadsheet: state.allowUnlistedWithSpreadsheet,
        includeExtraExcelColumns: state.includeExtraExcelColumns,
      };
      window.localStorage.setItem(STORAGE_KEY, JSON.stringify(serializable));
    } catch {
      // Ignore quota errors cleanly
    }
  }, [state]);

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
    setState((prev) => {
      const { nextState, addedCount } = commitDraftToExcel(prev);
      if (addedCount > 0) {
        setNotice(
          `${addedCount} lokasi dari simulasi rak berhasil disimpan ke Excel (total ${nextState.savedRecords.length} baris).`
        );
      }
      return nextState;
    });
  };

  const handleResetActiveRack = () => {
    setState((prev) => resetActiveRackSimulation(prev));
    setNotice(
      `Simulasi rak aktif direset. ${state.savedRecords.length} data yang sudah disimpan ke Excel tetap tersimpan.`
    );
  };

  const handleRemoveDraftByCoordinate = (coord: BoxCoordinate) => {
    setState((prev) => removeDraftByCoordinate(prev, coord));
    setNotice(
      `Isi box ${coord.locationCode} · RAK ${coord.rack} · ROW ${coord.row} · BOX ${coord.box} dihapus dari simulasi rak.`
    );
  };

  const handleRemoveSavedRecord = (placementId: string) => {
    const target = state.savedRecords.find((p) => p.id === placementId);
    setState((prev) => removeSavedRecordById(prev, placementId));
    if (target) {
      setNotice(
        `Baris ${target.oldLocation} (${target.newLocation}) dihapus dari hasil Excel.`
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
    setState((prev) => ({
      ...prev,
      spreadsheet: sheet,
    }));
    setNotice(
      `Berhasil mengimpor ${sheet.items.length} lokasi simpan lama dari ${sheet.fileName}.`
    );
  };

  const handleChangeSpreadsheetColumn = (nextColumn: string) => {
    setState((prev) => {
      if (!prev.spreadsheet) return prev;
      const descCol = prev.spreadsheet.columns.find(
        (col) =>
          col !== nextColumn &&
          /keterangan|uraian|judul|nama|arsip|unit|deskripsi/i.test(col)
      );
      const nextItems = extractSpreadsheetItems(
        prev.spreadsheet.rawRows,
        nextColumn,
        descCol
      );
      return {
        ...prev,
        spreadsheet: {
          ...prev.spreadsheet,
          selectedColumn: nextColumn,
          items: nextItems,
        },
      };
    });
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
    if (state.savedRecords.length === 0) return;
    const artifact = generateRelocationExcelArtifact(
      state.savedRecords,
      state.includeExtraExcelColumns
    );
    setGeneratedArtifact(artifact);
    setArtifactSignature(currentDataSignature);
    setNotice(
      `File ${artifact.fileName} berhasil dibuat (${artifact.rowCount} baris).`
    );
  };

  const handleDownloadExcel = () => {
    if (state.savedRecords.length === 0) return;
    let artifact = generatedArtifact;
    if (!artifact || isArtifactStale) {
      artifact = generateRelocationExcelArtifact(
        state.savedRecords,
        state.includeExtraExcelColumns
      );
      setGeneratedArtifact(artifact);
      setArtifactSignature(currentDataSignature);
    }
    downloadExcelArtifact(artifact);
    setNotice(`Mengunduh ${artifact.fileName}...`);
  };

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
          </nav>

          {/* Zone 3: Single Primary Action */}
          <div className="shrink-0">
            <button
              type="button"
              disabled={state.savedRecords.length === 0}
              onClick={handleDownloadExcel}
              className="inline-flex items-center gap-2 rounded-lg bg-[#0F4C81] px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap shrink-0"
            >
              <Download className="h-3.5 w-3.5" />
              Download Excel
            </button>
          </div>
        </div>
      </header>

      {/* Main Content */}
      <main id="top" className="mx-auto max-w-[1440px] space-y-5 px-4 py-5 sm:px-6">
        {/* Mobile Navigation Segmented Tabs */}
        <nav
          aria-label="Navigasi Mobile"
          className="grid grid-cols-3 gap-1 rounded-xl border border-slate-200 bg-white p-1 md:hidden"
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
            Simulasi Rak
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
            Hasil Excel
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
            Impor Data
          </button>
        </nav>

        {/* Metrics & Progress Strip */}
        <MetricsBar
          metrics={metrics}
          activeLocationCode={sanitizeLocationCode(activeLocationCode)}
          activeRack={activeRack}
          spreadsheetFileName={state.spreadsheet?.fileName}
        />

        {/* Tab Views */}
        {activeTab === 'rack' && (
          <div className="grid grid-cols-1 items-start gap-5 xl:grid-cols-12">
            <div className="xl:col-span-8">
              <RackVisualizer
                config={state.config}
                draftPlacements={state.draftPlacements}
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
                savedRecords={state.savedRecords}
                unsavedDraftCount={metrics.unsavedDraftCount}
                includeExtraColumns={state.includeExtraExcelColumns}
                generatedArtifact={generatedArtifact}
                isArtifactStale={isArtifactStale}
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
            savedRecords={state.savedRecords}
            unsavedDraftCount={metrics.unsavedDraftCount}
            includeExtraColumns={state.includeExtraExcelColumns}
            generatedArtifact={generatedArtifact}
            isArtifactStale={isArtifactStale}
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
            placements={allTrackedPlacements}
            allowUnlistedWithSpreadsheet={state.allowUnlistedWithSpreadsheet}
            onImportSpreadsheet={handleImportSpreadsheet}
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
      </main>

      {/* Interactive Box Input Modal */}
      <BoxInputModal
        coordinate={selectedCoordinate}
        state={state}
        onClose={() => setSelectedCoordinate(null)}
        onSubmitPlacement={handleSubmitPlacement}
        onRemovePlacement={handleRemoveDraftByCoordinate}
        onJumpToBox={handleJumpToBox}
        onOpenSpreadsheetTab={() => setActiveTab('import')}
      />

      {/* Explicit Confirmation Dialog ONLY for clearing Saved Excel Records */}
      <ConfirmDialog
        open={showClearExcelDialog}
        title="Kosongkan Semua Data Hasil Excel?"
        description={`Tindakan ini akan menghapus ${state.savedRecords.length} baris yang sudah tersimpan pada tabel Excel. Anda masih dapat membatalkannya menggunakan tombol Undo.`}
        confirmLabel="Ya, Kosongkan Excel"
        onConfirm={handleConfirmClearSavedExcel}
        onCancel={() => setShowClearExcelDialog(false)}
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

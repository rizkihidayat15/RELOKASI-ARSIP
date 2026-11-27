import React, { useEffect, useMemo, useState } from 'react';
import {
  Check,
  Minus,
  Plus,
  RotateCcw,
  Save,
  Search,
  SlidersHorizontal,
  Undo2,
} from 'lucide-react';
import type {
  BoxCoordinate,
  PlacementRecord,
  RackConfig,
} from '../types/archive';
import {
  computeBoxNumber,
  formatNewLocationCode,
  sanitizeLocationCode,
} from '../utils/relocationEngine';

type BoxFilter = 'all' | 'empty' | 'filled';

interface RackVisualizerProps {
  config: RackConfig;
  draftPlacements: PlacementRecord[];
  activeLocationCode: string;
  activeRack: number;
  startBoxNumber: number;
  continuousBoxNumbering: boolean;
  unsavedCount: number;
  canUndo: boolean;
  canResetRack: boolean;
  highlightedCoordinate: BoxCoordinate | null;
  onChangeLocationCode: (value: string) => void;
  onSelectRack: (rack: number) => void;
  onChangeStartBoxNumber: (value: number) => void;
  onToggleContinuousBoxNumbering: (continuous: boolean) => void;
  onClickBox: (coord: BoxCoordinate) => void;
  onSaveToExcel: () => void;
  onResetRack: () => void;
  onUndo: () => void;
  onUpdateConfig: (nextConfig: RackConfig) => void;
}

export const RackVisualizer: React.FC<RackVisualizerProps> = ({
  config,
  draftPlacements,
  activeLocationCode,
  activeRack,
  startBoxNumber,
  continuousBoxNumbering,
  unsavedCount,
  canUndo,
  canResetRack,
  highlightedCoordinate,
  onChangeLocationCode,
  onSelectRack,
  onChangeStartBoxNumber,
  onToggleContinuousBoxNumbering,
  onClickBox,
  onSaveToExcel,
  onResetRack,
  onUndo,
  onUpdateConfig,
}) => {
  const [filter, setFilter] = useState<BoxFilter>('all');
  const [searchQuery, setSearchQuery] = useState('');
  const [showConfig, setShowConfig] = useState(false);

  const [draftRows, setDraftRows] = useState(String(config.rowsPerRack));
  const [draftBoxes, setDraftBoxes] = useState(String(config.boxesPerRow));
  const [rackInputText, setRackInputText] = useState(String(activeRack));
  const [startBoxInputText, setStartBoxInputText] = useState(
    String(startBoxNumber)
  );

  useEffect(() => {
    setRackInputText(String(activeRack));
  }, [activeRack]);

  useEffect(() => {
    setStartBoxInputText(String(startBoxNumber));
  }, [startBoxNumber]);

  const cleanLocationCode = sanitizeLocationCode(activeLocationCode);

  const placementMap = useMemo(() => {
    const map = new Map<string, PlacementRecord>();
    for (const p of draftPlacements) {
      map.set(
        `${sanitizeLocationCode(p.locationCode)}-${p.rack}-${p.row}-${p.box}`,
        p
      );
    }
    return map;
  }, [draftPlacements]);

  const activeRackFilled = useMemo(() => {
    return draftPlacements.filter(
      (p) =>
        sanitizeLocationCode(p.locationCode) === cleanLocationCode &&
        p.rack === activeRack
    ).length;
  }, [draftPlacements, cleanLocationCode, activeRack]);

  const boxesPerRack = config.rowsPerRack * config.boxesPerRow;

  const rowNumbers = useMemo(
    () => Array.from({ length: config.rowsPerRack }, (_, i) => i + 1),
    [config.rowsPerRack]
  );
  const slotIndices = useMemo(
    () => Array.from({ length: config.boxesPerRow }, (_, i) => i),
    [config.boxesPerRow]
  );

  const endBoxNumber = useMemo(() => {
    return computeBoxNumber(
      config.rowsPerRack,
      config.boxesPerRow - 1,
      config.boxesPerRow,
      startBoxNumber,
      continuousBoxNumbering
    );
  }, [
    config.rowsPerRack,
    config.boxesPerRow,
    startBoxNumber,
    continuousBoxNumbering,
  ]);

  const normalizedQuery = searchQuery.trim().toLowerCase();

  const openConfigPanel = () => {
    setDraftRows(String(config.rowsPerRack));
    setDraftBoxes(String(config.boxesPerRow));
    setShowConfig((prev) => !prev);
  };

  const handleApplyConfig = (e: React.FormEvent) => {
    e.preventDefault();
    const nextConfig: RackConfig = {
      prefix: cleanLocationCode,
      rackCount: 1,
      rowsPerRack: Math.max(1, Math.min(20, parseInt(draftRows, 10) || 5)),
      boxesPerRow: Math.max(1, Math.min(30, parseInt(draftBoxes, 10) || 10)),
    };
    onUpdateConfig(nextConfig);
    setShowConfig(false);
  };

  const handleRackInputChange = (raw: string) => {
    const digitsOnly = raw.replace(/\D/g, '').slice(0, 3);
    setRackInputText(digitsOnly);
    if (!digitsOnly) return;
    const parsed = Math.max(1, Math.min(999, parseInt(digitsOnly, 10)));
    onSelectRack(parsed);
  };

  const handleStartBoxInputChange = (raw: string) => {
    const digitsOnly = raw.replace(/\D/g, '').slice(0, 5);
    setStartBoxInputText(digitsOnly);
    if (!digitsOnly) return;
    const parsed = Math.max(1, Math.min(99999, parseInt(digitsOnly, 10)));
    onChangeStartBoxNumber(parsed);
  };

  return (
    <section
      aria-label="Simulasi 1 Rak Utuh"
      className="rounded-xl border border-slate-200 bg-white"
    >
      {/* Top Toolbar: Header + Save / Reset / Undo Actions */}
      <div className="border-b border-slate-200 px-5 py-4">
        <div className="flex flex-wrap items-start justify-between gap-3">
          <div>
            <h2 className="text-base font-bold tracking-tight text-slate-900">
              Simulasi 1 Rak Utuh
            </h2>
            <p className="mt-0.5 text-xs text-slate-500">
              Masukkan Nomor Rak, isi box pada rak, klik{' '}
              <span className="font-semibold text-slate-700">Simpan</span> untuk
              memasukkan ke Excel, lalu{' '}
              <span className="font-semibold text-slate-700">Reset</span> untuk
              lanjut ke rak berikutnya.
            </p>
          </div>

          <button
            type="button"
            onClick={openConfigPanel}
            aria-expanded={showConfig}
            className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 bg-white px-3 py-1.5 text-xs font-medium text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap shrink-0"
          >
            <SlidersHorizontal className="h-3.5 w-3.5" />
            Atur Row & Box
          </button>
        </div>

        {/* Dedicated 1-Rack Input Bar: Lokasi Simpan + Nomor Rak + Mulai No. Box + Format Preview */}
        <div className="mt-4 grid grid-cols-1 gap-3 rounded-lg border border-slate-200 bg-slate-50/80 p-3.5 sm:grid-cols-12 sm:items-end">
          <div className="sm:col-span-3">
            <label
              htmlFor="input-lokasi-simpan"
              className="block text-xs font-semibold text-slate-700"
            >
              Lokasi Simpan
            </label>
            <input
              id="input-lokasi-simpan"
              type="text"
              value={activeLocationCode}
              onChange={(e) =>
                onChangeLocationCode(e.target.value.toUpperCase())
              }
              onBlur={() => {
                if (!activeLocationCode.trim()) {
                  onChangeLocationCode('RC2');
                }
              }}
              maxLength={16}
              placeholder="Contoh: RC2"
              className="mt-1 w-full rounded-lg border border-slate-300 bg-white px-3 py-2 font-mono text-sm font-semibold text-slate-900 uppercase placeholder:font-sans placeholder:font-normal placeholder:normal-case placeholder:text-slate-400 focus:border-[#0F4C81] focus:ring-2 focus:ring-[#0F4C81]/20 focus:outline-none"
            />
          </div>

          <div className="sm:col-span-3">
            <label
              htmlFor="input-nomor-rak"
              className="block text-xs font-semibold text-slate-700"
            >
              Nomor Rak
            </label>
            <div className="mt-1 flex items-center">
              <button
                type="button"
                aria-label="Kurangi nomor rak"
                disabled={activeRack <= 1}
                onClick={() => onSelectRack(Math.max(1, activeRack - 1))}
                className="flex h-[38px] w-8 shrink-0 items-center justify-center rounded-l-lg border border-r-0 border-slate-300 bg-white text-slate-600 transition-colors hover:bg-slate-100 disabled:opacity-40"
              >
                <Minus className="h-3.5 w-3.5" />
              </button>
              <input
                id="input-nomor-rak"
                type="text"
                inputMode="numeric"
                value={rackInputText}
                onChange={(e) => handleRackInputChange(e.target.value)}
                onBlur={() => {
                  if (!rackInputText.trim()) {
                    setRackInputText(String(activeRack));
                  }
                }}
                placeholder="Mis. 3"
                className="h-[38px] w-full min-w-0 border border-slate-300 bg-white px-2 text-center font-mono text-sm font-semibold text-slate-900 tabular-nums focus:z-10 focus:border-[#0F4C81] focus:ring-2 focus:ring-[#0F4C81]/20 focus:outline-none"
              />
              <button
                type="button"
                aria-label="Tambah nomor rak"
                disabled={activeRack >= 999}
                onClick={() => onSelectRack(Math.min(999, activeRack + 1))}
                className="flex h-[38px] w-8 shrink-0 items-center justify-center rounded-r-lg border border-l-0 border-slate-300 bg-white text-slate-600 transition-colors hover:bg-slate-100 disabled:opacity-40"
              >
                <Plus className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>

          <div className="sm:col-span-3">
            <label
              htmlFor="input-mulai-box"
              className="block text-xs font-semibold text-slate-700"
            >
              Mulai Nomor Box
            </label>
            <div className="mt-1 flex items-center">
              <button
                type="button"
                aria-label="Kurangi nomor box awal"
                disabled={startBoxNumber <= 1}
                onClick={() =>
                  onChangeStartBoxNumber(Math.max(1, startBoxNumber - 1))
                }
                className="flex h-[38px] w-8 shrink-0 items-center justify-center rounded-l-lg border border-r-0 border-slate-300 bg-white text-slate-600 transition-colors hover:bg-slate-100 disabled:opacity-40"
              >
                <Minus className="h-3.5 w-3.5" />
              </button>
              <input
                id="input-mulai-box"
                type="text"
                inputMode="numeric"
                value={startBoxInputText}
                onChange={(e) => handleStartBoxInputChange(e.target.value)}
                onBlur={() => {
                  if (!startBoxInputText.trim()) {
                    setStartBoxInputText(String(startBoxNumber));
                  }
                }}
                placeholder="Mis. 51"
                className="h-[38px] w-full min-w-0 border border-slate-300 bg-white px-2 text-center font-mono text-sm font-semibold text-slate-900 tabular-nums focus:z-10 focus:border-[#0F4C81] focus:ring-2 focus:ring-[#0F4C81]/20 focus:outline-none"
              />
              <button
                type="button"
                aria-label="Tambah nomor box awal"
                disabled={startBoxNumber >= 99999}
                onClick={() =>
                  onChangeStartBoxNumber(Math.min(99999, startBoxNumber + 1))
                }
                className="flex h-[38px] w-8 shrink-0 items-center justify-center rounded-r-lg border border-l-0 border-slate-300 bg-white text-slate-600 transition-colors hover:bg-slate-100 disabled:opacity-40"
              >
                <Plus className="h-3.5 w-3.5" />
              </button>
            </div>
          </div>

          <div className="sm:col-span-3">
            <p className="text-xs font-semibold text-slate-500">
              Format Nilai Hasil
            </p>
            <p
              data-testid="live-format-preview"
              title={`Rentang nomor box: ${startBoxNumber}–${endBoxNumber}`}
              className="mt-1 flex h-[38px] items-center truncate rounded-lg border border-slate-200 bg-white px-2.5 font-mono text-xs font-bold text-[#0F4C81] tabular-nums"
            >
              {cleanLocationCode}.{activeRack}.(ROW).(BOX)
            </p>
          </div>
        </div>

        {/* Cycle Action Bar: Simpan ke Excel + Reset Simulasi Rak + Undo */}
        <div className="mt-3.5 flex flex-wrap items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <button
              type="button"
              disabled={unsavedCount === 0}
              onClick={onSaveToExcel}
              data-testid="btn-simpan-excel"
              className="inline-flex items-center gap-1.5 rounded-lg bg-[#0F4C81] px-4 py-2 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] disabled:opacity-45 whitespace-nowrap shrink-0"
            >
              <Save className="h-3.5 w-3.5" />
              {unsavedCount > 0 ? `Simpan (${unsavedCount})` : 'Simpan'}
            </button>

            <button
              type="button"
              disabled={!canResetRack}
              onClick={onResetRack}
              data-testid="btn-reset-rak"
              title="Reset simulasi rak aktif tanpa menghapus data yang sudah disimpan ke Excel"
              className="inline-flex items-center gap-1.5 rounded-lg border border-amber-300 bg-amber-50/70 px-3.5 py-2 text-xs font-semibold text-amber-900 transition-colors hover:bg-amber-100/80 disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <RotateCcw className="h-3.5 w-3.5" />
              Reset
            </button>

            <button
              type="button"
              disabled={!canUndo}
              onClick={onUndo}
              className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 bg-white px-3 py-2 text-xs font-semibold text-slate-700 transition-colors hover:bg-slate-50 disabled:opacity-40 whitespace-nowrap shrink-0"
            >
              <Undo2 className="h-3.5 w-3.5" />
              Undo
            </button>
          </div>

          <p className="text-[11px] text-slate-500">
            Tombol <span className="font-semibold text-slate-700">Reset</span>{' '}
            hanya mengosongkan rak aktif — hasil yang sudah disimpan ke Excel
            tetap aman.
          </p>
        </div>

        {/* Collapsible Row & Box Dimensions Panel */}
        {showConfig && (
          <form
            onSubmit={handleApplyConfig}
            className="mt-3 space-y-3 rounded-lg border border-slate-200 bg-slate-50 p-3.5"
          >
            <div className="grid grid-cols-2 gap-3 sm:grid-cols-3">
              <div>
                <label
                  htmlFor="cfg-rows"
                  className="block text-xs font-medium text-slate-600"
                >
                  Jumlah Row / Rak
                </label>
                <input
                  id="cfg-rows"
                  type="number"
                  min={1}
                  max={20}
                  value={draftRows}
                  onChange={(e) => setDraftRows(e.target.value)}
                  className="mt-1 w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 font-mono text-sm text-slate-900 tabular-nums focus:border-[#0F4C81] focus:outline-none"
                />
              </div>
              <div>
                <label
                  htmlFor="cfg-boxes"
                  className="block text-xs font-medium text-slate-600"
                >
                  Jumlah Box / Row
                </label>
                <input
                  id="cfg-boxes"
                  type="number"
                  min={1}
                  max={30}
                  value={draftBoxes}
                  onChange={(e) => setDraftBoxes(e.target.value)}
                  className="mt-1 w-full rounded-md border border-slate-300 bg-white px-2.5 py-1.5 font-mono text-sm text-slate-900 tabular-nums focus:border-[#0F4C81] focus:outline-none"
                />
              </div>
              <div className="col-span-2 flex items-end sm:col-span-1">
                <button
                  type="submit"
                  className="w-full rounded-md bg-[#0F4C81] px-3 py-1.5 text-xs font-semibold text-white transition-colors hover:bg-[#0B3A63] whitespace-nowrap"
                >
                  Terapkan Ukuran
                </button>
              </div>
            </div>
            <label className="flex cursor-pointer items-center gap-2 text-xs text-slate-600 select-none">
              <input
                type="checkbox"
                checked={continuousBoxNumbering}
                onChange={(e) =>
                  onToggleContinuousBoxNumbering(e.target.checked)
                }
                className="h-3.5 w-3.5 rounded border-slate-300 text-[#0F4C81] focus:ring-[#0F4C81]"
              />
              <span>
                Penomoran box berlanjut antar row (contoh: Row 1 ={' '}
                {startBoxNumber}–{startBoxNumber + config.boxesPerRow - 1}, Row
                2 = {startBoxNumber + config.boxesPerRow}–
                {startBoxNumber + config.boxesPerRow * 2 - 1})
              </span>
            </label>
          </form>
        )}

        {/* Filter & Search Bar */}
        <div className="mt-3.5 flex flex-wrap items-center justify-between gap-3 border-t border-slate-100 pt-3.5">
          <div className="inline-flex items-center gap-1 rounded-lg bg-slate-100 p-1">
            {(
              [
                { id: 'all', label: 'Semua' },
                { id: 'empty', label: 'Kosong' },
                { id: 'filled', label: 'Terisi' },
              ] as const
            ).map((opt) => (
              <button
                key={opt.id}
                type="button"
                onClick={() => setFilter(opt.id)}
                className={`rounded-md px-3 py-1 text-xs font-medium transition-colors whitespace-nowrap shrink-0 ${
                  filter === opt.id
                    ? 'bg-white text-slate-900 shadow-xs'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                {opt.label}
              </button>
            ))}
          </div>

          <div className="relative min-w-[210px] flex-1 sm:max-w-xs">
            <Search
              aria-hidden="true"
              className="pointer-events-none absolute top-1/2 left-3 h-3.5 w-3.5 -translate-y-1/2 text-slate-400"
            />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              aria-label="Cari box atau lokasi pada rak ini"
              placeholder="Cari lokasi lama / kode baru..."
              className="w-full rounded-lg border border-slate-200 bg-slate-50 py-1.5 pr-3 pl-8 text-xs text-slate-900 placeholder:text-slate-400 focus:border-[#0F4C81] focus:bg-white focus:outline-none"
            />
          </div>
        </div>
      </div>

      {/* Single Complete Rack Architectural Grid */}
      <div className="p-5">
        <div className="mb-3 flex flex-wrap items-center justify-between gap-2 text-xs text-slate-500">
          <div className="font-semibold text-slate-800">
            <span className="font-mono text-[#0F4C81]">
              {cleanLocationCode}
            </span>
            <span className="mx-1.5 text-slate-300">·</span>
            <span>RAK {activeRack}</span>
            <span className="mx-1.5 text-slate-300">·</span>
            <span className="font-mono font-medium text-slate-700 tabular-nums">
              Box {startBoxNumber}–{endBoxNumber}
            </span>
            <span className="mx-2 font-normal text-slate-300">/</span>
            <span className="font-normal text-slate-600 tabular-nums">
              {activeRackFilled} terisi · {boxesPerRack - activeRackFilled}{' '}
              kosong
            </span>
          </div>
          <div className="flex flex-wrap items-center gap-4 text-xs text-slate-500">
            <span className="inline-flex items-center gap-1.5">
              <span
                aria-hidden="true"
                className="inline-block h-3 w-3 rounded border border-slate-300 bg-white"
              />
              Kosong
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span
                aria-hidden="true"
                className="inline-block h-3 w-3 rounded border border-[#0F4C81] bg-sky-50"
              />
              Belum Disimpan
            </span>
            <span className="inline-flex items-center gap-1.5">
              <span
                aria-hidden="true"
                className="inline-block h-3 w-3 rounded border border-emerald-600 bg-emerald-50"
              />
              Sudah Tersimulasi (Tabel Hasil)
            </span>
          </div>
        </div>

        <div className="space-y-3">
          {rowNumbers.map((rowNum) => {
            const rowBoxNumbers = slotIndices.map((slotIdx) =>
              computeBoxNumber(
                rowNum,
                slotIdx,
                config.boxesPerRow,
                startBoxNumber,
                continuousBoxNumbering
              )
            );
            const rowFilled = rowBoxNumbers.reduce((acc, boxNum) => {
              return (
                acc +
                (placementMap.has(
                  `${cleanLocationCode}-${activeRack}-${rowNum}-${boxNum}`
                )
                  ? 1
                  : 0)
              );
            }, 0);

            return (
              <div
                key={rowNum}
                className="rounded-lg border border-slate-200 bg-slate-50/70 p-3"
              >
                <div className="mb-2 flex items-center justify-between">
                  <span className="text-xs font-bold tracking-wide text-slate-700">
                    ROW {rowNum}
                    <span className="ml-2 font-mono text-[11px] font-normal text-slate-500 tabular-nums">
                      (Box {rowBoxNumbers[0]}–
                      {rowBoxNumbers[rowBoxNumbers.length - 1]})
                    </span>
                  </span>
                  <span className="font-mono text-[11px] text-slate-500 tabular-nums">
                    {rowFilled}/{config.boxesPerRow} terisi
                  </span>
                </div>

                <div
                  className="grid gap-1.5"
                  style={{
                    gridTemplateColumns:
                      config.boxesPerRow <= 10
                        ? `repeat(auto-fit, minmax(72px, 1fr))`
                        : `repeat(auto-fill, minmax(76px, 1fr))`,
                  }}
                >
                  {rowBoxNumbers.map((boxNum) => {
                    const key = `${cleanLocationCode}-${activeRack}-${rowNum}-${boxNum}`;
                    const placement = placementMap.get(key);
                    const newCode = formatNewLocationCode(
                      cleanLocationCode,
                      activeRack,
                      rowNum,
                      boxNum
                    );
                    const isFilled = Boolean(placement);
                    const isSaved = Boolean(placement?.savedToExcel);

                    const matchesFilter =
                      filter === 'all' ||
                      (filter === 'filled' && isFilled) ||
                      (filter === 'empty' && !isFilled);

                    const matchesSearch =
                      !normalizedQuery ||
                      newCode.toLowerCase().includes(normalizedQuery) ||
                      (placement?.oldLocation
                        .toLowerCase()
                        .includes(normalizedQuery) ??
                        false) ||
                      `box ${boxNum}`.includes(normalizedQuery);

                    const isDimmed = !matchesFilter || !matchesSearch;
                    const isHighlighted =
                      highlightedCoordinate !== null &&
                      sanitizeLocationCode(
                        highlightedCoordinate.locationCode
                      ) === cleanLocationCode &&
                      highlightedCoordinate.rack === activeRack &&
                      highlightedCoordinate.row === rowNum &&
                      highlightedCoordinate.box === boxNum;

                    const accessibleLabel = isFilled
                      ? `${cleanLocationCode} Rak ${activeRack} Row ${rowNum} Box ${boxNum}, Kode ${newCode}, Terisi lokasi lama ${placement!.oldLocation}${isSaved ? ', Sudah Tersimulasi (Tabel Hasil)' : ', Draft belum disimpan'}`
                      : `${cleanLocationCode} Rak ${activeRack} Row ${rowNum} Box ${boxNum}, Kode ${newCode}, Kosong`;

                    return (
                      <button
                        key={key}
                        type="button"
                        data-testid={`box-${activeRack}-${rowNum}-${boxNum}`}
                        data-location-code={newCode}
                        data-box-state={
                          !isFilled ? 'empty' : isSaved ? 'saved' : 'draft'
                        }
                        aria-label={accessibleLabel}
                        onClick={() =>
                          onClickBox({
                            locationCode: cleanLocationCode,
                            rack: activeRack,
                            row: rowNum,
                            box: boxNum,
                          })
                        }
                        className={`group relative flex min-h-[70px] flex-col justify-between rounded-lg border px-2 py-1.5 text-left transition-all duration-150 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[#0F4C81] ${
                          isDimmed ? 'opacity-30' : 'opacity-100'
                        } ${
                          isHighlighted
                            ? 'ring-2 ring-[#0F4C81] ring-offset-1'
                            : ''
                        } ${
                          !isFilled
                            ? 'border-slate-200 bg-white hover:border-[#0F4C81] hover:bg-sky-50/60'
                            : isSaved
                              ? 'border-emerald-600/55 bg-emerald-50/90 hover:border-emerald-700 hover:bg-emerald-100/80'
                              : 'border-[#0F4C81]/60 bg-sky-50/90 hover:border-[#0F4C81] hover:bg-sky-100/80'
                        }`}
                      >
                        <div className="flex w-full items-center justify-between gap-0.5">
                          <span
                            className={`text-[10px] font-semibold tracking-tight whitespace-nowrap ${
                              !isFilled
                                ? 'text-slate-500'
                                : isSaved
                                  ? 'text-emerald-800'
                                  : 'text-[#0F4C81]'
                            }`}
                          >
                            B{boxNum}
                          </span>
                          {isFilled ? (
                            isSaved ? (
                              <Check
                                aria-hidden="true"
                                className="h-3 w-3 shrink-0 text-emerald-700"
                              />
                            ) : (
                              <span
                                aria-hidden="true"
                                className="text-[9px] font-semibold text-[#0F4C81]"
                              >
                                Draft
                              </span>
                            )
                          ) : (
                            <Plus
                              aria-hidden="true"
                              className="h-3 w-3 shrink-0 text-slate-300 transition-colors group-hover:text-[#0F4C81]"
                            />
                          )}
                        </div>

                        <div className="my-0.5 min-w-0">
                          {isFilled ? (
                            <p
                              title={placement!.oldLocation}
                              className={`truncate font-mono text-sm font-bold tabular-nums ${
                                isSaved ? 'text-emerald-950' : 'text-slate-950'
                              }`}
                            >
                              {placement!.oldLocation}
                            </p>
                          ) : (
                            <p className="text-[11px] font-medium text-slate-400 group-hover:text-[#0F4C81]">
                              Kosong
                            </p>
                          )}
                        </div>

                        <span
                          className={`block truncate font-mono text-[9.5px] tabular-nums ${
                            !isFilled
                              ? 'text-slate-400'
                              : isSaved
                                ? 'text-emerald-700'
                                : 'text-[#0F4C81]'
                          }`}
                        >
                          {newCode}
                        </span>
                      </button>
                    );
                  })}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </section>
  );
};

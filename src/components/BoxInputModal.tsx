import React, { useEffect, useMemo, useRef, useState } from 'react';
import {
  AlertCircle,
  AlertTriangle,
  ArrowRight,
  Check,
  FileSpreadsheet,
  Search,
  Trash2,
  X,
} from 'lucide-react';
import type {
  BoxCoordinate,
  PlacementRecord,
  SimulationState,
  SpreadsheetItem,
} from '../types/archive';
import {
  formatBoxAddress,
  formatNewLocationCode,
  getAllTrackedPlacements,
  getPlacementByCoordinate,
  getPlacementByOldLocation,
  normalizeLocationCode,
  sanitizeLocationCode,
  searchOldLocations,
  validatePlacementInput,
} from '../utils/relocationEngine';

interface BoxInputModalProps {
  coordinate: BoxCoordinate | null;
  state: SimulationState;
  onClose: () => void;
  onSubmitPlacement: (
    coord: BoxCoordinate,
    oldLocation: string,
    confirmReplace: boolean
  ) => void;
  onRemovePlacement: (coord: BoxCoordinate) => void;
  onJumpToBox: (coord: BoxCoordinate) => void;
  onOpenSpreadsheetTab: () => void;
}

export const BoxInputModal: React.FC<BoxInputModalProps> = ({
  coordinate,
  state,
  onClose,
  onSubmitPlacement,
  onRemovePlacement,
  onJumpToBox,
  onOpenSpreadsheetTab,
}) => {
  const [inputValue, setInputValue] = useState('');
  const [boxNumberInput, setBoxNumberInput] = useState('');
  const [showSuggestions, setShowSuggestions] = useState(false);
  const [activeSuggestionIdx, setActiveSuggestionIdx] = useState(0);
  const [pendingReplaceConfirm, setPendingReplaceConfirm] = useState<string | null>(
    null
  );
  const [submitAttempted, setSubmitAttempted] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const effectiveCoordinate: BoxCoordinate | null = useMemo(() => {
    if (!coordinate) return null;
    const parsedBox = parseInt(boxNumberInput, 10);
    const resolvedBox =
      Number.isFinite(parsedBox) && parsedBox >= 1 ? parsedBox : coordinate.box;
    return {
      ...coordinate,
      box: resolvedBox,
    };
  }, [coordinate, boxNumberInput]);

  const currentPlacement: PlacementRecord | undefined = useMemo(() => {
    if (!effectiveCoordinate) return undefined;
    return getPlacementByCoordinate(state.draftPlacements, effectiveCoordinate);
  }, [effectiveCoordinate, state.draftPlacements]);

  const newLocationCode = useMemo(() => {
    if (!effectiveCoordinate) return '';
    return formatNewLocationCode(
      effectiveCoordinate.locationCode,
      effectiveCoordinate.rack,
      effectiveCoordinate.row,
      effectiveCoordinate.box
    );
  }, [effectiveCoordinate]);

  useEffect(() => {
    if (coordinate) {
      setInputValue('');
      setBoxNumberInput(String(coordinate.box));
      setPendingReplaceConfirm(null);
      setSubmitAttempted(false);
      setActiveSuggestionIdx(0);
      setShowSuggestions(Boolean(state.spreadsheet));
      const timer = setTimeout(() => {
        inputRef.current?.focus();
      }, 30);
      return () => clearTimeout(timer);
    }
  }, [coordinate, state.spreadsheet]);

  useEffect(() => {
    if (!coordinate) return;
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.preventDefault();
        onClose();
      }
    };
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [coordinate, onClose]);

  const suggestions: SpreadsheetItem[] = useMemo(() => {
    if (!state.spreadsheet) return [];
    return searchOldLocations(state.spreadsheet, inputValue, 12);
  }, [state.spreadsheet, inputValue]);

  const placedLookup = useMemo(() => {
    const map = new Map<string, PlacementRecord>();
    for (const p of getAllTrackedPlacements(state)) {
      map.set(normalizeLocationCode(p.oldLocation), p);
    }
    return map;
  }, [state]);

  const liveValidation = useMemo(() => {
    if (!effectiveCoordinate) return null;
    if (!inputValue.trim() && !submitAttempted) return null;
    return validatePlacementInput(state, effectiveCoordinate, inputValue);
  }, [effectiveCoordinate, state, inputValue, submitAttempted]);

  if (!effectiveCoordinate) return null;

  const handleSelectSuggestion = (item: SpreadsheetItem) => {
    setInputValue(item.oldLocation);
    setPendingReplaceConfirm(null);
    setSubmitAttempted(false);
    setShowSuggestions(false);
    inputRef.current?.focus();
  };

  const handlePrimarySubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setSubmitAttempted(true);
    const validation = validatePlacementInput(
      state,
      effectiveCoordinate,
      inputValue
    );

    if (
      validation.status === 'empty' ||
      validation.status === 'not_in_spreadsheet' ||
      validation.status === 'duplicate_old_location' ||
      validation.status === 'same_value_in_box'
    ) {
      setPendingReplaceConfirm(null);
      return;
    }

    if (validation.status === 'needs_replace_confirmation') {
      if (pendingReplaceConfirm === validation.normalizedOldLocation) {
        onSubmitPlacement(
          effectiveCoordinate,
          validation.normalizedOldLocation,
          true
        );
        onClose();
      } else {
        setPendingReplaceConfirm(validation.normalizedOldLocation);
      }
      return;
    }

    onSubmitPlacement(
      effectiveCoordinate,
      validation.normalizedOldLocation,
      false
    );
    onClose();
  };

  const handleInputKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!showSuggestions || suggestions.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setActiveSuggestionIdx((prev) => (prev + 1) % suggestions.length);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setActiveSuggestionIdx((prev) =>
        prev <= 0 ? suggestions.length - 1 : prev - 1
      );
    } else if (e.key === 'Tab' && suggestions[activeSuggestionIdx]) {
      e.preventDefault();
      handleSelectSuggestion(suggestions[activeSuggestionIdx]);
    } else if (
      e.key === 'Enter' &&
      suggestions[activeSuggestionIdx] &&
      normalizeLocationCode(inputValue) !==
        suggestions[activeSuggestionIdx].normalizedKey
    ) {
      e.preventDefault();
      handleSelectSuggestion(suggestions[activeSuggestionIdx]);
    }
  };

  const showNotFoundNotice =
    Boolean(state.spreadsheet) &&
    inputValue.trim().length > 0 &&
    suggestions.length === 0;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-slate-900/45 p-4 backdrop-blur-[2px]"
      role="presentation"
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="modal-box-input-title"
        className="w-full max-w-xl rounded-xl border border-slate-200 bg-white shadow-2xl"
      >
        {/* Modal Header */}
        <div className="flex items-start justify-between gap-4 border-b border-slate-200 px-6 py-4">
          <div>
            <h2
              id="modal-box-input-title"
              className="text-base font-bold tracking-tight text-slate-900"
            >
              INPUT LOKASI SIMPAN LAMA
            </h2>
            <p className="mt-0.5 text-xs text-slate-500">
              {formatBoxAddress(effectiveCoordinate)}
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            aria-label="Tutup dialog input lokasi simpan lama"
            className="rounded-lg p-1.5 text-slate-500 transition-colors hover:bg-slate-100 hover:text-slate-900 focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[#0F4C81]"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        {/* Selected Box Coordinates Summary (Nomor Box editable!) */}
        <div className="border-b border-slate-200 bg-slate-50/80 px-6 py-4">
          <dl className="grid grid-cols-2 gap-x-5 gap-y-3 sm:grid-cols-5 sm:items-end">
            <div>
              <dt className="text-xs text-slate-500">Lokasi Simpan</dt>
              <dd className="mt-0.5 font-mono text-base font-semibold text-slate-900">
                {sanitizeLocationCode(effectiveCoordinate.locationCode)}
              </dd>
            </div>
            <div>
              <dt className="text-xs text-slate-500">Nomor Rak</dt>
              <dd className="mt-0.5 font-mono text-base font-semibold text-slate-900 tabular-nums">
                {effectiveCoordinate.rack}
              </dd>
            </div>
            <div>
              <dt className="text-xs text-slate-500">Nomor Row</dt>
              <dd className="mt-0.5 font-mono text-base font-semibold text-slate-900 tabular-nums">
                {effectiveCoordinate.row}
              </dd>
            </div>
            <div>
              <dt className="text-xs text-slate-500">
                <label htmlFor="modal-box-number-input">Nomor Box</label>
              </dt>
              <dd className="mt-0.5">
                <input
                  id="modal-box-number-input"
                  type="text"
                  inputMode="numeric"
                  value={boxNumberInput}
                  onChange={(e) => {
                    const digits = e.target.value.replace(/\D/g, '').slice(0, 5);
                    setBoxNumberInput(digits);
                    setPendingReplaceConfirm(null);
                  }}
                  onBlur={() => {
                    if (!boxNumberInput.trim() && coordinate) {
                      setBoxNumberInput(String(coordinate.box));
                    }
                  }}
                  className="w-20 rounded-md border border-slate-300 bg-white px-2 py-0.5 font-mono text-sm font-semibold text-slate-900 tabular-nums focus:border-[#0F4C81] focus:outline-none"
                />
              </dd>
            </div>
            <div className="col-span-2 sm:col-span-1">
              <dt className="text-xs text-slate-500">Kode Lokasi Baru</dt>
              <dd
                data-testid="modal-new-location-code"
                className="mt-0.5 font-mono text-base font-bold text-[#0F4C81] tabular-nums"
              >
                {newLocationCode}
              </dd>
            </div>
          </dl>
        </div>

        {/* Form Body */}
        <form onSubmit={handlePrimarySubmit} className="px-6 py-5">
          {/* Current Occupant Notice */}
          {currentPlacement && (
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3 rounded-lg border border-emerald-200 bg-emerald-50/70 px-4 py-3">
              <div className="min-w-0">
                <p className="text-xs font-medium text-emerald-900">
                  Box ini sudah terisi oleh lokasi simpan lama:
                </p>
                <p className="mt-0.5 font-mono text-sm font-bold text-emerald-950 tabular-nums">
                  {currentPlacement.oldLocation}
                  <span className="mx-2 font-sans font-normal text-emerald-700">
                    →
                  </span>
                  {currentPlacement.newLocation}
                </p>
              </div>
              <button
                type="button"
                onClick={() => {
                  onRemovePlacement(effectiveCoordinate);
                  onClose();
                }}
                className="inline-flex items-center gap-1.5 rounded-lg border border-red-200 bg-white px-3 py-1.5 text-xs font-medium text-red-700 transition-colors hover:bg-red-50 whitespace-nowrap shrink-0"
              >
                <Trash2 className="h-3.5 w-3.5" />
                Hapus dari Box
              </button>
            </div>
          )}

          {/* Input Field with Search / Autocomplete */}
          <div>
            <div className="flex items-center justify-between gap-2">
              <label
                htmlFor="old-location-input"
                className="block text-sm font-semibold text-slate-800"
              >
                {currentPlacement
                  ? 'Ganti dengan Lokasi Simpan Lama Baru'
                  : 'Lokasi Simpan Lama'}
              </label>
              {state.spreadsheet ? (
                <span className="text-xs text-slate-500 tabular-nums">
                  Sumber: {state.spreadsheet.fileName} ({state.spreadsheet.items.length}{' '}
                  data)
                </span>
              ) : (
                <button
                  type="button"
                  onClick={() => {
                    onClose();
                    onOpenSpreadsheetTab();
                  }}
                  className="inline-flex items-center gap-1 text-xs font-medium text-[#0F4C81] hover:underline whitespace-nowrap shrink-0"
                >
                  <FileSpreadsheet className="h-3.5 w-3.5" />
                  Impor spreadsheet
                </button>
              )}
            </div>

            <div className="relative mt-2">
              <Search
                aria-hidden="true"
                className="pointer-events-none absolute top-1/2 left-3.5 h-4 w-4 -translate-y-1/2 text-slate-400"
              />
              <input
                ref={inputRef}
                id="old-location-input"
                type="text"
                autoComplete="off"
                value={inputValue}
                onChange={(e) => {
                  setInputValue(e.target.value);
                  setPendingReplaceConfirm(null);
                  setSubmitAttempted(false);
                  setActiveSuggestionIdx(0);
                  if (state.spreadsheet) {
                    setShowSuggestions(true);
                  }
                }}
                onFocus={() => {
                  if (state.spreadsheet) setShowSuggestions(true);
                }}
                onKeyDown={handleInputKeyDown}
                placeholder={
                  state.spreadsheet
                    ? 'Ketik kode atau angka lokasi lama dari spreadsheet (mis. 20)...'
                    : 'Ketik kode lokasi simpan lama secara langsung (mis. 201)...'
                }
                className="w-full rounded-lg border border-slate-300 bg-white py-2.5 pr-4 pl-10 font-mono text-sm text-slate-900 placeholder:font-sans placeholder:text-slate-400 focus:border-[#0F4C81] focus:ring-2 focus:ring-[#0F4C81]/20 focus:outline-none"
              />
            </div>

            {/* Live Mapping Preview */}
            {inputValue.trim() && (
              <div className="mt-2.5 flex items-center gap-2 text-xs text-slate-600">
                <span>Pemetaan:</span>
                <span className="font-mono font-semibold text-slate-900 tabular-nums">
                  {inputValue.trim()}
                </span>
                <ArrowRight className="h-3.5 w-3.5 text-slate-400" aria-hidden="true" />
                <span className="font-mono font-semibold text-[#0F4C81] tabular-nums">
                  {newLocationCode}
                </span>
              </div>
            )}

            {/* Autocomplete List from Imported Spreadsheet */}
            {state.spreadsheet && showSuggestions && suggestions.length > 0 && (
              <div
                role="listbox"
                aria-label="Daftar lokasi simpan lama dari spreadsheet"
                className="mt-2 max-h-48 overflow-y-auto rounded-lg border border-slate-200 bg-white divide-y divide-slate-100"
              >
                {suggestions.map((item, idx) => {
                  const placed = placedLookup.get(item.normalizedKey);
                  const isActive = idx === activeSuggestionIdx;
                  return (
                    <button
                      key={`${item.normalizedKey}-${item.rowNumber}`}
                      type="button"
                      role="option"
                      aria-selected={isActive}
                      onClick={() => handleSelectSuggestion(item)}
                      className={`flex w-full items-center justify-between gap-3 px-3.5 py-2 text-left text-sm transition-colors ${
                        isActive ? 'bg-sky-50' : 'hover:bg-slate-50'
                      }`}
                    >
                      <div className="min-w-0">
                        <span className="font-mono font-semibold text-slate-900 tabular-nums">
                          {item.oldLocation}
                        </span>
                        {item.description && (
                          <span className="ml-2 text-xs text-slate-500">
                            · {item.description}
                          </span>
                        )}
                      </div>
                      <span className="shrink-0 text-xs tabular-nums">
                        {placed ? (
                          <span className="font-medium text-amber-700">
                            Sudah di {placed.newLocation}
                          </span>
                        ) : (
                          <span className="text-emerald-700">Tersedia</span>
                        )}
                      </span>
                    </button>
                  );
                })}
              </div>
            )}

            {/* Not-Found Notice when Spreadsheet is Loaded */}
            {showNotFoundNotice && (
              <div
                role="status"
                className="mt-2.5 flex items-start gap-2.5 rounded-lg border border-amber-200 bg-amber-50 px-3.5 py-2.5 text-xs text-amber-900"
              >
                <AlertTriangle className="mt-0.5 h-4 w-4 shrink-0 text-amber-600" />
                <div>
                  <p className="font-semibold">
                    Data lokasi simpan lama "{inputValue.trim()}" tidak tersedia.
                  </p>
                  <p className="mt-0.5 text-amber-800">
                    {state.allowUnlistedWithSpreadsheet
                      ? 'Mode input di luar spreadsheet aktif. Anda tetap dapat memindahkan kode ini.'
                      : 'Tidak ditemukan kecocokan pada spreadsheet yang diimpor. Periksa kode atau aktifkan izin input manual pada tab Impor Spreadsheet.'}
                  </p>
                </div>
              </div>
            )}

            {/* Duplicate / Validation Error Messages */}
            {liveValidation &&
              (liveValidation.status === 'duplicate_old_location' ||
                liveValidation.status === 'same_value_in_box' ||
                liveValidation.status === 'empty' ||
                (liveValidation.status === 'not_in_spreadsheet' &&
                  !showNotFoundNotice)) && (
                <div
                  role="alert"
                  className="mt-3 flex items-start justify-between gap-3 rounded-lg border border-red-200 bg-red-50 px-3.5 py-2.5 text-xs text-red-800"
                >
                  <div className="flex items-start gap-2">
                    <AlertCircle className="mt-0.5 h-4 w-4 shrink-0 text-red-600" />
                    <p className="leading-relaxed">{liveValidation.message}</p>
                  </div>
                  {liveValidation.status === 'duplicate_old_location' && (
                    <button
                      type="button"
                      onClick={() => {
                        const existing = getPlacementByOldLocation(
                          getAllTrackedPlacements(state),
                          inputValue
                        );
                        if (existing) {
                          onJumpToBox({
                            locationCode: existing.locationCode,
                            rack: existing.rack,
                            row: existing.row,
                            box: existing.box,
                          });
                        }
                      }}
                      className="shrink-0 font-semibold text-red-900 underline hover:text-red-700 whitespace-nowrap"
                    >
                      Buka Box
                    </button>
                  )}
                </div>
              )}

            {/* Explicit Replace Confirmation Box */}
            {pendingReplaceConfirm && currentPlacement && (
              <div
                role="alert"
                className="mt-3 rounded-lg border border-amber-300 bg-amber-50 px-4 py-3 text-xs text-amber-950"
              >
                <p className="font-semibold text-amber-950">
                  Konfirmasi Penggantian Isi Box
                </p>
                <p className="mt-1 leading-relaxed text-amber-900">
                  Box{' '}
                  <span className="font-mono font-semibold">
                    {currentPlacement.newLocation}
                  </span>{' '}
                  saat ini ditempati oleh lokasi lama{' '}
                  <span className="font-mono font-semibold">
                    {currentPlacement.oldLocation}
                  </span>
                  . Ganti dengan{' '}
                  <span className="font-mono font-semibold">
                    {pendingReplaceConfirm}
                  </span>
                  ?
                </p>
              </div>
            )}
          </div>

          {/* Modal Footer Actions */}
          <div className="mt-6 flex items-center justify-end gap-3 border-t border-slate-100 pt-4">
            <button
              type="button"
              onClick={onClose}
              className="rounded-lg border border-slate-300 bg-white px-4 py-2 text-sm font-medium text-slate-700 transition-colors hover:bg-slate-50 whitespace-nowrap shrink-0"
            >
              Batalkan
            </button>
            <button
              type="submit"
              className="inline-flex items-center gap-2 rounded-lg bg-[#0F4C81] px-4 py-2 text-sm font-semibold text-white transition-colors hover:bg-[#0B3A63] focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-[#0F4C81] whitespace-nowrap shrink-0"
            >
              <Check className="h-4 w-4" />
              {pendingReplaceConfirm
                ? 'Konfirmasi Ganti Isi Box'
                : 'Pindahkan ke Box Ini'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

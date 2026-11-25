import type {
  BoxCoordinate,
  ImportedSpreadsheet,
  PlacementRecord,
  PlacementValidationResult,
  RackConfig,
  SimulationMetrics,
  SimulationState,
  SpreadsheetItem,
} from '../types/archive';

export const DEFAULT_RACK_CONFIG: RackConfig = {
  prefix: 'RC2',
  rackCount: 1,
  rowsPerRack: 5,
  boxesPerRow: 10,
};

export function createInitialSimulationState(
  config: RackConfig = DEFAULT_RACK_CONFIG
): SimulationState {
  return {
    config,
    draftPlacements: [],
    savedRecords: [],
    nextSequence: 1,
    history: [],
    spreadsheet: null,
    allowUnlistedWithSpreadsheet: false,
    includeExtraExcelColumns: false,
  };
}

export function normalizeLocationCode(value: string): string {
  return value.trim().replace(/\s+/g, ' ').toUpperCase();
}

export function sanitizeLocationCode(value: string): string {
  return value.trim().replace(/\s+/g, '').toUpperCase() || 'RC2';
}

export function formatNewLocationCode(
  locationCode: string,
  rack: number,
  row: number,
  box: number
): string {
  const cleanLoc = sanitizeLocationCode(locationCode);
  return `${cleanLoc}.${rack}.${row}.${box}`;
}

export function computeBoxNumber(
  row: number,
  slotIndex: number,
  boxesPerRow: number,
  startBoxNumber: number,
  continuousAcrossRows: boolean
): number {
  const safeStart = Math.max(1, Math.floor(startBoxNumber) || 1);
  if (continuousAcrossRows) {
    return safeStart + (row - 1) * boxesPerRow + slotIndex;
  }
  return safeStart + slotIndex;
}

export function formatBoxAddress(coord: BoxCoordinate): string {
  const cleanLoc = sanitizeLocationCode(coord.locationCode);
  return `${cleanLoc} · RAK ${coord.rack} · ROW ${coord.row} · BOX ${coord.box}`;
}

export function getPlacementByCoordinate(
  placements: PlacementRecord[],
  coord: BoxCoordinate
): PlacementRecord | undefined {
  const cleanLoc = sanitizeLocationCode(coord.locationCode);
  return placements.find(
    (p) =>
      sanitizeLocationCode(p.locationCode) === cleanLoc &&
      p.rack === coord.rack &&
      p.row === coord.row &&
      p.box === coord.box
  );
}

export function getPlacementByOldLocation(
  placements: PlacementRecord[],
  oldLocation: string
): PlacementRecord | undefined {
  const target = normalizeLocationCode(oldLocation);
  if (!target) return undefined;
  return placements.find((p) => normalizeLocationCode(p.oldLocation) === target);
}

export function getAllTrackedPlacements(
  state: SimulationState
): PlacementRecord[] {
  const seenIds = new Set<string>();
  const combined: PlacementRecord[] = [];
  for (const r of state.savedRecords) {
    seenIds.add(r.id);
    combined.push(r);
  }
  for (const d of state.draftPlacements) {
    if (!seenIds.has(d.id)) {
      combined.push(d);
    }
  }
  return combined;
}

export function findMatchingSpreadsheetItem(
  spreadsheet: ImportedSpreadsheet | null,
  input: string
): SpreadsheetItem | undefined {
  if (!spreadsheet) return undefined;
  const target = normalizeLocationCode(input);
  if (!target) return undefined;
  return spreadsheet.items.find((item) => item.normalizedKey === target);
}

export function searchOldLocations(
  spreadsheet: ImportedSpreadsheet | null,
  query: string,
  limit = 25
): SpreadsheetItem[] {
  if (!spreadsheet) return [];
  const trimmed = query.trim().toLowerCase();
  if (!trimmed) {
    return spreadsheet.items.slice(0, limit);
  }
  const exactOrStarts: SpreadsheetItem[] = [];
  const contains: SpreadsheetItem[] = [];

  for (const item of spreadsheet.items) {
    const val = item.oldLocation.toLowerCase();
    if (val === trimmed || val.startsWith(trimmed)) {
      exactOrStarts.push(item);
    } else if (val.includes(trimmed)) {
      contains.push(item);
    }
  }

  return [...exactOrStarts, ...contains].slice(0, limit);
}

export function validatePlacementInput(
  state: SimulationState,
  coord: BoxCoordinate,
  rawOldLocation: string
): PlacementValidationResult {
  const trimmed = rawOldLocation.trim();
  if (!trimmed) {
    return {
      status: 'empty',
      message: 'Masukkan nomor atau kode lokasi simpan lama terlebih dahulu.',
    };
  }

  let resolvedOldLocation = trimmed;
  if (state.spreadsheet) {
    const matched = findMatchingSpreadsheetItem(state.spreadsheet, trimmed);
    if (!matched && !state.allowUnlistedWithSpreadsheet) {
      return {
        status: 'not_in_spreadsheet',
        message: `Data lokasi simpan lama "${trimmed}" tidak tersedia pada spreadsheet yang diimpor.`,
      };
    }
    if (matched) {
      resolvedOldLocation = matched.oldLocation;
    }
  }

  const currentInBox = getPlacementByCoordinate(state.draftPlacements, coord);
  const existingInDraft = getPlacementByOldLocation(
    state.draftPlacements,
    resolvedOldLocation
  );
  const existingInSaved = getPlacementByOldLocation(
    state.savedRecords,
    resolvedOldLocation
  );

  const cleanCoordLoc = sanitizeLocationCode(coord.locationCode);

  if (existingInDraft) {
    const isSameBox =
      sanitizeLocationCode(existingInDraft.locationCode) === cleanCoordLoc &&
      existingInDraft.rack === coord.rack &&
      existingInDraft.row === coord.row &&
      existingInDraft.box === coord.box;

    if (isSameBox) {
      return {
        status: 'same_value_in_box',
        message: `Lokasi simpan lama "${resolvedOldLocation}" sudah ditempatkan pada box ini (${existingInDraft.newLocation}).`,
        existingPlacement: existingInDraft,
      };
    }

    return {
      status: 'duplicate_old_location',
      message: `Lokasi simpan lama "${resolvedOldLocation}" sudah ditempatkan pada ${formatBoxAddress(existingInDraft)} (${existingInDraft.newLocation}).`,
      existingPlacement: existingInDraft,
      inSavedExcel: Boolean(existingInDraft.savedToExcel),
    };
  }

  if (existingInSaved) {
    return {
      status: 'duplicate_old_location',
      message: `Lokasi simpan lama "${resolvedOldLocation}" sudah tersimpan di Excel pada ${formatBoxAddress(existingInSaved)} (${existingInSaved.newLocation}).`,
      existingPlacement: existingInSaved,
      inSavedExcel: true,
    };
  }

  if (currentInBox) {
    return {
      status: 'needs_replace_confirmation',
      message: `Box ${currentInBox.newLocation} saat ini terisi oleh lokasi lama "${currentInBox.oldLocation}". Yakin ingin menggantinya dengan "${resolvedOldLocation}"?`,
      currentPlacement: currentInBox,
      normalizedOldLocation: resolvedOldLocation,
    };
  }

  return {
    status: 'valid',
    normalizedOldLocation: resolvedOldLocation,
  };
}

function pushHistory(state: SimulationState): SimulationState['history'] {
  const snapshot = {
    draftPlacements: state.draftPlacements.map((p) => ({ ...p })),
    savedRecords: state.savedRecords.map((p) => ({ ...p })),
    nextSequence: state.nextSequence,
  };
  return [...state.history.slice(-49), snapshot];
}

export function applyPlacement(
  state: SimulationState,
  coord: BoxCoordinate,
  rawOldLocation: string,
  options?: { confirmReplace?: boolean; timestamp?: string }
): {
  nextState: SimulationState;
  result: PlacementValidationResult | { status: 'placed'; record: PlacementRecord };
} {
  const validation = validatePlacementInput(state, coord, rawOldLocation);

  if (
    validation.status === 'empty' ||
    validation.status === 'not_in_spreadsheet' ||
    validation.status === 'duplicate_old_location' ||
    validation.status === 'same_value_in_box'
  ) {
    return { nextState: state, result: validation };
  }

  if (
    validation.status === 'needs_replace_confirmation' &&
    !options?.confirmReplace
  ) {
    return { nextState: state, result: validation };
  }

  const finalOldLocation = validation.normalizedOldLocation;
  const cleanLoc = sanitizeLocationCode(coord.locationCode);
  const newLocation = formatNewLocationCode(
    cleanLoc,
    coord.rack,
    coord.row,
    coord.box
  );
  const placedAt = options?.timestamp ?? new Date().toISOString();
  const nextHistory = pushHistory(state);

  if (validation.status === 'needs_replace_confirmation') {
    const existing = validation.currentPlacement;
    const updatedRecord: PlacementRecord = {
      ...existing,
      oldLocation: finalOldLocation,
      locationCode: cleanLoc,
      newLocation,
      placedAt,
      savedToExcel: false,
    };
    const nextDrafts = state.draftPlacements.map((p) =>
      p.id === existing.id ? updatedRecord : p
    );
    // If the old draft was already committed to savedRecords with the same id, update it or leave saved intact?
    // Keep savedRecords synchronized if this exact id was already saved in this cycle!
    const nextSaved = existing.savedToExcel
      ? state.savedRecords.map((p) =>
          p.id === existing.id
            ? { ...updatedRecord, savedToExcel: true }
            : p
        )
      : state.savedRecords;

    const finalDrafts = existing.savedToExcel
      ? nextDrafts.map((p) =>
          p.id === existing.id ? { ...updatedRecord, savedToExcel: true } : p
        )
      : nextDrafts;

    return {
      nextState: {
        ...state,
        draftPlacements: finalDrafts,
        savedRecords: nextSaved,
        history: nextHistory,
      },
      result: { status: 'placed', record: updatedRecord },
    };
  }

  const record: PlacementRecord = {
    id: `${cleanLoc}-${coord.rack}-${coord.row}-${coord.box}-${state.nextSequence}`,
    oldLocation: finalOldLocation,
    newLocation,
    locationCode: cleanLoc,
    rack: coord.rack,
    row: coord.row,
    box: coord.box,
    sequence: state.nextSequence,
    placedAt,
    savedToExcel: false,
  };

  return {
    nextState: {
      ...state,
      draftPlacements: [...state.draftPlacements, record],
      nextSequence: state.nextSequence + 1,
      history: nextHistory,
    },
    result: { status: 'placed', record },
  };
}

/**
 * Commits all unsaved placements from the active rack simulation into savedRecords (Excel),
 * appending strictly below previous cycles in placement order.
 */
export function commitDraftToExcel(state: SimulationState): {
  nextState: SimulationState;
  addedCount: number;
} {
  const unsaved = state.draftPlacements
    .filter((p) => !p.savedToExcel)
    .sort((a, b) => a.sequence - b.sequence);

  if (unsaved.length === 0) {
    return { nextState: state, addedCount: 0 };
  }

  const nextHistory = pushHistory(state);
  let nextSeq = state.nextSequence;
  const committedMap = new Map<string, PlacementRecord>();
  const newlySaved: PlacementRecord[] = [];

  for (const item of unsaved) {
    const savedItem: PlacementRecord = {
      ...item,
      sequence: nextSeq++,
      savedToExcel: true,
    };
    committedMap.set(item.id, savedItem);
    newlySaved.push(savedItem);
  }

  const nextDrafts = state.draftPlacements.map((p) =>
    committedMap.has(p.id) ? committedMap.get(p.id)! : p
  );

  return {
    nextState: {
      ...state,
      draftPlacements: nextDrafts,
      savedRecords: [...state.savedRecords, ...newlySaved],
      nextSequence: nextSeq,
      history: nextHistory,
    },
    addedCount: newlySaved.length,
  };
}

/**
 * Resets ONLY the active 1-rack simulation (draftPlacements).
 * All previously saved Excel records (savedRecords) remain 100% intact!
 */
export function resetActiveRackSimulation(
  state: SimulationState
): SimulationState {
  if (state.draftPlacements.length === 0) return state;
  return {
    ...state,
    draftPlacements: [],
    history: pushHistory(state),
  };
}

/**
 * Clears all saved Excel records (separate explicit action inside Excel table).
 */
export function clearSavedExcelRecords(
  state: SimulationState
): SimulationState {
  if (state.savedRecords.length === 0) return state;
  const nextDrafts = state.draftPlacements.map((d) => ({
    ...d,
    savedToExcel: false,
  }));
  return {
    ...state,
    draftPlacements: nextDrafts,
    savedRecords: [],
    history: pushHistory(state),
  };
}

export function removeDraftByCoordinate(
  state: SimulationState,
  coord: BoxCoordinate
): SimulationState {
  const existing = getPlacementByCoordinate(state.draftPlacements, coord);
  if (!existing) return state;
  return {
    ...state,
    draftPlacements: state.draftPlacements.filter((p) => p.id !== existing.id),
    history: pushHistory(state),
  };
}

export function removeSavedRecordById(
  state: SimulationState,
  placementId: string
): SimulationState {
  const exists = state.savedRecords.some((p) => p.id === placementId);
  if (!exists) return state;

  return {
    ...state,
    savedRecords: state.savedRecords.filter((p) => p.id !== placementId),
    draftPlacements: state.draftPlacements.filter((p) => p.id !== placementId),
    history: pushHistory(state),
  };
}

export function undoLastAction(state: SimulationState): SimulationState {
  if (state.history.length === 0) return state;
  const previous = state.history[state.history.length - 1];
  return {
    ...state,
    draftPlacements: previous.draftPlacements,
    savedRecords: previous.savedRecords,
    nextSequence: previous.nextSequence,
    history: state.history.slice(0, -1),
  };
}

export function updateRackConfig(
  state: SimulationState,
  nextConfig: RackConfig
): SimulationState {
  const sanitized: RackConfig = {
    prefix: sanitizeLocationCode(nextConfig.prefix),
    rackCount: 1,
    rowsPerRack: Math.max(1, Math.min(20, Math.floor(nextConfig.rowsPerRack) || 1)),
    boxesPerRow: Math.max(1, Math.min(30, Math.floor(nextConfig.boxesPerRow) || 1)),
  };

  return {
    ...state,
    config: sanitized,
  };
}

export function extractSpreadsheetItems(
  rawRows: Record<string, string>[],
  selectedColumn: string,
  descriptionColumn?: string
): SpreadsheetItem[] {
  const seen = new Set<string>();
  const items: SpreadsheetItem[] = [];

  rawRows.forEach((row, index) => {
    const rawVal = (row[selectedColumn] ?? '').toString().trim();
    if (!rawVal) return;
    const normalizedKey = normalizeLocationCode(rawVal);
    if (seen.has(normalizedKey)) return;
    seen.add(normalizedKey);

    const descVal =
      descriptionColumn && descriptionColumn !== selectedColumn
        ? (row[descriptionColumn] ?? '').toString().trim()
        : undefined;

    items.push({
      oldLocation: rawVal,
      normalizedKey,
      rowNumber: index + 2,
      description: descVal || undefined,
    });
  });

  return items;
}

export function detectBestLocationColumn(columns: string[]): string {
  if (columns.length === 0) return '';
  const patterns = [
    /format\s*lama/i,
    /lokasi\s*(simpan\s*)?lama/i,
    /lokasi\s*lama/i,
    /kode\s*lama/i,
    /no\.?\s*box/i,
    /lokasi/i,
    /boks|box/i,
  ];
  for (const pattern of patterns) {
    const match = columns.find((col) => pattern.test(col));
    if (match) return match;
  }
  return columns[0];
}

export function calculateSimulationMetrics(
  state: SimulationState,
  activeLocationCode: string,
  activeRack: number
): SimulationMetrics {
  const cleanLoc = sanitizeLocationCode(activeLocationCode);
  const rackTotalBoxes = state.config.rowsPerRack * state.config.boxesPerRow;

  const activeRackDrafts = state.draftPlacements.filter(
    (p) =>
      sanitizeLocationCode(p.locationCode) === cleanLoc &&
      p.rack === activeRack
  );

  const rackFilledBoxes = activeRackDrafts.length;
  const rackEmptyBoxes = Math.max(0, rackTotalBoxes - rackFilledBoxes);
  const unsavedDraftCount = state.draftPlacements.filter(
    (p) => !p.savedToExcel
  ).length;
  const savedExcelCount = state.savedRecords.length;

  const allTracked = getAllTrackedPlacements(state);
  const totalRelocatedCount = allTracked.length;

  if (state.spreadsheet && state.spreadsheet.items.length > 0) {
    const spreadsheetKeys = new Set(
      state.spreadsheet.items.map((item) => item.normalizedKey)
    );
    const matchedRelocated = allTracked.filter((p) =>
      spreadsheetKeys.has(normalizeLocationCode(p.oldLocation))
    ).length;
    const spreadsheetTotal = state.spreadsheet.items.length;
    const spreadsheetRemaining = Math.max(0, spreadsheetTotal - matchedRelocated);
    const progressPercent = Math.min(
      100,
      Math.round((matchedRelocated / spreadsheetTotal) * 100)
    );
    return {
      rackTotalBoxes,
      rackFilledBoxes,
      rackEmptyBoxes,
      unsavedDraftCount,
      savedExcelCount,
      totalRelocatedCount,
      spreadsheetTotal,
      spreadsheetRemaining,
      progressPercent,
      progressLabel: `${matchedRelocated} dari ${spreadsheetTotal} data spreadsheet`,
    };
  }

  const progressPercent =
    rackTotalBoxes > 0
      ? Math.min(100, Math.round((rackFilledBoxes / rackTotalBoxes) * 100))
      : 0;

  return {
    rackTotalBoxes,
    rackFilledBoxes,
    rackEmptyBoxes,
    unsavedDraftCount,
    savedExcelCount,
    totalRelocatedCount,
    spreadsheetTotal: null,
    spreadsheetRemaining: null,
    progressPercent,
    progressLabel: `${rackFilledBoxes} dari ${rackTotalBoxes} box rak aktif`,
  };
}

export function buildExcelExportRows(
  savedRecords: PlacementRecord[],
  includeExtraColumns: boolean
): Record<string, string | number>[] {
  const ordered = [...savedRecords].sort((a, b) => a.sequence - b.sequence);
  return ordered.map((item): Record<string, string | number> => {
    if (!includeExtraColumns) {
      return {
        'FORMAT LAMA': item.oldLocation,
        'FORMAT BARU': item.newLocation,
      };
    }
    return {
      'FORMAT LAMA': item.oldLocation,
      'FORMAT BARU': item.newLocation,
      RAK: item.rack,
      ROW: item.row,
      BOX: item.box,
      'WAKTU PEMINDAHAN': new Date(item.placedAt).toLocaleString('id-ID'),
    };
  });
}

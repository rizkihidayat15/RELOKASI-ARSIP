import type {
  BoxCoordinate,
  ImportedSpreadsheet,
  LockedLocationInfo,
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
    externalLockedLocations: [],
    externalResultSourceName: null,
    externalResultSourceType: null,
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

/**
 * Parses a formatted new location code (e.g. "RC2.1.1.1") back into a BoxCoordinate.
 * Returns null if the code does not match the LOKASI.RAK.ROW.BOX structure.
 */
export function parseNewLocationCode(code: string): BoxCoordinate | null {
  const trimmed = code.trim();
  const match = trimmed.match(/^([A-Za-z0-9_-]+)\.(\d+)\.(\d+)\.(\d+)$/);
  if (!match) return null;
  const rack = parseInt(match[2], 10);
  const row = parseInt(match[3], 10);
  const box = parseInt(match[4], 10);
  if (
    !Number.isFinite(rack) ||
    !Number.isFinite(row) ||
    !Number.isFinite(box) ||
    rack < 1 ||
    row < 1 ||
    box < 1
  ) {
    return null;
  }
  return {
    locationCode: sanitizeLocationCode(match[1]),
    rack,
    row,
    box,
  };
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

/**
 * Builds a unified lookup map of ALL locked (already-used) old location codes
 * across active rack drafts, saved simulation records, and synced Google Sheets rows.
 */
export function getLockedLocationsMap(
  state: SimulationState
): Map<string, LockedLocationInfo> {
  const map = new Map<string, LockedLocationInfo>();

  // Lowest precedence: external Google Sheets or loaded result Excel rows
  for (const ext of state.externalLockedLocations ?? []) {
    const key = normalizeLocationCode(ext.oldLocation);
    if (!key) continue;
    const coord = ext.coordinate ?? parseNewLocationCode(ext.newLocation) ?? undefined;
    map.set(key, {
      ...ext,
      normalizedKey: key,
      coordinate: coord,
      source: ext.source === 'external_file' ? 'external_file' : 'google_sheets',
    });
  }

  // Local saved simulation records
  for (const saved of state.savedRecords) {
    const key = normalizeLocationCode(saved.oldLocation);
    if (!key) continue;
    map.set(key, {
      oldLocation: saved.oldLocation,
      normalizedKey: key,
      newLocation: saved.newLocation,
      source: 'saved',
      coordinate: {
        locationCode: saved.locationCode,
        rack: saved.rack,
        row: saved.row,
        box: saved.box,
      },
    });
  }

  // Active rack draft placements
  for (const draft of state.draftPlacements) {
    const key = normalizeLocationCode(draft.oldLocation);
    if (!key) continue;
    map.set(key, {
      oldLocation: draft.oldLocation,
      normalizedKey: key,
      newLocation: draft.newLocation,
      source: draft.savedToExcel ? 'saved' : 'draft',
      coordinate: {
        locationCode: draft.locationCode,
        rack: draft.rack,
        row: draft.row,
        box: draft.box,
      },
    });
  }

  return map;
}

/**
 * Parses rows read from a connected Google Spreadsheet or result Excel file (FORMAT LAMA | FORMAT BARU)
 * into LockedLocationInfo entries so codes already stored in the results spreadsheet are locked and counted.
 */
export function parseExternalSheetLocks(
  rows: string[][],
  source: 'google_sheets' | 'external_file' = 'google_sheets'
): LockedLocationInfo[] {
  if (!Array.isArray(rows) || rows.length === 0) return [];
  const firstRow = rows[0] ?? [];

  // Detect FORMAT LAMA and FORMAT BARU column indices dynamically if header row exists
  let oldColIdx = 0;
  let newColIdx = 1;
  let hasHeader = false;

  const detectedOld = firstRow.findIndex((cell) =>
    /^(format\s*lama|lokasi\s*(simpan\s*)?lama|kode\s*lama)$/i.test(
      String(cell ?? '').trim()
    )
  );
  const detectedNew = firstRow.findIndex((cell) =>
    /^(format\s*baru|lokasi\s*(simpan\s*)?baru|kode\s*baru)$/i.test(
      String(cell ?? '').trim()
    )
  );

  if (detectedOld >= 0 || detectedNew >= 0) {
    hasHeader = true;
    if (detectedOld >= 0) oldColIdx = detectedOld;
    if (detectedNew >= 0) newColIdx = detectedNew;
  }

  const startIdx = hasHeader ? 1 : 0;
  const seen = new Set<string>();
  const locks: LockedLocationInfo[] = [];

  for (let i = startIdx; i < rows.length; i++) {
    const oldVal = String(rows[i]?.[oldColIdx] ?? '').trim();
    const newVal = String(rows[i]?.[newColIdx] ?? '').trim();
    if (!oldVal || !newVal) continue;
    const key = normalizeLocationCode(oldVal);
    if (seen.has(key)) continue;
    seen.add(key);
    locks.push({
      oldLocation: oldVal,
      normalizedKey: key,
      newLocation: newVal,
      source,
      coordinate: parseNewLocationCode(newVal) ?? undefined,
    });
  }

  return locks;
}

/**
 * Returns the unified, deduplicated Tabel Hasil list combining rows read from
 * Google Sheets / loaded Excel result files with local savedRecords.
 * Every row in Tabel Hasil is treated as already simulated (savedToExcel: true).
 */
export function getCombinedResultRecords(
  state: SimulationState
): PlacementRecord[] {
  const externalList = state.externalLockedLocations ?? [];
  const orderedSaved = [...state.savedRecords].sort(
    (a, b) => a.sequence - b.sequence
  );

  if (externalList.length === 0) {
    return orderedSaved;
  }

  const savedByKey = new Map<string, PlacementRecord>();
  for (const saved of orderedSaved) {
    savedByKey.set(normalizeLocationCode(saved.oldLocation), saved);
  }

  const combined: PlacementRecord[] = [];
  const includedKeys = new Set<string>();
  let seq = 1;

  for (const ext of externalList) {
    const key = normalizeLocationCode(ext.oldLocation);
    if (!key || includedKeys.has(key)) continue;
    includedKeys.add(key);

    const localOverride = savedByKey.get(key);
    if (localOverride) {
      combined.push({
        ...localOverride,
        sequence: seq++,
        savedToExcel: true,
      });
      continue;
    }

    const coord = ext.coordinate ?? parseNewLocationCode(ext.newLocation);
    combined.push({
      id: `ext-${key}`,
      oldLocation: ext.oldLocation,
      newLocation: ext.newLocation,
      locationCode: coord?.locationCode ?? sanitizeLocationCode(state.config.prefix),
      rack: coord?.rack ?? 0,
      row: coord?.row ?? 0,
      box: coord?.box ?? 0,
      sequence: seq++,
      placedAt: '',
      savedToExcel: true,
    });
  }

  for (const saved of orderedSaved) {
    const key = normalizeLocationCode(saved.oldLocation);
    if (includedKeys.has(key)) continue;
    includedKeys.add(key);
    combined.push({
      ...saved,
      sequence: seq++,
      savedToExcel: true,
    });
  }

  return combined;
}

/**
 * Returns ALL simulated boxes across racks by unifying Tabel Hasil records
 * (savedRecords + external Google Sheets/file results with valid coordinates)
 * with active rack drafts. Any box already registered in Tabel Hasil is
 * immediately reflected as simulated (savedToExcel: true) on the rack grid.
 */
export function getEffectiveRackPlacements(
  state: SimulationState
): PlacementRecord[] {
  const byCoord = new Map<string, PlacementRecord>();

  for (const record of getCombinedResultRecords(state)) {
    if (record.rack < 1 || record.row < 1 || record.box < 1) continue;
    const coordKey = `${sanitizeLocationCode(record.locationCode)}-${record.rack}-${record.row}-${record.box}`;
    byCoord.set(coordKey, { ...record, savedToExcel: true });
  }

  for (const draft of state.draftPlacements) {
    const coordKey = `${sanitizeLocationCode(draft.locationCode)}-${draft.rack}-${draft.row}-${draft.box}`;
    byCoord.set(coordKey, draft);
  }

  return Array.from(byCoord.values());
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
  limit = 25,
  options?: { excludeLockedKeys?: Set<string> }
): SpreadsheetItem[] {
  if (!spreadsheet) return [];
  const excludeSet = options?.excludeLockedKeys;
  const trimmed = query.trim().toLowerCase();

  const pool = excludeSet
    ? spreadsheet.items.filter((item) => !excludeSet.has(item.normalizedKey))
    : spreadsheet.items;

  if (!trimmed) {
    return pool.slice(0, limit);
  }
  const exactOrStarts: SpreadsheetItem[] = [];
  const containsCode: SpreadsheetItem[] = [];
  const containsDesc: SpreadsheetItem[] = [];
  const hasLetters = /[a-z]/i.test(trimmed);

  for (const item of pool) {
    const val = item.oldLocation.toLowerCase();
    const desc = (item.description ?? '').toLowerCase();
    if (val === trimmed || val.startsWith(trimmed)) {
      exactOrStarts.push(item);
    } else if (val.includes(trimmed)) {
      containsCode.push(item);
    } else if (desc && desc.includes(trimmed) && (hasLetters || trimmed.length >= 4)) {
      containsDesc.push(item);
    }
  }

  return [...exactOrStarts, ...containsCode, ...containsDesc].slice(0, limit);
}

export function searchOldLocationsWithLockStatus(
  spreadsheet: ImportedSpreadsheet | null,
  lockedMap: Map<string, LockedLocationInfo>,
  query: string,
  options?: { hideLocked?: boolean; limit?: number }
): {
  visibleItems: SpreadsheetItem[];
  matchedLockedItems: Array<{ item: SpreadsheetItem; lock: LockedLocationInfo }>;
} {
  if (!spreadsheet) {
    return { visibleItems: [], matchedLockedItems: [] };
  }
  const limit = options?.limit ?? 12;
  const hideLocked = options?.hideLocked ?? true;

  const allMatches = searchOldLocations(spreadsheet, query, Math.max(limit * 3, 40));
  const matchedLockedItems: Array<{
    item: SpreadsheetItem;
    lock: LockedLocationInfo;
  }> = [];
  const visibleItems: SpreadsheetItem[] = [];

  for (const item of allMatches) {
    const lock = lockedMap.get(item.normalizedKey);
    if (lock) {
      matchedLockedItems.push({ item, lock });
      if (!hideLocked && visibleItems.length < limit) {
        visibleItems.push(item);
      }
    } else if (visibleItems.length < limit) {
      visibleItems.push(item);
    }
  }

  return { visibleItems, matchedLockedItems };
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

  const effectivePlacements = getEffectiveRackPlacements(state);
  const currentInBox = getPlacementByCoordinate(effectivePlacements, coord);
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
      message: `Kode "${resolvedOldLocation}" terkunci karena sudah ditempatkan pada ${formatBoxAddress(existingInDraft)} (${existingInDraft.newLocation}).`,
      existingPlacement: existingInDraft,
      lockedNewLocation: existingInDraft.newLocation,
      inSavedExcel: Boolean(existingInDraft.savedToExcel),
    };
  }

  if (existingInSaved) {
    const isSameBox =
      sanitizeLocationCode(existingInSaved.locationCode) === cleanCoordLoc &&
      existingInSaved.rack === coord.rack &&
      existingInSaved.row === coord.row &&
      existingInSaved.box === coord.box;

    if (isSameBox) {
      return {
        status: 'same_value_in_box',
        message: `Lokasi simpan lama "${resolvedOldLocation}" sudah tersimulasi & terdaftar di Tabel Hasil pada box ini (${existingInSaved.newLocation}).`,
        existingPlacement: existingInSaved,
      };
    }

    return {
      status: 'duplicate_old_location',
      message: `Kode "${resolvedOldLocation}" terkunci karena telah digunakan menjadi lokasi baru ${existingInSaved.newLocation} (${formatBoxAddress(existingInSaved)}).`,
      existingPlacement: existingInSaved,
      lockedNewLocation: existingInSaved.newLocation,
      inSavedExcel: true,
    };
  }

  const normalizedKey = normalizeLocationCode(resolvedOldLocation);
  const existingInExternal = (state.externalLockedLocations ?? []).find(
    (item) => normalizeLocationCode(item.oldLocation) === normalizedKey
  );
  if (existingInExternal) {
    const extCoord =
      existingInExternal.coordinate ??
      parseNewLocationCode(existingInExternal.newLocation);
    const isSameBox =
      extCoord !== null &&
      sanitizeLocationCode(extCoord.locationCode) === cleanCoordLoc &&
      extCoord.rack === coord.rack &&
      extCoord.row === coord.row &&
      extCoord.box === coord.box;

    if (isSameBox && currentInBox) {
      return {
        status: 'same_value_in_box',
        message: `Lokasi simpan lama "${resolvedOldLocation}" sudah tersimulasi & terdaftar di Tabel Hasil pada box ini (${existingInExternal.newLocation}).`,
        existingPlacement: currentInBox,
      };
    }

    const sourceText =
      existingInExternal.source === 'external_file'
        ? 'file Spreadsheet Hasil'
        : 'Google Sheets';
    return {
      status: 'duplicate_old_location',
      message: `Kode "${resolvedOldLocation}" terkunci karena telah terdaftar pada ${sourceText} menjadi lokasi baru ${existingInExternal.newLocation}.`,
      lockedNewLocation: existingInExternal.newLocation,
      inSavedExcel: true,
      inGoogleSheets: true,
    };
  }

  if (currentInBox) {
    const statusDesc = currentInBox.savedToExcel
      ? 'sudah tersimulasi & terdaftar di Tabel Hasil'
      : 'saat ini terisi';
    return {
      status: 'needs_replace_confirmation',
      message: `Box ${currentInBox.newLocation} ${statusDesc} oleh lokasi lama "${currentInBox.oldLocation}". Yakin ingin menggantinya dengan "${resolvedOldLocation}"?`,
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
    const isExternalEntry = existing.id.startsWith('ext-');
    const inDrafts = state.draftPlacements.some((p) => p.id === existing.id);
    const inSaved = state.savedRecords.some((p) => p.id === existing.id);

    const updatedRecord: PlacementRecord = {
      ...existing,
      id: isExternalEntry
        ? `${cleanLoc}-${coord.rack}-${coord.row}-${coord.box}-${state.nextSequence}`
        : existing.id,
      oldLocation: finalOldLocation,
      locationCode: cleanLoc,
      newLocation,
      placedAt,
      savedToExcel: inSaved,
    };

    const nextDrafts = inDrafts
      ? state.draftPlacements.map((p) =>
          p.id === existing.id ? updatedRecord : p
        )
      : [...state.draftPlacements, updatedRecord];

    const nextSaved = inSaved
      ? state.savedRecords.map((p) =>
          p.id === existing.id
            ? { ...updatedRecord, savedToExcel: true }
            : p
        )
      : state.savedRecords;

    const existingKey = normalizeLocationCode(existing.oldLocation);
    const nextExternal = (state.externalLockedLocations ?? []).filter(
      (ext) => normalizeLocationCode(ext.oldLocation) !== existingKey
    );

    return {
      nextState: {
        ...state,
        draftPlacements: nextDrafts,
        savedRecords: nextSaved,
        externalLockedLocations: nextExternal,
        nextSequence: isExternalEntry ? state.nextSequence + 1 : state.nextSequence,
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
  if (
    state.savedRecords.length === 0 &&
    (state.externalLockedLocations ?? []).length === 0
  ) {
    return state;
  }
  const nextDrafts = state.draftPlacements.map((d) => ({
    ...d,
    savedToExcel: false,
  }));
  return {
    ...state,
    draftPlacements: nextDrafts,
    savedRecords: [],
    externalLockedLocations: [],
    externalResultSourceName: null,
    externalResultSourceType: null,
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

/**
 * Removes any placement occupying the given box coordinate across active drafts,
 * local savedRecords, and external results (Tabel Hasil), freeing both the box
 * and the old location code.
 */
export function removePlacementByCoordinate(
  state: SimulationState,
  coord: BoxCoordinate
): SimulationState {
  const cleanLoc = sanitizeLocationCode(coord.locationCode);
  const matchesCoord = (p: {
    locationCode: string;
    rack: number;
    row: number;
    box: number;
  }) =>
    sanitizeLocationCode(p.locationCode) === cleanLoc &&
    p.rack === coord.rack &&
    p.row === coord.row &&
    p.box === coord.box;

  const nextDrafts = state.draftPlacements.filter((p) => !matchesCoord(p));
  const nextSaved = state.savedRecords.filter((p) => !matchesCoord(p));
  const nextExternal = (state.externalLockedLocations ?? []).filter((ext) => {
    const extCoord = ext.coordinate ?? parseNewLocationCode(ext.newLocation);
    return !extCoord || !matchesCoord(extCoord);
  });

  if (
    nextDrafts.length === state.draftPlacements.length &&
    nextSaved.length === state.savedRecords.length &&
    nextExternal.length === (state.externalLockedLocations ?? []).length
  ) {
    return state;
  }

  return {
    ...state,
    draftPlacements: nextDrafts,
    savedRecords: nextSaved,
    externalLockedLocations: nextExternal,
    history: pushHistory(state),
  };
}

export function removeSavedRecordById(
  state: SimulationState,
  placementId: string
): SimulationState {
  if (placementId.startsWith('ext-')) {
    const targetKey = placementId.slice(4);
    const nextExternal = (state.externalLockedLocations ?? []).filter(
      (ext) => normalizeLocationCode(ext.oldLocation) !== targetKey
    );
    if (nextExternal.length === (state.externalLockedLocations ?? []).length) {
      return state;
    }
    return {
      ...state,
      externalLockedLocations: nextExternal,
      history: pushHistory(state),
    };
  }

  const target = state.savedRecords.find((p) => p.id === placementId);
  if (!target) return state;
  const targetKey = normalizeLocationCode(target.oldLocation);

  return {
    ...state,
    savedRecords: state.savedRecords.filter((p) => p.id !== placementId),
    draftPlacements: state.draftPlacements.filter((p) => p.id !== placementId),
    externalLockedLocations: (state.externalLockedLocations ?? []).filter(
      (ext) => normalizeLocationCode(ext.oldLocation) !== targetKey
    ),
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

  const effectiveRackPlacements = getEffectiveRackPlacements(state);
  const activeRackFilledList = effectiveRackPlacements.filter(
    (p) =>
      sanitizeLocationCode(p.locationCode) === cleanLoc &&
      p.rack === activeRack
  );

  const rackFilledBoxes = activeRackFilledList.length;
  const rackEmptyBoxes = Math.max(0, rackTotalBoxes - rackFilledBoxes);
  const unsavedDraftCount = state.draftPlacements.filter(
    (p) => !p.savedToExcel
  ).length;
  const savedExcelCount = state.savedRecords.length;

  // Deduplicated set of old location keys committed in Spreadsheet Hasil (Google Sheets / result Excel + local savedRecords)
  const registeredResultKeys = new Set<string>();
  const externalKeys = new Set<string>();
  for (const ext of state.externalLockedLocations ?? []) {
    const key = normalizeLocationCode(ext.oldLocation);
    if (!key) continue;
    externalKeys.add(key);
    registeredResultKeys.add(key);
  }
  for (const saved of state.savedRecords) {
    const key = normalizeLocationCode(saved.oldLocation);
    if (!key) continue;
    registeredResultKeys.add(key);
  }

  // Active draft keys not yet committed into Spreadsheet Hasil
  const draftOnlyKeys = new Set<string>();
  for (const draft of state.draftPlacements) {
    const key = normalizeLocationCode(draft.oldLocation);
    if (!key) continue;
    if (!registeredResultKeys.has(key)) {
      draftOnlyKeys.add(key);
    }
  }

  const allTracked = getAllTrackedPlacements(state);
  const lockedMap = getLockedLocationsMap(state);
  const totalRelocatedCount = Math.max(lockedMap.size, allTracked.length);
  const spreadsheetResultsTotal = registeredResultKeys.size;
  const externalResultsCount = externalKeys.size;

  if (state.spreadsheet && state.spreadsheet.items.length > 0) {
    const templateKeys = new Set<string>();
    let templateRegisteredCount = 0;
    let templateDraftOnlyCount = 0;

    for (const item of state.spreadsheet.items) {
      templateKeys.add(item.normalizedKey);
      if (registeredResultKeys.has(item.normalizedKey)) {
        templateRegisteredCount += 1;
      } else if (draftOnlyKeys.has(item.normalizedKey)) {
        templateDraftOnlyCount += 1;
      }
    }

    let unmatchedResultCount = 0;
    for (const key of registeredResultKeys) {
      if (!templateKeys.has(key)) {
        unmatchedResultCount += 1;
      }
    }

    const templateTotal = state.spreadsheet.items.length;
    const templateSimulatedCount =
      templateRegisteredCount + templateDraftOnlyCount;
    const templateRemainingCount = Math.max(
      0,
      templateTotal - templateSimulatedCount
    );

    const registeredProgressPercent = Math.min(
      100,
      Math.round((templateRegisteredCount / templateTotal) * 100)
    );
    const progressPercent = Math.min(
      100,
      Math.round((templateSimulatedCount / templateTotal) * 100)
    );
    const draftProgressPercent = Math.max(
      0,
      progressPercent - registeredProgressPercent
    );

    return {
      rackTotalBoxes,
      rackFilledBoxes,
      rackEmptyBoxes,
      unsavedDraftCount,
      savedExcelCount,
      spreadsheetResultsTotal,
      externalResultsCount,
      totalRelocatedCount,
      templateTotal,
      templateRegisteredCount,
      templateDraftOnlyCount,
      templateSimulatedCount,
      templateRemainingCount,
      unmatchedResultCount,
      spreadsheetTotal: templateTotal,
      spreadsheetRemaining: templateRemainingCount,
      registeredProgressPercent,
      draftProgressPercent,
      progressPercent,
      progressLabel: `${templateSimulatedCount} dari ${templateTotal} kode template`,
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
    spreadsheetResultsTotal,
    externalResultsCount,
    totalRelocatedCount,
    templateTotal: null,
    templateRegisteredCount: null,
    templateDraftOnlyCount: null,
    templateSimulatedCount: null,
    templateRemainingCount: null,
    unmatchedResultCount: 0,
    spreadsheetTotal: null,
    spreadsheetRemaining: null,
    registeredProgressPercent: 0,
    draftProgressPercent: progressPercent,
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

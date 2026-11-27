import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyPlacement,
  buildExcelExportRows,
  calculateSimulationMetrics,
  clearSavedExcelRecords,
  commitDraftToExcel,
  computeBoxNumber,
  createInitialSimulationState,
  detectBestLocationColumn,
  extractSpreadsheetItems,
  formatNewLocationCode,
  getCombinedResultRecords,
  getEffectiveRackPlacements,
  getLockedLocationsMap,
  parseExternalSheetLocks,
  parseNewLocationCode,
  removeDraftByCoordinate,
  removePlacementByCoordinate,
  removeSavedRecordById,
  resetActiveRackSimulation,
  searchOldLocations,
  searchOldLocationsWithLockStatus,
  undoLastAction,
  validatePlacementInput,
} from './relocationEngine';
import type { ImportedSpreadsheet, SimulationState } from '../types/archive';

function createSampleSpreadsheet(): ImportedSpreadsheet {
  const rawRows = [
    { 'LOKASI SIMPAN LAMA': '201', KETERANGAN: 'Keuangan 2019' },
    { 'LOKASI SIMPAN LAMA': '202', KETERANGAN: 'Keuangan 2020' },
    { 'LOKASI SIMPAN LAMA': '205', KETERANGAN: 'Pengadaan 2021' },
    { 'LOKASI SIMPAN LAMA': '310', KETERANGAN: 'Umum 2022' },
  ];
  return {
    fileName: 'data_arsip.xlsx',
    sheetName: 'Sheet1',
    columns: ['LOKASI SIMPAN LAMA', 'KETERANGAN'],
    selectedColumn: 'LOKASI SIMPAN LAMA',
    rawRows,
    items: extractSpreadsheetItems(rawRows, 'LOKASI SIMPAN LAMA', 'KETERANGAN'),
    importedAt: '2026-11-24T10:00:00.000Z',
  };
}

test('formatNewLocationCode generates LOKASI.RAK.ROW.BOX correctly', () => {
  assert.equal(formatNewLocationCode('RC2', 1, 1, 1), 'RC2.1.1.1');
  assert.equal(formatNewLocationCode('RC2', 3, 2, 5), 'RC2.3.2.5');
  assert.equal(formatNewLocationCode(' rc3 ', 12, 4, 10), 'RC3.12.4.10');
});

test('computeBoxNumber supports custom starting box numbers in continuous and per-row modes', () => {
  // Continuous across rows starting at 51 (10 boxes/row)
  assert.equal(computeBoxNumber(1, 0, 10, 51, true), 51);
  assert.equal(computeBoxNumber(1, 9, 10, 51, true), 60);
  assert.equal(computeBoxNumber(2, 0, 10, 51, true), 61);
  assert.equal(computeBoxNumber(5, 9, 10, 51, true), 100);

  // Per-row reset starting at 51
  assert.equal(computeBoxNumber(1, 0, 10, 51, false), 51);
  assert.equal(computeBoxNumber(2, 0, 10, 51, false), 51);
  assert.equal(computeBoxNumber(2, 3, 10, 51, false), 54);
});

test('Simpan -> Reset -> Next Rack Cycle appends new records below previous cycles without losing Excel data', () => {
  let state = createInitialSimulationState();

  // Cycle 1: Rak 3
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 3, row: 1, box: 1 },
    '201'
  ).nextState;
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 3, row: 1, box: 2 },
    '202'
  ).nextState;

  // Before clicking Simpan, Excel savedRecords is still empty
  assert.equal(state.draftPlacements.length, 2);
  assert.equal(state.savedRecords.length, 0);

  // Click Simpan -> enters Excel
  const save1 = commitDraftToExcel(state);
  assert.equal(save1.addedCount, 2);
  state = save1.nextState;
  assert.equal(state.savedRecords.length, 2);

  // Clicking Simpan again does NOT duplicate already saved boxes
  const saveAgain = commitDraftToExcel(state);
  assert.equal(saveAgain.addedCount, 0);
  assert.equal(saveAgain.nextState.savedRecords.length, 2);

  // Click Reset -> clears active rack simulation ONLY, savedRecords stays intact!
  state = resetActiveRackSimulation(state);
  assert.equal(state.draftPlacements.length, 0);
  assert.equal(state.savedRecords.length, 2);

  // Cycle 2: Rak 4
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 4, row: 1, box: 1 },
    '203'
  ).nextState;
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 4, row: 2, box: 5 },
    '204'
  ).nextState;

  // Click Simpan -> appends below Cycle 1
  const save2 = commitDraftToExcel(state);
  assert.equal(save2.addedCount, 2);
  state = save2.nextState;

  const rows = buildExcelExportRows(state.savedRecords, false);
  assert.deepEqual(rows, [
    { 'FORMAT LAMA': '201', 'FORMAT BARU': 'RC2.3.1.1' },
    { 'FORMAT LAMA': '202', 'FORMAT BARU': 'RC2.3.1.2' },
    { 'FORMAT LAMA': '203', 'FORMAT BARU': 'RC2.4.1.1' },
    { 'FORMAT LAMA': '204', 'FORMAT BARU': 'RC2.4.2.5' },
  ]);
});

test('validation blocks duplicate old location against both active rack and previously saved Excel records', () => {
  let state = createInitialSimulationState();

  const emptyResult = validatePlacementInput(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '   '
  );
  assert.equal(emptyResult.status, 'empty');

  // Cycle 1: place 201, save to Excel, reset rack
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 3, row: 1, box: 1 },
    '201'
  ).nextState;
  state = commitDraftToExcel(state).nextState;
  state = resetActiveRackSimulation(state);

  // Cycle 2: attempting to place 201 again on Rak 4 must be blocked with saved Excel location
  const dupSaved = validatePlacementInput(
    state,
    { locationCode: 'RC2', rack: 4, row: 1, box: 1 },
    '201'
  );
  assert.equal(dupSaved.status, 'duplicate_old_location');
  if (dupSaved.status === 'duplicate_old_location') {
    assert.equal(dupSaved.inSavedExcel, true);
    assert.match(dupSaved.message, /RC2\.3\.1\.1/);
  }
});

test('occupied box requires explicit replacement confirmation and updates in place', () => {
  let state = createInitialSimulationState();
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '201'
  ).nextState;

  const unconfirmed = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '205'
  );
  assert.equal(unconfirmed.result.status, 'needs_replace_confirmation');
  assert.equal(unconfirmed.nextState.draftPlacements[0].oldLocation, '201');

  const confirmed = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '205',
    { confirmReplace: true }
  );
  assert.equal(confirmed.result.status, 'placed');
  assert.equal(confirmed.nextState.draftPlacements[0].oldLocation, '205');
});

test('spreadsheet search and strict validation work against actual imported data only', () => {
  const spreadsheet = createSampleSpreadsheet();
  const state = {
    ...createInitialSimulationState(),
    spreadsheet,
  };

  const matches = searchOldLocations(spreadsheet, '20');
  assert.deepEqual(
    matches.map((m) => m.oldLocation),
    ['201', '202', '205']
  );

  const notFound = validatePlacementInput(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '999'
  );
  assert.equal(notFound.status, 'not_in_spreadsheet');
});

test('undo, delete draft, delete saved, clearSavedExcelRecords, and metrics behave accurately', () => {
  let state = createInitialSimulationState({
    prefix: 'RC2',
    rackCount: 1,
    rowsPerRack: 2,
    boxesPerRow: 5,
  }); // 10 boxes per rack

  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 2, row: 1, box: 1 },
    '201'
  ).nextState;
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 2, row: 1, box: 2 },
    '202'
  ).nextState;

  let metrics = calculateSimulationMetrics(state, 'RC2', 2);
  assert.equal(metrics.rackTotalBoxes, 10);
  assert.equal(metrics.rackFilledBoxes, 2);
  assert.equal(metrics.rackEmptyBoxes, 8);
  assert.equal(metrics.unsavedDraftCount, 2);
  assert.equal(metrics.savedExcelCount, 0);

  state = commitDraftToExcel(state).nextState;
  metrics = calculateSimulationMetrics(state, 'RC2', 2);
  assert.equal(metrics.unsavedDraftCount, 0);
  assert.equal(metrics.savedExcelCount, 2);

  // Undo reverts Simpan
  state = undoLastAction(state);
  assert.equal(state.savedRecords.length, 0);
  assert.equal(state.draftPlacements.length, 2);

  // Delete draft box 2
  state = removeDraftByCoordinate(state, {
    locationCode: 'RC2',
    rack: 2,
    row: 1,
    box: 2,
  });
  assert.equal(state.draftPlacements.length, 1);

  // Save, delete by saved id, clear all
  state = commitDraftToExcel(state).nextState;
  assert.equal(state.savedRecords.length, 1);
  state = removeSavedRecordById(state, state.savedRecords[0].id);
  assert.equal(state.savedRecords.length, 0);

  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 2, row: 1, box: 1 },
    '201'
  ).nextState;
  state = commitDraftToExcel(state).nextState;
  state = clearSavedExcelRecords(state);
  assert.equal(state.savedRecords.length, 0);
});

test('detectBestLocationColumn prioritizes old location header names', () => {
  assert.equal(
    detectBestLocationColumn(['No', 'Uraian', 'Lokasi Simpan Lama', 'Tahun']),
    'Lokasi Simpan Lama'
  );
  assert.equal(detectBestLocationColumn(['Kode', 'Nama']), 'Kode');
});

test('locked old locations from simulation results and Google Sheets are hidden from default search and blocked on submit', () => {
  const spreadsheet = createSampleSpreadsheet();
  const sheetLocks = parseExternalSheetLocks([
    ['FORMAT LAMA', 'FORMAT BARU'],
    ['202', 'RC2.5.1.12'],
  ]);

  let state: SimulationState = {
    ...createInitialSimulationState(),
    spreadsheet,
    externalLockedLocations: sheetLocks,
  };

  // Relocate 201 and commit to saved results
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '201'
  ).nextState;
  state = commitDraftToExcel(state).nextState;
  state = resetActiveRackSimulation(state);

  const lockedMap = getLockedLocationsMap(state);
  assert.equal(lockedMap.get('201')?.newLocation, 'RC2.1.1.1');
  assert.equal(lockedMap.get('202')?.newLocation, 'RC2.5.1.12');
  assert.equal(lockedMap.get('202')?.source, 'google_sheets');

  // Default search (hideLocked: true) for '20' hides 201 and 202, leaving only 205
  const hiddenSearch = searchOldLocationsWithLockStatus(
    spreadsheet,
    lockedMap,
    '20',
    { hideLocked: true }
  );
  assert.deepEqual(
    hiddenSearch.visibleItems.map((i) => i.oldLocation),
    ['205']
  );
  assert.equal(hiddenSearch.matchedLockedItems.length, 2);

  // Searching directly for locked code '201' returns 0 visible available items but 1 matchedLockedItem
  const exactLockedSearch = searchOldLocationsWithLockStatus(
    spreadsheet,
    lockedMap,
    '201',
    { hideLocked: true }
  );
  assert.equal(exactLockedSearch.visibleItems.length, 0);
  assert.equal(exactLockedSearch.matchedLockedItems.length, 1);
  assert.equal(exactLockedSearch.matchedLockedItems[0].lock.newLocation, 'RC2.1.1.1');

  // Showing locked items (hideLocked: false) returns all 3 matches
  const showLockedSearch = searchOldLocationsWithLockStatus(
    spreadsheet,
    lockedMap,
    '20',
    { hideLocked: false }
  );
  assert.deepEqual(
    showLockedSearch.visibleItems.map((i) => i.oldLocation),
    ['201', '202', '205']
  );

  // Submitting Google-Sheets-locked code '202' is rejected
  const googleLockedValidation = validatePlacementInput(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 2 },
    '202'
  );
  assert.equal(googleLockedValidation.status, 'duplicate_old_location');
  if (googleLockedValidation.status === 'duplicate_old_location') {
    assert.equal(googleLockedValidation.inGoogleSheets, true);
    assert.equal(googleLockedValidation.lockedNewLocation, 'RC2.5.1.12');
  }
});

test('dashboard metrics synchronize template progress against Spreadsheet Hasil (Google Sheets + local saved + active draft) without double counting', () => {
  const spreadsheet = createSampleSpreadsheet(); // 4 items: 201, 202, 205, 310
  // External spreadsheet already has 201 (also saved locally), 202 (only in external sheet), and 999 (outside template)
  const externalLocks = parseExternalSheetLocks([
    ['FORMAT LAMA', 'FORMAT BARU'],
    ['201', 'RC2.1.1.1'],
    ['202', 'RC2.4.2.15'],
    ['999', 'RC2.9.1.1'],
  ]);

  let state: SimulationState = {
    ...createInitialSimulationState(),
    spreadsheet,
    externalLockedLocations: externalLocks,
  };

  // Place 205 as unsaved draft on an empty box on the active rack (RC2.1.1.1 is already occupied by 201 in Tabel Hasil)
  state = applyPlacement(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 2 },
    '205'
  ).nextState;

  // Also have 201 in local savedRecords to verify deduplication against externalLocks
  state = {
    ...state,
    savedRecords: [
      {
        id: 'saved-201',
        oldLocation: '201',
        newLocation: 'RC2.1.1.1',
        locationCode: 'RC2',
        rack: 1,
        row: 1,
        box: 9,
        sequence: 1,
        placedAt: '2026-01-01T00:00:00.000Z',
        savedToExcel: true,
      },
    ],
  };

  const metrics = calculateSimulationMetrics(state, 'RC2', 1);
  assert.equal(metrics.templateTotal, 4);
  assert.equal(metrics.templateRegisteredCount, 2); // 201 & 202 registered in Spreadsheet Hasil
  assert.equal(metrics.templateDraftOnlyCount, 1); // 205 in active rack draft only
  assert.equal(metrics.templateSimulatedCount, 3); // 201, 202, 205
  assert.equal(metrics.templateRemainingCount, 1); // 310 belum terdaftar / tersimulasi
  assert.equal(metrics.unmatchedResultCount, 1); // 999 outside template
  assert.equal(metrics.spreadsheetResultsTotal, 3); // 201, 202, 999 (deduplicated)
  assert.equal(metrics.registeredProgressPercent, 50); // 2/4
  assert.equal(metrics.draftProgressPercent, 25); // +1/4
  assert.equal(metrics.progressPercent, 75); // 3/4
});

test('entries registered in Tabel Hasil (savedRecords or external sheets/Excel) immediately appear as simulated boxes on the rack', () => {
  assert.deepEqual(parseNewLocationCode('RC2.3.2.15'), {
    locationCode: 'RC2',
    rack: 3,
    row: 2,
    box: 15,
  });
  assert.equal(parseNewLocationCode('INVALID'), null);

  const externalLocks = parseExternalSheetLocks(
    [
      ['FORMAT LAMA', 'FORMAT BARU'],
      ['201', 'RC2.1.1.1'],
      ['202', 'RC2.1.1.2'],
    ],
    'external_file'
  );

  let state: SimulationState = {
    ...createInitialSimulationState(),
    spreadsheet: createSampleSpreadsheet(),
    externalLockedLocations: externalLocks,
  };

  // Both rows from Tabel Hasil immediately show up in combined results AND as simulated boxes on Rak 1
  const combined = getCombinedResultRecords(state);
  assert.equal(combined.length, 2);
  assert.deepEqual(
    combined.map((r) => `${r.oldLocation}->${r.newLocation}`),
    ['201->RC2.1.1.1', '202->RC2.1.1.2']
  );

  const rackPlacements = getEffectiveRackPlacements(state);
  assert.equal(rackPlacements.length, 2);
  assert.equal(rackPlacements.every((p) => p.savedToExcel === true), true);

  // Active rack metrics reflect both simulated boxes immediately
  const metrics = calculateSimulationMetrics(state, 'RC2', 1);
  assert.equal(metrics.rackFilledBoxes, 2);
  assert.equal(metrics.rackEmptyBoxes, 48);

  // Attempting to place another code on RC2.1.1.1 asks for replace confirmation (cannot silently overwrite!)
  const overwriteCheck = validatePlacementInput(
    state,
    { locationCode: 'RC2', rack: 1, row: 1, box: 1 },
    '205'
  );
  assert.equal(overwriteCheck.status, 'needs_replace_confirmation');

  // Deleting box RC2.1.1.1 frees both the box and old code 201
  state = removePlacementByCoordinate(state, {
    locationCode: 'RC2',
    rack: 1,
    row: 1,
    box: 1,
  });
  assert.equal(getEffectiveRackPlacements(state).length, 1);
  assert.equal(getCombinedResultRecords(state).length, 1);
});



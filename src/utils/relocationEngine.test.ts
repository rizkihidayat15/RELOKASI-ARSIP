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
  removeDraftByCoordinate,
  removeSavedRecordById,
  resetActiveRackSimulation,
  searchOldLocations,
  undoLastAction,
  validatePlacementInput,
} from './relocationEngine';
import type { ImportedSpreadsheet } from '../types/archive';

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

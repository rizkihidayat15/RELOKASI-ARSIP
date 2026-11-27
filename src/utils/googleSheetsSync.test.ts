import test from 'node:test';
import assert from 'node:assert/strict';
import type { LockedLocationInfo, PlacementRecord } from '../types/archive';
import {
  buildSafeSpreadsheetUrl,
  computePendingGoogleSheetSyncCount,
  escapeSheetNameForA1,
  extractSpreadsheetGid,
  extractSpreadsheetId,
  planSpreadsheetRowSync,
} from './googleSheetsSync';

function makeRecord(
  sequence: number,
  oldLocation: string,
  newLocation: string,
  rack = 1,
  row = 1,
  box = sequence
): PlacementRecord {
  return {
    id: `rec-${sequence}`,
    oldLocation,
    newLocation,
    locationCode: newLocation.split('.')[0] || 'RC2',
    rack,
    row,
    box,
    sequence,
    placedAt: '2026-11-25T10:00:00.000Z',
    savedToExcel: true,
  };
}

test('extractSpreadsheetId parses Google Sheets URLs and raw IDs while rejecting unsafe URLs', () => {
  assert.equal(
    extractSpreadsheetId(
      'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=0'
    ),
    '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
  );
  assert.equal(
    extractSpreadsheetId(
      'https://docs.google.com/spreadsheets/u/0/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit?usp=sharing'
    ),
    '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
  );
  assert.equal(
    extractSpreadsheetId('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'),
    '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
  );
  assert.equal(extractSpreadsheetId('javascript:alert(1)'), null);
  assert.equal(
    extractSpreadsheetId('https://evil.example.com/spreadsheets/d/12345678901234567890'),
    null
  );
  assert.equal(extractSpreadsheetId('short'), null);
  assert.equal(
    buildSafeSpreadsheetUrl('1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'),
    'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit'
  );
});

test('extractSpreadsheetGid reads specific sheet tab gid from query or hash', () => {
  assert.equal(
    extractSpreadsheetGid(
      'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit?gid=987654#gid=987654'
    ),
    987654
  );
  assert.equal(
    extractSpreadsheetGid(
      'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit#gid=12345'
    ),
    12345
  );
  assert.equal(
    extractSpreadsheetGid(
      'https://docs.google.com/spreadsheets/d/1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms/edit'
    ),
    null
  );
});

test('escapeSheetNameForA1 quotes sheet names with spaces and single quotes safely', () => {
  assert.equal(escapeSheetNameForA1('Hasil Relokasi'), "'Hasil Relokasi'");
  assert.equal(escapeSheetNameForA1("Arsip '2026'"), "'Arsip ''2026'''");
});

test('planSpreadsheetRowSync writes header and appends all rows when target sheet is empty', () => {
  const records = [
    makeRecord(1, '201', 'RC2.1.1.1'),
    makeRecord(2, '202', 'RC2.1.1.2'),
  ];
  const plan = planSpreadsheetRowSync([], records, false);

  assert.equal(plan.needsHeader, true);
  assert.deepEqual(plan.headerRow, ['FORMAT LAMA', 'FORMAT BARU']);
  assert.equal(plan.updates.length, 0);
  assert.equal(plan.appends.length, 2);
  assert.equal(plan.unchangedCount, 0);
  assert.deepEqual(plan.appends[0].values, ['201', 'RC2.1.1.1']);
  assert.deepEqual(plan.appends[1].values, ['202', 'RC2.1.1.2']);
});

test('planSpreadsheetRowSync updates changed rows in place, skips unchanged rows, and appends new cycles below', () => {
  const existingRows = [
    ['FORMAT LAMA', 'FORMAT BARU'],
    ['201', 'RC2.1.1.1'],
    ['202', 'RC2.1.1.2'],
  ];

  const records = [
    makeRecord(1, '201', 'RC2.1.1.1'), // unchanged
    makeRecord(2, '202', 'RC2.3.2.65', 3, 2, 65), // updated location!
    makeRecord(3, '203', 'RC2.3.2.66', 3, 2, 66), // new row -> append below
  ];

  const plan = planSpreadsheetRowSync(existingRows, records, false);

  assert.equal(plan.needsHeader, false);
  assert.equal(plan.unchangedCount, 1);
  assert.equal(plan.existingDataRowCount, 2);

  assert.equal(plan.updates.length, 1);
  assert.equal(plan.updates[0].rowNumber, 3); // row 3 in 1-indexed sheet
  assert.equal(plan.updates[0].startColLetter, 'A');
  assert.equal(plan.updates[0].endColLetter, 'B');
  assert.deepEqual(plan.updates[0].values, ['202', 'RC2.3.2.65']);

  assert.equal(plan.appends.length, 1);
  assert.deepEqual(plan.appends[0].values, ['203', 'RC2.3.2.66']);
});

test('planSpreadsheetRowSync adapts to offset columns like NO | LOKASI SIMPAN LAMA | LOKASI SIMPAN BARU', () => {
  const existingRows = [
    ['NO', 'LOKASI SIMPAN LAMA', 'LOKASI SIMPAN BARU'],
    ['1', '201', 'RC2.1.1.1'],
    ['2', '202', 'RC2.1.1.2'],
  ];

  const records = [
    makeRecord(1, '201', 'RC2.1.1.1'),
    makeRecord(2, '202', 'RC2.2.1.9', 2, 1, 9),
    makeRecord(3, '203', 'RC2.2.1.10', 2, 1, 10),
  ];

  const plan = planSpreadsheetRowSync(existingRows, records, false);

  assert.equal(plan.needsHeader, false);
  assert.equal(plan.oldColIdx, 1);
  assert.equal(plan.newColIdx, 2);
  assert.equal(plan.unchangedCount, 1);
  assert.equal(plan.updates.length, 1);
  assert.equal(plan.updates[0].rowNumber, 3);
  assert.equal(plan.updates[0].startColLetter, 'C');
  assert.equal(plan.updates[0].endColLetter, 'C');
  assert.deepEqual(plan.updates[0].values, ['RC2.2.1.9']);

  assert.equal(plan.appends.length, 1);
  assert.equal(plan.appendStartColLetter, 'A');
  assert.equal(plan.appendEndColLetter, 'C');
  assert.deepEqual(plan.appends[0].values, [3, '203', 'RC2.2.1.10']);
});

test('planSpreadsheetRowSync supports optional extra columns cleanly', () => {
  const records = [makeRecord(1, '201', 'RC2.3.1.51', 3, 1, 51)];
  const plan = planSpreadsheetRowSync([], records, true);

  assert.equal(plan.needsHeader, true);
  assert.deepEqual(plan.headerRow, [
    'FORMAT LAMA',
    'FORMAT BARU',
    'RAK',
    'ROW',
    'BOX',
    'WAKTU SIMPAN',
  ]);
  assert.equal(plan.appends.length, 1);
  assert.equal(plan.appends[0].values[0], '201');
  assert.equal(plan.appends[0].values[1], 'RC2.3.1.51');
  assert.equal(plan.appends[0].values[2], 3);
  assert.equal(plan.appends[0].values[3], 1);
  assert.equal(plan.appends[0].values[4], 51);
});

test('computePendingGoogleSheetSyncCount accurately counts unsynced new and changed rows', () => {
  const externalLocks: LockedLocationInfo[] = [
    {
      oldLocation: '201',
      normalizedKey: '201',
      newLocation: 'RC2.1.1.1',
      source: 'google_sheets',
    },
    {
      oldLocation: '202',
      normalizedKey: '202',
      newLocation: 'RC2.1.1.2',
      source: 'google_sheets',
    },
  ];

  const records = [
    makeRecord(1, '201', 'RC2.1.1.1'), // already synced
    makeRecord(2, '202', 'RC2.2.1.5'), // updated!
    makeRecord(3, '203', 'RC2.2.1.6'), // new!
  ];

  const stats = computePendingGoogleSheetSyncCount(records, externalLocks, true);
  assert.equal(stats.pendingNewCount, 1);
  assert.equal(stats.pendingUpdateCount, 1);
  assert.equal(stats.totalPendingCount, 2);

  const whenDisconnected = computePendingGoogleSheetSyncCount(
    records,
    externalLocks,
    false
  );
  assert.equal(whenDisconnected.totalPendingCount, 0);
});

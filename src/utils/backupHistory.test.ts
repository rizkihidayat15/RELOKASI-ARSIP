import test from 'node:test';
import assert from 'node:assert/strict';
import type {
  ImportedSpreadsheet,
  LockedLocationInfo,
  PlacementRecord,
} from '../types/archive';
import {
  appendBackupHistoryEntry,
  buildTemplateBackupExportRows,
  createResultBackupEntry,
  createTemplateBackupEntry,
  filterBackupHistory,
  removeBackupHistoryEntry,
} from './backupHistory';

const SAMPLE_SHEET: ImportedSpreadsheet = {
  fileName: 'TEMPLATE_LOKASI_SIMPAN_LAMA.xlsx',
  sheetName: 'DATA_LOKASI_LAMA',
  columns: ['LOKASI SIMPAN LAMA', 'KETERANGAN'],
  selectedColumn: 'LOKASI SIMPAN LAMA',
  rawRows: [
    { 'LOKASI SIMPAN LAMA': '201', KETERANGAN: 'Keuangan' },
    { 'LOKASI SIMPAN LAMA': '202', KETERANGAN: 'Kepegawaian' },
  ],
  items: [
    {
      oldLocation: '201',
      normalizedKey: '201',
      rowNumber: 2,
      description: 'Keuangan',
    },
    {
      oldLocation: '202',
      normalizedKey: '202',
      rowNumber: 3,
      description: 'Kepegawaian',
    },
  ],
  importedAt: '2026-11-25T10:00:00.000Z',
};

const SAMPLE_RECORDS: PlacementRecord[] = [
  {
    id: 'r1',
    oldLocation: '201',
    newLocation: 'RC2.1.1.1',
    locationCode: 'RC2',
    rack: 1,
    row: 1,
    box: 1,
    sequence: 1,
    placedAt: '2026-11-25T10:05:00.000Z',
    savedToExcel: true,
  },
];

test('backup history prepends newest entries, respects max cap, filters, and deletes accurately', () => {
  const tplEntry = createTemplateBackupEntry({
    actionLabel: 'Impor Template Lokasi Lama',
    spreadsheet: SAMPLE_SHEET,
    lockedCount: 1,
    timestamp: '2026-11-25T10:00:00.000Z',
  });
  const resEntry = createResultBackupEntry({
    actionLabel: 'Update Google Sheets',
    records: SAMPLE_RECORDS,
    includeExtraColumns: false,
    targetSheetTitle: 'Master Relokasi 2026',
    timestamp: '2026-11-25T10:10:00.000Z',
  });

  let history = appendBackupHistoryEntry([], tplEntry, 2);
  history = appendBackupHistoryEntry(history, resEntry, 2);

  assert.equal(history.length, 2);
  assert.equal(history[0].category, 'result');
  assert.equal(history[1].category, 'template');

  assert.equal(filterBackupHistory(history, 'template', '').length, 1);
  assert.equal(filterBackupHistory(history, 'result', 'RC2.1.1.1').length, 1);
  assert.equal(
    filterBackupHistory(history, 'all', 'TEMPLATE_LOKASI_SIMPAN_LAMA').length,
    1
  );

  // Respects maxEntries retention cap
  const resEntry2 = createResultBackupEntry({
    actionLabel: 'Generate Excel',
    records: SAMPLE_RECORDS,
    includeExtraColumns: false,
  });
  history = appendBackupHistoryEntry(history, resEntry2, 2);
  assert.equal(history.length, 2);
  assert.equal(history[0].actionLabel, 'Generate Excel');

  history = removeBackupHistoryEntry(history, resEntry2.id);
  assert.equal(history.length, 1);
});

test('buildTemplateBackupExportRows annotates locked codes and supports remaining-only export', () => {
  const lockedMap = new Map<string, LockedLocationInfo>([
    [
      '201',
      {
        oldLocation: '201',
        normalizedKey: '201',
        newLocation: 'RC2.1.1.1',
        source: 'saved',
      },
    ],
  ]);

  const allRows = buildTemplateBackupExportRows(SAMPLE_SHEET, lockedMap, false);
  assert.equal(allRows.length, 2);
  assert.equal(allRows[0].STATUS, 'TERKUNCI (SUDAH DIGUNAKAN)');
  assert.equal(allRows[0]['LOKASI SIMPAN BARU'], 'RC2.1.1.1');
  assert.equal(allRows[1].STATUS, 'TERSEDIA');
  assert.equal(allRows[1]['LOKASI SIMPAN BARU'], '');

  const availableOnly = buildTemplateBackupExportRows(
    SAMPLE_SHEET,
    lockedMap,
    true
  );
  assert.equal(availableOnly.length, 1);
  assert.equal(availableOnly[0]['LOKASI SIMPAN LAMA'], '202');
});

import assert from 'node:assert/strict';
import test from 'node:test';
import {
  addBulkTemplateItems,
  addTemplateItem,
  createEmptyTemplateSpreadsheet,
  mergeSpreadsheetIntoTemplate,
  parseBulkTemplateInput,
  removeTemplateItem,
  updateTemplateItem,
} from './templateEditor.ts';
import {
  searchOldLocations,
  searchOldLocationsWithLockStatus,
} from './relocationEngine.ts';

test('addTemplateItem creates template if null, appends item, rejects duplicates, and immediately updates keyword search', () => {
  const first = addTemplateItem(null, '401', 'Arsip Pajak 2025');
  assert.equal(first.error, undefined);
  assert.equal(first.addedItem?.oldLocation, '401');
  assert.equal(first.nextSpreadsheet.items.length, 1);
  assert.equal(first.nextSpreadsheet.rawRows.length, 1);

  // Duplicate check (case/whitespace insensitive)
  const dup = addTemplateItem(first.nextSpreadsheet, ' 401 ', 'Duplikat');
  assert.ok(dup.error);
  assert.equal(dup.nextSpreadsheet.items.length, 1);

  // Add second item
  const second = addTemplateItem(
    first.nextSpreadsheet,
    'B-12',
    'Arsip Audit Internal'
  );
  assert.equal(second.error, undefined);
  assert.equal(second.nextSpreadsheet.items.length, 2);

  // Keyword search immediately finds by code AND by description keyword
  const byCode = searchOldLocations(second.nextSpreadsheet, '401');
  assert.deepEqual(
    byCode.map((i) => i.oldLocation),
    ['401']
  );

  const byKeyword = searchOldLocations(second.nextSpreadsheet, 'pajak');
  assert.deepEqual(
    byKeyword.map((i) => i.oldLocation),
    ['401']
  );

  const byKeywordAudit = searchOldLocationsWithLockStatus(
    second.nextSpreadsheet,
    new Map(),
    'audit'
  );
  assert.deepEqual(
    byKeywordAudit.visibleItems.map((i) => i.oldLocation),
    ['B-12']
  );
});

test('parseBulkTemplateInput and addBulkTemplateItems support newline, comma, semicolon, tab (Excel), and pipe formats with deduplication', () => {
  const parsed = parseBulkTemplateInput(
    [
      '401\tArsip Keuangan 2025',
      '402 | Arsip Kontrak',
      '403, 404; 405',
      '   ',
      '406',
    ].join('\n')
  );

  assert.deepEqual(parsed, [
    { oldLocation: '401', description: 'Arsip Keuangan 2025' },
    { oldLocation: '402', description: 'Arsip Kontrak' },
    { oldLocation: '403' },
    { oldLocation: '404' },
    { oldLocation: '405' },
    { oldLocation: '406' },
  ]);

  let sheet = createEmptyTemplateSpreadsheet();
  sheet = addTemplateItem(sheet, '401', 'Sudah Ada').nextSpreadsheet;

  const bulk = addBulkTemplateItems(sheet, [
    ...parsed,
    { oldLocation: '403', description: 'Duplikat dalam batch' },
  ]);

  assert.equal(bulk.addedCount, 5); // 402, 403, 404, 405, 406
  assert.deepEqual(bulk.duplicateCodes, ['401', '403']);
  assert.equal(bulk.nextSpreadsheet.items.length, 6);
  assert.equal(bulk.nextSpreadsheet.rawRows.length, 6);
});

test('mergeSpreadsheetIntoTemplate appends only unique items from an additional spreadsheet file', () => {
  let base = createEmptyTemplateSpreadsheet('TEMPLATE_A.xlsx');
  base = addTemplateItem(base, '201', 'Lama 201').nextSpreadsheet;
  base = addTemplateItem(base, '202', 'Lama 202').nextSpreadsheet;

  let incoming = createEmptyTemplateSpreadsheet('TAMBAHAN_B.xlsx');
  incoming = addTemplateItem(incoming, '202', 'Duplikat 202').nextSpreadsheet;
  incoming = addTemplateItem(incoming, '501', 'Baru 501').nextSpreadsheet;
  incoming = addTemplateItem(incoming, '502', 'Baru 502').nextSpreadsheet;

  const merged = mergeSpreadsheetIntoTemplate(base, incoming);
  assert.equal(merged.addedCount, 2);
  assert.equal(merged.skippedDuplicates, 1);
  assert.deepEqual(
    merged.nextSpreadsheet.items.map((i) => i.oldLocation),
    ['201', '202', '501', '502']
  );
  assert.equal(merged.nextSpreadsheet.fileName, 'TEMPLATE_A.xlsx');
});

test('updateTemplateItem updates code and keyword, protects locked codes from renaming, and removeTemplateItem deletes cleanly', () => {
  let sheet = createEmptyTemplateSpreadsheet();
  sheet = addTemplateItem(sheet, '201', 'Keuangan').nextSpreadsheet;
  sheet = addTemplateItem(sheet, '202', 'Umum').nextSpreadsheet;

  // Rename unlocked item 202 -> 202A + new keyword
  const edited = updateTemplateItem(sheet, '202', '202A', 'Arsip Khusus 2025');
  assert.equal(edited.error, undefined);
  assert.deepEqual(
    edited.nextSpreadsheet.items.map((i) => i.oldLocation),
    ['201', '202A']
  );
  assert.equal(
    searchOldLocations(edited.nextSpreadsheet, 'khusus')[0]?.oldLocation,
    '202A'
  );

  // Reject collision with existing code
  const collision = updateTemplateItem(
    edited.nextSpreadsheet,
    '202A',
    '201',
    'Tabrakan'
  );
  assert.ok(collision.error);

  // Locked item cannot change code, but CAN update description
  const lockedRename = updateTemplateItem(
    edited.nextSpreadsheet,
    '201',
    '201X',
    'Keuangan Baru',
    { isLocked: true }
  );
  assert.ok(lockedRename.error);

  const lockedDescEdit = updateTemplateItem(
    edited.nextSpreadsheet,
    '201',
    '201',
    'Keuangan Terverifikasi',
    { isLocked: true }
  );
  assert.equal(lockedDescEdit.error, undefined);
  assert.equal(
    lockedDescEdit.nextSpreadsheet.items[0]?.description,
    'Keuangan Terverifikasi'
  );

  // Remove unlocked item
  const removed = removeTemplateItem(lockedDescEdit.nextSpreadsheet, '202A');
  assert.equal(removed.removedItem?.oldLocation, '202A');
  assert.equal(removed.nextSpreadsheet.items.length, 1);
  assert.equal(removed.nextSpreadsheet.rawRows.length, 1);
});

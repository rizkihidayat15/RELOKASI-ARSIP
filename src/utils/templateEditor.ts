import type { ImportedSpreadsheet, SpreadsheetItem } from '../types/archive';
import { normalizeLocationCode } from './relocationEngine';

export interface TemplateEntryInput {
  oldLocation: string;
  description?: string;
}

export function getTemplateDescriptionColumn(
  spreadsheet: ImportedSpreadsheet
): string {
  const existing = spreadsheet.columns.find(
    (col) =>
      col !== spreadsheet.selectedColumn &&
      /keterangan|uraian|judul|nama|arsip|unit|deskripsi/i.test(col)
  );
  return existing ?? 'KETERANGAN';
}

export function createEmptyTemplateSpreadsheet(
  fileName = 'TEMPLATE_LOKASI_SIMPAN_LAMA.xlsx'
): ImportedSpreadsheet {
  return {
    fileName,
    sheetName: 'DATA_LOKASI_LAMA',
    columns: ['LOKASI SIMPAN LAMA', 'KETERANGAN'],
    selectedColumn: 'LOKASI SIMPAN LAMA',
    rawRows: [],
    items: [],
    importedAt: new Date().toISOString(),
  };
}

function buildRawRow(
  columns: string[],
  selectedColumn: string,
  descColumn: string,
  oldLocation: string,
  description?: string
): Record<string, string> {
  const row: Record<string, string> = {};
  for (const col of columns) {
    if (col === selectedColumn) {
      row[col] = oldLocation;
    } else if (col === descColumn) {
      row[col] = description ?? '';
    } else {
      row[col] = '';
    }
  }
  return row;
}

function ensureDescriptionColumn(
  spreadsheet: ImportedSpreadsheet,
  needsDesc: boolean
): { columns: string[]; descColumn: string } {
  const descColumn = getTemplateDescriptionColumn(spreadsheet);
  if (!needsDesc || spreadsheet.columns.includes(descColumn)) {
    return { columns: [...spreadsheet.columns], descColumn };
  }
  return {
    columns: [...spreadsheet.columns, descColumn],
    descColumn,
  };
}

export function addTemplateItem(
  spreadsheet: ImportedSpreadsheet | null,
  rawOldLocation: string,
  rawDescription?: string
): {
  nextSpreadsheet: ImportedSpreadsheet;
  addedItem?: SpreadsheetItem;
  error?: string;
} {
  const base = spreadsheet ?? createEmptyTemplateSpreadsheet();
  const oldLocation = rawOldLocation.trim();
  const description = (rawDescription ?? '').trim();

  if (!oldLocation) {
    return {
      nextSpreadsheet: base,
      error: 'Masukkan kode lokasi simpan lama terlebih dahulu.',
    };
  }

  const normalizedKey = normalizeLocationCode(oldLocation);
  const existing = base.items.find(
    (item) => item.normalizedKey === normalizedKey
  );
  if (existing) {
    return {
      nextSpreadsheet: base,
      error: `Kode lokasi simpan lama "${existing.oldLocation}" sudah ada di dalam template.`,
    };
  }

  const { columns, descColumn } = ensureDescriptionColumn(
    base,
    Boolean(description)
  );

  const nextRowNumber =
    base.items.reduce((max, item) => Math.max(max, item.rowNumber), 1) + 1;

  const addedItem: SpreadsheetItem = {
    oldLocation,
    normalizedKey,
    rowNumber: nextRowNumber,
    description: description || undefined,
  };

  const nextRawRows = base.rawRows.map((row) =>
    columns.includes(descColumn) && !(descColumn in row)
      ? { ...row, [descColumn]: '' }
      : { ...row }
  );
  nextRawRows.push(
    buildRawRow(
      columns,
      base.selectedColumn,
      descColumn,
      oldLocation,
      description
    )
  );

  return {
    nextSpreadsheet: {
      ...base,
      columns,
      rawRows: nextRawRows,
      items: [...base.items, addedItem],
    },
    addedItem,
  };
}

export function parseBulkTemplateInput(rawText: string): TemplateEntryInput[] {
  const entries: TemplateEntryInput[] = [];
  const lines = rawText.split(/\r?\n/);

  for (const rawLine of lines) {
    const line = rawLine.trim();
    if (!line) continue;

    // Tab-separated (direct copy-paste from Excel/Google Sheets)
    if (line.includes('\t')) {
      const parts = line.split('\t').map((p) => p.trim());
      const code = parts[0] ?? '';
      const desc = parts.slice(1).filter(Boolean).join(' ').trim();
      if (code) {
        entries.push({
          oldLocation: code,
          description: desc || undefined,
        });
      }
      continue;
    }

    // Pipe-separated: "401 | Arsip Keuangan"
    if (line.includes('|')) {
      const [codePart, ...rest] = line.split('|');
      const code = (codePart ?? '').trim();
      const desc = rest.join('|').trim();
      if (code) {
        entries.push({
          oldLocation: code,
          description: desc || undefined,
        });
      }
      continue;
    }

    // Comma or semicolon separated list of codes on one line: "401, 402, 403"
    if (line.includes(',') || line.includes(';')) {
      const tokens = line
        .split(/[,;]+/)
        .map((t) => t.trim())
        .filter(Boolean);
      for (const token of tokens) {
        entries.push({ oldLocation: token });
      }
      continue;
    }

    entries.push({ oldLocation: line });
  }

  return entries;
}

export function addBulkTemplateItems(
  spreadsheet: ImportedSpreadsheet | null,
  entries: TemplateEntryInput[]
): {
  nextSpreadsheet: ImportedSpreadsheet;
  addedCount: number;
  duplicateCodes: string[];
  invalidCount: number;
} {
  const base = spreadsheet ?? createEmptyTemplateSpreadsheet();
  const seenKeys = new Set(base.items.map((item) => item.normalizedKey));
  const duplicateCodes: string[] = [];
  let invalidCount = 0;

  const validToAdd: Array<{
    oldLocation: string;
    normalizedKey: string;
    description?: string;
  }> = [];

  for (const entry of entries) {
    const code = entry.oldLocation.trim();
    if (!code) {
      invalidCount += 1;
      continue;
    }
    const key = normalizeLocationCode(code);
    if (!key) {
      invalidCount += 1;
      continue;
    }
    if (seenKeys.has(key)) {
      duplicateCodes.push(code);
      continue;
    }
    seenKeys.add(key);
    const desc = (entry.description ?? '').trim();
    validToAdd.push({
      oldLocation: code,
      normalizedKey: key,
      description: desc || undefined,
    });
  }

  if (validToAdd.length === 0) {
    return {
      nextSpreadsheet: base,
      addedCount: 0,
      duplicateCodes,
      invalidCount,
    };
  }

  const hasAnyDesc = validToAdd.some((v) => Boolean(v.description));
  const { columns, descColumn } = ensureDescriptionColumn(base, hasAnyDesc);

  let nextRowNum =
    base.items.reduce((max, item) => Math.max(max, item.rowNumber), 1) + 1;

  const newItems: SpreadsheetItem[] = validToAdd.map((v) => ({
    oldLocation: v.oldLocation,
    normalizedKey: v.normalizedKey,
    rowNumber: nextRowNum++,
    description: v.description,
  }));

  const nextRawRows = base.rawRows.map((row) =>
    columns.includes(descColumn) && !(descColumn in row)
      ? { ...row, [descColumn]: '' }
      : { ...row }
  );

  for (const v of validToAdd) {
    nextRawRows.push(
      buildRawRow(
        columns,
        base.selectedColumn,
        descColumn,
        v.oldLocation,
        v.description
      )
    );
  }

  return {
    nextSpreadsheet: {
      ...base,
      columns,
      rawRows: nextRawRows,
      items: [...base.items, ...newItems],
    },
    addedCount: newItems.length,
    duplicateCodes,
    invalidCount,
  };
}

export function mergeSpreadsheetIntoTemplate(
  base: ImportedSpreadsheet | null,
  incoming: ImportedSpreadsheet
): {
  nextSpreadsheet: ImportedSpreadsheet;
  addedCount: number;
  skippedDuplicates: number;
} {
  if (!base) {
    return {
      nextSpreadsheet: incoming,
      addedCount: incoming.items.length,
      skippedDuplicates: 0,
    };
  }

  const result = addBulkTemplateItems(
    base,
    incoming.items.map((item) => ({
      oldLocation: item.oldLocation,
      description: item.description,
    }))
  );

  return {
    nextSpreadsheet: result.nextSpreadsheet,
    addedCount: result.addedCount,
    skippedDuplicates: result.duplicateCodes.length,
  };
}

export function updateTemplateItem(
  spreadsheet: ImportedSpreadsheet,
  originalNormalizedKey: string,
  rawNextOldLocation: string,
  rawNextDescription: string,
  options?: { isLocked?: boolean }
): {
  nextSpreadsheet: ImportedSpreadsheet;
  updatedItem?: SpreadsheetItem;
  error?: string;
} {
  const targetIndex = spreadsheet.items.findIndex(
    (item) => item.normalizedKey === originalNormalizedKey
  );
  if (targetIndex === -1) {
    return {
      nextSpreadsheet: spreadsheet,
      error: 'Item lokasi simpan lama tidak ditemukan pada template.',
    };
  }

  const originalItem = spreadsheet.items[targetIndex];
  const nextOldLocation = rawNextOldLocation.trim();
  const nextDescription = rawNextDescription.trim();

  if (!nextOldLocation) {
    return {
      nextSpreadsheet: spreadsheet,
      error: 'Kode lokasi simpan lama tidak boleh kosong.',
    };
  }

  const nextNormalizedKey = normalizeLocationCode(nextOldLocation);

  if (options?.isLocked && nextNormalizedKey !== originalNormalizedKey) {
    return {
      nextSpreadsheet: spreadsheet,
      error: `Kode "${originalItem.oldLocation}" sudah terkunci pada hasil simulasi. Hanya kolom keterangan yang dapat diubah.`,
    };
  }

  const duplicate = spreadsheet.items.find(
    (item, idx) =>
      idx !== targetIndex && item.normalizedKey === nextNormalizedKey
  );
  if (duplicate) {
    return {
      nextSpreadsheet: spreadsheet,
      error: `Kode "${duplicate.oldLocation}" sudah digunakan pada baris lain di template.`,
    };
  }

  const { columns, descColumn } = ensureDescriptionColumn(
    spreadsheet,
    Boolean(nextDescription)
  );

  const updatedItem: SpreadsheetItem = {
    ...originalItem,
    oldLocation: nextOldLocation,
    normalizedKey: nextNormalizedKey,
    description: nextDescription || undefined,
  };

  const nextItems = spreadsheet.items.map((item, idx) =>
    idx === targetIndex ? updatedItem : item
  );

  let updatedRawRow = false;
  const nextRawRows = spreadsheet.rawRows.map((row) => {
    const baseRow =
      columns.includes(descColumn) && !(descColumn in row)
        ? { ...row, [descColumn]: '' }
        : { ...row };
    const rowCodeKey = normalizeLocationCode(
      String(baseRow[spreadsheet.selectedColumn] ?? '')
    );
    if (!updatedRawRow && rowCodeKey === originalNormalizedKey) {
      updatedRawRow = true;
      return {
        ...baseRow,
        [spreadsheet.selectedColumn]: nextOldLocation,
        ...(columns.includes(descColumn)
          ? { [descColumn]: nextDescription }
          : {}),
      };
    }
    return baseRow;
  });

  if (!updatedRawRow) {
    nextRawRows.push(
      buildRawRow(
        columns,
        spreadsheet.selectedColumn,
        descColumn,
        nextOldLocation,
        nextDescription
      )
    );
  }

  return {
    nextSpreadsheet: {
      ...spreadsheet,
      columns,
      rawRows: nextRawRows,
      items: nextItems,
    },
    updatedItem,
  };
}

export function removeTemplateItem(
  spreadsheet: ImportedSpreadsheet,
  normalizedKey: string
): {
  nextSpreadsheet: ImportedSpreadsheet;
  removedItem?: SpreadsheetItem;
} {
  const removedItem = spreadsheet.items.find(
    (item) => item.normalizedKey === normalizedKey
  );
  if (!removedItem) {
    return { nextSpreadsheet: spreadsheet };
  }

  const nextItems = spreadsheet.items.filter(
    (item) => item.normalizedKey !== normalizedKey
  );

  const nextRawRows = spreadsheet.rawRows.filter(
    (row) =>
      normalizeLocationCode(
        String(row[spreadsheet.selectedColumn] ?? '')
      ) !== normalizedKey
  );

  return {
    nextSpreadsheet: {
      ...spreadsheet,
      items: nextItems,
      rawRows: nextRawRows,
    },
    removedItem,
  };
}

export function buildUpdatedTemplateExportRows(
  spreadsheet: ImportedSpreadsheet
): Array<Record<string, string>> {
  return spreadsheet.items.map((item) => ({
    'LOKASI SIMPAN LAMA': item.oldLocation,
    KETERANGAN: item.description ?? '',
  }));
}

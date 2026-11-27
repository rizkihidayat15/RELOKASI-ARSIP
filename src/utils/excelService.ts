import * as XLSX from 'xlsx';
import type {
  ImportedSpreadsheet,
  LockedLocationInfo,
  PlacementRecord,
} from '../types/archive';
import {
  buildExcelExportRows,
  detectBestLocationColumn,
  extractSpreadsheetItems,
  parseExternalSheetLocks,
} from './relocationEngine';

export const EXPORT_EXCEL_FILENAME = 'HASIL_LOKASI_SIMPAN_LAMA_DAN_BARU.xlsx';

export interface GeneratedExcelArtifact {
  fileName: string;
  rowCount: number;
  columnNames: string[];
  byteSize: number;
  generatedAt: string;
  buffer: ArrayBuffer;
}

export async function parseSpreadsheetFile(
  file: File
): Promise<ImportedSpreadsheet> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('File spreadsheet tidak memiliki lembar kerja (sheet).');
  }
  const worksheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(
    worksheet,
    { header: 1, defval: '' }
  );

  const nonEmptyRows = matrix.filter((row) =>
    row.some((cell) => String(cell ?? '').trim() !== '')
  );

  if (nonEmptyRows.length === 0) {
    throw new Error('File spreadsheet kosong. Pastikan berisi minimal 1 baris data.');
  }

  // Detect whether first row looks like headers or pure data
  const firstRow = nonEmptyRows[0].map((c) => String(c ?? '').trim());
  const hasMultipleRows = nonEmptyRows.length > 1;
  const firstRowLooksNumericOnly = firstRow.every(
    (cell) => cell === '' || /^\d+([.,-]\d+)*$/.test(cell)
  );

  let columns: string[];
  let dataRows: (string | number | boolean | null)[][];

  if (!hasMultipleRows || firstRowLooksNumericOnly) {
    const colCount = Math.max(...nonEmptyRows.map((r) => r.length), 1);
    columns = Array.from({ length: colCount }, (_, idx) =>
      idx === 0 ? 'LOKASI LAMA' : `KOLOM ${idx + 1}`
    );
    dataRows = nonEmptyRows;
  } else {
    columns = firstRow.map((header, idx) => header || `KOLOM ${idx + 1}`);
    dataRows = nonEmptyRows.slice(1);
  }

  // Deduplicate column names if needed
  const usedNames = new Map<string, number>();
  const uniqueColumns = columns.map((col) => {
    const count = usedNames.get(col) ?? 0;
    usedNames.set(col, count + 1);
    return count === 0 ? col : `${col} (${count + 1})`;
  });

  const rawRows: Record<string, string>[] = dataRows.map((row) => {
    const obj: Record<string, string> = {};
    uniqueColumns.forEach((colName, idx) => {
      obj[colName] = String(row[idx] ?? '').trim();
    });
    return obj;
  });

  const selectedColumn = detectBestLocationColumn(uniqueColumns);
  const descriptionColumn = uniqueColumns.find(
    (col) =>
      col !== selectedColumn &&
      /keterangan|uraian|judul|nama|arsip|unit|deskripsi/i.test(col)
  );

  const items = extractSpreadsheetItems(rawRows, selectedColumn, descriptionColumn);
  if (items.length === 0) {
    throw new Error(
      `Kolom "${selectedColumn}" tidak memiliki nilai lokasi simpan lama yang valid.`
    );
  }

  return {
    fileName: file.name,
    sheetName,
    columns: uniqueColumns,
    selectedColumn,
    rawRows,
    items,
    importedAt: new Date().toISOString(),
  };
}

export async function parseResultSpreadsheetFile(file: File): Promise<{
  fileName: string;
  sheetName: string;
  locks: LockedLocationInfo[];
}> {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const sheetName = workbook.SheetNames[0];
  if (!sheetName) {
    throw new Error('File hasil relokasi tidak memiliki lembar kerja (sheet).');
  }
  const worksheet = workbook.Sheets[sheetName];
  const matrix = XLSX.utils.sheet_to_json<(string | number | boolean | null)[]>(
    worksheet,
    { header: 1, defval: '' }
  );

  const stringRows: string[][] = matrix.map((row) =>
    row.map((cell) => String(cell ?? '').trim())
  );
  const locks = parseExternalSheetLocks(stringRows, 'external_file');

  if (locks.length === 0) {
    throw new Error(
      'Tidak ditemukan pasangan data FORMAT LAMA dan FORMAT BARU yang valid pada file hasil relokasi ini.'
    );
  }

  return {
    fileName: file.name,
    sheetName,
    locks,
  };
}

export function generateRelocationExcelArtifact(
  placements: PlacementRecord[],
  includeExtraColumns: boolean
): GeneratedExcelArtifact {
  const rows = buildExcelExportRows(placements, includeExtraColumns);
  const columnNames = includeExtraColumns
    ? ['FORMAT LAMA', 'FORMAT BARU', 'RAK', 'ROW', 'BOX', 'WAKTU PEMINDAHAN']
    : ['FORMAT LAMA', 'FORMAT BARU'];

  const worksheet = XLSX.utils.json_to_sheet(rows, { header: columnNames });

  worksheet['!cols'] = includeExtraColumns
    ? [
        { wch: 22 },
        { wch: 22 },
        { wch: 8 },
        { wch: 8 },
        { wch: 8 },
        { wch: 24 },
      ]
    : [{ wch: 24 }, { wch: 24 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'HASIL_SIMULASI');

  const arrayBuffer = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
  }) as ArrayBuffer;

  return {
    fileName: EXPORT_EXCEL_FILENAME,
    rowCount: rows.length,
    columnNames,
    byteSize: arrayBuffer.byteLength,
    generatedAt: new Date().toISOString(),
    buffer: arrayBuffer,
  };
}

export function downloadExcelArtifact(artifact: GeneratedExcelArtifact): void {
  const blob = new Blob([artifact.buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = artifact.fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadSpreadsheetTemplate(): void {
  const sampleRows = [
    { 'LOKASI SIMPAN LAMA': '201', KETERANGAN: 'Arsip Keuangan 2019' },
    { 'LOKASI SIMPAN LAMA': '202', KETERANGAN: 'Arsip Keuangan 2020' },
    { 'LOKASI SIMPAN LAMA': '203', KETERANGAN: 'Arsip Kepegawaian A' },
    { 'LOKASI SIMPAN LAMA': '204', KETERANGAN: 'Arsip Kepegawaian B' },
    { 'LOKASI SIMPAN LAMA': '205', KETERANGAN: 'Arsip Pengadaan 2021' },
    { 'LOKASI SIMPAN LAMA': '206', KETERANGAN: 'Arsip Pengadaan 2022' },
    { 'LOKASI SIMPAN LAMA': '210', KETERANGAN: 'Arsip Hukum & Kontrak' },
    { 'LOKASI SIMPAN LAMA': '301', KETERANGAN: 'Arsip Umum & Tata Usaha' },
  ];
  const worksheet = XLSX.utils.json_to_sheet(sampleRows);
  worksheet['!cols'] = [{ wch: 24 }, { wch: 30 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'DATA_LOKASI_LAMA');
  const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = 'TEMPLATE_LOKASI_SIMPAN_LAMA.xlsx';
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadUpdatedTemplateExcel(
  spreadsheet: ImportedSpreadsheet
): void {
  const rows = spreadsheet.items.map((item) => ({
    'LOKASI SIMPAN LAMA': item.oldLocation,
    KETERANGAN: item.description ?? '',
  }));
  const worksheet = XLSX.utils.json_to_sheet(rows, {
    header: ['LOKASI SIMPAN LAMA', 'KETERANGAN'],
  });
  worksheet['!cols'] = [{ wch: 24 }, { wch: 36 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    worksheet,
    spreadsheet.sheetName?.slice(0, 31) || 'DATA_LOKASI_LAMA'
  );
  const buffer = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
  }) as ArrayBuffer;
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const baseName = spreadsheet.fileName
    .replace(/\.(xlsx|xls|csv)$/i, '')
    .trim();
  const fileName = `${baseName || 'TEMPLATE_LOKASI_SIMPAN_LAMA'}.xlsx`;
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function formatBackupTimestampSlug(isoDate: string): string {
  const d = new Date(isoDate);
  if (Number.isNaN(d.getTime())) return 'BACKUP';
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}_${pad(d.getHours())}${pad(d.getMinutes())}`;
}

export function downloadTemplateBackupExcel(
  rows: Record<string, string>[],
  createdAt: string,
  suffix = 'TEMPLATE'
): void {
  const worksheet = XLSX.utils.json_to_sheet(rows, {
    header: [
      'LOKASI SIMPAN LAMA',
      'KETERANGAN',
      'STATUS',
      'LOKASI SIMPAN BARU',
    ],
  });
  worksheet['!cols'] = [{ wch: 22 }, { wch: 28 }, { wch: 26 }, { wch: 22 }];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'CADANGAN_TEMPLATE');
  const buffer = XLSX.write(workbook, {
    bookType: 'xlsx',
    type: 'array',
  }) as ArrayBuffer;
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = `CADANGAN_${suffix}_${formatBackupTimestampSlug(createdAt)}.xlsx`;
  document.body.appendChild(anchor);
  anchor.click();
  document.body.removeChild(anchor);
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export function downloadResultBackupExcel(
  records: PlacementRecord[],
  includeExtraColumns: boolean,
  createdAt: string
): void {
  const artifact = generateRelocationExcelArtifact(
    records,
    includeExtraColumns
  );
  downloadExcelArtifact({
    ...artifact,
    fileName: `CADANGAN_HASIL_${formatBackupTimestampSlug(createdAt)}.xlsx`,
  });
}

import type { LockedLocationInfo, PlacementRecord } from '../types/archive';

export interface DriveSpreadsheetItem {
  id: string;
  name: string;
  modifiedTime?: string;
  webViewLink?: string;
}

export interface ConnectedGoogleSheet {
  spreadsheetId: string;
  spreadsheetUrl: string;
  title: string;
  sheetName: string;
  availableSheets: string[];
  preferredGid?: number;
  lastSyncedAt?: string;
  lastSyncSummary?: {
    updatedCount: number;
    appendedCount: number;
    unchangedCount: number;
    totalSyncedRows: number;
  };
}

export interface PlannedRowUpdate {
  rowNumber: number;
  startColLetter: string;
  endColLetter: string;
  values: (string | number)[];
  oldLocation: string;
  newLocation: string;
}

export interface PlannedRowAppend {
  values: (string | number)[];
  oldLocation: string;
  newLocation: string;
}

export interface SpreadsheetSyncPlan {
  needsHeader: boolean;
  headerRow: string[];
  oldColIdx: number;
  newColIdx: number;
  appendStartColLetter: string;
  appendEndColLetter: string;
  updates: PlannedRowUpdate[];
  appends: PlannedRowAppend[];
  unchangedCount: number;
  existingDataRowCount: number;
}

export interface SpreadsheetSyncExecutionResult {
  updatedCount: number;
  appendedCount: number;
  unchangedCount: number;
  headerWritten: boolean;
  totalSheetRows: number;
}

export interface PendingGoogleSheetSyncStats {
  pendingNewCount: number;
  pendingUpdateCount: number;
  totalPendingCount: number;
}

export class WorkspaceAuthExpiredError extends Error {
  constructor(
    message = 'Sesi akses Google berakhir. Silakan masuk kembali dengan Google.'
  ) {
    super(message);
    this.name = 'WorkspaceAuthExpiredError';
  }
}

const SPREADSHEET_ID_PATTERN = /^[a-zA-Z0-9_-]{15,120}$/;

/**
 * Extracts a validated Google Spreadsheet ID from a pasted Google Sheets URL or raw ID.
 * Rejects non-Google hosts and unsafe URI schemes.
 */
export function extractSpreadsheetId(rawInput: string): string | null {
  const trimmed = rawInput.trim();
  if (!trimmed) return null;

  if (SPREADSHEET_ID_PATTERN.test(trimmed)) {
    return trimmed;
  }

  try {
    const parsed = new URL(trimmed);
    if (parsed.protocol !== 'https:' && parsed.protocol !== 'http:') {
      return null;
    }
    const host = parsed.hostname.toLowerCase();
    if (host !== 'docs.google.com' && host !== 'drive.google.com') {
      return null;
    }
    const pathMatch = parsed.pathname.match(
      /\/(?:spreadsheets\/(?:u\/\d+\/)?d|file\/d)\/([a-zA-Z0-9_-]{15,120})/
    );
    if (pathMatch?.[1] && SPREADSHEET_ID_PATTERN.test(pathMatch[1])) {
      return pathMatch[1];
    }
    const queryId = parsed.searchParams.get('id');
    if (queryId && SPREADSHEET_ID_PATTERN.test(queryId)) {
      return queryId;
    }
    return null;
  } catch {
    return null;
  }
}

/**
 * Extracts the specific sheet tab gid (if present in query ?gid= or hash #gid=)
 * so pasting a link to a specific tab always reads and updates THAT exact tab.
 */
export function extractSpreadsheetGid(rawInput: string): number | null {
  const trimmed = rawInput.trim();
  if (!trimmed) return null;
  try {
    const parsed = new URL(trimmed);
    const queryGid = parsed.searchParams.get('gid');
    if (queryGid && /^\d+$/.test(queryGid)) {
      return Number(queryGid);
    }
    const hashMatch = parsed.hash.match(/gid=(\d+)/);
    if (hashMatch?.[1]) {
      return Number(hashMatch[1]);
    }
    return null;
  } catch {
    return null;
  }
}

export function buildSafeSpreadsheetUrl(
  spreadsheetId: string,
  gid?: number
): string {
  const base = `https://docs.google.com/spreadsheets/d/${encodeURIComponent(spreadsheetId)}/edit`;
  return typeof gid === 'number' && Number.isFinite(gid)
    ? `${base}#gid=${gid}`
    : base;
}

export function escapeSheetNameForA1(sheetName: string): string {
  const safe = sheetName.trim() || 'Sheet1';
  return `'${safe.replace(/'/g, "''")}'`;
}

export function columnIndexToLetter(index: number): string {
  let n = Math.max(0, Math.floor(index)) + 1;
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function normalizeLocationKey(value: string): string {
  return value.trim().replace(/\s+/g, ' ').toUpperCase();
}

function formatRecordToRow(
  record: PlacementRecord,
  includeExtraColumns: boolean
): (string | number)[] {
  if (includeExtraColumns) {
    return [
      record.oldLocation,
      record.newLocation,
      record.rack,
      record.row,
      record.box,
      new Date(record.placedAt).toLocaleString('id-ID'),
    ];
  }
  return [record.oldLocation, record.newLocation];
}

function getHeaderRow(includeExtraColumns: boolean): string[] {
  return includeExtraColumns
    ? ['FORMAT LAMA', 'FORMAT BARU', 'RAK', 'ROW', 'BOX', 'WAKTU SIMPAN']
    : ['FORMAT LAMA', 'FORMAT BARU'];
}

function detectSheetColumnMapping(headerRow: string[] | undefined): {
  hasHeader: boolean;
  oldColIdx: number;
  newColIdx: number;
  hasLeadingNoColumn: boolean;
} {
  if (!headerRow || headerRow.length === 0) {
    return {
      hasHeader: false,
      oldColIdx: 0,
      newColIdx: 1,
      hasLeadingNoColumn: false,
    };
  }

  const detectedOld = headerRow.findIndex((cell) =>
    /^(format\s*lama|lokasi\s*(simpan\s*)?lama|kode\s*lama)$/i.test(
      String(cell ?? '').trim()
    )
  );
  const detectedNew = headerRow.findIndex((cell) =>
    /^(format\s*baru|lokasi\s*(simpan\s*)?baru|kode\s*baru)$/i.test(
      String(cell ?? '').trim()
    )
  );

  if (detectedOld >= 0 || detectedNew >= 0) {
    const oldColIdx =
      detectedOld >= 0 ? detectedOld : Math.max(0, detectedNew - 1);
    const newColIdx = detectedNew >= 0 ? detectedNew : oldColIdx + 1;
    const firstHeader = String(headerRow[0] ?? '').trim();
    const hasLeadingNoColumn =
      oldColIdx === 1 && /^(no\.?|nomor|#)$/i.test(firstHeader);
    return {
      hasHeader: true,
      oldColIdx,
      newColIdx,
      hasLeadingNoColumn,
    };
  }

  return {
    hasHeader: false,
    oldColIdx: 0,
    newColIdx: 1,
    hasLeadingNoColumn: false,
  };
}

/**
 * Pure planner that computes which rows to update in place (matching FORMAT LAMA)
 * and which rows to append at the bottom of the connected Google Sheet.
 * Automatically adapts to spreadsheets whose FORMAT LAMA / FORMAT BARU columns
 * are offset (e.g. NO | LOKASI SIMPAN LAMA | LOKASI SIMPAN BARU).
 */
export function planSpreadsheetRowSync(
  existingRows: string[][],
  records: PlacementRecord[],
  includeExtraColumns: boolean
): SpreadsheetSyncPlan {
  const ordered = [...records].sort((a, b) => a.sequence - b.sequence);
  const headerRow = getHeaderRow(includeExtraColumns);

  const nonEmptyRows = existingRows.filter((row) =>
    row.some((cell) => String(cell ?? '').trim() !== '')
  );
  const hasAnyContent = nonEmptyRows.length > 0;
  const colMap = hasAnyContent
    ? detectSheetColumnMapping(existingRows[0])
    : {
        hasHeader: false,
        oldColIdx: 0,
        newColIdx: 1,
        hasLeadingNoColumn: false,
      };
  const needsHeader = !hasAnyContent;

  const { oldColIdx, newColIdx, hasHeader, hasLeadingNoColumn } = colMap;
  const isStandardAdjacentLayout = oldColIdx === 0 && newColIdx === 1;

  const existingByOldLocation = new Map<
    string,
    { rowNumber: number; rowValues: string[] }
  >();

  let existingDataRowCount = 0;
  const startIndex = hasHeader ? 1 : 0;

  for (let i = startIndex; i < existingRows.length; i++) {
    const row = existingRows[i] ?? [];
    const oldVal = String(row[oldColIdx] ?? '').trim();
    if (!oldVal) continue;
    existingDataRowCount += 1;
    const key = normalizeLocationKey(oldVal);
    existingByOldLocation.set(key, {
      rowNumber: i + 1, // 1-indexed Google Sheets row number
      rowValues: row.map((v) => String(v ?? '')),
    });
  }

  const updates: PlannedRowUpdate[] = [];
  const appends: PlannedRowAppend[] = [];
  let unchangedCount = 0;
  let nextRowSequence = existingDataRowCount + 1;

  const maxTargetIdx = isStandardAdjacentLayout
    ? includeExtraColumns
      ? 5
      : 1
    : Math.max(oldColIdx, newColIdx);

  for (const record of ordered) {
    const key = normalizeLocationKey(record.oldLocation);
    const existing = existingByOldLocation.get(key);

    if (existing) {
      const existingNew = (existing.rowValues[newColIdx] ?? '').trim();
      const targetNew = record.newLocation.trim();

      if (isStandardAdjacentLayout) {
        const targetValues = formatRecordToRow(record, includeExtraColumns);
        const extraChanged =
          includeExtraColumns &&
          (String(existing.rowValues[2] ?? '').trim() !==
            String(targetValues[2]) ||
            String(existing.rowValues[3] ?? '').trim() !==
              String(targetValues[3]) ||
            String(existing.rowValues[4] ?? '').trim() !==
              String(targetValues[4]));

        if (existingNew !== targetNew || extraChanged) {
          updates.push({
            rowNumber: existing.rowNumber,
            startColLetter: 'A',
            endColLetter: includeExtraColumns ? 'F' : 'B',
            values: targetValues,
            oldLocation: record.oldLocation,
            newLocation: record.newLocation,
          });
        } else {
          unchangedCount += 1;
        }
      } else if (existingNew !== targetNew) {
        const colLetter = columnIndexToLetter(newColIdx);
        updates.push({
          rowNumber: existing.rowNumber,
          startColLetter: colLetter,
          endColLetter: colLetter,
          values: [record.newLocation],
          oldLocation: record.oldLocation,
          newLocation: record.newLocation,
        });
      } else {
        unchangedCount += 1;
      }
    } else if (isStandardAdjacentLayout) {
      appends.push({
        values: formatRecordToRow(record, includeExtraColumns),
        oldLocation: record.oldLocation,
        newLocation: record.newLocation,
      });
    } else {
      const paddedRow: (string | number)[] = Array.from(
        { length: maxTargetIdx + 1 },
        () => ''
      );
      if (hasLeadingNoColumn) {
        paddedRow[0] = nextRowSequence++;
      }
      paddedRow[oldColIdx] = record.oldLocation;
      paddedRow[newColIdx] = record.newLocation;
      appends.push({
        values: paddedRow,
        oldLocation: record.oldLocation,
        newLocation: record.newLocation,
      });
    }
  }

  return {
    needsHeader,
    headerRow,
    oldColIdx,
    newColIdx,
    appendStartColLetter: 'A',
    appendEndColLetter: columnIndexToLetter(maxTargetIdx),
    updates,
    appends,
    unchangedCount,
    existingDataRowCount,
  };
}

/**
 * Computes how many records in the local Tabel Hasil have not yet been written
 * (or have a different FORMAT BARU) compared to the connected Google Sheet snapshot.
 */
export function computePendingGoogleSheetSyncCount(
  combinedRecords: PlacementRecord[],
  externalLocks: LockedLocationInfo[],
  isConnectedToGoogleSheet: boolean
): PendingGoogleSheetSyncStats {
  if (!isConnectedToGoogleSheet || combinedRecords.length === 0) {
    return { pendingNewCount: 0, pendingUpdateCount: 0, totalPendingCount: 0 };
  }

  const sheetMap = new Map<string, string>();
  for (const ext of externalLocks) {
    if (ext.source !== 'google_sheets') continue;
    const key = normalizeLocationKey(ext.oldLocation);
    if (!key) continue;
    sheetMap.set(key, ext.newLocation.trim().toUpperCase());
  }

  let pendingNewCount = 0;
  let pendingUpdateCount = 0;

  for (const record of combinedRecords) {
    const key = normalizeLocationKey(record.oldLocation);
    if (!key) continue;
    const sheetNewLoc = sheetMap.get(key);
    if (sheetNewLoc === undefined) {
      pendingNewCount += 1;
    } else if (sheetNewLoc !== record.newLocation.trim().toUpperCase()) {
      pendingUpdateCount += 1;
    }
  }

  return {
    pendingNewCount,
    pendingUpdateCount,
    totalPendingCount: pendingNewCount + pendingUpdateCount,
  };
}

async function googleApiFetch<T>(
  url: string,
  accessToken: string,
  init?: RequestInit
): Promise<T> {
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(init?.body ? { 'Content-Type': 'application/json' } : {}),
      ...(init?.headers ?? {}),
    },
  });

  if (res.status === 401) {
    throw new WorkspaceAuthExpiredError();
  }

  const text = await res.text();
  let data: unknown = null;
  if (text) {
    try {
      data = JSON.parse(text);
    } catch {
      data = { error: { message: text } };
    }
  }

  if (!res.ok) {
    const apiMessage =
      (data as { error?: { message?: string } })?.error?.message ||
      `Permintaan Google API gagal (HTTP ${res.status}).`;

    // Only expire/clear auth token if OAuth scopes themselves are insufficient
    if (
      res.status === 403 &&
      /insufficient.*scope|ACCESS_TOKEN_SCOPE_INSUFFICIENT/i.test(apiMessage)
    ) {
      throw new WorkspaceAuthExpiredError(
        'Izin akses Google Sheets belum lengkap. Silakan masuk kembali dan setujui akses Google Sheets.'
      );
    }

    // Keep user signed in if this specific spreadsheet is view-only / unshared
    if (res.status === 403) {
      throw new Error(
        'Akun Google Anda tidak memiliki izin Edit pada spreadsheet ini. Pastikan file dibagikan sebagai Editor atau buat spreadsheet baru.'
      );
    }

    if (res.status === 404) {
      throw new Error(
        'Spreadsheet tidak ditemukan atau sudah dihapus. Periksa kembali link Google Spreadsheet tujuan.'
      );
    }

    throw new Error(apiMessage);
  }

  return data as T;
}

export async function listDriveSpreadsheets(
  accessToken: string
): Promise<DriveSpreadsheetItem[]> {
  const query = encodeURIComponent(
    "mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
  );
  const fields = encodeURIComponent(
    'files(id,name,modifiedTime,webViewLink)'
  );
  const url = `https://www.googleapis.com/drive/v3/files?q=${query}&orderBy=modifiedTime%20desc&pageSize=25&fields=${fields}`;

  const data = await googleApiFetch<{ files?: DriveSpreadsheetItem[] }>(
    url,
    accessToken
  );
  return (data.files ?? []).filter((f) => SPREADSHEET_ID_PATTERN.test(f.id));
}

export async function fetchSpreadsheetMetadata(
  spreadsheetId: string,
  accessToken: string,
  preferredSheetName?: string,
  preferredGid?: number | null
): Promise<ConnectedGoogleSheet> {
  const fields = encodeURIComponent(
    'spreadsheetId,spreadsheetUrl,properties.title,sheets.properties(sheetId,title)'
  );
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}?fields=${fields}`;
  const data = await googleApiFetch<{
    spreadsheetId: string;
    spreadsheetUrl?: string;
    properties?: { title?: string };
    sheets?: Array<{ properties?: { sheetId?: number; title?: string } }>;
  }>(url, accessToken);

  const sheetEntries = (data.sheets ?? [])
    .map((s) => ({
      sheetId: s.properties?.sheetId,
      title: s.properties?.title?.trim() ?? '',
    }))
    .filter((s) => Boolean(s.title));

  const resolvedSheets =
    sheetEntries.length > 0 ? sheetEntries.map((s) => s.title) : ['Sheet1'];

  const matchedByGid =
    typeof preferredGid === 'number'
      ? sheetEntries.find((s) => s.sheetId === preferredGid)
      : undefined;

  const selectedSheet =
    matchedByGid?.title ||
    (preferredSheetName && resolvedSheets.includes(preferredSheetName)
      ? preferredSheetName
      : resolvedSheets[0]);

  const resolvedGid = sheetEntries.find(
    (s) => s.title === selectedSheet
  )?.sheetId;

  return {
    spreadsheetId: data.spreadsheetId,
    spreadsheetUrl: buildSafeSpreadsheetUrl(data.spreadsheetId, resolvedGid),
    title: data.properties?.title || 'Spreadsheet Tanpa Judul',
    sheetName: selectedSheet,
    availableSheets: resolvedSheets,
    preferredGid: resolvedGid,
  };
}

export async function createRelocationSpreadsheet(
  accessToken: string,
  title = 'HASIL_LOKASI_SIMPAN_LAMA_DAN_BARU'
): Promise<ConnectedGoogleSheet> {
  const url = 'https://sheets.googleapis.com/v4/spreadsheets';
  const data = await googleApiFetch<{
    spreadsheetId: string;
    properties?: { title?: string };
    sheets?: Array<{ properties?: { sheetId?: number; title?: string } }>;
  }>(url, accessToken, {
    method: 'POST',
    body: JSON.stringify({
      properties: { title },
      sheets: [{ properties: { title: 'Hasil Relokasi' } }],
    }),
  });

  const sheetName =
    data.sheets?.[0]?.properties?.title || 'Hasil Relokasi';
  const sheetId = data.sheets?.[0]?.properties?.sheetId;

  // Initialize header row immediately so the spreadsheet is ready for inspection
  const headerRange = encodeURIComponent(
    `${escapeSheetNameForA1(sheetName)}!A1:B1`
  );
  await googleApiFetch(
    `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(data.spreadsheetId)}/values/${headerRange}?valueInputOption=RAW`,
    accessToken,
    {
      method: 'PUT',
      body: JSON.stringify({
        values: [['FORMAT LAMA', 'FORMAT BARU']],
      }),
    }
  );

  return {
    spreadsheetId: data.spreadsheetId,
    spreadsheetUrl: buildSafeSpreadsheetUrl(data.spreadsheetId, sheetId),
    title: data.properties?.title || title,
    sheetName,
    availableSheets: [sheetName],
    preferredGid: sheetId,
  };
}

export async function readSheetRows(
  spreadsheetId: string,
  sheetName: string,
  accessToken: string
): Promise<string[][]> {
  const range = encodeURIComponent(
    `${escapeSheetNameForA1(sheetName)}!A:Z`
  );
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(spreadsheetId)}/values/${range}`;
  const data = await googleApiFetch<{ values?: string[][] }>(url, accessToken);
  return data.values ?? [];
}

/**
 * Executes an update-or-append sync against the connected Google Spreadsheet.
 * Uses valueInputOption=RAW to preserve leading zeros and prevent formula injection.
 */
export async function syncRecordsToGoogleSheet(
  target: ConnectedGoogleSheet,
  records: PlacementRecord[],
  includeExtraColumns: boolean,
  accessToken: string
): Promise<SpreadsheetSyncExecutionResult> {
  const existingRows = await readSheetRows(
    target.spreadsheetId,
    target.sheetName,
    accessToken
  );

  const plan = planSpreadsheetRowSync(
    existingRows,
    records,
    includeExtraColumns
  );

  const escapedSheet = escapeSheetNameForA1(target.sheetName);

  // 1. Write header row if sheet is completely empty
  let headerWritten = false;
  if (plan.needsHeader) {
    const headerRange = encodeURIComponent(
      `${escapedSheet}!A1:${plan.appendEndColLetter}1`
    );
    await googleApiFetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(target.spreadsheetId)}/values/${headerRange}?valueInputOption=RAW`,
      accessToken,
      {
        method: 'PUT',
        body: JSON.stringify({
          values: [plan.headerRow],
        }),
      }
    );
    headerWritten = true;
  }

  // 2. Batch-update any existing rows whose FORMAT BARU changed
  if (plan.updates.length > 0) {
    const batchUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(target.spreadsheetId)}/values:batchUpdate`;
    await googleApiFetch(batchUrl, accessToken, {
      method: 'POST',
      body: JSON.stringify({
        valueInputOption: 'RAW',
        data: plan.updates.map((u) => ({
          range: `${escapedSheet}!${u.startColLetter}${u.rowNumber}:${u.endColLetter}${u.rowNumber}`,
          values: [u.values],
        })),
      }),
    });
  }

  // 3. Append all brand-new rows below existing data
  if (plan.appends.length > 0) {
    const appendRange = encodeURIComponent(
      `${escapedSheet}!${plan.appendStartColLetter}:${plan.appendEndColLetter}`
    );
    const appendUrl = `https://sheets.googleapis.com/v4/spreadsheets/${encodeURIComponent(target.spreadsheetId)}/values/${appendRange}:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS`;
    await googleApiFetch(appendUrl, accessToken, {
      method: 'POST',
      body: JSON.stringify({
        values: plan.appends.map((a) => a.values),
      }),
    });
  }

  return {
    updatedCount: plan.updates.length,
    appendedCount: plan.appends.length,
    unchangedCount: plan.unchangedCount,
    headerWritten,
    totalSheetRows: plan.existingDataRowCount + plan.appends.length,
  };
}

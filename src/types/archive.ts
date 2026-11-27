export interface RackConfig {
  prefix: string;
  rackCount: number;
  rowsPerRack: number;
  boxesPerRow: number;
}

export interface BoxCoordinate {
  locationCode: string;
  rack: number;
  row: number;
  box: number;
}

export interface PlacementRecord {
  id: string;
  oldLocation: string;
  newLocation: string;
  locationCode: string;
  rack: number;
  row: number;
  box: number;
  sequence: number;
  placedAt: string;
  savedToExcel?: boolean;
}

export interface SpreadsheetItem {
  oldLocation: string;
  normalizedKey: string;
  rowNumber: number;
  description?: string;
}

export interface ImportedSpreadsheet {
  fileName: string;
  sheetName: string;
  columns: string[];
  selectedColumn: string;
  rawRows: Record<string, string>[];
  items: SpreadsheetItem[];
  importedAt: string;
}

export interface SimulationSnapshot {
  draftPlacements: PlacementRecord[];
  savedRecords: PlacementRecord[];
  nextSequence: number;
}

export interface LockedLocationInfo {
  oldLocation: string;
  normalizedKey: string;
  newLocation: string;
  source: 'draft' | 'saved' | 'google_sheets' | 'external_file';
  coordinate?: BoxCoordinate;
}

export type BackupHistoryCategory = 'template' | 'result';

export interface BackupHistoryEntry {
  id: string;
  category: BackupHistoryCategory;
  actionLabel: string;
  createdAt: string;
  summary: string;
  fileName?: string;
  itemCount: number;
  lockedCount?: number;
  includeExtraColumns?: boolean;
  targetSheetTitle?: string;
  targetSheetTab?: string;
  targetSheetUrl?: string;
  templateSnapshot?: ImportedSpreadsheet;
  recordsSnapshot?: PlacementRecord[];
}

export interface SimulationState {
  config: RackConfig;
  /** Placements currently shown in the active 1-rack simulation */
  draftPlacements: PlacementRecord[];
  /** Cumulative placements committed via the "Simpan" button for Excel/Sheets export */
  savedRecords: PlacementRecord[];
  /** Existing old->new mappings read back from connected Google Sheets or local result Excel to lock codes & drive dashboard sync */
  externalLockedLocations: LockedLocationInfo[];
  externalResultSourceName?: string | null;
  externalResultSourceType?: 'google_sheets' | 'local_file' | null;
  nextSequence: number;
  history: SimulationSnapshot[];
  spreadsheet: ImportedSpreadsheet | null;
  allowUnlistedWithSpreadsheet: boolean;
  includeExtraExcelColumns: boolean;
}

export interface SimulationMetrics {
  rackTotalBoxes: number;
  rackFilledBoxes: number;
  rackEmptyBoxes: number;
  unsavedDraftCount: number;
  savedExcelCount: number;
  /** Deduplicated total rows across local saved Excel + connected Google Sheets / loaded result file */
  spreadsheetResultsTotal: number;
  /** Count of rows read directly from connected Google Sheets or loaded result Excel file */
  externalResultsCount: number;
  totalRelocatedCount: number;
  /** Template-specific synchronized progress metrics (null when no template is loaded) */
  templateTotal: number | null;
  templateRegisteredCount: number | null;
  templateDraftOnlyCount: number | null;
  templateSimulatedCount: number | null;
  templateRemainingCount: number | null;
  /** Registered result codes in Spreadsheet Hasil that are not present in the active template */
  unmatchedResultCount: number;
  spreadsheetTotal: number | null;
  spreadsheetRemaining: number | null;
  registeredProgressPercent: number;
  draftProgressPercent: number;
  progressPercent: number;
  progressLabel: string;
}

export type PlacementValidationResult =
  | { status: 'valid'; normalizedOldLocation: string }
  | { status: 'empty'; message: string }
  | { status: 'not_in_spreadsheet'; message: string }
  | {
      status: 'duplicate_old_location';
      message: string;
      existingPlacement?: PlacementRecord;
      lockedNewLocation: string;
      inSavedExcel: boolean;
      inGoogleSheets?: boolean;
    }
  | {
      status: 'same_value_in_box';
      message: string;
      existingPlacement: PlacementRecord;
    }
  | {
      status: 'needs_replace_confirmation';
      message: string;
      currentPlacement: PlacementRecord;
      normalizedOldLocation: string;
    };

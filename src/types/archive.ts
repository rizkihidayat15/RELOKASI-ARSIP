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

export interface SimulationState {
  config: RackConfig;
  /** Placements currently shown in the active 1-rack simulation */
  draftPlacements: PlacementRecord[];
  /** Cumulative placements committed via the "Simpan" button for Excel export */
  savedRecords: PlacementRecord[];
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
  totalRelocatedCount: number;
  spreadsheetTotal: number | null;
  spreadsheetRemaining: number | null;
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
      existingPlacement: PlacementRecord;
      inSavedExcel: boolean;
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

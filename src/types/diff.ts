/// 差分比較の種類
export type DiffType = "added" | "removed" | "modified" | "unchanged";

/// 比較対象の設定
export interface DiffOptions {
  compareValues: boolean;
  compareFormats: boolean;
  compareFormulas: boolean;
  compareComments: boolean;
}

/// セル差分情報
export interface CellDiff {
  row: number;
  col: number;
  type: DiffType;
  oldValue?: any;
  newValue?: any;
  oldFormula?: string;
  newFormula?: string;
  formatChanges?: {
    font?: boolean;
    background?: boolean;
    border?: boolean;
    numberFormat?: boolean;
    alignment?: boolean;
  };
  oldFormat?: any;
  newFormat?: any;
  oldComment?: string;
  newComment?: string;
}

/// シート差分情報
export interface SheetDiff {
  sheetName: string;
  type: DiffType;
  cells: CellDiff[];
  summary: {
    totalChanges: number;
    addedCells: number;
    removedCells: number;
    modifiedCells: number;
  };
}

/// ワークブック差分情報
export interface WorkbookDiff {
  fileName1: string;
  fileName2: string;
  sheets: SheetDiff[];
  summary: {
    totalSheets: number;
    addedSheets: number;
    removedSheets: number;
    modifiedSheets: number;
    totalCellChanges: number;
  };
}

/// エラー情報の型定義
export interface ErrorInfo {
  message: string;
  details?: string;
  type: "file" | "format" | "read" | "compare" | "unknown";
  suggestions?: string[];
}

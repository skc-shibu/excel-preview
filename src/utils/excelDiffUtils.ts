import * as GC from "@grapecity/spread-sheets";
import {
  CellDiff,
  SheetDiff,
  WorkbookDiff,
  DiffOptions,
  DiffType,
} from "../types/diff";

/// 比較から除外するシート名のリスト
const EXCLUDED_SHEET_NAMES = [
  "Evaluation Version",
  "evaluation version",
  "EVALUATION VERSION",
];

/// シート名が除外対象かどうかを判定
const isExcludedSheet = (sheetName: string): boolean => {
  return EXCLUDED_SHEET_NAMES.some((excludedName) =>
    sheetName.toLowerCase().includes(excludedName.toLowerCase())
  );
};

/// 2つのワークブックを比較して差分を生成
export const compareWorkbooks = (
  workbook1: GC.Spread.Sheets.Workbook,
  workbook2: GC.Spread.Sheets.Workbook,
  fileName1: string,
  fileName2: string,
  options: DiffOptions
): WorkbookDiff => {
  const sheets1 = getAllSheets(workbook1);
  const sheets2 = getAllSheets(workbook2);

  const sheetDiffs: SheetDiff[] = [];
  const allSheetNames: string[] = [];

  // シート名を収集
  sheets1.forEach((s) => {
    if (!allSheetNames.includes(s.name)) {
      allSheetNames.push(s.name);
    }
  });
  sheets2.forEach((s) => {
    if (!allSheetNames.includes(s.name)) {
      allSheetNames.push(s.name);
    }
  });

  let totalCellChanges = 0;
  let addedSheets = 0;
  let removedSheets = 0;
  let modifiedSheets = 0;

  for (let i = 0; i < allSheetNames.length; i++) {
    const sheetName = allSheetNames[i];
    const sheet1 = sheets1.find((s) => s.name === sheetName);
    const sheet2 = sheets2.find((s) => s.name === sheetName);

    let sheetDiff: SheetDiff;

    if (!sheet1 && sheet2) {
      // シートが追加された
      sheetDiff = {
        sheetName,
        type: "added",
        cells: [],
        summary: {
          totalChanges: 0,
          addedCells: 0,
          removedCells: 0,
          modifiedCells: 0,
        },
      };
      addedSheets++;
    } else if (sheet1 && !sheet2) {
      // シートが削除された
      sheetDiff = {
        sheetName,
        type: "removed",
        cells: [],
        summary: {
          totalChanges: 0,
          addedCells: 0,
          removedCells: 0,
          modifiedCells: 0,
        },
      };
      removedSheets++;
    } else if (sheet1 && sheet2) {
      // シートが存在する - セル比較を実行
      sheetDiff = compareSheets(
        sheet1.worksheet,
        sheet2.worksheet,
        sheetName,
        options
      );
      if (sheetDiff.summary.totalChanges > 0) {
        modifiedSheets++;
      }
      totalCellChanges += sheetDiff.summary.totalChanges;
    } else {
      continue;
    }

    sheetDiffs.push(sheetDiff);
  }

  return {
    fileName1,
    fileName2,
    sheets: sheetDiffs,
    summary: {
      totalSheets: allSheetNames.length,
      addedSheets,
      removedSheets,
      modifiedSheets,
      totalCellChanges,
    },
  };
};

/// 2つのシートを比較して差分を生成
const compareSheets = (
  sheet1: GC.Spread.Sheets.Worksheet,
  sheet2: GC.Spread.Sheets.Worksheet,
  sheetName: string,
  options: DiffOptions
): SheetDiff => {
  const cellDiffs: CellDiff[] = [];

  // 使用範囲を取得（簡易版）
  const maxRow = Math.max(sheet1.getRowCount(), sheet2.getRowCount());
  const maxCol = Math.max(sheet1.getColumnCount(), sheet2.getColumnCount());

  let addedCells = 0;
  let removedCells = 0;
  let modifiedCells = 0;

  // セル単位で比較（範囲を制限）
  const checkRows = Math.min(maxRow, 100); // 最大100行まで
  const checkCols = Math.min(maxCol, 50); // 最大50列まで

  for (let row = 0; row < checkRows; row++) {
    for (let col = 0; col < checkCols; col++) {
      const cellDiff = compareCells(sheet1, sheet2, row, col, options);
      if (cellDiff && cellDiff.type !== "unchanged") {
        cellDiffs.push(cellDiff);

        switch (cellDiff.type) {
          case "added":
            addedCells++;
            break;
          case "removed":
            removedCells++;
            break;
          case "modified":
            modifiedCells++;
            break;
        }
      }
    }
  }

  return {
    sheetName,
    type: cellDiffs.length > 0 ? "modified" : "unchanged",
    cells: cellDiffs,
    summary: {
      totalChanges: cellDiffs.length,
      addedCells,
      removedCells,
      modifiedCells,
    },
  };
};

/// 2つのセルを比較して差分を生成
const compareCells = (
  sheet1: GC.Spread.Sheets.Worksheet,
  sheet2: GC.Spread.Sheets.Worksheet,
  row: number,
  col: number,
  options: DiffOptions
): CellDiff | null => {
  try {
    const value1 = sheet1.getValue(row, col);
    const value2 = sheet2.getValue(row, col);
    const formula1 = sheet1.getFormula(row, col);
    const formula2 = sheet2.getFormula(row, col);

    let hasChanges = false;
    let diffType: DiffType = "unchanged";

    // 値の比較
    const hasValue1 = value1 !== null && value1 !== undefined && value1 !== "";
    const hasValue2 = value2 !== null && value2 !== undefined && value2 !== "";

    if (options.compareValues) {
      if (!hasValue1 && hasValue2) {
        diffType = "added";
        hasChanges = true;
      } else if (hasValue1 && !hasValue2) {
        diffType = "removed";
        hasChanges = true;
      } else if (hasValue1 && hasValue2 && value1 !== value2) {
        diffType = "modified";
        hasChanges = true;
      }
    }

    // 数式の比較
    if (options.compareFormulas && formula1 !== formula2) {
      diffType = diffType === "unchanged" ? "modified" : diffType;
      hasChanges = true;
    }

    // 書式の比較（簡易版）
    let formatChanges: CellDiff["formatChanges"];
    if (options.compareFormats) {
      formatChanges = compareFormats(sheet1, sheet2, row, col);
      if (
        formatChanges &&
        Object.values(formatChanges).some((changed) => changed)
      ) {
        diffType = diffType === "unchanged" ? "modified" : diffType;
        hasChanges = true;
      }
    }

    if (!hasChanges) {
      return null;
    }

    return {
      row,
      col,
      type: diffType,
      oldValue: value1,
      newValue: value2,
      oldFormula: formula1 || undefined,
      newFormula: formula2 || undefined,
      formatChanges,
    };
  } catch (error) {
    console.warn(`セル比較エラー (${row}, ${col}):`, error);
    return null;
  }
};

/// 書式の比較（簡易版）
const compareFormats = (
  sheet1: GC.Spread.Sheets.Worksheet,
  sheet2: GC.Spread.Sheets.Worksheet,
  row: number,
  col: number
): CellDiff["formatChanges"] | undefined => {
  try {
    const style1 = sheet1.getStyle(row, col);
    const style2 = sheet2.getStyle(row, col);

    if (!style1 && !style2) {
      return undefined;
    }

    const changes: CellDiff["formatChanges"] = {};

    // 背景色の比較
    if (style1?.backColor !== style2?.backColor) {
      changes.background = true;
    }

    // フォントの比較（簡易）
    const font1 = JSON.stringify(style1?.font || {});
    const font2 = JSON.stringify(style2?.font || {});
    if (font1 !== font2) {
      changes.font = true;
    }

    return Object.keys(changes).length > 0 ? changes : undefined;
  } catch {
    return undefined;
  }
};

/// ワークブックからすべてのシートを取得
const getAllSheets = (
  workbook: GC.Spread.Sheets.Workbook
): Array<{ name: string; worksheet: GC.Spread.Sheets.Worksheet }> => {
  const sheets: Array<{ name: string; worksheet: GC.Spread.Sheets.Worksheet }> =
    [];
  const sheetCount = workbook.getSheetCount();

  for (let i = 0; i < sheetCount; i++) {
    const worksheet = workbook.getSheet(i);
    if (worksheet) {
      const sheetName = worksheet.name();
      // 除外対象のシートでない場合のみ追加
      if (!isExcludedSheet(sheetName)) {
        sheets.push({
          name: sheetName,
          worksheet,
        });
      }
    }
  }

  return sheets;
};

/// セル位置を文字列に変換（例: A1, B2）
export const getCellAddress = (row: number, col: number): string => {
  let columnName = "";
  let tempCol = col;

  while (tempCol >= 0) {
    columnName = String.fromCharCode(65 + (tempCol % 26)) + columnName;
    tempCol = Math.floor(tempCol / 26) - 1;
  }

  return `${columnName}${row + 1}`;
};

/// 差分タイプに応じた色を取得
export const getDiffColor = (type: DiffType): string => {
  switch (type) {
    case "added":
      return "#d4edda"; // 緑
    case "removed":
      return "#f8d7da"; // 赤
    case "modified":
      return "#fff3cd"; // 黄
    default:
      return "transparent";
  }
};

/// 差分タイプに応じたアイコンを取得
export const getDiffIcon = (type: DiffType): string => {
  switch (type) {
    case "added":
      return "➕";
    case "removed":
      return "➖";
    case "modified":
      return "✏️";
    default:
      return "";
  }
};

/// 2つのワークブック（JSON）を比較して差分を生成
export const compareWorkbookJSON = (
  workbookJSON1: any,
  workbookJSON2: any,
  fileName1: string,
  fileName2: string,
  options: DiffOptions
): WorkbookDiff => {
  const sheets1 = getJSONSheets(workbookJSON1);
  const sheets2 = getJSONSheets(workbookJSON2);

  const sheetDiffs: SheetDiff[] = [];
  const allSheetNames: string[] = [];

  // シート名を収集
  sheets1.forEach((s) => {
    if (!allSheetNames.includes(s.name)) {
      allSheetNames.push(s.name);
    }
  });
  sheets2.forEach((s) => {
    if (!allSheetNames.includes(s.name)) {
      allSheetNames.push(s.name);
    }
  });

  let totalCellChanges = 0;
  let addedSheets = 0;
  let removedSheets = 0;
  let modifiedSheets = 0;

  for (let i = 0; i < allSheetNames.length; i++) {
    const sheetName = allSheetNames[i];
    const sheet1 = sheets1.find((s) => s.name === sheetName);
    const sheet2 = sheets2.find((s) => s.name === sheetName);

    let sheetDiff: SheetDiff;

    if (!sheet1 && sheet2) {
      // シートが追加された
      sheetDiff = {
        sheetName,
        type: "added",
        cells: [],
        summary: {
          totalChanges: 0,
          addedCells: 0,
          removedCells: 0,
          modifiedCells: 0,
        },
      };
      addedSheets++;
    } else if (sheet1 && !sheet2) {
      // シートが削除された
      sheetDiff = {
        sheetName,
        type: "removed",
        cells: [],
        summary: {
          totalChanges: 0,
          addedCells: 0,
          removedCells: 0,
          modifiedCells: 0,
        },
      };
      removedSheets++;
    } else if (sheet1 && sheet2) {
      // シートが存在する - セル比較を実行
      sheetDiff = compareSheetsJSON(
        sheet1.data,
        sheet2.data,
        sheetName,
        options
      );
      if (sheetDiff.summary.totalChanges > 0) {
        modifiedSheets++;
      }
      totalCellChanges += sheetDiff.summary.totalChanges;
    } else {
      continue;
    }

    sheetDiffs.push(sheetDiff);
  }

  return {
    fileName1,
    fileName2,
    sheets: sheetDiffs,
    summary: {
      totalSheets: allSheetNames.length,
      addedSheets,
      removedSheets,
      modifiedSheets,
      totalCellChanges,
    },
  };
};

/// JSONワークブックからシート情報を取得
const getJSONSheets = (
  workbookJSON: any
): Array<{ name: string; data: any }> => {
  console.log("Debug - getJSONSheets called with workbook:", workbookJSON);

  const sheets: Array<{ name: string; data: any }> = [];

  if (workbookJSON && workbookJSON.sheets) {
    console.log(
      "Debug - Found sheets object:",
      Object.keys(workbookJSON.sheets)
    );

    Object.keys(workbookJSON.sheets).forEach((sheetKey) => {
      const sheetData = workbookJSON.sheets[sheetKey];
      if (sheetData) {
        const sheetName = sheetData.name || sheetKey;
        console.log(
          `Debug - Processing sheet: ${sheetName} (excluded: ${isExcludedSheet(
            sheetName
          )})`
        );

        // 除外対象のシートでない場合のみ追加
        if (!isExcludedSheet(sheetName)) {
          sheets.push({
            name: sheetName,
            data: sheetData,
          });
        }
      }
    });
  } else {
    console.log(
      "Debug - No sheets found in workbook. Structure:",
      Object.keys(workbookJSON || {})
    );
  }

  console.log(
    `Debug - Returning ${sheets.length} sheets:`,
    sheets.map((s) => s.name)
  );
  return sheets;
};

/// 2つのJSONシートを比較
const compareSheetsJSON = (
  sheetData1: any,
  sheetData2: any,
  sheetName: string,
  options: DiffOptions
): SheetDiff => {
  console.log(`Debug - Comparing sheet: ${sheetName}`);
  console.log(`Debug - Options:`, options);

  const cellDiffs: CellDiff[] = [];

  // データテーブルが存在しない場合の処理
  if (!sheetData1?.data?.dataTable && !sheetData2?.data?.dataTable) {
    return {
      sheetName,
      type: "unchanged",
      cells: [],
      summary: {
        totalChanges: 0,
        addedCells: 0,
        removedCells: 0,
        modifiedCells: 0,
      },
    };
  }

  console.log(
    `Debug - Sheet1 dataTable keys (first 10):`,
    Object.keys(sheetData1?.data?.dataTable || {}).slice(0, 10)
  );
  if (
    sheetData1?.data?.dataTable &&
    Object.keys(sheetData1.data.dataTable).length > 0
  ) {
    const firstKey = Object.keys(sheetData1.data.dataTable)[0];
    console.log(
      `Debug - Sheet1 first cell data:`,
      sheetData1.data.dataTable[firstKey]
    );
  }

  console.log(
    `Debug - Sheet2 dataTable keys (first 10):`,
    Object.keys(sheetData2?.data?.dataTable || {}).slice(0, 10)
  );
  if (
    sheetData2?.data?.dataTable &&
    Object.keys(sheetData2.data.dataTable).length > 0
  ) {
    const firstKey = Object.keys(sheetData2.data.dataTable)[0];
    console.log(
      `Debug - Sheet2 first cell data:`,
      sheetData2.data.dataTable[firstKey]
    );
  }

  // 両方のシートから実際に存在するセルキーを収集
  const sheet1Keys = Object.keys(sheetData1?.data?.dataTable || {});
  const sheet2Keys = Object.keys(sheetData2?.data?.dataTable || {});

  // 全てのセルキーを統合
  const uniqueKeys = new Set([...sheet1Keys, ...sheet2Keys]);
  const allCellKeys = Array.from(uniqueKeys);

  console.log(`Debug - Total unique cell keys to compare:`, allCellKeys.length);
  console.log(`Debug - Sample cell keys:`, allCellKeys.slice(0, 5));

  // 比較対象のセルを反復処理
  for (const cellKey of allCellKeys) {
    const row = parseInt(cellKey, 10);

    // 行データを取得
    const rowData1 = sheetData1?.data?.dataTable?.[cellKey];
    const rowData2 = sheetData2?.data?.dataTable?.[cellKey];

    if (!rowData1 && !rowData2) {
      continue; // 両方とも存在しない行はスキップ
    }

    // 行データが存在する場合、その行の全ての列をチェック
    const colKeys1 = Object.keys(rowData1 || {});
    const colKeys2 = Object.keys(rowData2 || {});
    const uniqueColKeys = new Set([...colKeys1, ...colKeys2]);
    const allColKeys = Array.from(uniqueColKeys);

    // デバッグ：最初の数行の詳細な比較ログ
    if (row < 5) {
      console.log(`Debug - Row ${row} comparison:`);
      console.log(`  - Sheet1 columns:`, colKeys1.slice(0, 10));
      console.log(`  - Sheet2 columns:`, colKeys2.slice(0, 10));
      console.log(`  - Total columns to compare:`, allColKeys.length);
    }

    for (const colKey of allColKeys) {
      const col = parseInt(colKey, 10);

      const cellData1 = rowData1?.[colKey];
      const cellData2 = rowData2?.[colKey];

      // デバッグ：最初の数セルの詳細なデータ
      if (row < 3 && col < 3) {
        console.log(`Debug - Cell [${row},${col}] data:`);
        console.log(`  - Sheet1:`, cellData1);
        console.log(`  - Sheet2:`, cellData2);
      }

      // 値の比較
      if (options.compareValues) {
        const value1 = getCellValueFromJSON(cellData1);
        const value2 = getCellValueFromJSON(cellData2);

        // デバッグ：値の詳細
        if (row < 3 && col < 3) {
          console.log(`  - Extracted values: "${value1}" vs "${value2}"`);
          console.log(`  - Value types: ${typeof value1} vs ${typeof value2}`);
          console.log(`  - Values equal:`, value1 === value2);
        }

        if (value1 !== value2) {
          console.log(
            `Debug - Found value difference at [${row},${col}]: "${value1}" vs "${value2}"`
          );

          cellDiffs.push({
            row,
            col,
            type: !cellData1 ? "added" : !cellData2 ? "removed" : "modified",
            oldValue: value1,
            newValue: value2,
          });
        }
      }

      // 数式の比較
      if (options.compareFormulas) {
        const formula1 = getCellFormulaFromJSON(cellData1);
        const formula2 = getCellFormulaFromJSON(cellData2);

        if (formula1 !== formula2) {
          console.log(
            `Debug - Found formula difference at [${row},${col}]: "${formula1}" vs "${formula2}"`
          );

          cellDiffs.push({
            row,
            col,
            type: !formula1 ? "added" : !formula2 ? "removed" : "modified",
            oldFormula: formula1 || undefined,
            newFormula: formula2 || undefined,
          });
        }
      }

      // 書式の比較
      if (options.compareFormats) {
        const format1 = getCellFormatFromJSON(cellData1);
        const format2 = getCellFormatFromJSON(cellData2);
        const formatChanges = compareFormatsJSON(format1, format2);

        if (
          formatChanges &&
          Object.values(formatChanges).some((changed) => changed)
        ) {
          console.log(
            `Debug - Found format difference at [${row},${col}]:`,
            formatChanges
          );
          console.log(`  - Old format:`, format1);
          console.log(`  - New format:`, format2);

          cellDiffs.push({
            row,
            col,
            type: "modified",
            formatChanges,
            oldFormat: format1,
            newFormat: format2,
          });
        }
      }
    }
  }

  console.log(
    `Debug - Sheet ${sheetName} comparison result: ${cellDiffs.length} differences found`
  );

  return {
    sheetName,
    type: cellDiffs.length > 0 ? "modified" : "unchanged",
    cells: cellDiffs,
    summary: {
      totalChanges: cellDiffs.length,
      addedCells: 0,
      removedCells: 0,
      modifiedCells: 0,
    },
  };
};

/// JSONからセル値を取得
const getCellValueFromJSON = (cellData: any): any => {
  if (!cellData) return null;

  // 値は cellData.value に格納されている
  return cellData.value ?? null;
};

/// JSONからセルの数式を取得
const getCellFormulaFromJSON = (cellData: any): string | null => {
  if (!cellData) return null;

  // 数式は cellData.formula に格納されている
  return cellData.formula ?? null;
};

/// JSONからセルの書式情報を取得
const getCellFormatFromJSON = (cellData: any): any => {
  if (!cellData || !cellData.style) return {};

  const style = cellData.style;
  return {
    // 数値書式
    numberFormat: style.formatter || style.numberFormat || null,
    // フォント関連
    fontFamily: style.fontFamily || null,
    fontSize: style.fontSize || null,
    fontWeight: style.fontWeight || null,
    fontStyle: style.fontStyle || null,
    textDecoration: style.textDecoration || null,
    foreColor: style.foreColor || style.color || null,
    // 背景色
    backColor: style.backColor || style.backgroundColor || null,
    // 罫線
    borderLeft: style.borderLeft || null,
    borderRight: style.borderRight || null,
    borderTop: style.borderTop || null,
    borderBottom: style.borderBottom || null,
    // 配置
    hAlign: style.hAlign || style.textAlign || null,
    vAlign: style.vAlign || style.verticalAlign || null,
  };
};

/// 2つの書式オブジェクトを比較
const compareFormatsJSON = (format1: any, format2: any): any => {
  const changes = {
    numberFormat: false,
    font: false,
    background: false,
    border: false,
    alignment: false,
  };

  // 数値書式の比較
  if (format1.numberFormat !== format2.numberFormat) {
    changes.numberFormat = true;
  }

  // フォント関連の比較
  if (
    format1.fontFamily !== format2.fontFamily ||
    format1.fontSize !== format2.fontSize ||
    format1.fontWeight !== format2.fontWeight ||
    format1.fontStyle !== format2.fontStyle ||
    format1.textDecoration !== format2.textDecoration ||
    format1.foreColor !== format2.foreColor
  ) {
    changes.font = true;
  }

  // 背景色の比較
  if (format1.backColor !== format2.backColor) {
    changes.background = true;
  }

  // 罫線の比較
  if (
    format1.borderLeft !== format2.borderLeft ||
    format1.borderRight !== format2.borderRight ||
    format1.borderTop !== format2.borderTop ||
    format1.borderBottom !== format2.borderBottom
  ) {
    changes.border = true;
  }

  // 配置の比較
  if (format1.hAlign !== format2.hAlign || format1.vAlign !== format2.vAlign) {
    changes.alignment = true;
  }

  return changes;
};

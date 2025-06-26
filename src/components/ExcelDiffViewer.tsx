import React, { useState, useRef, useEffect } from "react";
import { SpreadSheets, Worksheet } from "@grapecity/spread-sheets-react";
import * as GC from "@grapecity/spread-sheets";
import * as ExcelIO from "@grapecity/spread-excelio";
import { DiffOptions, WorkbookDiff, ErrorInfo } from "../types/diff";
import {
  compareWorkbookJSON,
  getDiffColor,
  getDiffIcon,
  getCellAddress,
} from "../utils/excelDiffUtils";
import "./ExcelDiffViewer.css";

/// ExcelDiffViewerコンポーネント - 2つのExcelファイルの差分比較機能を提供
const ExcelDiffViewer: React.FC = () => {
  const [file1Name, setFile1Name] = useState<string>("");
  const [file2Name, setFile2Name] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isComparing, setIsComparing] = useState<boolean>(false);
  const [error, setError] = useState<ErrorInfo | null>(null);
  const [showErrorDetails, setShowErrorDetails] = useState<boolean>(false);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const [showFullscreenSidebar, setShowFullscreenSidebar] =
    useState<boolean>(true);
  const [diffResult, setDiffResult] = useState<WorkbookDiff | null>(null);
  const [currentDiffIndex, setCurrentDiffIndex] = useState<number>(0);
  const [selectedCellInfo, setSelectedCellInfo] = useState<any>(null);
  const [isDiffExecuted, setIsDiffExecuted] = useState<boolean>(false);
  const [optionsChanged, setOptionsChanged] = useState<boolean>(false);

  const spreadRef1 = useRef<GC.Spread.Sheets.Workbook | null>(null);
  const spreadRef2 = useRef<GC.Spread.Sheets.Workbook | null>(null);
  const workbookData1 = useRef<any>(null);
  const workbookData2 = useRef<any>(null);
  const isSyncing = useRef<boolean>(false);

  const [diffOptions, setDiffOptions] = useState<DiffOptions>({
    compareValues: true,
    compareFormats: true,
    compareFormulas: false,
    compareComments: false,
  });

  /// ESCキーでの全画面解除
  useEffect(() => {
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape" && isFullscreen) {
        setIsFullscreen(false);
      }
    };

    document.addEventListener("keydown", handleKeyDown);
    return () => {
      document.removeEventListener("keydown", handleKeyDown);
    };
  }, [isFullscreen]);

  /// 全画面状態変更時のSpreadSheetsサイズ更新
  useEffect(() => {
    if (spreadRef1.current && spreadRef2.current) {
      const timeoutId = setTimeout(() => {
        try {
          spreadRef1.current!.refresh();
          spreadRef1.current!.invalidateLayout();
          spreadRef2.current!.refresh();
          spreadRef2.current!.invalidateLayout();
        } catch (error) {
          console.warn("SpreadSheets resize error:", error);
        }
      }, 150);

      return () => clearTimeout(timeoutId);
    }
  }, [isFullscreen]);

  /// サイドバー表示状態変更時のSpreadSheetsサイズ更新
  useEffect(() => {
    if (isFullscreen && spreadRef1.current && spreadRef2.current) {
      const timeoutId = setTimeout(() => {
        try {
          spreadRef1.current!.refresh();
          spreadRef1.current!.invalidateLayout();
          spreadRef2.current!.refresh();
          spreadRef2.current!.invalidateLayout();
        } catch (error) {
          console.warn("SpreadSheets sidebar resize error:", error);
        }
      }, 150);

      return () => clearTimeout(timeoutId);
    }
  }, [showFullscreenSidebar, isFullscreen]);

  /// SpreadSheetの基本設定
  const setupSpreadSheet = (spread: GC.Spread.Sheets.Workbook) => {
    spread.suspendPaint();
    spread.options.tabStripVisible = true;
    spread.options.newTabVisible = false;
    spread.options.tabNavigationVisible = true;
    spread.resumePaint();
  };

  /// シート同期とスクロール同期の設定
  const setupSheetSync = (
    spread: GC.Spread.Sheets.Workbook,
    workbookNumber: 1 | 2
  ) => {
    console.log(`Setting up sheet sync for workbook ${workbookNumber}`);

    // シート変更イベントの設定
    spread.bind(
      GC.Spread.Sheets.Events.ActiveSheetChanged,
      (sender: any, args: any) => {
        // 他方のワークブックのシートも同期
        const otherSpread =
          workbookNumber === 1 ? spreadRef2.current : spreadRef1.current;
        if (otherSpread && otherSpread !== sender) {
          const newSheetIndex = args.newSheet;
          const currentSheetIndex = otherSpread.getActiveSheetIndex();

          // 同じインデックスのシートが存在する場合のみ同期
          if (
            newSheetIndex !== currentSheetIndex &&
            newSheetIndex < otherSpread.getSheetCount()
          ) {
            try {
              otherSpread.setActiveSheetIndex(newSheetIndex);
            } catch (error) {
              console.warn("シート同期エラー:", error);
            }
          }
        }
      }
    );

    // スクロール同期イベントの設定（TopRowChangedとLeftColumnChangedを使用）
    spread.bind(
      GC.Spread.Sheets.Events.TopRowChanged,
      (sender: any, args: any) => {
        console.log(`TopRowChanged triggered - Workbook ${workbookNumber}`, {
          sheetName: args.sheetName,
          isSyncing: isSyncing.current,
          eventArgs: args,
        });

        if (isSyncing.current) {
          console.log("Skipping sync due to isSyncing flag");
          return;
        }

        const otherSpread =
          workbookNumber === 1 ? spreadRef2.current : spreadRef1.current;

        if (otherSpread) {
          const activeSheet = spread.getActiveSheet();
          const otherActiveSheet = otherSpread.getActiveSheet();

          if (activeSheet && otherActiveSheet) {
            // TopRowChangedイベントからの情報を使用
            const newTopRow =
              args.newTopRow !== undefined
                ? args.newTopRow
                : activeSheet.getViewportTopRow(0);
            const currentLeftColumn = activeSheet.getViewportLeftColumn(0);

            console.log(
              `Syncing vertical scroll from workbook ${workbookNumber}`,
              {
                newTopRow: newTopRow,
                currentLeftColumn: currentLeftColumn,
                eventNewTopRow: args.newTopRow,
                eventOldTopRow: args.oldTopRow,
                targetWorkbook: workbookNumber === 1 ? 2 : 1,
              }
            );

            isSyncing.current = true;
            try {
              console.log(
                `Syncing to row: ${newTopRow}, col: ${currentLeftColumn}`
              );

              // 複数の方法を試す
              try {
                // 方法1: showCellを使用
                otherActiveSheet.showCell(
                  newTopRow,
                  currentLeftColumn,
                  GC.Spread.Sheets.VerticalPosition.top,
                  GC.Spread.Sheets.HorizontalPosition.left
                );
                console.log("showCell method completed");
              } catch (showCellError) {
                console.warn("showCell failed, trying showRow:", showCellError);
                // 方法2: showRowを使用
                otherActiveSheet.showRow(
                  newTopRow,
                  GC.Spread.Sheets.VerticalPosition.top
                );
                console.log("showRow method completed");
              }

              console.log("Vertical sync completed successfully");
            } catch (error) {
              console.warn("垂直スクロール同期エラー:", error);
            } finally {
              setTimeout(() => {
                isSyncing.current = false;
                console.log("isSyncing flag reset to false");
              }, 50);
            }
          } else {
            console.warn("Active sheets not found for vertical sync");
          }
        } else {
          console.warn("Other spread not found for vertical sync");
        }
      }
    );

    // 水平スクロール同期
    spread.bind(
      GC.Spread.Sheets.Events.LeftColumnChanged,
      (sender: any, args: any) => {
        console.log(
          `LeftColumnChanged triggered - Workbook ${workbookNumber}`,
          {
            sheetName: args.sheetName,
            isSyncing: isSyncing.current,
            eventArgs: args,
          }
        );

        if (isSyncing.current) {
          console.log("Skipping sync due to isSyncing flag");
          return;
        }

        const otherSpread =
          workbookNumber === 1 ? spreadRef2.current : spreadRef1.current;

        if (otherSpread) {
          const activeSheet = spread.getActiveSheet();
          const otherActiveSheet = otherSpread.getActiveSheet();

          if (activeSheet && otherActiveSheet) {
            // LeftColumnChangedイベントからの情報を使用
            const newLeftColumn =
              args.newLeftCol !== undefined
                ? args.newLeftCol
                : activeSheet.getViewportLeftColumn(0);
            const currentTopRow = activeSheet.getViewportTopRow(0);

            console.log(
              `Syncing horizontal scroll from workbook ${workbookNumber}`,
              {
                newLeftColumn: newLeftColumn,
                currentTopRow: currentTopRow,
                eventNewLeftCol: args.newLeftCol,
                eventOldLeftCol: args.oldLeftCol,
                targetWorkbook: workbookNumber === 1 ? 2 : 1,
              }
            );

            isSyncing.current = true;
            try {
              console.log(
                `Syncing to row: ${currentTopRow}, col: ${newLeftColumn}`
              );

              // 複数の方法を試す
              try {
                // 方法1: showCellを使用
                otherActiveSheet.showCell(
                  currentTopRow,
                  newLeftColumn,
                  GC.Spread.Sheets.VerticalPosition.top,
                  GC.Spread.Sheets.HorizontalPosition.left
                );
                console.log("showCell method completed");
              } catch (showCellError) {
                console.warn(
                  "showCell failed, trying showColumn:",
                  showCellError
                );
                // 方法2: showColumnを使用
                otherActiveSheet.showColumn(
                  newLeftColumn,
                  GC.Spread.Sheets.HorizontalPosition.left
                );
                console.log("showColumn method completed");
              }

              console.log("Horizontal sync completed successfully");
            } catch (error) {
              console.warn("水平スクロール同期エラー:", error);
            } finally {
              setTimeout(() => {
                isSyncing.current = false;
                console.log("isSyncing flag reset to false");
              }, 50);
            }
          } else {
            console.warn("Active sheets not found for horizontal sync");
          }
        } else {
          console.warn("Other spread not found for horizontal sync");
        }
      }
    );

    console.log(`Sheet sync setup completed for workbook ${workbookNumber}`);

    // セル選択イベントの設定
    spread.bind(
      GC.Spread.Sheets.Events.SelectionChanged,
      (sender: any, args: any) => {
        if (diffResult) {
          const activeSheet = spread.getActiveSheet();
          const selection = activeSheet.getSelections()[0];
          if (selection) {
            const row = selection.row;
            const col = selection.col;
            const sheetName = activeSheet.name();

            // 選択されたセルの差分情報を検索
            const cellDiff = findCellDiff(diffResult, sheetName, row, col);
            setSelectedCellInfo(cellDiff);
          }
        }
      }
    );
  };

  /// セル差分情報を検索
  const findCellDiff = (
    diff: WorkbookDiff,
    sheetName: string,
    row: number,
    col: number
  ) => {
    const sheet = diff.sheets.find((s) => s.sheetName === sheetName);
    if (!sheet) return null;

    const cellDiff = sheet.cells.find((c) => c.row === row && c.col === col);
    if (!cellDiff) return null;

    return {
      ...cellDiff,
      sheetName,
      address: getCellAddress(row, col),
    };
  };

  /// SpreadSheets初期化完了時の処理
  const onWorkbook1Initialized = (spread: GC.Spread.Sheets.Workbook) => {
    console.log("Workbook 1 initialized");
    spreadRef1.current = spread;
    setupSpreadSheet(spread);
    setupSheetSync(spread, 1);
  };

  const onWorkbook2Initialized = (spread: GC.Spread.Sheets.Workbook) => {
    console.log("Workbook 2 initialized");
    spreadRef2.current = spread;
    setupSpreadSheet(spread);
    setupSheetSync(spread, 2);
  };

  /// エラー設定のヘルパー関数
  const setErrorInfo = (
    message: string,
    type: ErrorInfo["type"],
    details?: string,
    suggestions?: string[]
  ) => {
    setError({
      message,
      details,
      type,
      suggestions,
    });
  };

  /// ファイル選択時の処理
  const handleFile1Select = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!validateFile(file)) return;

    setError(null);
    setFile1Name(file.name);
    loadExcelFile(file, 1);
  };

  const handleFile2Select = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!validateFile(file)) return;

    setError(null);
    setFile2Name(file.name);
    loadExcelFile(file, 2);
  };

  /// ファイル検証
  const validateFile = (file: File): boolean => {
    // ファイルサイズのチェック（50MB制限 - 2ファイルなので制限を下げる）
    const maxSize = 50 * 1024 * 1024; // 50MB
    if (file.size > maxSize) {
      setErrorInfo(
        "ファイルサイズが大きすぎます",
        "file",
        `ファイルサイズ: ${(file.size / 1024 / 1024).toFixed(
          2
        )}MB（制限: 50MB）`,
        ["より小さなファイルを選択してください"]
      );
      return false;
    }

    // サポートされているファイル形式の確認
    const supportedFormats = [".xlsx", ".xls"];
    const fileExtension = file.name
      .toLowerCase()
      .substring(file.name.lastIndexOf("."));

    if (!supportedFormats.includes(fileExtension)) {
      setErrorInfo(
        "サポートされていないファイル形式です",
        "format",
        `ファイル形式: ${fileExtension}`,
        ["xlsx、xls形式のファイルを選択してください"]
      );
      return false;
    }

    return true;
  };

  /// Excelファイルの読み込み処理
  const loadExcelFile = async (file: File, fileNumber: 1 | 2) => {
    const spreadRef = fileNumber === 1 ? spreadRef1 : spreadRef2;

    if (!spreadRef.current) {
      setErrorInfo(
        "スプレッドシートが初期化されていません",
        "unknown",
        "コンポーネントの初期化エラー",
        ["ページを再読み込みしてください"]
      );
      return;
    }

    setIsLoading(true);

    try {
      const excelIO = new ExcelIO.IO();

      await new Promise<void>((resolve, reject) => {
        excelIO.open(
          file,
          (json: any) => {
            if (json.name === "Error") {
              reject(
                new Error(json.message || "ファイルの読み込みに失敗しました。")
              );
              return;
            }

            try {
              spreadRef.current!.fromJSON(json);
              if (fileNumber === 1) {
                workbookData1.current = json;
              } else {
                workbookData2.current = json;
              }
              resolve();
            } catch (error) {
              reject(error);
            }
          },
          (error: any) => {
            reject(
              new Error(error.message || "ファイルの読み込みに失敗しました。")
            );
          }
        );
      });

      // ファイル読み込み完了（自動実行は行わない）
    } catch (error) {
      console.error("Excel file loading error:", error);
      handleLoadError(error);
    } finally {
      setIsLoading(false);
    }
  };

  /// 読み込みエラーの処理
  const handleLoadError = (error: any) => {
    let errorMessage = "ファイルの読み込み中にエラーが発生しました";
    let errorDetails = "";
    let suggestions: string[] = [];

    if (error instanceof Error) {
      errorMessage = error.message;
      errorDetails = error.stack || error.toString();

      if (
        error.message.includes("password") ||
        error.message.includes("protected")
      ) {
        suggestions = [
          "パスワードで保護されたファイルはサポートされていません",
        ];
      } else if (
        error.message.includes("corrupt") ||
        error.message.includes("invalid")
      ) {
        suggestions = ["ファイルが破損している可能性があります"];
      } else {
        suggestions = [
          "別のファイルで試してください",
          "ブラウザを再読み込みして再度お試しください",
        ];
      }
    }

    setErrorInfo(errorMessage, "read", errorDetails, suggestions);
  };

  /// 差分比較の実行
  const performDiffComparison = async () => {
    if (!workbookData1.current || !workbookData2.current) {
      console.warn("Both workbooks must be loaded before comparison");
      return;
    }

    setIsComparing(true);
    setError(null);

    try {
      // 少し遅延を入れてローディング表示を確実に見せる
      await new Promise((resolve) => setTimeout(resolve, 100));

      const diff = compareWorkbookJSON(
        workbookData1.current,
        workbookData2.current,
        file1Name,
        file2Name,
        diffOptions
      );
      setDiffResult(diff);
      setCurrentDiffIndex(0);

      if (diff.sheets.length > 0) {
        highlightDifferences(diff);
      }
    } catch (error) {
      console.error("Diff comparison error:", error);
      setErrorInfo(
        "差分比較中にエラーが発生しました",
        "unknown",
        error instanceof Error ? error.message : String(error),
        ["ファイルを再読み込みして再度お試しください"]
      );
    } finally {
      setIsComparing(false);
      setOptionsChanged(false); // 比較完了時にオプション変更フラグをリセット
    }
  };

  /// 手動での差分比較実行
  const executeDiffComparison = async () => {
    if (!workbookData1.current || !workbookData2.current) {
      setErrorInfo(
        "両方のファイルを選択してください",
        "file",
        "差分比較を実行するには2つのExcelファイルが必要です",
        ["ファイル1とファイル2の両方を選択してください"]
      );
      return;
    }

    setIsDiffExecuted(true);
    await performDiffComparison();
  };

  /// 差分のハイライト表示
  const highlightDifferences = (diff: WorkbookDiff) => {
    if (!spreadRef1.current || !spreadRef2.current) return;

    try {
      diff.sheets.forEach((sheetDiff) => {
        const sheet1 = findSheetByName(
          spreadRef1.current!,
          sheetDiff.sheetName
        );
        const sheet2 = findSheetByName(
          spreadRef2.current!,
          sheetDiff.sheetName
        );

        sheetDiff.cells.forEach((cellDiff) => {
          const color = getDiffColor(cellDiff.type);

          if (sheet1) {
            sheet1.getRange(cellDiff.row, cellDiff.col, 1, 1).backColor(color);
          }
          if (sheet2) {
            sheet2.getRange(cellDiff.row, cellDiff.col, 1, 1).backColor(color);
          }
        });
      });

      spreadRef1.current.refresh();
      spreadRef2.current.refresh();
    } catch (error) {
      console.warn("ハイライト表示エラー:", error);
    }
  };

  /// シート名でシートを検索
  const findSheetByName = (
    workbook: GC.Spread.Sheets.Workbook,
    sheetName: string
  ): GC.Spread.Sheets.Worksheet | null => {
    const sheetCount = workbook.getSheetCount();
    for (let i = 0; i < sheetCount; i++) {
      const sheet = workbook.getSheet(i);
      if (sheet && sheet.name() === sheetName) {
        return sheet;
      }
    }
    return null;
  };

  /// リセット処理
  const handleReset = () => {
    setFile1Name("");
    setFile2Name("");
    setError(null);
    setDiffResult(null);
    setCurrentDiffIndex(0);
    setIsFullscreen(false);
    setShowFullscreenSidebar(true);
    setIsDiffExecuted(false);
    setIsComparing(false);
    setOptionsChanged(false);
    workbookData1.current = null;
    workbookData2.current = null;

    if (spreadRef1.current) {
      spreadRef1.current.clearSheets();
      spreadRef1.current.addSheet(0);
    }
    if (spreadRef2.current) {
      spreadRef2.current.clearSheets();
      spreadRef2.current.addSheet(0);
    }
  };

  /// 差分オプションの変更
  const handleDiffOptionChange = (option: keyof DiffOptions) => {
    const newOptions = { ...diffOptions, [option]: !diffOptions[option] };
    setDiffOptions(newOptions);

    // 既に比較が実行されている場合、オプション変更フラグを立てる
    if (isDiffExecuted) {
      setOptionsChanged(true);
    }
  };

  /// エラー詳細の表示切り替え
  const toggleErrorDetails = () => {
    setShowErrorDetails(!showErrorDetails);
  };

  /// サイドバーの表示切り替え
  const toggleFullscreenSidebar = () => {
    setShowFullscreenSidebar(!showFullscreenSidebar);
  };

  /// 全画面表示の切り替え
  const toggleFullscreen = () => {
    setIsFullscreen(!isFullscreen);

    // 全画面表示の切り替え後にSpreadSheetsのサイズを更新
    setTimeout(() => {
      if (spreadRef1.current && spreadRef2.current) {
        try {
          // より確実なサイズ更新のために複数回実行
          const updateSizes = () => {
            spreadRef1.current!.refresh();
            spreadRef1.current!.invalidateLayout();
            spreadRef2.current!.refresh();
            spreadRef2.current!.invalidateLayout();
          };

          updateSizes();

          // 追加で少し遅延してもう一度実行
          setTimeout(updateSizes, 100);
        } catch (error) {
          console.warn("Toggle fullscreen resize error:", error);
        }
      }
    }, 150);
  };

  /// 指定されたセルにジャンプ
  const jumpToCell = (cellDiff: any) => {
    if (!spreadRef1.current || !spreadRef2.current) return;

    try {
      // 対象シートを検索
      const sheet1 = findSheetByName(spreadRef1.current, cellDiff.sheetName);
      const sheet2 = findSheetByName(spreadRef2.current, cellDiff.sheetName);

      // シートをアクティブにする
      if (sheet1) {
        const sheetIndex1 = getSheetIndex(
          spreadRef1.current,
          cellDiff.sheetName
        );
        if (sheetIndex1 >= 0) {
          spreadRef1.current.setActiveSheetIndex(sheetIndex1);
        }
      }

      if (sheet2) {
        const sheetIndex2 = getSheetIndex(
          spreadRef2.current,
          cellDiff.sheetName
        );
        if (sheetIndex2 >= 0) {
          spreadRef2.current.setActiveSheetIndex(sheetIndex2);
        }
      }

      // セルを選択してスクロール
      setTimeout(() => {
        if (sheet1) {
          sheet1.setSelection(cellDiff.row, cellDiff.col, 1, 1);
          sheet1.showCell(
            cellDiff.row,
            cellDiff.col,
            GC.Spread.Sheets.VerticalPosition.center,
            GC.Spread.Sheets.HorizontalPosition.center
          );
        }
        if (sheet2) {
          sheet2.setSelection(cellDiff.row, cellDiff.col, 1, 1);
          sheet2.showCell(
            cellDiff.row,
            cellDiff.col,
            GC.Spread.Sheets.VerticalPosition.center,
            GC.Spread.Sheets.HorizontalPosition.center
          );
        }
      }, 100);
    } catch (error) {
      console.warn("セルジャンプエラー:", error);
    }
  };

  /// シートのインデックスを取得
  const getSheetIndex = (
    workbook: GC.Spread.Sheets.Workbook,
    sheetName: string
  ): number => {
    const sheetCount = workbook.getSheetCount();
    for (let i = 0; i < sheetCount; i++) {
      const sheet = workbook.getSheet(i);
      if (sheet && sheet.name() === sheetName) {
        return i;
      }
    }
    return -1;
  };

  /// エラータイプに応じたアイコンの取得
  const getErrorIcon = (type: ErrorInfo["type"]) => {
    switch (type) {
      case "file":
        return "📁";
      case "format":
        return "📄";
      case "read":
        return "📊";
      case "compare":
        return "🔍";
      default:
        return "⚠️";
    }
  };

  /// 次の差分へジャンプ
  const jumpToNextDiff = () => {
    if (!diffResult) return;

    const allDiffs = diffResult.sheets.flatMap((sheet) =>
      sheet.cells.map((cell) => ({ ...cell, sheetName: sheet.sheetName }))
    );

    if (allDiffs.length === 0) return;

    const nextIndex = (currentDiffIndex + 1) % allDiffs.length;
    setCurrentDiffIndex(nextIndex);

    // セルにジャンプ
    jumpToCell(allDiffs[nextIndex]);
  };

  /// 前の差分へジャンプ
  const jumpToPrevDiff = () => {
    if (!diffResult) return;

    const allDiffs = diffResult.sheets.flatMap((sheet) =>
      sheet.cells.map((cell) => ({ ...cell, sheetName: sheet.sheetName }))
    );

    if (allDiffs.length === 0) return;

    const prevIndex =
      currentDiffIndex === 0 ? allDiffs.length - 1 : currentDiffIndex - 1;
    setCurrentDiffIndex(prevIndex);

    // セルにジャンプ
    jumpToCell(allDiffs[prevIndex]);
  };

  return (
    <div className={`excel-diff-viewer ${isFullscreen ? "fullscreen" : ""}`}>
      {!isFullscreen && (
        <div className="diff-upload-section">
          <div className="diff-upload-controls">
            <div className="file-upload-group">
              <input
                type="file"
                id="file1-input"
                accept=".xlsx,.xls"
                onChange={handleFile1Select}
                className="file-input"
              />
              <label htmlFor="file1-input" className="file-label">
                📁 ファイル1を選択
              </label>
              {file1Name && <span className="file-name">📊 {file1Name}</span>}
            </div>

            <div className="file-upload-group">
              <input
                type="file"
                id="file2-input"
                accept=".xlsx,.xls"
                onChange={handleFile2Select}
                className="file-input"
              />
              <label htmlFor="file2-input" className="file-label">
                📁 ファイル2を選択
              </label>
              {file2Name && <span className="file-name">📊 {file2Name}</span>}
            </div>

            {(file1Name || file2Name) && (
              <button
                onClick={handleReset}
                className="reset-button"
                type="button"
              >
                リセット
              </button>
            )}
          </div>

          <div className="diff-options">
            <h3>比較オプション</h3>
            <div className="options-grid">
              <label className="option-item">
                <input
                  type="checkbox"
                  checked={diffOptions.compareValues}
                  onChange={() => handleDiffOptionChange("compareValues")}
                />
                <span>値を比較</span>
              </label>
              <label className="option-item">
                <input
                  type="checkbox"
                  checked={diffOptions.compareFormats}
                  onChange={() => handleDiffOptionChange("compareFormats")}
                />
                <span>書式を比較</span>
              </label>
              <label className="option-item">
                <input
                  type="checkbox"
                  checked={diffOptions.compareFormulas}
                  onChange={() => handleDiffOptionChange("compareFormulas")}
                />
                <span>数式を比較</span>
              </label>
              <label className="option-item">
                <input
                  type="checkbox"
                  checked={diffOptions.compareComments}
                  onChange={() => handleDiffOptionChange("compareComments")}
                />
                <span>コメントを比較</span>
              </label>
            </div>
          </div>

          {file1Name && file2Name && !isDiffExecuted && (
            <div className="diff-execute-section">
              <button
                onClick={executeDiffComparison}
                className="execute-diff-button"
                type="button"
                disabled={isComparing}
              >
                {isComparing ? "比較中..." : "🔍 差分比較を実行"}
              </button>
              <p className="execute-help-text">
                両方のファイルを読み込み、比較オプションを設定した後に実行してください
              </p>
            </div>
          )}

          {(isLoading || isComparing) && (
            <div className="loading">
              <div className="loading-spinner"></div>
              <span>{isComparing ? "差分比較中..." : "処理中..."}</span>
            </div>
          )}

          {error && (
            <div className="error">
              <div className="error-header">
                <span className="error-icon">{getErrorIcon(error.type)}</span>
                <span className="error-message">{error.message}</span>
              </div>

              {error.suggestions && error.suggestions.length > 0 && (
                <div className="error-suggestions">
                  <h4>対処法:</h4>
                  <ul>
                    {error.suggestions.map((suggestion, index) => (
                      <li key={index}>{suggestion}</li>
                    ))}
                  </ul>
                </div>
              )}

              {error.details && (
                <div className="error-details">
                  <button
                    onClick={toggleErrorDetails}
                    className="details-toggle"
                    type="button"
                  >
                    {showErrorDetails ? "詳細を隠す" : "詳細を表示"}
                    <span
                      className={`arrow ${showErrorDetails ? "up" : "down"}`}
                    >
                      ▼
                    </span>
                  </button>
                  {showErrorDetails && (
                    <div className="error-details-content">
                      <pre>{error.details}</pre>
                    </div>
                  )}
                </div>
              )}
            </div>
          )}

          {diffResult && (
            <div className="diff-summary">
              <h3>差分サマリー</h3>
              <div className="summary-stats">
                <div className="stat-item">
                  <span className="stat-label">総変更数:</span>
                  <span className="stat-value">
                    {diffResult.summary.totalCellChanges}
                  </span>
                </div>
                <div className="stat-item">
                  <span className="stat-label">変更シート:</span>
                  <span className="stat-value">
                    {diffResult.summary.modifiedSheets}
                  </span>
                </div>
              </div>

              <div className="diff-actions">
                <button
                  onClick={executeDiffComparison}
                  className={`re-execute-diff-button ${
                    optionsChanged ? "options-changed" : ""
                  }`}
                  type="button"
                  disabled={isComparing}
                >
                  {isComparing
                    ? "比較中..."
                    : optionsChanged
                    ? "🔄 オプション変更 - 再比較が必要"
                    : "🔄 差分比較を再実行"}
                </button>
                {optionsChanged && (
                  <p className="options-changed-notice">
                    比較オプションが変更されました。再比較ボタンを押して結果を更新してください。
                  </p>
                )}
              </div>

              {diffResult.summary.totalCellChanges > 0 && (
                <div className="diff-navigation">
                  <button
                    onClick={jumpToPrevDiff}
                    className="nav-button"
                    type="button"
                  >
                    ← 前の差分
                  </button>
                  <span className="diff-counter">
                    {currentDiffIndex + 1} /{" "}
                    {diffResult.sheets.flatMap((s) => s.cells).length}
                  </span>
                  <button
                    onClick={jumpToNextDiff}
                    className="nav-button"
                    type="button"
                  >
                    次の差分 →
                  </button>
                </div>
              )}
            </div>
          )}

          {selectedCellInfo && (
            <div className="cell-detail">
              <h3>セル詳細情報</h3>
              <div className="detail-grid">
                <div className="detail-item">
                  <span className="detail-label">位置:</span>
                  <span className="detail-value">
                    {selectedCellInfo.address} ({selectedCellInfo.sheetName})
                  </span>
                </div>
                <div className="detail-item">
                  <span className="detail-label">変更種別:</span>
                  <span
                    className={`detail-value diff-type-${selectedCellInfo.type}`}
                  >
                    {getDiffIcon(selectedCellInfo.type)}{" "}
                    {selectedCellInfo.type === "added"
                      ? "追加"
                      : selectedCellInfo.type === "removed"
                      ? "削除"
                      : selectedCellInfo.type === "modified"
                      ? "変更"
                      : "不明"}
                  </span>
                </div>
                {selectedCellInfo.oldValue !== undefined && (
                  <div className="detail-item">
                    <span className="detail-label">元の値:</span>
                    <span className="detail-value">
                      {String(selectedCellInfo.oldValue || "（空）")}
                    </span>
                  </div>
                )}
                {selectedCellInfo.newValue !== undefined && (
                  <div className="detail-item">
                    <span className="detail-label">新しい値:</span>
                    <span className="detail-value">
                      {String(selectedCellInfo.newValue || "（空）")}
                    </span>
                  </div>
                )}
                {selectedCellInfo.oldFormula && (
                  <div className="detail-item">
                    <span className="detail-label">元の数式:</span>
                    <span className="detail-value formula">
                      {selectedCellInfo.oldFormula}
                    </span>
                  </div>
                )}
                {selectedCellInfo.newFormula && (
                  <div className="detail-item">
                    <span className="detail-label">新しい数式:</span>
                    <span className="detail-value formula">
                      {selectedCellInfo.newFormula}
                    </span>
                  </div>
                )}
                {selectedCellInfo.formatChanges && (
                  <div className="detail-item">
                    <span className="detail-label">書式変更:</span>
                    <span className="detail-value">
                      {Object.entries(selectedCellInfo.formatChanges)
                        .filter(([_, changed]) => changed)
                        .map(([type, _]) => {
                          switch (type) {
                            case "font":
                              return "フォント";
                            case "background":
                              return "背景色";
                            case "border":
                              return "罫線";
                            case "numberFormat":
                              return "数値書式";
                            case "alignment":
                              return "配置";
                            default:
                              return type;
                          }
                        })
                        .join(", ")}
                    </span>
                  </div>
                )}
                {selectedCellInfo.oldFormat?.numberFormat &&
                  selectedCellInfo.newFormat?.numberFormat &&
                  selectedCellInfo.oldFormat.numberFormat !==
                    selectedCellInfo.newFormat.numberFormat && (
                    <div className="detail-item">
                      <span className="detail-label">元の表示形式:</span>
                      <span className="detail-value format">
                        {selectedCellInfo.oldFormat.numberFormat || "標準"}
                      </span>
                    </div>
                  )}
                {selectedCellInfo.oldFormat?.numberFormat &&
                  selectedCellInfo.newFormat?.numberFormat &&
                  selectedCellInfo.oldFormat.numberFormat !==
                    selectedCellInfo.newFormat.numberFormat && (
                    <div className="detail-item">
                      <span className="detail-label">新しい表示形式:</span>
                      <span className="detail-value format">
                        {selectedCellInfo.newFormat.numberFormat || "標準"}
                      </span>
                    </div>
                  )}
              </div>
            </div>
          )}
        </div>
      )}

      <div className="diff-spreadsheet-container">
        <div className="spreadsheet-header">
          <div className="spreadsheet-title">
            <span className="file-icon">🔍</span>
            <span className="title-text">Excel差分比較</span>
          </div>
          <div className="spreadsheet-controls">
            {(file1Name || file2Name) && (
              <button
                onClick={toggleFullscreen}
                className="fullscreen-button"
                type="button"
              >
                {isFullscreen ? "🗗 縮小" : "🗖 全画面"}
              </button>
            )}
            {isFullscreen && (
              <>
                <button
                  onClick={toggleFullscreenSidebar}
                  className="sidebar-toggle-button"
                  type="button"
                  title={
                    showFullscreenSidebar
                      ? "サイドバーを隠す"
                      : "サイドバーを表示"
                  }
                >
                  {showFullscreenSidebar ? "◀" : "▶"} 詳細
                </button>
                <button
                  onClick={handleReset}
                  className="reset-button-fullscreen"
                  type="button"
                >
                  🔄 リセット
                </button>
              </>
            )}
          </div>
        </div>

        <div
          className={`spreadsheet-content ${
            isFullscreen ? "fullscreen-layout" : ""
          } ${isFullscreen && showFullscreenSidebar ? "show-sidebar" : ""}`}
        >
          <div className="spreadsheet-panels">
            <div className="spreadsheet-panel">
              <div className="panel-header">
                <span className="panel-title">
                  📊 {file1Name || "ファイル1"}
                </span>
              </div>
              <SpreadSheets
                workbookInitialized={onWorkbook1Initialized}
                hostStyle={{
                  width: "100%",
                  height: isFullscreen ? "calc(100vh - 120px)" : "500px",
                  border: "1px solid #ccc",
                  borderRadius: isFullscreen ? "0" : "4px",
                }}
              >
                <Worksheet />
              </SpreadSheets>
            </div>

            <div className="spreadsheet-panel">
              <div className="panel-header">
                <span className="panel-title">
                  📊 {file2Name || "ファイル2"}
                </span>
              </div>
              <SpreadSheets
                workbookInitialized={onWorkbook2Initialized}
                hostStyle={{
                  width: "100%",
                  height: isFullscreen ? "calc(100vh - 120px)" : "500px",
                  border: "1px solid #ccc",
                  borderRadius: isFullscreen ? "0" : "4px",
                }}
              >
                <Worksheet />
              </SpreadSheets>
            </div>
          </div>

          {isFullscreen && showFullscreenSidebar && (
            <div className="fullscreen-sidebar">
              <div className="sidebar-content">
                {/* 比較オプション */}
                <div className="sidebar-section">
                  <h3>比較オプション</h3>
                  <div className="sidebar-options">
                    <label className="sidebar-option-item">
                      <input
                        type="checkbox"
                        checked={diffOptions.compareValues}
                        onChange={() => handleDiffOptionChange("compareValues")}
                      />
                      <span>値を比較</span>
                    </label>
                    <label className="sidebar-option-item">
                      <input
                        type="checkbox"
                        checked={diffOptions.compareFormats}
                        onChange={() =>
                          handleDiffOptionChange("compareFormats")
                        }
                      />
                      <span>書式を比較</span>
                    </label>
                    <label className="sidebar-option-item">
                      <input
                        type="checkbox"
                        checked={diffOptions.compareFormulas}
                        onChange={() =>
                          handleDiffOptionChange("compareFormulas")
                        }
                      />
                      <span>数式を比較</span>
                    </label>
                    <label className="sidebar-option-item">
                      <input
                        type="checkbox"
                        checked={diffOptions.compareComments}
                        onChange={() =>
                          handleDiffOptionChange("compareComments")
                        }
                      />
                      <span>コメントを比較</span>
                    </label>
                  </div>

                  {file1Name && file2Name && (
                    <div className="sidebar-execute-section">
                      <button
                        onClick={executeDiffComparison}
                        className="sidebar-execute-button"
                        type="button"
                        disabled={isComparing}
                      >
                        {isComparing ? "比較中..." : "🔍 差分比較実行"}
                      </button>
                    </div>
                  )}
                </div>

                {/* 差分サマリー */}
                {diffResult && (
                  <div className="sidebar-section">
                    <h3>差分サマリー</h3>
                    <div className="sidebar-summary-stats">
                      <div className="sidebar-stat-item">
                        <span className="sidebar-stat-label">総変更数:</span>
                        <span className="sidebar-stat-value">
                          {diffResult.summary.totalCellChanges}
                        </span>
                      </div>
                      <div className="sidebar-stat-item">
                        <span className="sidebar-stat-label">変更シート:</span>
                        <span className="sidebar-stat-value">
                          {diffResult.summary.modifiedSheets}
                        </span>
                      </div>
                    </div>

                    {diffResult.summary.totalCellChanges > 0 && (
                      <div className="sidebar-navigation">
                        <button
                          onClick={jumpToPrevDiff}
                          className="sidebar-nav-button"
                          type="button"
                        >
                          ← 前の差分
                        </button>
                        <span className="sidebar-diff-counter">
                          {currentDiffIndex + 1} /{" "}
                          {diffResult.sheets.flatMap((s) => s.cells).length}
                        </span>
                        <button
                          onClick={jumpToNextDiff}
                          className="sidebar-nav-button"
                          type="button"
                        >
                          次の差分 →
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {/* セル詳細情報 */}
                {selectedCellInfo && (
                  <div className="sidebar-section">
                    <h3>セル詳細情報</h3>
                    <div className="sidebar-detail-grid">
                      <div className="sidebar-detail-item">
                        <span className="sidebar-detail-label">位置:</span>
                        <span className="sidebar-detail-value">
                          {selectedCellInfo.address} (
                          {selectedCellInfo.sheetName})
                        </span>
                      </div>
                      <div className="sidebar-detail-item">
                        <span className="sidebar-detail-label">変更種別:</span>
                        <span
                          className={`sidebar-detail-value diff-type-${selectedCellInfo.type}`}
                        >
                          {getDiffIcon(selectedCellInfo.type)}{" "}
                          {selectedCellInfo.type === "added"
                            ? "追加"
                            : selectedCellInfo.type === "removed"
                            ? "削除"
                            : selectedCellInfo.type === "modified"
                            ? "変更"
                            : "不明"}
                        </span>
                      </div>
                      {selectedCellInfo.oldValue !== undefined && (
                        <div className="sidebar-detail-item">
                          <span className="sidebar-detail-label">元の値:</span>
                          <span className="sidebar-detail-value">
                            {String(selectedCellInfo.oldValue || "（空）")}
                          </span>
                        </div>
                      )}
                      {selectedCellInfo.newValue !== undefined && (
                        <div className="sidebar-detail-item">
                          <span className="sidebar-detail-label">
                            新しい値:
                          </span>
                          <span className="sidebar-detail-value">
                            {String(selectedCellInfo.newValue || "（空）")}
                          </span>
                        </div>
                      )}
                      {selectedCellInfo.oldFormula && (
                        <div className="sidebar-detail-item">
                          <span className="sidebar-detail-label">
                            元の数式:
                          </span>
                          <span className="sidebar-detail-value formula">
                            {selectedCellInfo.oldFormula}
                          </span>
                        </div>
                      )}
                      {selectedCellInfo.newFormula && (
                        <div className="sidebar-detail-item">
                          <span className="sidebar-detail-label">
                            新しい数式:
                          </span>
                          <span className="sidebar-detail-value formula">
                            {selectedCellInfo.newFormula}
                          </span>
                        </div>
                      )}
                      {selectedCellInfo.formatChanges && (
                        <div className="sidebar-detail-item">
                          <span className="sidebar-detail-label">
                            書式変更:
                          </span>
                          <span className="sidebar-detail-value">
                            {Object.entries(selectedCellInfo.formatChanges)
                              .filter(([_, changed]) => changed)
                              .map(([type, _]) => {
                                switch (type) {
                                  case "font":
                                    return "フォント";
                                  case "background":
                                    return "背景色";
                                  case "border":
                                    return "罫線";
                                  case "numberFormat":
                                    return "数値書式";
                                  case "alignment":
                                    return "配置";
                                  default:
                                    return type;
                                }
                              })
                              .join(", ")}
                          </span>
                        </div>
                      )}
                      {selectedCellInfo.oldFormat?.numberFormat &&
                        selectedCellInfo.newFormat?.numberFormat &&
                        selectedCellInfo.oldFormat.numberFormat !==
                          selectedCellInfo.newFormat.numberFormat && (
                          <div className="sidebar-detail-item">
                            <span className="sidebar-detail-label">
                              元の表示形式:
                            </span>
                            <span className="sidebar-detail-value format">
                              {selectedCellInfo.oldFormat.numberFormat ||
                                "標準"}
                            </span>
                          </div>
                        )}
                      {selectedCellInfo.oldFormat?.numberFormat &&
                        selectedCellInfo.newFormat?.numberFormat &&
                        selectedCellInfo.oldFormat.numberFormat !==
                          selectedCellInfo.newFormat.numberFormat && (
                          <div className="sidebar-detail-item">
                            <span className="sidebar-detail-label">
                              新しい表示形式:
                            </span>
                            <span className="sidebar-detail-value format">
                              {selectedCellInfo.newFormat.numberFormat ||
                                "標準"}
                            </span>
                          </div>
                        )}
                    </div>
                  </div>
                )}
              </div>
            </div>
          )}
        </div>
      </div>

      {isFullscreen && (
        <div className="fullscreen-help">
          <span>ESCキーまたは「縮小」ボタンで全画面を解除</span>
        </div>
      )}
    </div>
  );
};

export default ExcelDiffViewer;

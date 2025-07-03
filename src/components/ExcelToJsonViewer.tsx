import React, {
  useState,
  useCallback,
  startTransition,
  useEffect,
} from "react";
import * as ExcelIO from "@grapecity/spread-excelio";
import { diffLines } from "diff";
import "./ExcelToJsonViewer.css";

/// JSON差分比較のオプション
export interface JsonDiffOptions {
  compareMode: "full" | "sheets_only";
  ignoreOrder: boolean;
  ignoreEmptyValues: boolean;
  maxDepth: number;
}

/// JSON差分の種類
export type JsonDiffType = "added" | "removed" | "modified";

/// JSON差分の詳細
export interface JsonDiff {
  path: (string | number)[];
  type: JsonDiffType;
  oldValue?: any;
  newValue?: any;
}

/// JSON差分比較の結果
export interface JsonDiffResult {
  differences: JsonDiff[];
  summary: {
    total: number;
    added: number;
    removed: number;
    modified: number;
  };
}

/// LineDiff インターフェース
interface LineDiff {
  left: string;
  right: string;
  type: "added" | "removed" | "unchanged";
}

/**
 * 2つのJSONオブジェクトを再帰的に比較し、差分を検出します。
 * @param obj1 比較元オブジェクト
 * @param obj2 比較先オブジェクト
 * @param options 比較オプション
 * @returns 差分結果
 */
export const compareJsonObjects = (
  obj1: any,
  obj2: any,
  options: JsonDiffOptions
): JsonDiffResult => {
  const differences: JsonDiff[] = [];
  const summary = { total: 0, added: 0, removed: 0, modified: 0 };

  const compare = (
    o1: any,
    o2: any,
    path: (string | number)[] = [],
    depth: number = 0
  ) => {
    if (depth > options.maxDepth) {
      return;
    }

    if (o1 === o2) {
      return;
    }

    if (
      o1 === null ||
      o2 === null ||
      typeof o1 !== "object" ||
      typeof o2 !== "object"
    ) {
      if (
        !(
          options.ignoreEmptyValues &&
          ((o1 === null && o2 === undefined) ||
            (o1 === undefined && o2 === null))
        )
      ) {
        differences.push({
          path,
          type: "modified",
          oldValue: o1,
          newValue: o2,
        });
        summary.modified++;
      }
      return;
    }

    const keys1 = Object.keys(o1);
    const keys2 = Object.keys(o2);
    const allKeys = new Set([...keys1, ...keys2]);

    allKeys.forEach((key) => {
      const newPath = [...path, key];
      const val1 = o1[key];
      const val2 = o2[key];

      if (keys1.includes(key) && !keys2.includes(key)) {
        differences.push({ path: newPath, type: "removed", oldValue: val1 });
        summary.removed++;
      } else if (!keys1.includes(key) && keys2.includes(key)) {
        differences.push({ path: newPath, type: "added", newValue: val2 });
        summary.added++;
      } else {
        compare(val1, val2, newPath, depth + 1);
      }
    });
  };

  const data1 =
    options.compareMode === "sheets_only" && obj1.sheets ? obj1.sheets : obj1;
  const data2 =
    options.compareMode === "sheets_only" && obj2.sheets ? obj2.sheets : obj2;

  compare(data1, data2);

  summary.total = summary.added + summary.removed + summary.modified;
  return { differences, summary };
};

/**
 * JSON差分結果をCSV形式の文字列に変換します。
 * @param diffResult 差分結果
 * @returns CSV文字列
 */
export const exportJsonDiffToCsv = (diffResult: JsonDiffResult): string => {
  const header = "Path,Type,Old Value,New Value\n";
  const rows = diffResult.differences.map((diff) => {
    const path = diff.path.join(".");
    const oldValue =
      diff.oldValue !== undefined ? `"${JSON.stringify(diff.oldValue)}"` : "";
    const newValue =
      diff.newValue !== undefined ? `"${JSON.stringify(diff.newValue)}"` : "";
    return `${path},${diff.type},${oldValue},${newValue}`;
  });
  return header + rows.join("\n");
};

/**
 * JSON差分結果をJSON形式の文字列に変換します。
 * @param diffResult 差分結果
 * @returns JSON文字列
 */
export const exportJsonDiffToJson = (diffResult: JsonDiffResult): string => {
  return JSON.stringify(diffResult, null, 2);
};

/// ExcelToJsonViewerコンポーネント - ExcelファイルをJSON形式に変換して表示・比較
const ExcelToJsonViewer: React.FC = () => {
  // 基本状態
  const [fileName, setFileName] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [jsonData, setJsonData] = useState<any>(null);
  const [jsonString, setJsonString] = useState<string>("");
  const [showRawJson, setShowRawJson] = useState<boolean>(false);
  const [copySuccess, setCopySuccess] = useState<boolean>(false);
  const [showSheetsOnly, setShowSheetsOnly] = useState<boolean>(true);
  const [inputResetKey, setInputResetKey] = useState<number>(0);

  // 差分比較用の状態
  const [isDiffMode, setIsDiffMode] = useState<boolean>(false);
  const [fileName1, setFileName1] = useState<string>("");
  const [fileName2, setFileName2] = useState<string>("");
  const [jsonData1, setJsonData1] = useState<any>(null);
  const [jsonData2, setJsonData2] = useState<any>(null);
  const [jsonString1, setJsonString1] = useState<string>("");
  const [jsonString2, setJsonString2] = useState<string>("");
  const [diffResult, setDiffResult] = useState<JsonDiffResult | null>(null);
  const [isComparing, setIsComparing] = useState<boolean>(false);
  const [diffOptions, setDiffOptions] = useState<JsonDiffOptions>({
    compareMode: "sheets_only",
    ignoreOrder: false,
    ignoreEmptyValues: true,
    maxDepth: 10,
  });
  const [lineDiffs, setLineDiffs] = useState<LineDiff[] | null>(null);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);

  /// ESCキーで全画面モードを終了
  useEffect(() => {
    const handleEscKey = (event: KeyboardEvent) => {
      if (event.key === "Escape" && isFullscreen) {
        setIsFullscreen(false);
      }
    };

    if (isFullscreen) {
      document.addEventListener("keydown", handleEscKey);
      document.body.style.overflow = "hidden"; // スクロール防止
    } else {
      document.body.style.overflow = ""; // スクロール復元
    }

    return () => {
      document.removeEventListener("keydown", handleEscKey);
      document.body.style.overflow = ""; // クリーンアップ
    };
  }, [isFullscreen]);

  /// モード切り替え処理
  const toggleMode = useCallback(() => {
    const newIsDiffMode = !isDiffMode;
    startTransition(() => {
      setIsDiffMode(newIsDiffMode);
      // モード変更時にリセット
      if (newIsDiffMode) {
        // 差分比較モードのリセット
        setFileName1("");
        setFileName2("");
        setJsonData1(null);
        setJsonData2(null);
        setJsonString1("");
        setJsonString2("");
        setDiffResult(null);
        setIsComparing(false);
        setLineDiffs(null);
      } else {
        // 単一ファイルモードのリセット
        setFileName("");
        setJsonData(null);
        setJsonString("");
        setShowRawJson(false);
        setCopySuccess(false);
        setShowSheetsOnly(true);
        setIsProcessing(false);
      }
      setError("");
      setInputResetKey((prev) => prev + 1);
    });
  }, [isDiffMode]);

  /// 全画面表示の切り替え
  const toggleFullscreen = useCallback(() => {
    setIsFullscreen(!isFullscreen);
  }, [isFullscreen]);

  /// ファイル選択時の処理（単一ファイル）
  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!validateFile(file)) return;

    setError("");
    setFileName(file.name);
    loadExcelFile(file);
  };

  /// ファイル選択時の処理（差分比較用）
  const handleFileSelect1 = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!validateFile(file)) return;

    setError("");
    setFileName1(file.name);
    loadExcelFileForDiff(file, 1);
  };

  const handleFileSelect2 = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!validateFile(file)) return;

    setError("");
    setFileName2(file.name);
    loadExcelFileForDiff(file, 2);
  };

  /// ファイル検証
  const validateFile = (file: File): boolean => {
    // ファイルサイズのチェック（100MB制限）
    const maxSize = 100 * 1024 * 1024; // 100MB
    if (file.size > maxSize) {
      setError(
        `ファイルサイズが大きすぎます。ファイルサイズ: ${(
          file.size /
          1024 /
          1024
        ).toFixed(2)}MB（制限: 100MB）`
      );
      return false;
    }

    // サポートされているファイル形式の確認
    const supportedFormats = [".xlsx", ".xls"];
    const fileExtension = file.name
      .toLowerCase()
      .substring(file.name.lastIndexOf("."));

    if (!supportedFormats.includes(fileExtension)) {
      setError(
        `サポートされていないファイル形式です。ファイル形式: ${fileExtension}（サポート形式: xlsx、xls）`
      );
      return false;
    }

    return true;
  };

  /// Excelファイルの読み込み処理
  const loadExcelFile = async (file: File) => {
    setIsLoading(true);
    setJsonData(null);
    setJsonString("");

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
              // JSONデータを保存
              setJsonData(json);
              // 見やすい形式でJSONを文字列化（初期読み込み時は常にtrue）
              updateJsonString(json, true);
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
    } catch (error) {
      console.error("Excel file loading error:", error);
      handleLoadError(error);
    } finally {
      setIsLoading(false);
    }
  };

  /// 差分比較用のExcelファイル読み込み処理
  const loadExcelFileForDiff = async (file: File, fileNumber: 1 | 2) => {
    setIsLoading(true);
    setError("");

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
              // ファイル番号に応じてデータを保存
              if (fileNumber === 1) {
                setJsonData1(json);
                updateJsonStringForDiff(
                  json,
                  1,
                  diffOptions.compareMode === "sheets_only"
                );
              } else {
                setJsonData2(json);
                updateJsonStringForDiff(
                  json,
                  2,
                  diffOptions.compareMode === "sheets_only"
                );
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
    } catch (error) {
      console.error("Excel file loading error:", error);
      handleLoadError(error);
    } finally {
      setIsLoading(false);
    }
  };

  /// JSONデータの文字列化
  const updateJsonString = (data: any, sheetsOnly: boolean) => {
    if (sheetsOnly && data && data.sheets) {
      // sheets以下の情報のみを抽出
      const sheetsOnlyJson = {
        sheets: data.sheets,
      };
      const formattedJson = JSON.stringify(sheetsOnlyJson, null, 2);
      setJsonString(formattedJson);
    } else {
      // 全体のJSONを表示
      const formattedJson = JSON.stringify(data, null, 2);
      setJsonString(formattedJson);
    }
  };

  /// 差分比較用のJSONデータの文字列化
  const updateJsonStringForDiff = (
    data: any,
    fileNumber: 1 | 2,
    sheetsOnly: boolean
  ) => {
    const processedData =
      sheetsOnly && data && data.sheets ? { sheets: data.sheets } : data;

    const formattedJson = JSON.stringify(processedData, null, 2);

    if (fileNumber === 1) {
      setJsonString1(formattedJson);
    } else {
      setJsonString2(formattedJson);
    }
  };

  /// sheets表示モードの切り替え
  const toggleSheetsOnlyMode = async () => {
    const newShowSheetsOnly = !showSheetsOnly;

    // sheetsのみモードにしようとしているが、sheetsが存在しない場合はエラーを表示
    if (newShowSheetsOnly && (!jsonData || !jsonData.sheets)) {
      setError(
        "シートデータが見つかりません。このファイルにはシート情報が含まれていない可能性があります。"
      );
      return;
    }

    setIsProcessing(true);
    setError(""); // エラーをクリア

    try {
      // 処理を非同期化して、UI更新の時間を確保
      await new Promise((resolve) => setTimeout(resolve, 100));

      setShowSheetsOnly(newShowSheetsOnly);
      if (jsonData) {
        updateJsonString(jsonData, newShowSheetsOnly);
      }
    } catch (error) {
      console.error("表示モード切り替えエラー:", error);
      setError("表示モードの切り替えに失敗しました");
    } finally {
      setIsProcessing(false);
    }
  };

  /// RAW表示/構造表示の切り替え
  const toggleRawJsonDisplay = async () => {
    setIsProcessing(true);

    try {
      // 大きなJSONデータの場合、ハイライト処理に時間がかかる可能性があるため非同期処理
      await new Promise((resolve) => setTimeout(resolve, 50));

      setShowRawJson(!showRawJson);
    } catch (error) {
      console.error("表示モード切り替えエラー:", error);
      setError("表示モードの切り替えに失敗しました");
    } finally {
      setIsProcessing(false);
    }
  };

  /// 読み込みエラーの処理
  const handleLoadError = (error: any) => {
    let errorMessage = "ファイルの読み込み中にエラーが発生しました";

    if (error instanceof Error) {
      errorMessage = error.message;

      if (
        error.message.includes("password") ||
        error.message.includes("protected")
      ) {
        errorMessage = "パスワードで保護されたファイルはサポートされていません";
      } else if (
        error.message.includes("corrupt") ||
        error.message.includes("invalid")
      ) {
        errorMessage = "ファイルが破損している可能性があります";
      }
    }

    setError(errorMessage);
  };

  /// JSONをクリップボードにコピー
  const copyToClipboard = async () => {
    try {
      await navigator.clipboard.writeText(jsonString);
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    } catch (err) {
      console.error("クリップボードへのコピーに失敗しました:", err);
      // フォールバック: テキストエリアを使った方法
      const textArea = document.createElement("textarea");
      textArea.value = jsonString;
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      document.body.removeChild(textArea);
      setCopySuccess(true);
      setTimeout(() => setCopySuccess(false), 2000);
    }
  };

  /// ファイルとしてダウンロード
  const downloadJson = () => {
    const blob = new Blob([jsonString], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const suffix = showSheetsOnly ? "_sheets_only" : "_spreadjs";
    a.download = `${fileName.replace(/\.[^/.]+$/, "")}${suffix}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  /// 差分比較の実行
  const performDiffComparison = useCallback(async () => {
    if (!jsonData1 || !jsonData2) {
      setError("両方のファイルを選択してください。");
      return;
    }

    setIsComparing(true);
    setError("");

    try {
      // 時間のかかる処理を非同期で実行
      const result = await new Promise<JsonDiffResult>((resolve) => {
        setTimeout(() => {
          const diff = compareJsonObjects(jsonData1, jsonData2, diffOptions);
          resolve(diff);
        }, 50); // UIの更新を許可するために少し待つ
      });

      // 行単位の差分を取得
      const lineParts = diffLines(jsonString1, jsonString2);
      const rows: LineDiff[] = [];
      lineParts.forEach((part) => {
        const lines = part.value.split("\n");
        if (lines[lines.length - 1] === "") lines.pop();
        lines.forEach((line) => {
          if (part.added) {
            rows.push({ left: "", right: line, type: "added" });
          } else if (part.removed) {
            rows.push({ left: line, right: "", type: "removed" });
          } else {
            rows.push({ left: line, right: line, type: "unchanged" });
          }
        });
      });

      // 行数ベースの統計情報を計算
      const addedCount = rows.filter((row) => row.type === "added").length;
      const removedCount = rows.filter((row) => row.type === "removed").length;
      const modifiedCount = 0; // 行単位の差分では「変更」は追加+削除として扱われるため0

      // 統計情報を行数ベースで更新
      const updatedResult: JsonDiffResult = {
        ...result,
        summary: {
          total: addedCount + removedCount + modifiedCount,
          added: addedCount,
          removed: removedCount,
          modified: modifiedCount,
        },
      };

      // 状態更新をstartTransition内で実行
      startTransition(() => {
        setDiffResult(updatedResult);
        setLineDiffs(rows);
      });
    } catch (err) {
      console.error("Comparison error:", err);
      startTransition(() => {
        setError(
          err instanceof Error
            ? err.message
            : "比較中に不明なエラーが発生しました。"
        );
        setDiffResult(null);
        setLineDiffs(null);
      });
    } finally {
      setIsComparing(false);
    }
  }, [jsonData1, jsonData2, jsonString1, jsonString2, diffOptions]);

  /// 差分比較オプションの変更ハンドラ
  const handleDiffOptionChange = useCallback(
    (option: keyof JsonDiffOptions, value: any) => {
      startTransition(() => {
        setDiffOptions((prev: JsonDiffOptions) => ({
          ...prev,
          [option]: value,
        }));

        // 比較モードが変更された場合、JSONデータを再生成
        if (option === "compareMode") {
          const isSheetOnly = value === "sheets_only";
          if (jsonData1) {
            updateJsonStringForDiff(jsonData1, 1, isSheetOnly);
          }
          if (jsonData2) {
            updateJsonStringForDiff(jsonData2, 2, isSheetOnly);
          }
        }

        // 差分結果をクリア（再比較が必要）
        setDiffResult(null);
        setLineDiffs(null);
      });
    },
    [jsonData1, jsonData2]
  );

  /// 差分結果のダウンロード
  const downloadDiffResult = (format: "json" | "csv") => {
    if (!diffResult) return;

    const content =
      format === "json"
        ? exportJsonDiffToJson(diffResult)
        : exportJsonDiffToCsv(diffResult);

    const mimeType = format === "json" ? "application/json" : "text/csv";
    const fileExtension = format === "json" ? ".json" : ".csv";

    const blob = new Blob([content], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    const downloadFileName = `diff_${fileName1}_vs_${fileName2}_${new Date().getTime()}${fileExtension}`;
    a.download = downloadFileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  /// リセット処理
  const handleReset = useCallback(() => {
    startTransition(() => {
      if (isDiffMode) {
        // 差分比較モードのリセット
        setFileName1("");
        setFileName2("");
        setJsonData1(null);
        setJsonData2(null);
        setJsonString1("");
        setJsonString2("");
        setDiffResult(null);
        setIsComparing(false);
        setLineDiffs(null);
      } else {
        // 単一ファイルモードのリセット
        setFileName("");
        setJsonData(null);
        setJsonString("");
        setShowRawJson(false);
        setCopySuccess(false);
        setShowSheetsOnly(true);
        setIsProcessing(false);
      }
      setError("");
      setInputResetKey((prev) => prev + 1);
    });
  }, [isDiffMode]);

  /// JSONデータの概要情報を取得
  const getJsonSummary = () => {
    if (!jsonData) return null;

    // sheetsの数を取得（配列またはオブジェクトに対応）
    let sheetCount = 0;
    if (jsonData.sheets) {
      if (Array.isArray(jsonData.sheets)) {
        sheetCount = jsonData.sheets.length;
      } else if (typeof jsonData.sheets === "object") {
        sheetCount = Object.keys(jsonData.sheets).length;
      }
    }

    const summary = {
      version: jsonData.version || "不明",
      sheetCount: sheetCount,
      fileSize: `${(jsonString.length / 1024).toFixed(2)} KB`,
      hasStyles: jsonData.customList || jsonData.namedStyles ? true : false,
      hasNames: jsonData.names && jsonData.names.length > 0,
      displayMode: showSheetsOnly ? "sheetsのみ" : "全体",
    };

    return summary;
  };

  /// sheets表示時のシート構造を取得
  const getSheetsOnlyStructure = () => {
    if (!jsonData || !jsonData.sheets || !showSheetsOnly) return null;

    // sheetsが配列の場合
    if (Array.isArray(jsonData.sheets)) {
      return jsonData.sheets.map((sheet: any, index: number) => {
        return {
          index,
          name: sheet.name || `Sheet${index + 1}`,
          visible: sheet.visible !== false,
          rowCount: sheet.rowCount || 0,
          columnCount: sheet.columnCount || 0,
          hasData: sheet.data ? true : false,
          hasStyles: sheet.styles
            ? Object.keys(sheet.styles).length > 0
            : false,
        };
      });
    }

    // sheetsがオブジェクトの場合
    if (typeof jsonData.sheets === "object") {
      return Object.entries(jsonData.sheets).map(
        ([sheetName, sheet]: [string, any], index: number) => {
          return {
            index,
            name: sheet.name || sheetName,
            visible: sheet.visible !== false,
            rowCount: sheet.rowCount || 0,
            columnCount: sheet.columnCount || 0,
            hasData: sheet.data ? true : false,
            hasStyles: sheet.styles
              ? Object.keys(sheet.styles).length > 0
              : false,
          };
        }
      );
    }

    return null;
  };

  /// JSONを色分けして表示するためのハイライト処理
  const highlightJson = (jsonString: string): string => {
    return jsonString
      .replace(/"([^"]+)":/g, '<span class="json-key">"$1":</span>')
      .replace(/:\s*"([^"]*)"/g, ': <span class="json-string">"$1"</span>')
      .replace(/:\s*(\d+\.?\d*)/g, ': <span class="json-number">$1</span>')
      .replace(/:\s*(true|false)/g, ': <span class="json-boolean">$1</span>')
      .replace(/:\s*(null)/g, ': <span class="json-null">$1</span>')
      .replace(/([{}[\]])/g, '<span class="json-bracket">$1</span>');
  };

  const summary = getJsonSummary();
  const sheetsStructure = getSheetsOnlyStructure();

  return (
    <div className="excel-to-json-viewer">
      <div className="upload-section">
        <div className="mode-selector">
          <button
            onClick={toggleMode}
            className={`mode-button ${
              isDiffMode ? "diff-mode" : "single-mode"
            }`}
            type="button"
          >
            {isDiffMode ? "📊 単一ファイル" : "🔄 差分比較"}
          </button>
          <span className="mode-description">
            {isDiffMode
              ? "2つのExcelファイルをJSON形式で比較"
              : "1つのExcelファイルをJSON形式に変換"}
          </span>
        </div>

        {!isDiffMode ? (
          // 単一ファイルモード
          <div className="upload-controls">
            <div className="file-upload-group">
              <input
                key={inputResetKey}
                type="file"
                id="excel-file-input"
                accept=".xlsx,.xls"
                onChange={handleFileSelect}
                className="file-input"
              />
              <label htmlFor="excel-file-input" className="file-label">
                📁 Excelファイルを選択
              </label>
              {fileName && <span className="file-name">📊 {fileName}</span>}
            </div>

            {fileName && (
              <button
                onClick={handleReset}
                className="reset-button"
                type="button"
              >
                🔄 リセット
              </button>
            )}
          </div>
        ) : (
          // 差分比較モード
          <div className="diff-upload-controls">
            <div className="file-upload-group">
              <input
                key={`${inputResetKey}-1`}
                type="file"
                id="excel-file-input-1"
                accept=".xlsx,.xls"
                onChange={handleFileSelect1}
                className="file-input"
              />
              <label htmlFor="excel-file-input-1" className="file-label">
                📁 ファイル1を選択
              </label>
              {fileName1 && <span className="file-name">📊 {fileName1}</span>}
            </div>

            <div className="file-upload-group">
              <input
                key={`${inputResetKey}-2`}
                type="file"
                id="excel-file-input-2"
                accept=".xlsx,.xls"
                onChange={handleFileSelect2}
                className="file-input"
              />
              <label htmlFor="excel-file-input-2" className="file-label">
                📁 ファイル2を選択
              </label>
              {fileName2 && <span className="file-name">📊 {fileName2}</span>}
            </div>

            <div className="diff-controls">
              {(fileName1 || fileName2) && (
                <button
                  onClick={handleReset}
                  className="reset-button"
                  type="button"
                >
                  🔄 リセット
                </button>
              )}

              {fileName1 && fileName2 && (
                <button
                  onClick={performDiffComparison}
                  className="control-button"
                  type="button"
                  disabled={isComparing}
                >
                  {isComparing ? "比較中..." : "🔍 差分比較"}
                </button>
              )}
            </div>
          </div>
        )}

        {isLoading && (
          <div className="loading">
            <div className="loading-spinner"></div>
            <span>{isDiffMode ? "JSON変換中..." : "JSON変換中..."}</span>
          </div>
        )}

        {isComparing && (
          <div className="loading">
            <div className="loading-spinner"></div>
            <span>差分比較中...</span>
          </div>
        )}

        {error && (
          <div className="error">
            <span className="error-icon">⚠️</span>
            <span className="error-message">{error}</span>
          </div>
        )}

        {isDiffMode && (
          <div className="diff-options">
            <h3>🔧 比較オプション</h3>
            <div className="options-grid">
              <div className="option-item">
                <label>
                  <input
                    type="radio"
                    name="compareMode"
                    value="sheets_only"
                    checked={diffOptions.compareMode === "sheets_only"}
                    onChange={(e) =>
                      handleDiffOptionChange("compareMode", e.target.value)
                    }
                  />
                  sheetsのみ比較
                </label>
              </div>
              <div className="option-item">
                <label>
                  <input
                    type="radio"
                    name="compareMode"
                    value="full"
                    checked={diffOptions.compareMode === "full"}
                    onChange={(e) =>
                      handleDiffOptionChange("compareMode", e.target.value)
                    }
                  />
                  全JSON比較
                </label>
              </div>
              <div className="option-item">
                <label>
                  <input
                    type="checkbox"
                    checked={diffOptions.ignoreOrder}
                    onChange={(e) =>
                      handleDiffOptionChange("ignoreOrder", e.target.checked)
                    }
                  />
                  配列の順序を無視
                </label>
              </div>
              <div className="option-item">
                <label>
                  <input
                    type="checkbox"
                    checked={diffOptions.ignoreEmptyValues}
                    onChange={(e) =>
                      handleDiffOptionChange(
                        "ignoreEmptyValues",
                        e.target.checked
                      )
                    }
                  />
                  空値を無視
                </label>
              </div>
            </div>
          </div>
        )}

        {summary && !isDiffMode && (
          <div className="json-summary">
            <h3>📋 ファイル情報</h3>
            <div className="summary-grid">
              <div className="summary-item">
                <span className="summary-label">SpreadJSバージョン:</span>
                <span className="summary-value">{summary.version}</span>
              </div>
              <div className="summary-item">
                <span className="summary-label">シート数:</span>
                <span className="summary-value">{summary.sheetCount}</span>
              </div>
              <div className="summary-item">
                <span className="summary-label">JSON サイズ:</span>
                <span className="summary-value">{summary.fileSize}</span>
              </div>
              <div className="summary-item">
                <span className="summary-label">表示モード:</span>
                <span className="summary-value">{summary.displayMode}</span>
              </div>
              <div className="summary-item">
                <span className="summary-label">スタイル情報:</span>
                <span className="summary-value">
                  {summary.hasStyles ? "✅ あり" : "❌ なし"}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">名前定義:</span>
                <span className="summary-value">
                  {summary.hasNames ? "✅ あり" : "❌ なし"}
                </span>
              </div>
            </div>
          </div>
        )}

        {isDiffMode && diffResult && (
          <div className="diff-summary">
            <h3>📊 差分比較結果</h3>
            <div className="summary-grid">
              <div className="summary-item">
                <span className="summary-label">ファイル1:</span>
                <span className="summary-value file-name" title={fileName1}>
                  {fileName1}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">ファイル2:</span>
                <span className="summary-value file-name" title={fileName2}>
                  {fileName2}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">比較モード:</span>
                <span className="summary-value">
                  {diffOptions.compareMode === "sheets_only"
                    ? "シートのみ"
                    : "完全なJSON"}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">総差分数:</span>
                <span className="summary-value">
                  {diffResult.summary.total}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">追加:</span>
                <span className="summary-value added">
                  {diffResult.summary.added}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">削除:</span>
                <span className="summary-value removed">
                  {diffResult.summary.removed}
                </span>
              </div>
              <div className="summary-item">
                <span className="summary-label">変更:</span>
                <span className="summary-value modified">
                  {diffResult.summary.modified}
                </span>
              </div>
            </div>
          </div>
        )}
      </div>

      {jsonData && !isDiffMode && (
        <div className="json-output-section">
          <div className="output-controls">
            <h3>🔧 SpreadJS JSON出力</h3>
            <div className="control-buttons">
              <button
                onClick={toggleSheetsOnlyMode}
                className="control-button"
                type="button"
                disabled={!jsonData || isLoading || isProcessing}
              >
                {isProcessing
                  ? "切り替え中..."
                  : showSheetsOnly
                  ? "全データ表示"
                  : "シートのみ表示"}
              </button>
              <button
                onClick={toggleRawJsonDisplay}
                className="control-button"
                type="button"
                disabled={isProcessing}
              >
                {isProcessing
                  ? "切り替え中..."
                  : showRawJson
                  ? "構造表示"
                  : "RAW表示"}
              </button>
              <button
                onClick={copyToClipboard}
                className={`control-button ${copySuccess ? "success" : ""}`}
                type="button"
                disabled={isProcessing}
              >
                {copySuccess ? "コピー完了" : "コピー"}
              </button>
              <button
                onClick={downloadJson}
                className="control-button"
                type="button"
                disabled={isProcessing}
              >
                ダウンロード
              </button>
            </div>
          </div>

          {showRawJson ? (
            <div className="json-raw-output">
              <div className="json-controls">
                <div className="json-controls-left">
                  <span className="json-size-info">
                    サイズ: {(jsonString.length / 1024).toFixed(2)} KB
                  </span>
                  <span className="json-lines-info">
                    行数: {jsonString.split("\n").length}
                  </span>
                </div>
                <div className="json-controls-right">
                  <button
                    onClick={() => {
                      const jsonElement =
                        document.querySelector(".json-content");
                      if (jsonElement) {
                        jsonElement.scrollTop = 0;
                      }
                    }}
                    className="json-scroll-top"
                    type="button"
                  >
                    ⬆️ トップに戻る
                  </button>
                </div>
              </div>
              <pre
                className="json-content highlighted"
                dangerouslySetInnerHTML={{
                  __html: highlightJson(jsonString),
                }}
              />
            </div>
          ) : (
            <div className="json-structured-output">
              <div className="json-structure">
                {showSheetsOnly && sheetsStructure ? (
                  <div className="structure-section">
                    <h4>📊 シート構造 (sheetsのみ)</h4>
                    <div className="sheets-list">
                      {sheetsStructure.map((sheet: any) => (
                        <div key={sheet.index} className="sheet-item">
                          <div className="sheet-header">
                            <span className="sheet-name">{sheet.name}</span>
                            <span className="sheet-size">
                              ({sheet.rowCount} × {sheet.columnCount})
                            </span>
                          </div>
                          <div className="sheet-details">
                            <span className="detail-item">
                              表示: {sheet.visible ? "✅ 表示" : "❌ 非表示"}
                            </span>
                            <span className="detail-item">
                              データ: {sheet.hasData ? "✅ あり" : "❌ なし"}
                            </span>
                            <span className="detail-item">
                              スタイル:{" "}
                              {sheet.hasStyles ? "✅ あり" : "❌ なし"}
                            </span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>
                ) : (
                  <>
                    {jsonData.sheets && (
                      <div className="structure-section">
                        <h4>📊 シート構造</h4>
                        <div className="sheets-list">
                          {/* sheetsが配列の場合 */}
                          {Array.isArray(jsonData.sheets) &&
                            jsonData.sheets.map((sheet: any, index: number) => (
                              <div key={index} className="sheet-item">
                                <div className="sheet-header">
                                  <span className="sheet-name">
                                    {sheet.name || `Sheet${index + 1}`}
                                  </span>
                                  {sheet.rowCount && sheet.columnCount && (
                                    <span className="sheet-size">
                                      ({sheet.rowCount} × {sheet.columnCount})
                                    </span>
                                  )}
                                </div>
                                {sheet.data && (
                                  <div className="sheet-details">
                                    <span className="detail-item">
                                      データセル:{" "}
                                      {
                                        Object.keys(sheet.data.dataTable || {})
                                          .length
                                      }
                                    </span>
                                    {sheet.styles && (
                                      <span className="detail-item">
                                        スタイル:{" "}
                                        {Object.keys(sheet.styles).length}
                                      </span>
                                    )}
                                    {sheet.spans && (
                                      <span className="detail-item">
                                        結合セル:{" "}
                                        {Object.keys(sheet.spans).length}
                                      </span>
                                    )}
                                  </div>
                                )}
                              </div>
                            ))}

                          {/* sheetsがオブジェクトの場合 */}
                          {!Array.isArray(jsonData.sheets) &&
                            typeof jsonData.sheets === "object" &&
                            Object.entries(jsonData.sheets).map(
                              (
                                [sheetName, sheet]: [string, any],
                                index: number
                              ) => (
                                <div key={sheetName} className="sheet-item">
                                  <div className="sheet-header">
                                    <span className="sheet-name">
                                      {sheet.name || sheetName}
                                    </span>
                                    {sheet.rowCount && sheet.columnCount && (
                                      <span className="sheet-size">
                                        ({sheet.rowCount} × {sheet.columnCount})
                                      </span>
                                    )}
                                  </div>
                                  {sheet.data && (
                                    <div className="sheet-details">
                                      <span className="detail-item">
                                        データセル:{" "}
                                        {
                                          Object.keys(
                                            sheet.data.dataTable || {}
                                          ).length
                                        }
                                      </span>
                                      {sheet.styles && (
                                        <span className="detail-item">
                                          スタイル:{" "}
                                          {Object.keys(sheet.styles).length}
                                        </span>
                                      )}
                                      {sheet.spans && (
                                        <span className="detail-item">
                                          結合セル:{" "}
                                          {Object.keys(sheet.spans).length}
                                        </span>
                                      )}
                                    </div>
                                  )}
                                </div>
                              )
                            )}
                        </div>
                      </div>
                    )}

                    <div className="structure-section">
                      <h4>🔧 JSON プロパティ</h4>
                      <div className="properties-list">
                        {Object.keys(jsonData).map((key) => (
                          <div key={key} className="property-item">
                            <span className="property-key">{key}</span>
                            <span className="property-type">
                              {Array.isArray(jsonData[key])
                                ? `Array[${jsonData[key].length}]`
                                : typeof jsonData[key]}
                            </span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </>
                )}
              </div>
            </div>
          )}
        </div>
      )}

      {isDiffMode && diffResult && (
        <div className="diff-result-section">
          <div className="diff-controls">
            <h3>📋 差分詳細</h3>
            <div className="control-buttons">
              <button
                onClick={toggleFullscreen}
                className="control-button"
                type="button"
                title={isFullscreen ? "全画面を終了 (ESC)" : "全画面で表示"}
              >
                {isFullscreen ? "🗗 全画面終了" : "🗖 全画面表示"}
              </button>
              <button
                onClick={() => downloadDiffResult("json")}
                className="control-button"
                type="button"
              >
                📁 JSON出力
              </button>
              <button
                onClick={() => downloadDiffResult("csv")}
                className="control-button"
                type="button"
              >
                📊 CSV出力
              </button>
            </div>
          </div>

          {lineDiffs && (
            <div className="diff-table-wrapper">
              {lineDiffs.length === 0 ? (
                <div className="no-differences">
                  <span>✅ 違いはありません</span>
                </div>
              ) : (
                <table className="diff-table">
                  <thead>
                    <tr>
                      <th>ファイル1</th>
                      <th>ファイル2</th>
                    </tr>
                  </thead>
                  <tbody>
                    {lineDiffs.map((row, index) => (
                      <tr key={index} className={`diff-${row.type}`}>
                        <td className="left-cell">{row.left || ""}</td>
                        <td className="right-cell">{row.right || ""}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}
        </div>
      )}

      {isDiffMode && (fileName1 || fileName2) && !diffResult && (
        <div className="diff-preview-section">
          <h3>📋 ファイル確認</h3>
          <div className="file-preview-grid">
            {fileName1 && (
              <div className="file-preview">
                <h4>📊 ファイル1: {fileName1}</h4>
                {jsonData1 ? (
                  <div className="preview-info">
                    <span>✅ 読み込み完了</span>
                    <span>
                      サイズ: {(jsonString1.length / 1024).toFixed(2)} KB
                    </span>
                  </div>
                ) : (
                  <div className="preview-info">
                    <span>⏳ 読み込み中...</span>
                  </div>
                )}
              </div>
            )}
            {fileName2 && (
              <div className="file-preview">
                <h4>📊 ファイル2: {fileName2}</h4>
                {jsonData2 ? (
                  <div className="preview-info">
                    <span>✅ 読み込み完了</span>
                    <span>
                      サイズ: {(jsonString2.length / 1024).toFixed(2)} KB
                    </span>
                  </div>
                ) : (
                  <div className="preview-info">
                    <span>⏳ 読み込み中...</span>
                  </div>
                )}
              </div>
            )}
          </div>
        </div>
      )}

      {/* 全画面表示モード */}
      {isFullscreen && isDiffMode && diffResult && (
        <div className="diff-fullscreen-overlay">
          <div className="diff-fullscreen-content">
            <div className="diff-fullscreen-header">
              <h2>📋 差分詳細 - 全画面表示</h2>
              <div className="diff-fullscreen-controls">
                <span className="fullscreen-hint">ESCキーで終了</span>
                <button
                  onClick={toggleFullscreen}
                  className="control-button fullscreen-close-button"
                  type="button"
                  title="全画面を終了 (ESC)"
                >
                  ✕ 閉じる
                </button>
              </div>
            </div>

            <div className="diff-fullscreen-info">
              <div className="fullscreen-summary">
                <span className="summary-item">
                  <span className="summary-label">総差分数:</span>
                  <span className="summary-value">
                    {diffResult.summary.total}
                  </span>
                </span>
                <span className="summary-item">
                  <span className="summary-label">追加:</span>
                  <span className="summary-value added">
                    {diffResult.summary.added}
                  </span>
                </span>
                <span className="summary-item">
                  <span className="summary-label">削除:</span>
                  <span className="summary-value removed">
                    {diffResult.summary.removed}
                  </span>
                </span>
                <span className="summary-item">
                  <span className="summary-label">変更:</span>
                  <span className="summary-value modified">
                    {diffResult.summary.modified}
                  </span>
                </span>
              </div>
            </div>

            <div className="diff-fullscreen-table-wrapper">
              {lineDiffs && lineDiffs.length === 0 ? (
                <div className="no-differences">
                  <span>✅ 違いはありません</span>
                </div>
              ) : (
                <table className="diff-table diff-fullscreen-table">
                  <thead>
                    <tr>
                      <th>ファイル1: {fileName1}</th>
                      <th>ファイル2: {fileName2}</th>
                    </tr>
                  </thead>
                  <tbody>
                    {lineDiffs?.map((row, index) => (
                      <tr key={index} className={`diff-${row.type}`}>
                        <td className="left-cell">{row.left || ""}</td>
                        <td className="right-cell">{row.right || ""}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default ExcelToJsonViewer;

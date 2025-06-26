import React, { useState, useRef, useEffect } from "react";
import { SpreadSheets, Worksheet } from "@grapecity/spread-sheets-react";
import * as GC from "@grapecity/spread-sheets";
import * as ExcelIO from "@grapecity/spread-excelio";
import "./ExcelViewer.css";

/// エラー情報の型定義
interface ErrorInfo {
  message: string;
  details?: string;
  type: "file" | "format" | "read" | "unknown";
  suggestions?: string[];
}

/// ExcelViewerコンポーネント - Excelファイルのアップロードとプレビュー機能を提供
const ExcelViewer: React.FC = () => {
  const [fileName, setFileName] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [error, setError] = useState<ErrorInfo | null>(null);
  const [showErrorDetails, setShowErrorDetails] = useState<boolean>(false);
  const [isFullscreen, setIsFullscreen] = useState<boolean>(false);
  const spreadRef = useRef<GC.Spread.Sheets.Workbook | null>(null);

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
    if (spreadRef.current) {
      // 状態変更後のDOM更新を待ってからリサイズ処理を実行
      const timeoutId = setTimeout(() => {
        try {
          // SpreadSheetsのレイアウトを更新
          spreadRef.current!.refresh();
          // 追加でinvalidateLayoutも呼び出してより確実にサイズ更新
          spreadRef.current!.invalidateLayout();
        } catch (error) {
          console.warn("SpreadSheets resize error:", error);
        }
      }, 150); // より確実にするため少し長めの遅延

      return () => clearTimeout(timeoutId);
    }
  }, [isFullscreen]);

  /// SpreadSheetsの初期化完了時の処理
  const onWorkbookInitialized = (spread: GC.Spread.Sheets.Workbook) => {
    spreadRef.current = spread;
    // スプレッドシートの基本設定
    spread.suspendPaint();
    spread.options.tabStripVisible = true;
    spread.options.newTabVisible = false;
    spread.options.tabNavigationVisible = true;
    spread.resumePaint();
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
  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    // ファイルサイズのチェック（100MB制限）
    const maxSize = 100 * 1024 * 1024; // 100MB
    if (file.size > maxSize) {
      setErrorInfo(
        "ファイルサイズが大きすぎます",
        "file",
        `ファイルサイズ: ${(file.size / 1024 / 1024).toFixed(
          2
        )}MB（制限: 100MB）`,
        [
          "より小さなファイルを選択してください",
          "不要なシートやデータを削除してファイルサイズを縮小してください",
        ]
      );
      return;
    }

    // サポートされているファイル形式の確認
    const supportedFormats = [".xlsx", ".xls", ".csv"];
    const fileExtension = file.name
      .toLowerCase()
      .substring(file.name.lastIndexOf("."));

    if (!supportedFormats.includes(fileExtension)) {
      setErrorInfo(
        "サポートされていないファイル形式です",
        "format",
        `ファイル形式: ${fileExtension}`,
        [
          "xlsx、xls、csv形式のファイルを選択してください",
          "Excelで別の形式で保存し直してください",
        ]
      );
      return;
    }

    setError(null);
    setShowErrorDetails(false);
    setFileName(file.name);
    loadExcelFile(file);
  };

  /// Excelファイルの読み込み処理
  const loadExcelFile = async (file: File) => {
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

      // ファイルをSpreadSheetsに読み込み
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

      let errorMessage = "ファイルの読み込み中にエラーが発生しました";
      let errorDetails = "";
      let suggestions: string[] = [];

      if (error instanceof Error) {
        errorMessage = error.message;
        errorDetails = error.stack || error.toString();

        // エラーメッセージに基づいて具体的な提案を生成
        if (
          error.message.includes("password") ||
          error.message.includes("protected")
        ) {
          suggestions = [
            "パスワードで保護されたファイルはサポートされていません",
            "パスワード保護を解除してから再度お試しください",
          ];
        } else if (
          error.message.includes("corrupt") ||
          error.message.includes("invalid")
        ) {
          suggestions = [
            "ファイルが破損している可能性があります",
            "元のファイルを確認して、正常なExcelファイルを選択してください",
            "別のアプリケーションでファイルを開いて修復を試してください",
          ];
        } else if (
          error.message.includes("format") ||
          error.message.includes("unsupported")
        ) {
          suggestions = [
            "ファイル形式がサポートされていない可能性があります",
            "Excel形式（.xlsx）で保存し直してください",
          ];
        } else {
          suggestions = [
            "別のファイルで試してください",
            "ファイルが使用中でないか確認してください",
            "ブラウザを再読み込みして再度お試しください",
          ];
        }
      }

      setErrorInfo(errorMessage, "read", errorDetails, suggestions);
    } finally {
      setIsLoading(false);
    }
  };

  /// ファイル選択のリセット
  const handleReset = () => {
    setFileName("");
    setError(null);
    setShowErrorDetails(false);
    setIsFullscreen(false);
    if (spreadRef.current) {
      spreadRef.current.clearSheets();
      spreadRef.current.addSheet(0);
    }
  };

  /// エラー詳細の表示切り替え
  const toggleErrorDetails = () => {
    setShowErrorDetails(!showErrorDetails);
  };

  /// 全画面表示の切り替え
  const toggleFullscreen = () => {
    setIsFullscreen(!isFullscreen);
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
      default:
        return "⚠️";
    }
  };

  return (
    <div className={`excel-viewer ${isFullscreen ? "fullscreen" : ""}`}>
      {!isFullscreen && (
        <div className="upload-section">
          <div className="upload-controls">
            <input
              type="file"
              id="file-input"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileSelect}
              className="file-input"
            />
            <label htmlFor="file-input" className="file-label">
              📁 Excelファイルを選択
            </label>
            {fileName && (
              <button
                onClick={handleReset}
                className="reset-button"
                type="button"
              >
                リセット
              </button>
            )}
          </div>

          {fileName && (
            <div className="file-info">
              <span className="file-name">📊 {fileName}</span>
            </div>
          )}

          {isLoading && (
            <div className="loading">
              <div className="loading-spinner"></div>
              <span>ファイルを読み込み中...</span>
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
        </div>
      )}

      <div className="spreadsheet-container">
        <div className="spreadsheet-header">
          {fileName && (
            <div className="spreadsheet-title">
              <span className="file-icon">📊</span>
              <span className="title-text">{fileName}</span>
            </div>
          )}
          <div className="spreadsheet-controls">
            {fileName && (
              <button
                onClick={toggleFullscreen}
                className="fullscreen-button"
                type="button"
                title={isFullscreen ? "全画面を解除 (ESC)" : "全画面表示"}
              >
                {isFullscreen ? "🗗" : "🗖"}
                {isFullscreen ? "縮小" : "全画面"}
              </button>
            )}
            {isFullscreen && (
              <button
                onClick={handleReset}
                className="reset-button-fullscreen"
                type="button"
                title="リセット"
              >
                🔄 リセット
              </button>
            )}
          </div>
        </div>

        <SpreadSheets
          workbookInitialized={onWorkbookInitialized}
          hostStyle={{
            width: "100%",
            height: isFullscreen ? "calc(100vh - 60px)" : "600px",
            border: "1px solid #ccc",
            borderRadius: isFullscreen ? "0" : "4px",
          }}
        >
          <Worksheet />
        </SpreadSheets>
      </div>

      {isFullscreen && (
        <div className="fullscreen-help">
          <span>ESCキーまたは「縮小」ボタンで全画面を解除</span>
        </div>
      )}
    </div>
  );
};

export default ExcelViewer;

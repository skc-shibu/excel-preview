import React, { useState } from "react";
import * as ExcelIO from "@grapecity/spread-excelio";
import "./ExcelToJsonViewer.css";

/// ExcelToJsonViewerコンポーネント - ExcelファイルをJSON形式に変換して表示
const ExcelToJsonViewer: React.FC = () => {
  const [fileName, setFileName] = useState<string>("");
  const [isLoading, setIsLoading] = useState<boolean>(false);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [error, setError] = useState<string>("");
  const [jsonData, setJsonData] = useState<any>(null);
  const [jsonString, setJsonString] = useState<string>("");
  const [showRawJson, setShowRawJson] = useState<boolean>(false);
  const [copySuccess, setCopySuccess] = useState<boolean>(false);
  const [showSheetsOnly, setShowSheetsOnly] = useState<boolean>(true);

  /// ファイル選択時の処理
  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;

    if (!validateFile(file)) return;

    setError("");
    setFileName(file.name);
    loadExcelFile(file);
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

  /// リセット処理
  const handleReset = () => {
    setFileName("");
    setError("");
    setJsonData(null);
    setJsonString("");
    setShowRawJson(false);
    setCopySuccess(false);
    setShowSheetsOnly(true);
    setIsProcessing(false);
  };

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
        <div className="upload-controls">
          <div className="file-upload-group">
            <input
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

        {isLoading && (
          <div className="loading">
            <div className="loading-spinner"></div>
            <span>JSON変換中...</span>
          </div>
        )}

        {error && (
          <div className="error">
            <span className="error-icon">⚠️</span>
            <span className="error-message">{error}</span>
          </div>
        )}

        {summary && (
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
      </div>

      {jsonData && (
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
    </div>
  );
};

export default ExcelToJsonViewer;

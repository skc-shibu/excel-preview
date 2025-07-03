import React, { useState } from "react";
import "./App.css";
import ExcelViewer from "./components/ExcelViewer";
import ExcelToJsonViewer from "./components/ExcelToJsonViewer";

/// アプリケーションのモード定義
type AppMode = "viewer" | "json";

/// Appコンポーネント - アプリケーションのメインコンポーネント
const App: React.FC = () => {
  const [mode, setMode] = useState<AppMode>("viewer");

  /// モードに応じたコンポーネントのレンダリング
  const renderModeComponent = () => {
    switch (mode) {
      case "viewer":
        return <ExcelViewer />;
      case "json":
        return <ExcelToJsonViewer />;
      default:
        return <ExcelViewer />;
    }
  };

  return (
    <div className="App">
      <header className="App-header">
        <h1>Excel ファイルビューア</h1>
        <p>SpreadJS を使用してExcelファイルをプレビューできます</p>

        <div className="mode-selector">
          <button
            className={`mode-button ${mode === "viewer" ? "active" : ""}`}
            onClick={() => setMode("viewer")}
            type="button"
          >
            📊 通常表示
          </button>
          <button
            className={`mode-button ${mode === "json" ? "active" : ""}`}
            onClick={() => setMode("json")}
            type="button"
          >
            🔧 JSON変換
          </button>
        </div>
      </header>
      <main className="App-main">{renderModeComponent()}</main>
    </div>
  );
};

export default App;

import React, { useState } from "react";
import "./App.css";
import ExcelViewer from "./components/ExcelViewer";
import ExcelDiffViewer from "./components/ExcelDiffViewer";

/// アプリケーションのモード定義
type AppMode = "viewer" | "diff";

/// Appコンポーネント - アプリケーションのメインコンポーネント
const App: React.FC = () => {
  const [mode, setMode] = useState<AppMode>("viewer");

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
            className={`mode-button ${mode === "diff" ? "active" : ""}`}
            onClick={() => setMode("diff")}
            type="button"
          >
            🔍 差分比較
          </button>
        </div>
      </header>
      <main className="App-main">
        {mode === "viewer" ? <ExcelViewer /> : <ExcelDiffViewer />}
      </main>
    </div>
  );
};

export default App;

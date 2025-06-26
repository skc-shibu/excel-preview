import React from "react";
import "./App.css";
import ExcelViewer from "./components/ExcelViewer";

/// Appコンポーネント - アプリケーションのメインコンポーネント
const App: React.FC = () => {
  return (
    <div className="App">
      <header className="App-header">
        <h1>Excel ファイルビューア</h1>
        <p>SpreadJS を使用してExcelファイルをプレビューできます</p>
      </header>
      <main className="App-main">
        <ExcelViewer />
      </main>
    </div>
  );
};

export default App;

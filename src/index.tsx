import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css";
import App from "./App";

/// SpreadJSライブラリのインポート
import * as GC from "@grapecity/spread-sheets";

/// SpreadJSの日本語化ユーティリティ
import { applyJapaneseCulture } from "./utils/spreadJSLocalization";

/// SpreadJSの日本語リソースの読み込みと日本語ロケール設定
(async () => {
  // @ts-ignore
  await import("@mescius/spread-sheets-resources-ja");
  applyJapaneseCulture(GC);
})();

const root = ReactDOM.createRoot(
  document.getElementById("root") as HTMLElement
);
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);

import * as GC from "@grapecity/spread-sheets";

/// SpreadJSの日本語設定を適用する
export const applyJapaneseCulture = (GC_Instance: typeof GC): void => {
  try {
    // CultureManagerが存在するかチェック
    if (GC_Instance.Spread.Common && GC_Instance.Spread.Common.CultureManager) {
      GC_Instance.Spread.Common.CultureManager.culture("ja-jp");
      console.log("SpreadJSの日本語ロケールが正常に適用されました");
    } else {
      console.warn("CultureManagerが見つかりません");
    }
  } catch (error) {
    console.error("日本語設定の適用に失敗しました:", error);
  }
};

/// SpreadJSの総合的な日本語化設定を適用する
export const setupSpreadJSJapanese = (
  GC_Instance: typeof GC,
  spread?: GC.Spread.Sheets.Workbook
): void => {
  // ロケール設定
  applyJapaneseCulture(GC_Instance);
};

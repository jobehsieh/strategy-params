# 交接檔（handoff.md）

> 任何 Agent、任何電腦接手前必讀；收工時必更新。

## ⏯️ 目前做到哪

- 已完成 `策略最佳化參數彙整_20260929.xlsx`，存放於專案根目錄。
- 已從 13 個 `_BEST.wsp` 的 MultiCharts「策略回測績效報告」擷取「平倉權益曲線」，圖檔位於 `screenshots\<策略名稱>.png`。
- 已重跑 `scripts\build_summary.py --source "H:\我的雲端硬碟\Muti-Agents 0924"`，13 張圖皆嵌入「平倉權益曲線圖」分頁。
- 最終驗證：13 個策略、38 輪、106 個 WFO 視窗、23 個來源檔、13 張內嵌圖片；成果檔大小 787,848 bytes。
- MultiCharts 程式參數一致性比對維持 13/13 一致；主表第 18 欄與 `wsp參數比對`分頁均由腳本產生。
- Typeless 與 DeskIn 曾為避免 UI 輸入干擾而暫停，收工前已恢復；`DeskIn_Service` 為 Running / Automatic。

## 🚦 目前狀態

- ✅ 階段三完成：參數彙整、WFO 明細、wsp 比對、13 張權益曲線與最終 Excel 均已完成。
- ✅ H 槽專案資料夾已有最新 Excel 與 13 張 PNG。
- ✅ `AGENTS.md`、`handoff.md` 原檔有保留並在本次收工更新。
- ℹ️ `screenshots/*.png` 依 `.gitignore` 不進版控；Excel `策略最佳化參數彙整_20260929.xlsx` 是本次新增產物。
- ℹ️ L3 Obsidian 未啟用，無可用 Obsidian MCP，故本次不更新 L3。

## ➡️ 下一步

1. 進入階段四：決定下一批要納入的策略或最佳化工作區。
2. 若來源報告有更新，重跑 `python scripts\build_summary.py --source "H:\我的雲端硬碟\Muti-Agents 0924"`，再核對 13/38/106/23 與圖片數。
3. 若需重新擷取權益曲線，先載入對應 `_BEST.wsp`，開啟主工具列「策略績效報告」，確認左側選中「平倉權益曲線」後擷取。

## ⚠️ 注意事項

- MultiCharts 同時開啟過多工作區時，新工作區可能只顯示空白；關閉已完成的工作區並選「否」不儲存，再重新載入即可。
- 報告按鈕位於主工具列第一列，既有 1214x645 視窗下約在 `(896, 72)`；座標會隨視窗尺寸改變，操作前必須重新觀察。
- `*_BEST.wsp` 的 input 現值已與主表最佳參數核對一致，不要再手動改參數。
- 「所有參數設定值」與「R3 WFO推薦參數」是不同口徑，切勿混用。
- 策略 7 的樣本外結果不穩定；策略 12 的 Multiplier 與 TakeProfit 位於搜尋邊界；策略 1 缺少可估算月報酬率的資料區間。
- Codex 附加工作樹 `C:\Users\Administrator\.codex\worktrees\multicharts-equity-curves\最佳化參數彙整` 的 `.git` 指向 H 槽已失效；後續 Git 操作請直接使用 `H:\我的雲端硬碟\最佳化參數彙整`。

## 🕐 最後更新

- 時間：2026-09-29（收工）
- 更新者：Codex @ DESKTOP-9IROG8R
- 本次重點：完成 13 張權益曲線擷取、最終 Excel 重建與 H 槽同步；Typeless、DeskIn 已恢復。
- Git push：待推
- L3 Obsidian：未啟用，略過

# 投資模擬器 Trading Sim

用**虛擬資金**練習加密貨幣與股票 ETF 的出入金、換匯、買賣與損益追蹤。全程不連接任何真實帳戶，也不會有實際金流。

> **線上版：** [trading-sim.charleschang0306.workers.dev](https://trading-sim.charleschang0306.workers.dev/)

## 專案內容

| 檔案 | 說明 |
|---|---|
| `trading-simulator.html` | 單檔網頁版主程式，雙擊即可在瀏覽器離線使用 |
| `index.html` | 與主程式同步的首頁入口，供 Cloudflare Workers / 靜態託管使用 |
| `投資模擬器_範本.xlsx` | Excel 基準幣記帳範本 |
| `build_xlsx.py` | 將網頁版備份 JSON 轉成填好資料的 Excel |
| `README.md` | 專案說明 |

## 網頁版功能

### 資產與帳務

- **多幣別錢包**：TWD / USD 現金錢包分開管理。
- **出入金**：可選幣別入金/出金，淨投入以當下匯率換算為基準幣。
- **換匯**：TWD ⇄ USD，支援匯率與手續費；USD 換回 TWD 時認列匯率損益。
- **買賣交易**：支援 BTC / ETH / SOL / VOO / VT / 0050，也可自訂資產。
- **自動換匯**：買進時若報價幣別餘額不足，可勾選自動從基準幣換匯。
- **成本法**：持倉採移動平均成本，並保留買進當時基準幣成本，用於拆分交易/匯率損益。
- **現金利息 / 賣出稅**：可設定現金年利率、ETF/加密賣出稅，並手動計入 30 天利息。

### DCA 與分析

- **定期定額 DCA**：設定標的、金額、週/月頻率、起訖日；到期可手動執行，開啟頁面時也會自動補跑到期計畫（補跑以當下價格近似）。
- **績效指標**：XIRR、簡單年化、年化波動、Sharpe、最大回撤。
- **視覺化分析**：回撤曲線、月報酬熱力圖。
- **基準對照**：以資產參數近似的買進持有基準比較（模擬參考，非真實歷史回測）。
- **蒙地卡羅模擬**：依目前持倉參數估計未來終值分位數（P10/P25/P50/P75/P90），支援每月追加投入。

### 行情與資料

- **價格模式**：每檔資產可選手動、隨機漫步、即時 API。
  - 加密貨幣：Binance（備援 CoinGecko）
  - 台股 ETF：TWSE 每日收盤價
  - 美股 ETF：Stooq（必要時可在設定填自訂 CORS 代理）
- **交易紀錄**：完整流水帳，支援篩選、搜尋、刪除後重算、CSV 匯出。
- **備份還原**：JSON 匯出/匯入，可跨裝置搬移資料。
- **PWA 支援**：含 manifest、行動裝置 meta、icon；Service Worker 快取在 HTTP(S) 環境啟用。

## 使用方式

1. 開啟[線上版](https://trading-sim.charleschang0306.workers.dev/)，或雙擊本機 `trading-simulator.html`。
2. 先到 **設定 → 載入範例資料** 快速體驗，或從 **交易 / 出入金** 開始入金。
3. 需要換幣時使用 **換匯**，再進行買賣或設定 DCA。
4. 到 **分析** 查看 XIRR、回撤、熱力圖、基準對照與蒙地卡羅結果。

## 資料保存

- 資料儲存在瀏覽器 `localStorage`（`twsim.v3`），不會上傳。
- 舊版 `twsim.v1 / twsim.v2` 資料會在載入時自動遷移。
- 換裝置或清理瀏覽器前，請使用 **設定 → 備份 JSON**。

## Excel 範本

`投資模擬器_範本.xlsx` 提供基準幣總帳視圖，適合對帳與自訂公式。  
注意：Excel 版不包含完整多幣別錢包、換匯損益拆分、DCA、XIRR/蒙地卡羅等網頁版進階功能。

可將網頁版備份 JSON 轉成 Excel：

```powershell
python build_xlsx.py trading-sim-backup.json
```

## 部署

本專案已部署到 Cloudflare Workers：<https://trading-sim.charleschang0306.workers.dev/>

若要自行部署靜態版本，可將 repo 連接到 Cloudflare Pages/Workers，使用根目錄 `/` 作為輸出目錄，無需 build command。

## 注意事項

- 本工具僅供模擬與教育用途，數字不構成投資建議。
- DCA 補跑、基準對照、蒙地卡羅為近似模擬/估計，不代表真實歷史成交或未來表現。
- 美股即時報價可能受 CORS 或來源限制影響，可改用手動/隨機模式或自訂代理。

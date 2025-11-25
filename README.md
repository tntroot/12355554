<style>
/* 建議在本地環境或支援 HTML 渲染的 Markdown 檢視器使用 */
body { font-family: 'Meiryo UI', 'Microsoft JhengHei', 'Segoe UI', sans-serif; color: #222; line-height: 1.6; }
h1 { color: #0f9ed5; font-weight: 700; }
h2 { color: #1565a9; margin-top: 1.2rem; }
h3 { color: #333; margin-top: 1rem; }
code, pre { background: #f6f8fa; padding: 6px 10px; border-radius: 6px; font-size: 0.95rem; }
.note { background: #f0f7ff; border-left: 4px solid #0f9ed5; padding: 8px 12px; margin: 8px 0; border-radius: 4px; }
.small { font-size: 0.9rem; color: #555; }
</style>

# workRewards — addBtnModel1（上班紀錄自動化）

簡介
- addBtnModel1 模組為專案的核心：讀取「輸入資料」工作表，將當天出勤或請假資料新增或更新至對應月份工作表（yyyy_mm月），並計算加班、補齊中間缺漏日期、調整格式與樣式。

概觀功能
- 驗證輸入欄位（姓名、日期、上下班時間或請假資訊）
- 依輸入日期自動建立月份工作表，並做標題與表頭樣式
- 新增/更新單筆出勤紀錄（含請假）
- 自動補齊月份內缺漏日期（標為休假）
- 根據工作地點讀取設定（無效工時、上班提前、下班延後）
- 計算並顯示該月份加班總計（H1:I2）
- 自動設定欄寬、文字換行與列高

目錄
- 必要檔案
- 資料欄位對應
- 月份表結構
- 子程序與函式說明
- 使用方式（快速上手）
- 部署與注意事項
- 已知限制與改進建議
- 版本紀錄

必要檔案
- addBtnModel1.bas（主流程）
- functionModel.bas（工具函式：getLastDate、dayOff、writeData、writeLeaveData、wsMonthFormat）
- countOverTime.bas（計算加班合計）
- Excel 活頁簿：須含名為「輸入資料」之工作表

輸入資料（「輸入資料」工作表）欄位
- E4：姓名（必填）
- D6：年（例：2025）
- E6：月（例：11）
- F6：日（例：25）
- D8：上班時間（格式：HHMM，例如 0830 或 830）
- E8：下班時間（格式：HHMM）
- C8：工作地點（會以此至 J6:J8 範圍搜尋設定）
- D9：備註（會填入月份表 F 欄）
- D12：請假類型（如有）
- F12：請假原因（如有）
- 設定表（工作地點對應設定）：
  - J 欄：工作地點名稱
  - K 欄：invalidHours（無效工時）
  - L 欄：adjStart（上班提前分鐘）
  - M 欄：adjEnd（下班延後分鐘）

月份表（yyyy_mm月）欄位與樣式
- A1：合併標題，範例："2025_11月 上班紀錄"
- A2:G2：表頭內容 → ["日期", "上班時間", "下班時間", "工作地點", "加班時數 (小時)", "備註", "結果"]
- 資料行：
  - A：日期（yyyy/mm/dd）
  - B：上班時間（hh:mm）
  - C：下班時間（hh:mm）
  - D：工作地點
  - E：加班時數（公式計算）
  - F：備註
  - G：結果（輸入摘要或請假文字）

重要公式（E 欄：加班時數）
- 範例（在 writeData 寫入的 FormulaR1C1）：
  - FLOOR.MATH(MAX(0, (RC[-2] - RC[-3]) * 24 - 9 - invalidHours - IF((RC[-2]-RC[-3])*24 >= 9.5, 0.5, 0)), 0.5)
- 意義：當日工時扣掉正常 9 小時與無效工時，若時數≥9.5 再扣 0.5 小時；結果以 0.5 小時為單位向下取整。

子程序 / 函式（重點）
- 添加()（addBtnModel1）
  - 主入口，檢查輸入、解析日期、建立或取得月份表、找出最後一筆資料、判斷是否已存在該日期，呼叫 writeData 或 writeLeaveData，最後執行加班總計計算。
- getLastDate(wsMonth, lastDate, nextRow)（functionModel）
  - 取得月份表最後一筆日期，並回傳下一筆要寫入的 nextRow（起始為第 3 列）。
- dayOff(wsMonth, lastDate, nextRow, ii)（functionModel）
  - 在 nextRow 插入休假行，並設定背景色與文字 ("休假")。
- writeData(ParamArray)（functionModel）
  - 寫入一般出勤資料（包含把上/下班時間寫進 B,C 列，填工作地點、備註與加班公式），並套用格式。
- writeLeaveData(ParamArray)（functionModel）
  - 寫入請假資料（合併 B:F，填入請假類型），並於 G 欄填入結果提示。
- 計算加班(fistUseWsMonth As Boolean)（countOverTime）
  - 計算整個月份的 E 欄總和並寫入 H2:I2；若新建立月份表則會同時建立標題 H1。

使用方式（快速上手）
1. 將三個 .bas 檔匯入 Excel 巨集模組（VBE）。
2. 開啟工作表「輸入資料」，填好姓名、日期、時間、工作地點等欄位。
3. 在 Excel 執行巨集：選擇 添加() 或為表單按鈕綁定 添加 程式。
4. 檢查對應月份工作表（例：2025_11月）內容是否正確，確認加班總表 H1:I2。

部署與注意事項
- Excel 必須允許巨集執行（Macro enabled）。
- D8/E8 格式允許數字或文字（0830/830/08:30），模組會以 Format(..., "0000") 轉換為 4 位數並解析為時分。
- 若工作地點找不到對應設定（J6:J8），則 invalidHours/adjStart/adjEnd 會以 0 作為預設。
- 若請假欄位有輸入（D12 或 F12），上/下班時間可留空（系統會寫入請假內容，而不是出勤時間）。

已知限制
- 跨日班次（下班時間屬於隔日）尚未自動處理；可能導致工時負值或不準確。建議在 writeData 添加判斷：若 endTime < startTime，則 endTime = endTime + 1。
- 部分函式（如 wsMonthFormat、getLastDate）在專案中尚有未完整或待修正之處，需要檢查合併儲存格、邏輯與回傳值。
- 若使用者手動變動合併儲存格或樣式，程式可能會解除合併或覆寫格式。

建議改進（優先順序）
1. 支援跨日班次（endTime < startTime → 加 1 天）。
2. 增加 UserForm 友善輸入介面與資料驗證。
3. 加入匯出功能：月報、匯總統計（工作地點、出勤/請假統計）。
4. 加強錯誤處理與輸入驗證，避免非預期 Null/Empty 值。
5. 補上完整單元測試（模擬常見情境）。

範例：新增一筆出勤（可複製於輸入表）
- E4: 王小明
- D6: 2025，E6: 11，F6: 25
- D8: 0830，E8: 1730
- C8: 公司A
- D9: 備註文字（可選）
- 執行 添加 → 自動寫入 2025_11月

版本與變更紀錄
- v1.0 — 初始版本：支援新增/更新/請假/加班合計/補缺日期
- v1.1 — 建議新增：跨日班次支援、UserForm UI、輸入格式容錯性提高

聯絡與維護
- 若需要幫助或回報錯誤，請聯絡專案維護者或在 commit message 註明修改者、用途與變更說明。

.small 本文件以 addBtnModel1 為主，詳細執行細節以實際模組程式碼為準；請在修改程式時同步更新 README。

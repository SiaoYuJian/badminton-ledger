# 羽球社儲值金查詢

輸入姓名 → 顯示 2026 年打球記錄、費用、儲值與目前剩餘儲值金。純靜態網頁,可放 GitHub Pages / Netlify。

## 結構
- `scripts/build_data.py` — 讀 `「2026場地費明細」`,每人輸出一個 JSON 到 `web/data/`
- `web/index.html` — 查詢頁(單檔,手機優先)
- `web/data/` — 產出資料:`meta.json`(salt、更新日)+ 每人 `<sha256>.json`

## 更新資料
1. 從 OneDrive 下載最新 Excel
2. `python3 scripts/build_data.py path/to/社費支出.xlsx`
3. 部署:`./scripts/deploy.sh`(把 `web/` 推到 `gh-pages` 分支,GitHub Pages 由該分支發佈)

本機預覽:`cd web && python3 -m http.server 8899`,開 http://localhost:8899
(直接雙擊 index.html 無法讀 JSON,需經 http)

## 計算規則
餘額 = 該人所有列加總(`盤點` 年初結轉 + `收入` 儲值 + `支出` 場地費)。
**不讀分攤表樞紐的「總計」**——樞紐表忘了重新整理就會過期。

## 隱私
- 不產生總名單;檔名是 `sha256(salt + 姓名)`,網頁端用輸入的姓名算出檔名再讀取。
- 這只能防「列出所有人」,**知道姓名的人仍可查到該人餘額**(本專案刻意只用姓名驗證)。
- `ref/`、`*.xlsx` 已列入 `.gitignore`,請勿提交原始帳本。

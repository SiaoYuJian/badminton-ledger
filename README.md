# 羽球社儲值金查詢

輸入姓名 → 顯示 2026 年打球記錄、費用、儲值與目前剩餘儲值金。純靜態網頁,可放 GitHub Pages / Netlify。

## 結構
- `scripts/build_data.py` — 讀 `「2026場地費明細」`,每人輸出一個 JSON 到 `web/data/`
- `web/index.html` — 查詢頁(單檔,手機優先)
- `web/data/` — 產出資料:`meta.json`(salt、更新日)+ 每人 `<sha256>.json`

## 更新資料(帳本有變動時)
```bash
# 1. 從 OneDrive 下載最新 Excel,放到 ref/社費支出.xlsx(此資料夾不會被提交)
# 2. 重新產生資料
python3 scripts/build_data.py ref/社費支出.xlsx
# 3. 部署
./scripts/deploy.sh
```
- `build_data.py` 讀取 `「2026場地費明細」`,更新 `web/data/`,並印出:社員數、資料到哪一天、
  「與公式不符」的場地費清單(供核對帳本)、餘額最低的 5 位。
- `deploy.sh` 會自動:① 把 `web/` 的變更 commit ② 推到 `main` ③ 把 `web/` 推到 `gh-pages` 分支。
  **不必手動 `git add/commit`**。若看到 `Everything up-to-date`,代表沒有新資料可推
  (多半是忘了做步驟 1~2,或 Excel 沒有存檔)。
- 推送後 GitHub Pages 約需 30~60 秒建置;瀏覽器若看到舊資料,請強制重新整理。
- 確認是否更新:網頁標題下方的「最後更新」時間,或 `…/data/meta.json` 的 `updatedAt`。
- 只更新資料**不需要**改版號;改頁面功能才更新版號(見下方「版本」)。

### 本機預覽
```bash
cd web && python3 -m http.server 8899   # 開 http://localhost:8899
```
直接雙擊 index.html 無法讀 JSON,需經 http。8899 未被本機 docker 使用;若被占用請換其他埠。

## 計算規則
餘額 = 該人所有列加總(`盤點` 年初結轉 + `收入` 儲值 + `支出` 場地費)。
**不讀分攤表樞紐的「總計」**——樞紐表忘了重新整理就會過期。

## 隱私
- 不產生總名單;檔名是 `sha256(salt + 姓名)`,網頁端用輸入的姓名算出檔名再讀取。
- 這只能防「列出所有人」,**知道姓名的人仍可查到該人餘額**(本專案刻意只用姓名驗證)。
- `ref/`、`*.xlsx` 已列入 `.gitignore`,請勿提交原始帳本。

## 版本
目前版本 **v1.0.0**(git tag `v1.0.0`,頁面頁尾也會顯示)。之後改版請更新頁尾版號並建立新 tag。

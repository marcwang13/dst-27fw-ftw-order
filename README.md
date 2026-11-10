# DESCENTE 27FW 鞋類 經銷線上訂貨

經銷商打開網站挑鞋款、依鞋碼（mm）填數量。有接上收單網址時，用 Email 收驗證碼登入後送出；還沒接上時是傳單模式，下載 Excel 訂單本後用 LINE 或 Email 傳給業務。

交單截止日：**2026/11/13**（含當天）。

網站：https://marcwang13.github.io/dst-27fw-ftw-order/

目前 `config.js` 的 `SUBMIT_URL` 是空的，所以是傳單模式。試算表和 Apps Script 部署好之後，把網址貼進去再推上 main。

## 檔案

- `index.html`：下單頁、我的訂單、訂單彙整
- `data.js`：商品（一筆一個款色）
- `imgs.js`：縮圖
- `big/`：點開看的大圖
- `config.js`：`SUBMIT_URL`
- `apps-script/Code.gs`：貼到試算表的 Apps Script
- `tools/gen_ftw.py`：用原廠訂單本重產商品和圖片

訂單編號前綴 `FTW27-`。這站和服裝站都在 `marcwang13.github.io`，瀏覽器儲存用 `ftw-` 開頭，跟服裝站的 `dst-` 分開。

## 商品資料怎麼來的

來源是原廠檔「27Q3Q4 DST FTW GLOBAL ORDER SHEET (EMERS, DJ) 2026.11」的工作表 `27FW`。

- 一列一個款色，款號在 `확정 STYLE NO.`，色號在 `COLOR`，色碼 SKU 是款號接色號。
- 鞋碼看表頭 220 到 300。灰底的格子是這個顏色沒開的尺寸，網站上是斜線，不能填。
- 售價只用 `TW-SRP (NTD)`，當成台灣建議售價。表上每個款色都有這個數字。`FOB TTL` 是壞掉的公式，沒有採用，網站也不顯示供貨價。
- 中文品名是型號加品類（跑鞋、競速鞋、懶人鞋、休閒鞋、拖鞋）。顏色用台灣慣用的叫法。
- 上市日用 `LAUNCH DATE`，預計出貨是上市日加 14 天（表上的 GLOBAL ETD）。
- 分頁：跑鞋、休閒鞋、拖鞋。
- 圖片是表上 IMAGE 欄嵌的款式圖。同一列有兩張時，用比較大的那張。KINETIC MJ 黑色那列另外嵌了一張不同鞋型的小圖，沒有放進網站。
- 工廠名稱留在原廠表，網站只寫產地。
- 原廠表的下單日、PRE/REGULAR 欄是工廠用的，沒有寫進網站、Excel、信件或這份說明。經銷看到的截止日只有 2026/11/13。

這次是 19 個款色、10 個款號。

## 更新商品

```bash
python3 tools/gen_ftw.py /path/to/order-sheet.xlsx
```

需要 Python 3、`openpyxl`、`Pillow`。腳本會重寫 `data.js`、`imgs.js` 和 `big/`。新色號或新品名如果沒有對照，程式會停下來，到 `tools/gen_ftw.py` 的 `COLOR_ZH`、`NAME` 補上再跑。

改完推上 main，GitHub Pages 大約一分鐘後換頁。

## 後端（還沒做，留給之後）

1. 用 `marc_wang@emers.com` 開試算表，名稱：**DESCENTE 27FW 鞋類 線上訂單**。
2. 擴充功能 → Apps Script，把 `apps-script/Code.gs` 整份貼上，存檔。
3. 專案設定 → 指令碼屬性，新增 `ADMIN_KEY`。跟其他訂貨站同一組。不要寫進這個 repo、不要放進前端。
4. 部署 → 新增部署作業 → 網頁應用程式。執行身分：我；存取：任何人。授權後複製網址。
5. 網址貼進 `config.js` 的 `SUBMIT_URL`，推上 main。之後改程式用「部署 → 管理部署作業 → 新版本」，網址才不會變。
6. 用測試單走一遍：驗證碼登入、送出、改單、客人刪除、管理者登入、訂單彙整的刪除訂單、一鍵統整交單。測完把測試單刪掉。

通知信收件人（和另外四套相同）：

`marc_wang@emers.com,zero_wang@emers.com,zoe_chang@emers.com,vince_tian@emers.com,david_wu@emers.com,benson_wang@emers.com,max_lee@emers.com,kai_weng@emers.com`

另外寄給下單的客人。試算表分頁是「訂單」「明細」，每張單再依店名開一頁，尺寸橫排是鞋碼。

一鍵統整交單的折數是 5 折（`HO_DISC = 0.5`），可以在交單檔總表 B4 改，三個分頁的批發價會一起重算。這季商品都在 FTW(鞋)。

## 交接

- 範本是服裝站 `marcwang13/dst-27fw-order`。畫面、登入、Excel 訂單本、通知信、訂單彙整（每張單有「刪除訂單」）、一鍵統整交單都沿用那套。
- 拿掉的是上身／下身尺寸。鞋碼沒有前季銷售佔比。
- 管理密碼只放 Apps Script 指令碼屬性，repo 是公開的。
- 若 GitHub Pages 還沒開：Settings → Pages → Branch 選 `main`、資料夾 `/ (root)`。

# 奶茶定位所（self-media-positioning）

奶茶色系、可愛活潑的自媒體定位小測驗。使用者是不寫程式的課程講師，請用繁體中文、白話解釋，避免術語。

## 檔案
- `index.html`：測驗網站本體（單一檔案，HTML/CSS/JS 全在裡面，無建置步驟）
  - 題目 `const QUIZ`、奶茶人格 `const PERSONAS`、賽道 `const NICHES`、色票在最上方 `:root`
  - 結果頁：吉祥物互動、特質雷達圖、賽道推薦、定位句、右下角浮動分享按鈕、Email 收名單表單、課程推薦
- `admin.html`：網頁後台（儀表板、名單、序列信、課程產品、寄信設定、網站設定＝LINE 與 AdSense 廣告），用密碼登入
- `privacy.html`：隱私權政策（AdSense 審核必備）；`og.png`：分享到 LINE/FB 時的預覽圖（1200×630）
- 好友配對：分享連結帶 `?f=主口味-副口味&n=暱稱`；來源追蹤：網址加 `?src=ig` 等代號，存進名單「來源」欄
- `config.js`：後台網址 `ENDPOINT`（Google Apps Script 網頁應用程式網址），留空時寄信／課程推薦會自動隱藏
- `backend/Code.gs`：Google 試算表 + Apps Script 後台程式（名單、Resend 寄信、每小時序列信、退訂、後台 API）
  - 這個檔案要貼到 Google Apps Script 才會生效，改完要提醒使用者「重新部署成新版本」
- `backend/後台設定教學.md`：後台設定步驟

## 上線
- GitHub：https://github.com/daxingcourse/self-media-positioning ，`main` 分支推上去後 GitHub Pages 自動部署
- 網址：https://daxingcourse.github.io/self-media-positioning/ （後台 /admin.html）

## 注意
- 預覽：直接用瀏覽器打開 index.html 即可
- 不要把 Resend 金鑰或後台密碼寫進任何檔案（它們存在 Apps Script 的指令碼屬性）

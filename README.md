# Work Planner

本機工作排程與每日工作紀錄工具，支援看板管理、工作紀錄、瀏覽器本機儲存，以及本機 Excel 同步。

## 功能

- 工作排程看板
- 每日工作紀錄
- 最近工作紀錄查看與編輯
- 本機 Excel 同步
- JSON 匯入 / 匯出

## 使用方式

### 使用執行檔

到 GitHub Releases 下載最新版本的執行檔。

執行 `工作紀錄工具.exe` 後，工具會啟動本機服務並自動開啟網頁。

如果瀏覽器沒有自動開啟，可手動打開：

- http://localhost:8787/工作排程.html
- http://localhost:8787/工作紀錄.html

### 使用 Node.js

```
node local-excel-server.js
```
然後手動開啟：

- http://localhost:8787/工作排程.html
- http://localhost:8787/工作紀錄.html

### 檔案說明

- 工作排程.html：工作排程看板
- 工作紀錄.html：每日工作紀錄與最近紀錄列表
- local-excel-server.js：本機 Excel 同步服務

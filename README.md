# Work Planner

本機工作排程與每日工作紀錄工具，支援看板管理、工作紀錄、瀏覽器本機儲存，以及本機 Excel 同步。

## Files

- `工作排程.html`：工作排程看板
- `工作紀錄.html`：每日工作紀錄與最近紀錄列表
- `local-excel-server.js`：本機 Excel 同步服務

## Run

```bash
node local-excel-server.js
```

Then open:

- `http://localhost:8787/工作排程.html`
- `http://localhost:8787/工作紀錄.html`

## Data

The local Excel file is generated as `工作資料.xls`. This file may contain private work data and should not be committed.

The packaged executable `工作紀錄工具.exe` is also excluded from source control. If needed, publish executable builds through GitHub Releases.

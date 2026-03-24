# sstramkak

Smart 5G Dashboard — a single-page web application that syncs data with Google Sheets via a Google Apps Script Web App.

---

## Google Sheets Sync — Google Apps Script

All data is pushed to / pulled from Google Sheets through a Google Apps Script (GAS) Web App.  
The script source lives in [`gas/Code.gs`](gas/Code.gs).

### Sheets managed

| Sheet name    | Key fields (selection)                                      |
|---------------|-------------------------------------------------------------|
| Sales         | id, name, phone, amount, agent, branch, date, …            |
| Customers     | id, name, phone, agent, branch, date, …                    |
| **TopUp**     | id, customerId, name, phone, amount, agent, branch, date, **endDate**, tariff, remark, tuStatus, lat, lng |
| Terminations  | id, customerId, name, phone, reason, agent, branch, date, … |
| OutCoverage   | id, name, phone, agent, branch, date, …                    |
| Promotions    | id, title, startDate, endDate, remark, …                   |
| **Deposits**  | id, agent, branch, cash, **creditAmount**, riel, **cashDetail**, amount, date, **submittedAt**, remark, status, approvedBy, approvedAt |
| KPI           | id, agent, branch, date, …                                 |
| Items         | id, name, price, …                                         |
| Coverage      | id, lat, lng, …                                            |
| Staff         | id, username, password, role, status, …                    |

### TopUp — `endDate` column (expire date)

The TopUp form has an **Expiry Date** field (`tu-end-date`).  
The value is saved in the record as **`endDate`** and must be stored in the **`endDate`** column of the *TopUp* sheet.

The Apps Script (`gas/Code.gs`) automatically adds the `endDate` column if it is missing when the first sync is performed — no manual sheet editing is required.

### Deposits — approval columns (`status`, `approvedBy`, `approvedAt`)

When a supervisor approves a deposit the record gains three fields: `status` (set to `"approved"`), `approvedBy`, and `approvedAt`.  
The Apps Script automatically creates these columns in the *Deposits* sheet the first time a sync containing them is performed.

**Status column is always written.** All deposit records — including legacy ones loaded from earlier versions of the sheet — are normalised in the app so that `status` defaults to `"pending"`, `approvedBy` defaults to `""`, and `approvedAt` defaults to `""`. This guarantees the three columns are always populated in Google Sheets after the next sync.

**Targeted status updates.** When a deposit is approved, the app uses the `updateRow` GAS action to update **only** the `status`, `approvedBy`, and `approvedAt` columns of the matching row, without touching any other data in the sheet.  If the targeted update fails (e.g. the sheet row cannot be found), the app automatically falls back to a full sync so the change is never silently lost.

### Deposits — `cash` column (Cash $)

The `cash` column stores the USD cash amount for each deposit record.  
If the Deposits sheet is missing this column (e.g. it was created before the cash field was introduced), the app automatically back-fills the value from the deposit total when loading records, and the column will be created in the sheet the next time *Sync Up* is run.  
**No manual sheet editing is required — run *Sync Up* once from the admin panel to add the missing `cash` column.**

---

## How to deploy / redeploy the Apps Script Web App

1. Open [Google Apps Script](https://script.google.com) and open the project linked to your Google Spreadsheet (or create a new standalone script and bind it to the sheet via the **Triggers** menu in the left sidebar).
2. Replace the contents of `Code.gs` with the file at [`gas/Code.gs`](gas/Code.gs) in this repository.
3. Click **Deploy → Manage deployments → New deployment** (or edit an existing one):
   - **Type:** Web App
   - **Execute as:** Me
   - **Who has access:** Anyone
4. Click **Deploy** and copy the generated Web App URL.
5. If the URL has changed, update `GS_URL` in `app.js` (line ~79):
   ```js
   const GS_URL = 'https://script.google.com/macros/s/<deployment-id>/exec';
   ```
6. Save and reload the web app — sync should work immediately.

> **Note:** When updating an existing deployment, select **"New version"** from the version dropdown in the deployment editor for the changes to take effect.

---

## Sync overview

The web app communicates with the GAS Web App via `POST` requests:

```json
{ "sheet": "TopUp", "action": "sync", "data": [ { "id": "...", "endDate": "2025-12-31", ... } ] }
```

Supported `action` values:

| Action      | Description                                                           |
|-------------|-----------------------------------------------------------------------|
| `sync`      | Replace all rows in the sheet with `data`                             |
| `read`      | Return all rows as a JSON array of objects                            |
| `delete`    | Delete the row whose `id` matches `data.id`                           |
| `append`    | Append a single record row                                            |
| `updateRow` | Update specific fields of the row whose `id` matches `data.id`; only the keys present in `data` are written, all other columns are left untouched |

The Apps Script automatically creates missing sheets and adds new header columns on the first sync, so it is **backward-compatible** with existing spreadsheet data.
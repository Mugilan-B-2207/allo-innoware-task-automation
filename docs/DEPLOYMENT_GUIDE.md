# Google Sheets & Apps Script Deployment Guide

> **Live Prototype Note:** The live Google Sheet prototype is already deployed and verified (`11/11 PASS` in `Test Results`). Use the steps below if you want to deploy a fresh copy into another Google Sheet from scratch.

---

## Option A: Manual Google Sheets Deployment (Fastest — 5 Minutes)

1. **Create a new Google Spreadsheet**
   - Go to [https://sheets.new](https://sheets.new) signed into your Google account.
   - Rename it to **`Allo Innoware - Recurring Task Automation`**.

2. **Open Apps Script Editor**
   - Click **Extensions → Apps Script**.
   - Delete the default empty `Code.gs` content.

3. **Add All 10 `.gs` Files + Manifest**
   - Paste [`Code.gs`](../appsscript/Code.gs) into `Code.gs`.
   - Click **`+` → Script** and create these 9 script files matching [`appsscript/`](../appsscript/):
     - `Config` ← [`Config.gs`](../appsscript/Config.gs)
     - `Utils` ← [`Utils.gs`](../appsscript/Utils.gs)
     - `Setup` ← [`Setup.gs`](../appsscript/Setup.gs)
     - `TaskEngine` ← [`TaskEngine.gs`](../appsscript/TaskEngine.gs)
     - `CompletionService` ← [`CompletionService.gs`](../appsscript/CompletionService.gs)
     - `DashboardService` ← [`DashboardService.gs`](../appsscript/DashboardService.gs)
     - `AlertService` ← [`AlertService.gs`](../appsscript/AlertService.gs)
     - `EscalationService` ← [`EscalationService.gs`](../appsscript/EscalationService.gs)
     - `TestRunner` ← [`TestRunner.gs`](../appsscript/TestRunner.gs)
   - Click **Project Settings (gear icon)** → check **Show "appsscript.json" manifest file in editor** → replace its content with [`appsscript.json`](../appsscript/appsscript.json).
   - Click **Save (Ctrl+S)**.

4. **Initialize the System & Authorize Permissions**
   - Reload the Google Spreadsheet tab. Wait ~3 seconds for the **`Allo Operations`** custom menu to appear.
   - Click **`Allo Operations → Initialize System`**.
   - When prompted, click **Continue → Advanced → Go to ... (unsafe) → Allow** to grant Spreadsheet & MailApp permissions.
   - Verify all 6 tabs are created (`Employee Dashboard`, `Manager Dashboard`, `Task Master`, `Cycle Log`, `Alert Log`, `Configuration`) with dummy tasks `TASK-001` through `TASK-004`.

5. **Configure Real Emails (Optional for Live Email Demo)**
   - Open **`Configuration`** tab and update:
     - `Manager Email` (`B7`)
     - `Leadership Email` (`B8`)
     - Keep `Test Mode Enabled = TRUE` while running tests, or set `Test Mode Enabled = FALSE` and `Email Notifications Enabled = TRUE` to send real Gmail emails.
   - Open **`Task Master`** tab and update placeholder employee emails (`E5:E8`) and manager/leadership emails (`O5:P8`).

6. **Run the Automated Test Suite**
   - Click **`Allo Operations → Run Test Cases`**.
   - Inspect the newly created **`Test Results`** tab showing **11/11 PASS**.
   - *(Optional)* Click **`Allo Operations → Reset Demo Data`** afterwards if you want a clean 4-task queue before recording your demo.

7. **Install Automation Triggers**
   - Click **`Allo Operations → Install Automation Triggers`**.
   - Verify in Apps Script Editor → **Triggers (clock icon)** that 4 non-duplicate triggers are installed:
     - `onSheetEditTrigger` (On edit)
     - `scheduledMorningCycleJob` (Daily 8–9 AM)
     - `scheduledEmployeeReminderJob` (Daily 9–10 AM)
     - `scheduledQuarterHourMonitorJob` (Every 15 minutes)

8. **Verify Live Workflow**
   - Follow [`docs/ASSIGNMENT_DEMO_SCRIPT.md`](./ASSIGNMENT_DEMO_SCRIPT.md).

---

## Option B: `clasp` CLI Deployment (If Using Apps Script CLI)

```powershell
# 1. Install clasp globally and log in
npm install -g @google/clasp
clasp login

# 2. Inside d:\Allo\appsscript, create a new Google Sheet + bound Apps Script project
cd d:\Allo\appsscript
clasp create --type sheets --title "Allo Innoware - Recurring Task Automation" --rootDir .

# 3. Push all .gs files and appsscript.json
clasp push

# 4. Open the created Google Sheet in your browser
clasp open --addon
```
*(Or copy `appsscript/.clasp.json.example` to `appsscript/.clasp.json`, paste the Script ID of an existing Google Sheet's Apps Script project, and run `clasp push`.)*

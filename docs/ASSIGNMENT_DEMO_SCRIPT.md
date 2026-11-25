# Live Assignment Demonstration Script (For HR / Hiring Manager Review)

Use this 11-step walkthrough to demonstrate the entire system end-to-end inside Google Sheets without needing real employees or an external ticket portal.

### Pre-Demo Prep
1. Run **`Allo Operations → Reset Demo Data`** so `Employee Dashboard` starts cleanly with the 4 active tasks (`TASK-001` through `TASK-004`).
2. Open **`Configuration`** and confirm `Test Mode Enabled = TRUE` (or set `FALSE` + `Email Notifications Enabled = TRUE` with your own email address to show live Gmail messages arriving).

---

### Live Walkthrough (Steps 1–11)

1. **Employee A's Daily Task Appears:**
   - Open **`Employee Dashboard`**. Point out `TASK-001-YYYYMMDD-001` (`Daily Operations Update`, `Employee A`, `Pending`, `Checkbox Confirmation + Optional WhatsApp Link`).
2. **Employee A Confirms WhatsApp Update (No Ticket ID):**
   - Check `Completion Confirmation` (`K6`), optionally paste a sample WhatsApp message link in `Proof Link` (`L6`), and change `Submission Action` (`P6`) to **`Submit Completion`**.
   - *Point out:* No Ticket ID is required; `TASK-001` closes in `Cycle Log` and the next working-day cycle immediately appears on `Employee Dashboard`.
3. **Employee B Enters Portal Ticket ID (`TASK-002`):**
   - On `TASK-002` (`Portal Task A`), first demonstrate validation by typing `"done"` in `Ticket ID` and selecting `Submit Completion` -> blocked with clear `VALIDATION ERROR` in Notes!
   - Now enter `TKT-1042` in `Ticket ID` and select **`Submit Completion`**.
4. **Current Cycle Closes & 5. Next Cycle Generates Automatically:**
   - Open **`Cycle Log`** to show `TASK-002` Cycle `1` is permanently preserved as `Completed` with `TKT-1042`, `Completed By`, `Completed At`, and `Next Due Date = Today + 7 days`.
   - Return to **`Employee Dashboard`** to show Cycle `2` (`TASK-002-...-002`) is already queued with `Status = Pending`.
6. **Employee C's Task Becomes Overdue (`TASK-003`):**
   - In **`Cycle Log`**, temporarily edit `TASK-003`'s `Due Date` to yesterday (`YYYY-MM-DD`) to simulate a missed deadline, then click **`Allo Operations → Run Pending Check`**.
   - Show `TASK-003` turns red **`Overdue`** (`Days Overdue = 1`) on both `Employee Dashboard` and `Manager Dashboard`.
7. **Manager Receives Same-Day Pending Alert:**
   - Click **`Allo Operations → Run Alert Check`**.
   - Open **`Alert Log`** (and Gmail inbox if live emails are enabled) to show **`Manager First Alert`** recorded for `TASK-003` with `Manager Alert Sent At` stamped in `Cycle Log`.
8. **Manager Acknowledges the Issue:**
   - Open **`Manager Dashboard`**. Check **`Acknowledge`** (`Col J`) on `TASK-003`, add note `"Followed up with Employee C"`, and enter tomorrow as `Expected Resolution Date`.
   - *Point out:* Status updates to amber **`Acknowledged`** (`Acknowledged By` & `At` stamped), **while the task remains open until Employee C actually completes it**.
9. **Leadership Receives Escalation After 24h SLA:**
   - In **`Cycle Log`**, backdate `TASK-003`'s `Manager Alert Sent At` (`Col P`) by 25 hours, then click **`Allo Operations → Run Alert Check`**.
   - Open **`Alert Log`** and **`Manager Dashboard`** to show **`Leadership Escalation`** sent and status turned dark red **`Escalated`**.
10. **Employee Completion Resolves the Escalation:**
    - Back on **`Employee Dashboard`**, enter `TKT-2099` on `TASK-003` and select **`Submit Completion`**.
    - Show `Cycle Log` marks it `Completed`, `Alert Log` records a one-time **`Completion Notification`**, and `Manager Dashboard` clears the escalation risk.
11. **Idempotency (No Duplicate Alerts or Cycles):**
    - Click **`Allo Operations → Generate Due Cycles`** and **`Allo Operations → Run Alert Check`** twice in a row.
    - Show `Cycle Log` and `Alert Log` row counts remain completely unchanged (**0 duplicates**).

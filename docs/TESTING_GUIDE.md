# Testing & Verification Guide

## 1. Execution Scope & Verification Matrix

| Verification Layer | Status | What Was Executed |
| :--- | :--- | :--- |
| **Live Google Sheet & Bound Apps Script** | **Executed & Passed (11/11 PASS)** | Ran **`Allo Operations → Run Test Cases`** inside the live Google Sheet (`Asia/Kolkata`) with Editor access enabled. Verified all 7 tabs, formulas, KPI totals (`TOTAL OVERDUE: 1` / `Tasks Overdue (Total): 1`), and `Test Results`. |
| **Local Node Mock Suite (`node tests/run_tests.js`)** | **Executed & Passed (11/11 + 3 checks)** | Ran the exact 10 `.gs` files in V8 `vm` against mocked `SpreadsheetApp`, `ScriptApp`, `MailApp`, `LockService`, `Utilities`, and `Session`. Verified cycle math, validations, SLA transitions, deduplication, trigger idempotency, and `MailApp` dispatch. |
| **Real Gmail Delivery** | **Configurable for Evaluator Testing** | Verified `MailApp.sendEmail` payload generation (`4` HTML alerts dispatched when `Test Mode Enabled = FALSE` and `Email Notifications Enabled = TRUE`). To send real emails during review, swap the `@example.com` placeholders in `Configuration!B7:B8` and `Task Master!E5:E8`, set `Test Mode Enabled = FALSE`, and click `Allo Operations → Run Alert Check`. |

---

## 2. Automated Test Cases (`TestRunner.gs`)

| Test ID | Scenario | Verified Assertion |
| :--- | :--- | :--- |
| **TEST-01** | Recurring Cycle (`TASK-002` 7-day) | Closes `TASK-002-20261123-001`, logs completion timestamp, spawns `TASK-002-20261130-002` due `2026-11-30` (`Completion Date + 7 days`). |
| **TEST-02** | Single-Point Completion (`TASK-003`) | Submits `TKT-15099` solely from `Employee Dashboard`; closes cycle in `Cycle Log`; refreshes `Employee Dashboard` with next cycle due `2026-12-08`. |
| **TEST-03** | Daily WhatsApp Task (`TASK-001`) | Confirms via checkbox + optional WhatsApp link without Ticket ID; spawns next working-day cycle (`2026-11-24`). |
| **TEST-04** | Same-Day Manager Alert (`TASK-004`) | Verifies `Email Notifications Enabled = FALSE` suppresses dispatch; when enabled, `17:05` sets `Overdue` and `17:20` logs `Manager First Alert` in `Alert Log`. |
| **TEST-05** | Manager Acknowledgement | Manager acknowledges on `Manager Dashboard`; `Acknowledged By` & `At` recorded (`delay 40m`); task transitions to `Acknowledged` (never `Completed`). |
| **TEST-06** | Leadership Escalation | At `>24h` after Manager alert (`2026-11-24 17:25`), escalates `TASK-004` to Leadership with reason (`Days Overdue: 2` on `2026-11-25`); verifies `0` completed tasks are escalated. |
| **TEST-07** | Full Escalation Timing Progression | Tests 7 sequential windows on `TASK-001`: `09:05` Reminder → `16:00` Pending → `17:05` Overdue → `17:20` Manager Alert → `10:05` Next-Day Follow-up → `14:05` Early Unacked Leadership Escalation → `14:30` Post-Escalation Completion Notice. |
| **TEST-08** | Duplicate Prevention | Repeats `generateDueCycles()`, `runAlertAndEscalationCheck()`, and `installProjectTriggers()` twice -> `0` duplicate cycles, `0` duplicate alerts, `0` duplicate triggers. |
| **EDGE-01** | Invalid Proof Guard | Rejects generic `"done"` for Portal task and unchecked checkbox for Daily task. |
| **EDGE-02** | Unauthorized & Duplicate Submit | Blocks `Employee C` submitting `Employee B`'s task (`TASK-002`) and blocks re-submitting a closed Cycle ID. |
| **EDGE-03** | Calendar, Inactive & Missing Email | Verifies Friday→Tuesday skip across weekend + Monday holiday, `Jan 31 + 30d = Mar 02`, `Active = FALSE` filtering, and graceful `FAILED` log on empty email. |

---

## 3. Running Tests Locally vs Inside Google Sheets

- **Locally (Node.js):**
  ```bash
  node tests/run_tests.js
  ```
- **Inside Google Sheets:**
  Click **`Allo Operations → Run Test Cases`** and inspect the **`Test Results`** tab.

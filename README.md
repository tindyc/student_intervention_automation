# 📊 Student Intervention Tracker & Sync Automator

A zero-config, fully automated Google Sheets tracking system that aggregates student data from Zoho CRM, Cypher LMS, and GitHub into a unified dashboard, powered by an automated Risk Engine.

## 💼 The Business Case: Why This Exists

**The Challenge:** Educational support teams waste countless hours manually downloading CSVs, cross-referencing Zoho CRM with Cypher LMS, and checking individual GitHub profiles just to figure out which students are falling behind. This fragmented process leads to delayed interventions and administrative burnout.

**The Solution:** This Automator eliminates platform-hopping. By pulling all critical data streams into a single, automated dashboard, it transforms how student success is managed:
* ⏳ **Eliminates Manual Labor:** Replaces hours of manual data entry and spreadsheet merging with a single click. 
* 🎯 **Enables Proactive Interventions:** The built-in **Risk Engine** automatically evaluates project deadlines, resubmissions, and LMS inactivity to instantly color-code at-risk students.
* 🔄 **Creates a Single Source of Truth:** Facilitators no longer need to log into three different platforms to get a 360-degree view of a student's progress.

---

## ✨ Key Features
* **📧 Automated CRM Import:** Automatically searches your Gmail for the latest scheduled Zoho CRM reports and seamlessly updates the tracker safely without duplicating records.
* **🎓 Cypher LMS Integration:** Matches raw Cypher LMS exports to students (via Email or Learner ID) to track progress, current lessons, and days since last activity.
* **🐙 GitHub Activity Tracker:** Pings the GitHub API to find student usernames based on email and tracks their most recent commit activity.
* **🚦 Automated Risk Engine:** Evaluates project deadlines, resubmission statuses, and LMS inactivity to automatically identify and flag at-risk students.
* **🛡️ Data Safety & Auditing:** Features crash-locks, duplicate file prevention, identity-conflict resolution, and a dedicated `Sync Audit Log` to track every addition, update, or removal.

---

## 🚀 Quick Setup Guide

💡 **Zero-Config Spreadsheets:** You do not need to manually create any sheets, tabs, or columns! The script will automatically build the entire spreadsheet architecture the first time you run it. The *only* manual data entry required is pasting your Cypher LMS export into the `Cypher Raw Data` tab.

**Step 1: Configure Your CRM Report & Run Your First Sync**
The script will automatically prompt you to configure it the very first time you run it.
1. In the Google Sheet, go to the top menu: **📊 Student Tracker → ▶️ Run Full Sync**.
2. A popup prompt will appear asking for your CRM report name. Enter the filename of your Zoho CRM scheduled report (e.g., `L5_student_data_Tindy`) and click OK.
3. Ensure your Zoho CRM is actively sending this report to your connected Gmail account at least once every 7 days.

*(Note: If you ever need to change this CRM filename later on, you can update it via **📊 Student Tracker → ⚙️ Change CRM Report Name**).*

**Step 2: Run a System Check (Optional)**
If you ever want to verify your setup is still working without running a full sync:
1. Go to **📊 Student Tracker → 🧪 Run Setup Check**.
2. This will verify that your email query works, that emails are arriving on time, and that the required sheets are properly formatted.

---

## 📖 Daily User Guide

### Option A: The "One-Click" Full Sync (Recommended)
This handles everything in one go: CRM updates, LMS tracking, GitHub fetching, and Risk evaluation.
1. Open the auto-generated **Cypher Raw Data** tab.
2. Paste your latest Cypher LMS CSV export into this sheet (starting at cell A1).
3. Go to the menu: **📊 Student Tracker → ▶️ Run Full Sync**.
4. Wait for the sync to complete. Check the **Sync Dashboard** tab for a summary of changes.

### Option B: Modular Syncing
If you only need to update specific platforms to prepare for an intervention meeting, you can run them individually:
* **🔄 Sync CRM Only:** Fetches the latest CRM email and updates core student details.
* **🎓 Sync Cypher LMS Activity:** Only processes the latest data pasted in the `Cypher Raw Data` sheet.
* **🐙 Update GitHub Activity:** Re-checks GitHub for recent activity. *(Note: GitHub API rate limits apply, so use this sparingly).*

---

## 📂 Sheet Architecture Reference

*(Note: All tabs below are generated automatically by the script)*

| Tab Name | Purpose |
| :--- | :--- |
| **Student Tracker** | The main interactive tracker. Includes automated columns and manual columns (e.g., Notes). |
| **Sync Dashboard** | A quick overview of the last sync times, records updated, and total at-risk students. |
| **Sync Audit Log** | A historical log showing exactly when a student was ADDED, UPDATED, or REMOVED. |
| **CRM Raw Data** | Temporary storage. The script uses this to read the latest CRM email. *Do not edit manually.* |
| **Cypher Raw Data** | Paste your Cypher CSV export here before running an LMS or Full Sync. This is the only manual data requirement. |
| **LMS Unmatched** | Lists students found in the LMS export who could not be matched to anyone in the CRM. |

---

## 🛠️ Troubleshooting & Utilities

If the script encounters an error, check the **📊 Student Tracker** menu for these built-in fixes:

* **🔓 Clear Crash Lock:** If your internet drops or a Google limit is hit mid-sync, the script "locks" itself to prevent data corruption. Once you ensure the tracker data looks okay, click this to unlock the system.
* **🔓 Allow Duplicate CRM File:** The script normally prevents you from syncing the exact same CRM file twice to save processing time. Click this to force it to re-read the last file.
* **🧹 Reset GitHub API Cache:** If GitHub data seems stuck or outdated, click this to clear the script's memory and force a fresh lookup on the next sync.
* **⚠️ Mass Deletion Warning:** If the CRM sync attempts to delete more than 15% of your cohort at once, it will pause and ask for confirmation. This protects you in case Zoho accidentally sends a blank or incomplete report.
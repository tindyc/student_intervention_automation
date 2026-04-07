# 📊 Student Intervention Tracker & Sync Automator

A zero-config, fully automated Google Sheets tracking system that aggregates student data from Zoho CRM, Cypher LMS, and GitHub into a unified dashboard, powered by an automated Risk Engine.

## 💼 The Business Case: Why This Exists

**The Challenge:** Educational support teams waste countless hours manually downloading CSVs, cross-referencing Zoho CRM with Cypher LMS, and checking individual GitHub profiles just to figure out which students are falling behind. This fragmented process leads to delayed interventions and administrative burnout.

**The Solution:** This Automator eliminates platform-hopping. By pulling all critical data streams into a single, automated dashboard, it transforms how student success is managed:
* ⏳ **Eliminates Manual Labor:** Replaces hours of manual data entry and spreadsheet merging with a single click. 
* 🎯 **Enables Proactive Interventions:** The built-in **Risk Engine** automatically evaluates project deadlines, resubmissions, and LMS inactivity to instantly color-code at-risk students and tell you exactly *why* they are at risk.
* 🔄 **Creates a Single Source of Truth:** Facilitators no longer need to log into three different platforms to get a 360-degree view of a student's progress.

---

## ✨ Key Features

* **💬 One-Click Action Links:** The tracker automatically turns raw data into direct communication and review links. No more searching for students across platforms:
  * **Discord:** Click "Message" in the tracker to instantly open a direct message with the student.
  * **Zoho CRM:** Click a student's Full Name to jump straight to their live CRM profile.
  * **GitHub:** Click their generated GitHub Profile link to immediately view their code and recent commits.
* **🚦 Automated Risk Engine:** Evaluates project deadlines, resubmissions, and LMS inactivity. It automatically color-codes at-risk students and generates specific text flags in the "Auto Risk Reason" column (e.g., *"Deadline missed"*, *"Resub overdue"*, *"LMS Inactive 14+ days"*).
* **🧠 Dynamic Course Adaptation (L3 vs L5):** Automatically detects if your cohort is studying a Level 3 (L3) or Level 5 (L5) curriculum based on the CRM file name, automatically generating the correct tracking columns, modules, and project deadlines for that specific course.
* **📧 Automated CRM Import:** Automatically searches your Gmail for the latest scheduled Zoho CRM reports and seamlessly updates the tracker safely without duplicating records.
* **🎓 Cypher LMS Integration:** Matches raw Cypher LMS exports to students (via Email or Learner ID) to track progress, current lessons, and days since last activity.
* **🐙 GitHub Activity Tracker:** Pings the GitHub API to find student usernames based on email and tracks their most recent commit activity.

---

## 🚀 Quick Setup Guide

💡 **Zero-Config Spreadsheets:** You do not need to manually create any sheets, tabs, or columns! The script will automatically build the entire spreadsheet architecture the first time you run it. The *only* manual data entry required is pasting your Cypher LMS export into the `Cypher Raw Data` tab.

**Step 1: Set Up & Name Your Report in Zoho CRM**
Before running the script, ensure your Zoho CRM scheduled report is configured correctly:
1. Schedule your student data report in Zoho CRM to be emailed to your connected Gmail account at least once every 7 days.
2. **CRITICAL NAMING RULE:** The exported CSV filename *must* contain either **"L3"** or **"L5"** so the script knows which curriculum columns to build. (e.g., Name your report `L3_student_data_Marko` or `L5_Active_Students_Tindy`).
3. Ensure Zoho is sending the email with the standard subject line (`Zoho CRM - Report Scheduler`) and a CSV attachment.

**Step 2: Configure the Script & Run Your First Sync**
The script will automatically prompt you to link your newly created CRM report the very first time you run it.
1. In the Google Sheet, go to the top menu: **📊 Student Tracker → ▶️ Run Full Sync**.
2. A popup prompt will appear asking for your CRM report name. Enter the exact partial filename you chose in Step 1 (e.g., `L3_student_data_Marko`) and click OK. 
3. The script will find the email, detect the L3/L5 level from your filename, and build the entire spreadsheet automatically!

*(Note: If you ever need to change this CRM filename later on, you can update it via **📊 Student Tracker → ⚙️ Change CRM Report Name**).*

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

## 🛡️ Built-In Data Safety & Validation Checks

Because this script manages live student data, it is heavily fortified with automated checks to prevent data corruption, accidental deletions, and schema breaks:

* **Mass Deletion Safeguard:** If a Zoho CRM export is accidentally filtered or blank, the script detects if >15% of your cohort is about to be deleted. It will pause the sync and ask for manual confirmation before proceeding, preventing catastrophic data loss.
* **Identity Conflict Resolution:** If a student's Email matches but their CRM Record ID has changed, the script halts to prevent merging two different students together. It also scans the incoming CRM file for internal duplicate IDs or Emails before syncing.
* **Preflight System Checks:** Before touching any data, the script verifies that the CRM email isn't stale (older than 7 days) and ensures you haven't accidentally deleted critical columns (like "Email" or "Record Id") from the tracker.
* **Crash & Timeout Locks:** Google scripts timeout after 6 minutes. This script monitors its own runtime and safely exits at 4.5 minutes to prevent "partial writes". If a catastrophic crash does occur mid-write, it activates a "Crash Lock" to prevent overlapping syncs until you verify the data.

---

## 📂 Sheet Architecture Reference

*(Note: All tabs below are generated automatically by the script)*

| Tab Name | Purpose |
| :--- | :--- |
| **Student Tracker** | The main interactive tracker. Includes automated columns (with clickable Discord/GitHub/CRM links) and the `Auto Risk Reason` output. **Columns dynamically adapt to include specific L3 modules or L5 projects based on your course setup.** |
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
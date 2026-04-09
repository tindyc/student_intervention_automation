# 📊 Student Intervention Automation

A fault-tolerant, fully automated data pipeline built in Google Apps Script that consolidates student data from Zoho CRM, Cypher LMS, and GitHub into a unified tracking system with a real-time risk evaluation engine.

---

## 🚀 Quick User Guide (Start Here)

### First-Time Setup
1. Run: **📊 Student Tracker → ▶️ Run Full Sync**
2. Enter your CRM report filename when prompted (e.g. `L5_student_data`)
3. The system will automatically:
   - Detect course type (L3 / L5)
   - Build all sheets and columns
   - Sync your data

---

### Daily Usage (Recommended Flow)

1. Open **Cypher Raw Data**
2. Paste latest LMS export (CSV) into cell A1
3. Run: **📊 Student Tracker → ▶️ Run Full Sync**
4. View results in:
   - **Student Tracker** (main view)
   - **Sync Dashboard** (summary)

---

### Optional (Modular Sync)

- **🔄 Sync CRM Only** → Updates student records  
- **🎓 Sync LMS Only** → Updates activity + progress  
- **🐙 GitHub Sync Only** → Updates coding activity  

---

## 💼 Business Problem

Facilitation teams often rely on fragmented systems:
- Zoho CRM (student records)
- Cypher LMS (learning activity)
- GitHub (technical progress)

This leads to:
- Manual CSV downloads and merging  
- Delayed identification of at-risk students  
- High administrative overhead  

---

## 🚀 Solution

This system replaces a fragmented, manual workflow with a single automated process.

Instead of:
- Logging into multiple platforms  
- Downloading and merging CSV files  
- Manually checking each student’s activity  

Everything is now handled in just  **one click** !!!

With a single sync:
1. CRM data is pulled automatically from Gmail  
2. LMS activity is merged into each student record  
3. GitHub activity is updated  
4. Risk levels are calculated instantly  

The result:
- A fully updated tracker  
- Clear visibility of every student  
- Automatic identification of at-risk students  

All in one place, with no manual cross-checking required.

---

## 🧠 Architecture Overview

This project is structured as a multi-step data pipeline:

- **Ingestion:** Gmail → CRM CSV  
- **Transformation:** Clean and standardise data  
- **Sync Engine:** Add, update, and remove students  
- **Enrichment:** LMS + GitHub data  
- **Analytics:** Risk Engine  
- **Output:** Tracker + Dashboard + Logs  

### Design Principles
- Safe to run multiple times  
- Only update what has changed  
- Stop on errors instead of guessing  
- Keep each stage independent  

---

## 🔬 How the Data Sync Works

This system safely combines data from multiple sources while making sure nothing is duplicated, overwritten incorrectly, or lost.

---

### 🧱 How Students Are Matched

Each student is identified using:

- **Record Id** (main identifier from CRM)  
- **Email** (used only if ID is missing)  

How it works:
- Match using Record Id first  
- If missing, fallback to Email  
- If both exist but don’t match, stop the sync  

This prevents mixing up different students.

---

### 🔁 Safe Syncing

You can run the sync as many times as needed.

Each run:
- Updates existing students  
- Adds new students  
- Removes students no longer in CRM  

The result is always a clean and up-to-date tracker.

---

### 🔍 How Changes Are Detected

Before updating, each row is compared with existing data.

- No change → skip  
- Change → update  

This keeps performance high and avoids unnecessary writes.

---

### 🧬 Handling Different Data Formats

CRM and LMS exports are not always consistent.

The system:
- Matches similar column names automatically  
- Cleans values before comparing  
- Ignores unknown fields safely  

If a critical field is missing, the sync stops.

---

### 🔗 Combining Multiple Data Sources

Data is merged into one row per student.

Matching order:
1. Record Id  
2. Email  

Handled cases:
- Missing data  
- Duplicate records  
- Mismatched entries  

If unsure, the system stops instead of guessing.

---

### 🧾 Safe Write Process

All updates are prepared before writing.

Steps:
1. Build updated data in memory  
2. Group changes into batches  
3. Write in controlled operations  

This ensures:
- No partial updates  
- No broken sheets  
- No inconsistent data  

---

### 🚨 Data Safety Checks

#### Prevents Large Accidental Deletions

If too many students are about to be removed:
- The sync stops  
- Manual confirmation is required  

---

#### Prevents Student Mix-Ups

If:
- Same Email  
- Different Record Id  

The sync stops immediately.

---

#### Prevents Outdated Data

If CRM data is older than 7 days:
- Sync is blocked  

---

## 🧠 Risk Engine

Automatically identifies students who may need support.

Checks:
- Missed deadlines  
- Resubmission windows  
- LMS inactivity  

Examples:
- Deadline passed → "Deadline missed"  
- Resub overdue → "Resub overdue"  
- No activity → "LMS inactive 7/14/30+ days"  

Outputs:
- Colour-coded rows  
- Clear reasons for risk  

---

## ⚡ Performance Optimisation

- Only changed rows are updated  
- Data processed in memory  
- Batch writes reduce API calls  
- Caching reduces repeated work  
- Script stops early to avoid timeout  

---

## ✨ Key Features

- One-click full sync  
- Dynamic schema (L3 vs L5 courses)  
- Clickable action links:
  - CRM profile  
  - Discord DM  
  - GitHub profile  
- Audit logging (add/update/remove history)  
- LMS unmatched tracking  

---

## 🛠️ Built-In Tools & Controls (Menu Features)

The tracker includes a custom Google Sheets menu to manage the system without touching code.

### 📊 Sync Controls

- **▶️ Run Full Sync**  
  Runs the full pipeline (CRM + LMS + GitHub + Risk Engine)

- **🔄 Sync CRM Only**  
  Updates student records from CRM  

- **🎓 Sync LMS Only**  
  Updates learning activity and progress  

- **🐙 GitHub Sync Only**  
  Refreshes GitHub activity data  

---

### ⚙️ Configuration & Setup

- **🧪 Run Setup Check**  
  Verifies system configuration before syncing  

- **⚙️ Change CRM Report Name**  
  Updates the CRM file identifier used for Gmail search  

---

### 🛡️ Recovery & Safety Tools

- **🔓 Clear Crash Lock**  
  Unlocks the system after a failed or interrupted sync  
  (Prevents running on partially written data)

- **🔓 Allow Duplicate CRM File**  
  Allows reprocessing the same CRM file if needed  
  (Normally blocked to prevent unnecessary re-syncs)

- **🧹 Reset GitHub API Cache**  
  Clears cached GitHub data to force fresh API calls  

---

### 💡 Why This Matters

These controls allow non-technical users to:
- Recover safely from errors  
- Avoid duplicate processing  
- Maintain system integrity without editing code  

This turns the script into a usable internal tool rather than just a backend process.

---

## 📊 Impact

| Before | After |
|------|--------|
| Manual data aggregation | Fully automated pipeline |
| Multiple systems | Single dashboard |
| Reactive support | Proactive intervention |
| Hours of admin work | One-click sync |

---

## 🧩 Technical Highlights

- Multi-source data integration  
- Safe re-runnable sync logic  
- Row-level change detection  
- Controlled write operations  
- Schema validation  
- Rule-based risk evaluation  
- Performance optimisation under runtime limits  

---

## ⚖️ Tradeoffs

- GitHub API rate limits (handled with caching)  
- Apps Script runtime limit (~6 minutes)  
- LMS matching depends on consistent identifiers  

---

## 🛠️ Tech Stack

- Google Apps Script (JavaScript)  
- Google Sheets  
- Gmail API  
- GitHub API  

---

## 📌 Summary

A production-style data pipeline built within a constrained environment, focused on data integrity, reliability, and automation.

Designed to replace manual workflows with a safe, scalable system for real-time student tracking and intervention.
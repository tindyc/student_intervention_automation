# 📊 Student Intervention Tracker & Sync Automator

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

Educational support teams often rely on fragmented systems:
- Zoho CRM (student records)
- Cypher LMS (learning activity)
- GitHub (technical progress)

This leads to:
- Manual CSV downloads and merging  
- Delayed identification of at-risk students  
- High administrative overhead  

---

## 🚀 Solution

This system automates the entire workflow into a single pipeline:

1. Fetch CRM data from Gmail  
2. Sync and reconcile student records  
3. Enrich with LMS + GitHub data  
4. Apply Risk Engine  
5. Output dashboard + audit logs  

---

## 🧠 Architecture Overview

This project is designed as an **idempotent, multi-stage data pipeline**:

- **Ingestion:** Gmail → CRM CSV  
- **Transformation:** schema + alias resolution  
- **Sync Engine:** deterministic add/update/delete  
- **Enrichment:** LMS + GitHub  
- **Analytics:** Risk Engine  
- **Output:** dashboard + logs  

### Design Principles
- Idempotent execution (safe re-runs)  
- Deterministic diffing (hash-based)  
- Fail-fast validation  
- Modular processing stages  

---

## 🛡️ Data Integrity & Reliability

Built in Safeguards:

- **Deterministic Row Hashing**
  - Only writes rows that actually changed  

- **Transaction-Safe Writes**
  - Prevents partial updates  

- **Crash Recovery Lock**
  - Blocks unsafe re-runs after failure  

- **Mass Deletion Protection**
  - Detects abnormal removals and stops execution  

- **Schema Drift Detection**
  - Flags missing or renamed columns  

- **Identity Conflict Protection**
  - Prevents merging incorrect student records  

- **Manual Column Protection**
  - Preserves user-entered notes  

- **Stale Data Rejection**
  - Blocks CRM files older than 7 days  

---

## ⚡ Performance Optimisations

- Batch write operations  
- Row-level diffing (hashing)  
- LRU caching (dates + API calls)  
- Runtime guard (4.5 min cutoff)  
- GitHub API rate-limit handling  

---

## 🚦 Risk Engine

Automatically evaluates student risk based on:

- Missed deadlines  
- Resubmission windows  
- LMS inactivity  

### Outputs
- Colour-coded risk levels  
- Auto-generated reasons:
  - "Deadline missed"  
  - "Resub overdue"  
  - "LMS inactive 7/14/30+ days "  

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

## 📊 Impact

| Before | After |
|------|--------|
| Manual data aggregation | Fully automated pipeline |
| Multiple systems | Single dashboard |
| Reactive support | Proactive intervention |
| Hours of admin work | One-click sync |

---

## 🧩 Technical Highlights

- Multi-source data integration (CRM, LMS, REST API)  
- Idempotent sync architecture  
- Deterministic diffing system  
- Transaction-safe write engine  
- Schema validation + enforcement  
- Rule-based analytics engine  
- Performance optimisation under runtime constraints  

---

## ⚖️ Tradeoffs

- GitHub API rate limits (handled via caching)  
- Apps Script runtime limit (~6 minutes)  
- LMS matching depends on consistent identifiers  

---

## 🛠️ Tech Stack

- Google Apps Script (JavaScript runtime)  
- Google Sheets (data store + UI)  
- Gmail API (data ingestion)  
- GitHub REST API (activity tracking)  

---

## 📌 Summary

A production-style data pipeline built within a constrained environment, with strong emphasis on data integrity, fault tolerance, and automation.

Designed to replace manual workflows with a reliable, scalable system for real-time student intervention.
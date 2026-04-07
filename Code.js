// ================= CONFIG & GLOBALS =================
const GITHUB_TOKEN = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
const DEBUG_LOGS = false;
const VERBOSE_LOGS = false; // Set to true for deep JSON/Batch stringify logs

const START_TIME = Date.now();
const MAX_RUNTIME = 4.5 * 60 * 1000; // 4.5 Minutes max runtime to safely exit before hard Google limit

function shouldStop() { 
  if (Date.now() - START_TIME > MAX_RUNTIME) {
    throw new Error("❌ TIME LIMIT HIT - Execution aborted to prevent partial writes and maintain data integrity.");
  }
}

// email search query generator
function buildCrmSearchQuery(fileIdentifier, strict = true) {
  const baseQuery = `subject:"Zoho CRM - Report Scheduler" has:attachment`;
  // Force strict filename matching to prevent pulling backups or similar reports
  return strict 
    ? `${baseQuery} newer_than:7d filename:${fileIdentifier}` 
    : `${baseQuery} filename:${fileIdentifier}`;
}

const MAX_FILE_AGE_MS = 7 * 24 * 60 * 60 * 1000; // 7 Days hard limit for file freshness
let globalAmbiguousDateCount = 0;
let globalUnknownHeaders = 0;
let globalAliasCollisions = 0;
let HEADER_CACHE = {}; // Sheet-level session cache

const CONFIG = {
  EXPECTED_DATE_FORMAT: "UK", // "UK" (DD/MM/YYYY) or "US" (MM/DD/YYYY)
  SHEETS: {
    TRACKER: "Student Tracker",
    CRM_RAW: "CRM Raw Data",
    LMS_RAW: "Cypher Raw Data",
    DASHBOARD: "Sync Dashboard",
    AUDIT_LOG: "Sync Audit Log",
    UNMATCHED: "LMS Unmatched"
  },
  CORE_FIELDS: {
    ID: "Record Id", 
    EMAIL: "Email",
    NAME: "Full Name"
  },
  PROJECTS: ["P1", "P2", "P3", "P4", "P5"], // L5
  L3_MODULES: ["PI", "OOP", "ST", "PM", "JS", "MC", "RT"], // L3 
  RISK: {
    RESUB_GRACE_DAYS: 7,
    SUBMITTED_STATUSES: ["submitted on time", "submitted late"]
  }
};

// ================= FEATURE FLAGS & DYNAMIC CONFIG =================

const FEATURES = {
  PROJECT_TRACKING: true,
  SPECIALISATION: true,
  L3_MODULE_TRACKING: false,
  LMS_TRACKING: true,
  GITHUB_TRACKING: true,
  RISK_ENGINE: true,
  ALLOW_OVERWRITE_NON_MANUAL: true 
};

const COURSE_PRESETS = {
  L5: {
    PROJECT_TRACKING: true,   
    SPECIALISATION: true,     
    L3_MODULE_TRACKING: false,
    LMS_TRACKING: true,
    GITHUB_TRACKING: true,
    RISK_ENGINE: true
  },
  L3: {
    PROJECT_TRACKING: false,  
    SPECIALISATION: false,    
    L3_MODULE_TRACKING: true, 
    LMS_TRACKING: true,       
    GITHUB_TRACKING: true,    
    RISK_ENGINE: true         
  }
};

function normHeader(h) { return h ? h.toString().toLowerCase().trim() : ""; }
function normEmail(e) { return e ? e.toString().toLowerCase().trim() : ""; }
function normId(id) { 
  if (!id) return "";
  return String(id).trim().replace(/^zcrm_/, ''); 
}
const NEVER = "Never";
function isNever(val) {
  if (val === null || val === undefined) return false;
  return String(val).trim().toLowerCase() === "never";
}

// Performance caching for commonly used headers
const H = {
  ID: normHeader(CONFIG.CORE_FIELDS.ID),
  EMAIL: normHeader(CONFIG.CORE_FIELDS.EMAIL),
  NAME: normHeader(CONFIG.CORE_FIELDS.NAME),
  COURSE_CODE: normHeader("Course Of Interest Code"),
  PATHWAY: normHeader("Pathway"),
  PROJ_SUB: normHeader("Projects submitted"),
  AUTO_RISK_REASON: normHeader("Auto Risk Reason"),
  RISK_NOTES: normHeader("Risk Notes"),
  LMS_ACT: normHeader("LMS Last Activity"),
  LMS_DAYS: normHeader("Days Since LMS Activity"),
  LMS_PROG: normHeader("Progress"),
  LMS_LES: normHeader("Current Lesson"),
  GH_UNAME: normHeader("GitHub Username"),
  GH_ACT: normHeader("Last GitHub Activity"),
  GH_PROF: normHeader("GitHub Profile"),
  SUB_STATUS: normHeader("submission_status"),
  DISC_ID: normHeader("discord user id"),
  P5_DEADLINE: normHeader("P5_submission_deadline"),
  SPEC_DEADLINE: normHeader("Specialization Selection Deadline")
};

const STRICT_HEADER_MAP = {
  "record id": H.ID,
  "email": H.EMAIL,
  "full name": H.NAME,
  "course of interest code": H.COURSE_CODE
};

// Globals to be initialized dynamically
let PROJECT_FIELDS = [];
let CRM_FIELDS = [];
let COLUMN_OWNERS = {};
let ALL_KNOWN_HEADERS = new Set();
const _loggedUnknowns = new Set();

function loadCourseType() {
  const props = PropertiesService.getDocumentProperties();
  let fileIdentifier = props.getProperty("CRM_FILE_IDENTIFIER");
  let courseType = props.getProperty("COURSE_TYPE");
  
  if (!fileIdentifier) {
    try {
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt(
        "⚙️ Tracker Setup",
        "Enter part of your CRM report filename.\n\nExample:\nL5_student_data\n\nThis is used to automatically find your report in Gmail.",
        ui.ButtonSet.OK
      );
      
      fileIdentifier = response.getResponseText().trim();
      if (!fileIdentifier) {
        ui.alert("⚠️ No name provided. Defaulting to search for 'student_data'. You can change this in the menu later.");
        fileIdentifier = "student_data";
      }
    } catch (e) {
      throw new Error("⚠️ Setup incomplete: CRM report name not configured. Please run 'Run Setup Check' from the Student Tracker menu to initialize.");
    }
    
    if (/L3/i.test(fileIdentifier)) {
      courseType = "L3";
    } else {
      courseType = "L5"; 
    }
    
    props.setProperty("CRM_FILE_IDENTIFIER", fileIdentifier);
    props.setProperty("COURSE_TYPE", courseType);
  }
  
  applyCoursePreset(courseType || "L5");
  buildGlobals();
}

function applyCoursePreset(courseType) {
  const preset = COURSE_PRESETS[courseType];
  if (!preset) {
    console.warn(`⚠️ Unknown course preset: ${courseType}. Defaulting to L5.`);
    Object.assign(FEATURES, COURSE_PRESETS["L5"]);
  } else {
    Object.assign(FEATURES, preset);
  }
}

function buildGlobals() {
  PROJECT_FIELDS = FEATURES.PROJECT_TRACKING ? CONFIG.PROJECTS.map(p => ({
    dColName: normHeader(`${p}_submission_deadline`),
    rColName: normHeader(`${p}_resubmission_deadline`),
    sColName: normHeader(`${p}_submission_status`)
  })) : [];

  CRM_FIELDS = [
    CONFIG.CORE_FIELDS.NAME, "Course Of Interest Code", "Lead Status", "Discord Nickname", "Discord User ID",
    CONFIG.CORE_FIELDS.EMAIL, "Tag", "Cohort Facilitator", "Additional Learning Needs / LLDD", CONFIG.CORE_FIELDS.ID
  ];

  if (FEATURES.PROJECT_TRACKING) {
    CRM_FIELDS.push("Projects submitted", "Pathway");
    CONFIG.PROJECTS.forEach(p => {
      CRM_FIELDS.push(`${p}_submission_deadline`, `${p}_late_submission_deadline`, `${p}_resubmission_deadline`, `${p}_submission_status`);
    });
  }

  if (FEATURES.SPECIALISATION) {
    CRM_FIELDS.push("Specialization Programme Name", "Specialization Selection Deadline");
  }

  // Set up Manual Columns
  const manualSet = new Set(["Notes", "Risk Notes"].map(normHeader));
  if (FEATURES.L3_MODULE_TRACKING) {
    CONFIG.L3_MODULES.forEach(mod => {
      manualSet.add(normHeader(mod));
      manualSet.add(normHeader(`${mod} notes`));
    });
  }

  COLUMN_OWNERS = {
    CRM: new Set(CRM_FIELDS.map(normHeader)),
    LMS: FEATURES.LMS_TRACKING ? new Set(["LMS Last Activity", "Days Since LMS Activity", "Progress", "Current Lesson"].map(normHeader)) : new Set(),
    GITHUB: FEATURES.GITHUB_TRACKING ? new Set(["GitHub Username", "Last GitHub Activity", "GitHub Profile"].map(normHeader)) : new Set(),
    MANUAL: manualSet
  };

  ALL_KNOWN_HEADERS = new Set([
    ...COLUMN_OWNERS.CRM, 
    ...COLUMN_OWNERS.LMS, 
    ...COLUMN_OWNERS.GITHUB, 
    ...COLUMN_OWNERS.MANUAL, 
    H.AUTO_RISK_REASON
  ]);
}

// ================= GLOBAL UTILITIES =================

function showProgress(ss, message, title = "Sync", timeout = -1) {
  console.log(`▶️ ${title}: ${message}`);
  try { if (ss) { ss.toast(message, title, timeout); SpreadsheetApp.flush(); } } catch (e) {}
}

function resolveAlias(norm) {
  if (!norm) return norm;
  if (norm.includes("course") && norm.includes("interest")) return H.COURSE_CODE;
  if (norm.includes("discord") && norm.includes("user")) return H.DISC_ID;
  if (norm.includes("specialization") && norm.includes("program")) return normHeader("Specialization Programme Name");
  if (norm.includes("additional learning") || norm.includes("lldd")) return normHeader("Additional Learning Needs / LLDD");
  if ((norm.includes("record") || norm.includes("learner") || norm.includes("final student")) && norm.includes("id")) return H.ID;
  return norm;
}

function getHeaderMap(sheetName, headersNorm, isTracker = false) {
  if (Object.keys(HEADER_CACHE).length > 50) HEADER_CACHE = {}; 
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, JSON.stringify(headersNorm));
  const txtHash = rawHash.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
  const key = sheetName + "_" + isTracker + "_" + headersNorm.length + "_" + txtHash;
  
  if (!HEADER_CACHE[key]) {
    HEADER_CACHE[key] = buildHeaderMap(headersNorm, isTracker);
  }
  return HEADER_CACHE[key];
}

function buildHeaderMap(headers, isTracker = false) {
  const map = {};
  const CRITICAL_ALIASES = new Set([H.ID, H.EMAIL]); 
  
  headers.forEach((h, i) => {
    let norm = normHeader(h);
    
    // Strict schema mapping first for reliability
    if (!isTracker && STRICT_HEADER_MAP[norm]) {
        norm = STRICT_HEADER_MAP[norm];
    } else {
        norm = resolveAlias(norm); // Fuzzy fallback
    }

    if (norm) {
      if (isTracker && map[norm] !== undefined) {
        if (CRITICAL_ALIASES.has(norm)) {
           throw new Error(`❌ FATAL ALIAS COLLISION: Multiple columns mapped to critical field "${norm}". Fix sheet headers to prevent data corruption.`);
        } else {
           globalAliasCollisions++;
           if (DEBUG_LOGS) console.warn(`⚠️ Alias collision detected in Tracker: "${norm}" mapped to col ${map[norm]}, skipping col ${i}.`);
           return; 
        }
      }
      map[norm] = i;
      if (isTracker && DEBUG_LOGS && !ALL_KNOWN_HEADERS.has(norm) && !_loggedUnknowns.has(norm)) {
        globalUnknownHeaders++;
        if (_loggedUnknowns.size > 200) return; 
        console.warn(`⚠️ Schema Drift Alert: Unknown header detected in Tracker -> "${norm}"`);
        _loggedUnknowns.add(norm);
      }
    }
  });

  return map;
}

function normalizeForComparison(v) {
  if (v === null || v === undefined || v === "") return "";
  if (v instanceof Date) return `DATE:${v.getTime()}`;
  if (typeof v === "number") return `NUM:${v}`;
  if (typeof v === "boolean") return `BOOL:${v}`;
  if (typeof v === "object") {
     const sortedKeys = Object.keys(v).sort();
     const sortedObj = {};
     for(let k of sortedKeys) sortedObj[k] = v[k];
     return `OBJ:${JSON.stringify(sortedObj)}`;
  }
  const str = String(v).trim().toLowerCase(); 
  if (str.startsWith("=")) return `FORMULA:${str.replace(/\s+/g, "")}`;
  return `STR:${str}`;
}

function isDifferent(a, b) {
  return normalizeForComparison(a) !== normalizeForComparison(b);
}

function fastHash(row) {
  if (!row || !Array.isArray(row)) return "";
  let hash = "";
  for (let i = 0; i < row.length; i++) {
    let v = row[i];
    if (v instanceof Date) hash += i + ":D:" + v.getTime() + "|";
    else if (typeof v === "number") hash += i + ":N:" + v + "|";
    else if (typeof v === "boolean") hash += i + ":B:" + v + "|";
    else if (v === null || v === undefined || v === "") hash += i + ":|";
    else if (typeof v === "object") {
        const sortedKeys = Object.keys(v).sort();
        const sortedObj = {};
        for(let k of sortedKeys) sortedObj[k] = v[k];
        hash += i + ":O:" + JSON.stringify(sortedObj) + "|";
    }
    else {
      const str = String(v).trim().toLowerCase();
      hash += i + ":" + (str.startsWith("=") ? "F:" + str.replace(/\s+/g, "") : "S:" + str) + "|";
    }
  }
  return hash;
}

function logDiff(origRow, newRow, headerMap, identifier) {
  if (!DEBUG_LOGS) return;
  let diffs = [];
  Object.keys(headerMap).forEach(key => {
    const idx = headerMap[key];
    const orig = origRow[idx];
    const newVal = newRow[idx];
    if (isDifferent(orig, newVal)) { diffs.push(`[${key}]: '${orig}' -> '${newVal}'`); }
  });
  if (diffs.length > 0) { console.log(`🔄 UPDATED [${identifier}]:\n  ` + diffs.join("\n  ")); }
}

function enforceCacheLimit(cacheObj, limit) {
  const keys = Object.keys(cacheObj);
  if (keys.length > limit) {
    const overflow = keys.length - limit;
    const keysToRemove = Object.entries(cacheObj).sort((a,b) => (a[1]?.ts || 0) - (b[1]?.ts || 0)).slice(0, overflow).map(e => e[0]);
    keysToRemove.forEach(k => delete cacheObj[k]);
  }
}

// ================= ROW MODEL =================
function createRowModel(row, headerMap, readOnlyFields = new Set()) {
  return {
    get(field) {
      const idx = headerMap[field];
      return idx !== undefined ? row[idx] : undefined;
    },
    set(field, value) {
      if (readOnlyFields.has(field)) {
         return; 
      }
      const idx = headerMap[field];
      if (idx !== undefined) {
        if (typeof row[idx] === "string" && row[idx].startsWith("=") && !(typeof value === "string" && value.startsWith("="))) {
           return;
        }
        row[idx] = value;
      }
    },
    safeSet(field, value) {
      if (isDifferent(this.get(field), value)) {
        this.set(field, value);
      }
    },
    has(field) {
      return headerMap[field] !== undefined;
    },
    colIndex(field) {
      return headerMap[field];
    },
    get raw() { 
      return row; 
    }
  };
}

let _dateCache = new Map();

function parseRobustDate(rawDate) {
  if (rawDate === null || rawDate === undefined) return null;
  if (rawDate instanceof Date) return rawDate;
  
  const str = String(rawDate).trim();
  if (str === "" || str.toLowerCase() === "null") return null; 
  if (isNever(str)) return NEVER; 
  
  if (_dateCache.has(str)) {
    const cached = _dateCache.get(str);
    _dateCache.delete(str);
    _dateCache.set(str, cached); // True LRU bump
    return cached;
  }
  
  if (_dateCache.size > 1000) {
    const keys = _dateCache.keys();
    for (let i = 0; i < 200; i++) _dateCache.delete(keys.next().value);
  }
  
  let parsedDate = null;
  if (str.includes('T')) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) parsedDate = d;
  } else {
    const parts = str.split(" ");
    const datePart = parts[0], timePart = parts[1] || "00:00:00";
    let y, m, d;
    const tParts = timePart.split(":");
    const th = parseInt(tParts[0], 10) || 0;
    const tm = parseInt(tParts[1], 10) || 0;
    const ts = parseInt(tParts[2], 10) || 0;
    if (datePart.includes("-") || datePart.includes("/")) {
      const separator = datePart.includes("-") ? "-" : "/";
      const dParts = datePart.split(separator);
      if (dParts.length === 3) {
        const p1 = parseInt(dParts[0], 10);
        const p2 = parseInt(dParts[1], 10);
        const p3 = parseInt(dParts[2], 10);
        if (p1 > 1000) { y = p1; m = p2; d = p3; } 
        else {
          y = p3;
          if (p1 > 12 && p2 <= 12) { d = p1; m = p2; } 
          else if (p2 > 12 && p1 <= 12) { m = p1; d = p2; } 
          else { 
              if (CONFIG.EXPECTED_DATE_FORMAT === "US") { m = p1; d = p2; } 
              else { d = p1; m = p2; }
              globalAmbiguousDateCount++;
          } 
        }
        if (d && m && y) parsedDate = new Date(y, m - 1, d, th, tm, ts);
      }
    }
  }
  
  if (!parsedDate || isNaN(parsedDate.getTime())) {
      if (DEBUG_LOGS && VERBOSE_LOGS) console.warn(`⚠️ Invalid date encountered: '${str}'`);
      return null;
  }
  _dateCache.set(str, parsedDate);
  return parsedDate;
}

function calculateDaysSince(dateVal) {
  if (!dateVal || dateVal === "" || isNever(dateVal) || dateVal === "INVALID_DATE") return null; 
  let d = parseRobustDate(dateVal); 
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  const today = new Date(); today.setHours(0,0,0,0);
  const past = new Date(d); past.setHours(0,0,0,0);
  const days = Math.floor((today - past) / (1000 * 60 * 60 * 24));
  if (days < 0) return null; 
  return days;
}

// ================= PREFLIGHT & TRANSACTION MANAGER =================

function runPreflightChecks(isSetupCheck = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getDocumentProperties();

  const fileId = props.getProperty("CRM_FILE_IDENTIFIER");
  if (!fileId) {
    throw new Error("⚠️ Setup incomplete:\n\nCRM report name not configured.\n\nGo to:\nStudent Tracker → ⚙️ Change CRM Report Name");
  }

  const threads = GmailApp.search(buildCrmSearchQuery(fileId, true), 0, 1);

  if (!threads.length) {
    throw new Error(`⚠️ No CRM emails found from the last 7 days.\n\nLooking for CSV attachment containing:\n"${fileId}"\n\nFix:\n• Check your inbox\n• Or update report name via the menu`);
  }

  const tracker = ss.getSheetByName(CONFIG.SHEETS.TRACKER);
  if (!tracker) {
    if (isSetupCheck) throw new Error("⚠️ Tracker sheet missing.\n\nRun a Full Sync once to generate the structure.");
    return true; 
  }

  if (tracker.getLastRow() === 0 || tracker.getLastColumn() === 0) {
    if (isSetupCheck) throw new Error("⚠️ Tracker not initialised.\n\nRun a Full Sync once to generate the structure.");
    return true; 
  }

  const headers = tracker.getRange(1, 1, 1, tracker.getLastColumn()).getValues()[0].map(normHeader);

  if (!headers.includes(H.EMAIL) || !headers.includes(H.ID)) {
    throw new Error("❌ Critical columns missing (Email / Record Id).\n\nDo not rename or delete core headers.");
  }

  return true;
}

function checkCrashLock() {
  const props = PropertiesService.getScriptProperties();
  const lockStr = props.getProperty("SYNC_TRANSACTION_ACTIVE");
  if (lockStr) {
    const info = JSON.parse(lockStr);
    const TEN_MINUTES = 10 * 60 * 1000;
    if (Date.now() - new Date(info.startedAt).getTime() > TEN_MINUTES) {
        if (DEBUG_LOGS) console.warn("⚠️ Cleared stale crash lock from a previous timed-out execution.");
        props.deleteProperty("SYNC_TRANSACTION_ACTIVE");
        return;
    }
    throw new Error(`❌ CRITICAL: Previous sync crashed during the [${info.stage}] Write/Delete phase at ${info.startedAt}. Manual inspection of the Tracker is required to clear partial duplicates. Clear the crash lock from the custom menu when safe to proceed.`);
  }
}

function verifySyncContext() {
   const props = PropertiesService.getScriptProperties();
   if (!props.getProperty("ACTIVE_SYNC_CONTEXT")) {
       throw new Error("❌ CRITICAL: Module executed outside of a secure sync context. Aborting to prevent data corruption.");
   }
}

function executeSecureWrite(stageName, writeAction) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("SYNC_TRANSACTION_ACTIVE")) throw new Error(`❌ CRITICAL: Nested transaction detected during [${stageName}]. Architecture violation.`);
  
  props.setProperty("SYNC_TRANSACTION_ACTIVE", JSON.stringify({ startedAt: new Date().toISOString(), stage: stageName }));
  let success = false;
  try { writeAction(); success = true; } 
  finally { if (success) props.deleteProperty("SYNC_TRANSACTION_ACTIVE"); }
}

// ================= MODULAR RISK ENGINE =================
function applyRiskEngine() {
  verifySyncContext();
  if (!FEATURES.RISK_ENGINE) return { redCount: 0, aborted: false };

  const trackerSheet = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  const data = getSheetDataWithFormulas(trackerSheet);
  if (data.length <= 1) return { redCount: 0, aborted: false };
  
  let hRow = data[0].map(normHeader);
  let autoReasonColIdx = hRow.indexOf(H.AUTO_RISK_REASON);
  let notesColIdx = hRow.indexOf(H.RISK_NOTES);
  
  let headersModified = false;
  let currentCols = data[0].length;
  
  if (autoReasonColIdx === -1) {
    autoReasonColIdx = currentCols++;
    data[0].push("Auto Risk Reason");
    for (let i = 1; i < data.length; i++) data[i].push("");
    headersModified = true;
  }
  if (notesColIdx === -1) {
    notesColIdx = currentCols++;
    data[0].push("Risk Notes");
    for (let i = 1; i < data.length; i++) data[i].push("");
    headersModified = true;
  }
  
  let headerMap = getHeaderMap(CONFIG.SHEETS.TRACKER, hRow, true); 

  if (headersModified) {
    const maxCols = trackerSheet.getMaxColumns();
    if (maxCols < currentCols) trackerSheet.insertColumnsAfter(maxCols, currentCols - maxCols);
    trackerSheet.getRange(1, 1, 1, currentCols).setValues([data[0]]).setFontWeight("bold");
    
    hRow = data[0].map(normHeader);
    headerMap = getHeaderMap(CONFIG.SHEETS.TRACKER, hRow, true);
  }
  
  const headers = data[0];
  const numRows = Math.max(data.length - 1, 0);
  let bg = numRows > 0 ? trackerSheet.getRange(2, 1, numRows, headers.length).getBackgrounds() : [];
  
  let redCount = 0;
  let forceWriteRows = new Set();
  let anyChange = false;
  let processedRows = 0; 
  
  const nextData = data.map(r => r.slice());
  let rowsToWrite = [];
  
  const manualColIndexes = [];
  COLUMN_OWNERS.MANUAL.forEach(field => {
      const idx = headerMap[field];
      if (idx !== undefined) manualColIndexes.push(idx);
  });

  for (let i = 1; i < nextData.length; i++) {
    shouldStop();
    processedRows++;
    
    const student = createRowModel(nextData[i], headerMap);
    const originalBgRow = bg[i - 1] || [];
    const workingBgRow = originalBgRow.slice();
    const origRowCopy = student.raw.slice();
    const origBgCopy = workingBgRow.slice();
    
    const origHash = fastHash(origRowCopy);
    const origBgHash = fastHash(origBgCopy);
    
    const { isRed, totalRisk, flags } = evaluateRowRisk(student, workingBgRow);
    const flagsText = flags.length ? [...new Set(flags.map(f => f.trim()))].join(" | ") : "";
    
    student.safeSet(H.AUTO_RISK_REASON, flagsText);
    
    if (isRed) redCount++;
    
    const newHash = fastHash(student.raw);
    const newBgHash = fastHash(workingBgRow);
    
    const valChanged = origHash !== newHash;
    const bgChanged = origBgHash !== newBgHash;
    
    if (valChanged || bgChanged) {
       anyChange = true;
       forceWriteRows.add(i);
       const sheetRow = i + 1;
       const safeBg = workingBgRow;
       
       if (valChanged && bgChanged) {
           rowsToWrite.push({ row: sheetRow, values: nextData[i].slice(), bg: safeBg });
       } else if (valChanged) {
           rowsToWrite.push({ row: sheetRow, values: nextData[i].slice(), bg: null });
       } else if (bgChanged) {
           rowsToWrite.push({ row: sheetRow, values: null, bg: safeBg });
       }
       if (bgChanged) bg[i-1] = workingBgRow.slice();
    }
  }
  
  if (processedRows === 0 && nextData.length > 1) {
    throw new Error("❌ CRITICAL: No Risk Engine rows successfully processed. Aborting to prevent blank overwrite.");
  }
  
  if (!anyChange && rowsToWrite.length === 0 && processedRows > 0) {
    return { redCount: redCount, aborted: false };
  }

  if (rowsToWrite.length > 0) {
     executeSecureWrite("RISK_ENGINE", () => {
       writeRowsInBatches(trackerSheet, rowsToWrite, headerMap, manualColIndexes);
     });
  }
  
  return { redCount: redCount, aborted: false };
}

function evaluateDeadlineRisk(student, bgRow) {
  let riskScore = 0;
  let reasons = [];
  const today = new Date(); today.setHours(0,0,0,0);
  let worstDeadline = null, worstCol = null;
  let isResub = false;
  
  const CRM_COLUMNS = COLUMN_OWNERS.CRM;
  const RISK_COLORS = new Set(["#ea9999", "#f6b26b", "#ffe599", "#fff2cc"]);

  // Unconditionally wipe Risk colors ONLY on CRM-owned columns, preserving manual highlight colors
  for (const pFields of PROJECT_FIELDS) {
    const dCol = student.colIndex(pFields.dColName);
    const rCol = student.colIndex(pFields.rColName);
    if (dCol !== undefined && CRM_COLUMNS.has(pFields.dColName) && RISK_COLORS.has(bgRow[dCol])) bgRow[dCol] = "#ffffff";
    if (rCol !== undefined && CRM_COLUMNS.has(pFields.rColName) && RISK_COLORS.has(bgRow[rCol])) bgRow[rCol] = "#ffffff";
  }

  for (const pFields of PROJECT_FIELDS) {
    if (!student.has(pFields.dColName)) continue;
    const dCol = student.colIndex(pFields.dColName);
    const rCol = student.colIndex(pFields.rColName);
    const d = parseRobustDate(student.get(pFields.dColName));
    const resubDate = parseRobustDate(student.get(pFields.rColName));
    const statusRaw = student.get(pFields.sColName);
    const status = statusRaw ? statusRaw.toString().toLowerCase().trim() : "";
    let activeDate = null, activeCol = null;
    let currentIsResub = false;

    if (resubDate instanceof Date && !isNaN(resubDate.getTime())) {
      const graceDate = new Date(resubDate); graceDate.setDate(graceDate.getDate() + CONFIG.RISK.RESUB_GRACE_DAYS);
      if (today.getTime() <= graceDate.getTime()) { activeDate = resubDate; activeCol = rCol; currentIsResub = true; }
    } else if (d instanceof Date && !isNaN(d.getTime()) && !CONFIG.RISK.SUBMITTED_STATUSES.some(s => status.includes(s))) {
      activeDate = d; activeCol = dCol;
    }
    if (activeDate && (!worstDeadline || activeDate < worstDeadline)) { 
      worstDeadline = activeDate; worstCol = activeCol; isResub = currentIsResub; 
    }
  }

  if (worstDeadline && worstCol !== null && worstCol !== undefined) {
    const diff = Math.floor((worstDeadline - today) / (1000 * 60 * 60 * 24));
    let targetColor = "#ffffff";
    if (diff < 0) { 
      targetColor = "#ea9999"; riskScore += isResub ? 3 : 4; reasons.push(isResub ? "Resub overdue" : "Deadline missed");
    } else if (diff <= 7) { 
      targetColor = "#f6b26b"; riskScore += 2; reasons.push("Deadline ≤ 7 days");
    } else if (diff <= 14) { 
      targetColor = "#ffe599"; riskScore += 1;
    } else if (diff <= 30) { targetColor = "#fff2cc"; }
    bgRow[worstCol] = targetColor;
  }
  return { score: riskScore, reasons: reasons };
}

function evaluateLmsRisk(student, bgRow) {
  let riskScore = 0;
  let reasons = [];
  const RISK_COLORS = new Set(["#ea9999", "#f6b26b", "#ffe599", "#fff2cc"]);
  
  if (student.has(H.LMS_ACT)) {
    const days = calculateDaysSince(student.get(H.LMS_ACT));
    let targetColor = "#ffffff";
    
    const tActCol = student.colIndex(H.LMS_ACT);
    const tDaysCol = student.colIndex(H.LMS_DAYS);
    
    if (tActCol !== undefined && RISK_COLORS.has(bgRow[tActCol])) bgRow[tActCol] = "#ffffff";
    if (tDaysCol !== undefined && RISK_COLORS.has(bgRow[tDaysCol])) bgRow[tDaysCol] = "#ffffff";

    if (days === null) {} 
    else if (days > 30) { targetColor = "#ea9999"; riskScore += 2; reasons.push("LMS Inactive 30+ days"); } 
    else if (days > 14) { targetColor = "#f6b26b"; riskScore += 1; reasons.push("LMS Inactive 14+ days"); } 
    else if (days > 7) { targetColor = "#ffe599"; }
    
    if (targetColor !== "#ffffff") {
        if (tActCol !== undefined) bgRow[tActCol] = targetColor;
        if (tDaysCol !== undefined) bgRow[tDaysCol] = targetColor;
    }
  }
  return { score: riskScore, reasons: reasons };
}

function evaluateRowRisk(student, bgRow) {
  while (bgRow.length < student.raw.length) bgRow.push("#ffffff");
  
  let deadlineRisk = 0;
  let lmsRisk = 0;
  let reasons = [];

  if (FEATURES.PROJECT_TRACKING) {
    const pRisk = evaluateDeadlineRisk(student, bgRow);
    deadlineRisk += pRisk.score;
    reasons.push(...pRisk.reasons);
  }

  if (FEATURES.LMS_TRACKING) {
    const lRisk = evaluateLmsRisk(student, bgRow);
    lmsRisk += lRisk.score;
    reasons.push(...lRisk.reasons);
  }

  let isRed = false;
  let totalRisk = Math.min(deadlineRisk + lmsRisk, 4);

  if (student.has(H.NAME)) {
    const nameColIdx = student.colIndex(H.NAME);
    const RISK_COLORS = new Set(["#ea9999", "#f6b26b", "#ffe599", "#fff2cc"]);
    let targetColor = bgRow[nameColIdx] || "#ffffff"; 
    
    if (RISK_COLORS.has(targetColor)) targetColor = "#ffffff";
    
    if (deadlineRisk >= 3 || totalRisk >= 4) { targetColor = "#ea9999"; isRed = true; } 
    else if (totalRisk >= 2) { targetColor = "#f6b26b"; } 
    else if (totalRisk === 1) { targetColor = "#ffe599"; } 
    
    bgRow[nameColIdx] = targetColor;
  }
  return { isRed, totalRisk, flags: reasons };
}

// ================= DATA VALIDATION =================
function enforceDateValidation() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  const lastCol = sheet.getLastColumn(), lastRow = sheet.getLastRow();
  if (lastRow <= 1 || lastCol === 0) return;
  const headersNorm = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normHeader);
  const dateCols = headersNorm.map((hNorm, i) => hNorm.includes("deadline") ? i : undefined).filter(i => i !== undefined);
  if (dateCols.length === 0) return;
  const rule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).setHelpText("❌ Invalid Input: Please enter a valid date format (DD/MM/YYYY).").build();
  dateCols.forEach(col => sheet.getRange(2, col + 1, lastRow - 1).setDataValidation(rule));
}

function normalizeDeadlineDates() {
  const sheet = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  const lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return;
  const headersNorm = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(normHeader);
  const dateCols = headersNorm.map((hNorm, i) => hNorm.includes("deadline") ? i : undefined).filter(i => i !== undefined);
  if (dateCols.length === 0) return;
  dateCols.forEach(colIdx => {
    const range = sheet.getRange(2, colIdx + 1, lastRow - 1, 1);
    const values = range.getValues();
    let changed = false;
    for (let r = 0; r < values.length; r++) {
       const orig = values[r][0];
       const parsed = parseRobustDate(orig);
       if (parsed !== orig && parsed instanceof Date) { values[r][0] = parsed; changed = true; } 
       else if (parsed === null && orig !== "") { values[r][0] = ""; changed = true; }
    }
    if (changed) range.setValues(values);
  });
}

// ================= HELPERS =================
function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function mapStatus(raw) {
  if (!raw) return "";
  const clean = raw.toString().toLowerCase().trim();
  if (/late/.test(clean)) return "Submitted late";
  if (/on\s*time/.test(clean)) return "Submitted on time";
  if (/not|pending|incomplete/.test(clean)) return "Not yet submitted";
  return raw;
}

function applyDiscordLink(val){
  const idClean = val.toString().trim().replace(/[^\d]/g,'');
  return `=HYPERLINK("https://discord.com/users/${idClean}", "Message")`;
}

function applyCrmProfileLink(name, recordId) {
  if (!name || !recordId) return name;
  const idClean = String(recordId).trim().replace(/^zcrm_/, '');
  const safeName = String(name).replace(/"/g, '""');
  return `=HYPERLINK("https://crm.zoho.com/crm/org41701914/tab/Contacts/${idClean}", "${safeName}")`;
}

function getSheetDataWithFormulas(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const formulas = range.getFormulas();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      if (formulas[r][c]) values[r][c] = formulas[r][c];
    }
  }
  return values;
}

function writeRowsInBatches(sheet, rows, headerMap, manualColIndexes = null){
  if(!rows.length) return;
  rows.sort((a,b)=>a.row-b.row);
  const maxRowRequired = Math.max(...rows.map(r => r.row));
  const maxRows = sheet.getMaxRows();
  if (maxRowRequired > maxRows) {
      sheet.insertRowsAfter(maxRows, Math.min(50, maxRowRequired - maxRows + 5));
  }
  
  let runs = [];
  for (const r of rows) {
    const hasVal = r.values !== null && r.values !== undefined;
    const hasBg = r.bg !== null && r.bg !== undefined;
    const lastRun = runs[runs.length - 1];
    if (lastRun && lastRun.endRow === r.row - 1 && lastRun.hasVal === hasVal && lastRun.hasBg === hasBg) {
      if (hasVal) lastRun.values.push(r.values);
      if (hasBg) lastRun.bg.push(r.bg);
      lastRun.endRow = r.row;
    } else {
      runs.push({ startRow: r.row, endRow: r.row, hasVal: hasVal, hasBg: hasBg, values: hasVal ? [r.values] : [], bg: hasBg ? [r.bg] : [] });
    }
  }
  for (const run of runs) {
    const executeWrite = () => {
      const startRow = Number(run.startRow);
      const numRows = Number(run.endRow - run.startRow + 1);
      const numCols = run.hasVal ? Number(run.values[0].length) : (run.hasBg ? Number(run.bg[0].length) : 0);
      
      if (numCols > 0) {
        if (sheet.getMaxRows() < startRow) {
           throw new Error(`❌ CRITICAL: Sheet structure mutated externally during write phase (Row ${startRow} out of bounds). Aborting to prevent data corruption.`);
        }
        if (sheet.getLastColumn() !== numCols) {
           throw new Error(`❌ CRITICAL: Sheet column count changed mid-execution. Expected ${numCols}, got ${sheet.getLastColumn()}. Aborting write.`);
        }
        
        const targetRange = sheet.getRange(startRow, 1, numRows, numCols);
        
        if (run.hasVal && FEATURES.ALLOW_OVERWRITE_NON_MANUAL === false) {
           const liveValues = targetRange.getValues();
           const liveFormulas = targetRange.getFormulas();
           for (let r = 0; r < numRows; r++) {
               if (liveValues[r].length !== run.values[r].length) throw new Error("❌ Column mismatch detected.");
               for (let cIdx = 0; cIdx < run.values[r].length; cIdx++) {
                   const liveVal = liveFormulas[r][cIdx] ? liveFormulas[r][cIdx] : liveValues[r][cIdx];
                   if (run.values[r][cIdx] !== liveVal) run.values[r][cIdx] = liveVal;
               }
           }
        } else if (run.hasVal && manualColIndexes && manualColIndexes.length > 0) {
          const liveValues = targetRange.getValues();
          const liveFormulas = targetRange.getFormulas();
          for (let r = 0; r < numRows; r++) {
            if (liveValues[r].length !== run.values[r].length) {
                throw new Error("❌ Column mismatch detected during JIT merge. Sheet structure may have changed. Aborting write to prevent data corruption.");
            }
            manualColIndexes.forEach(cIdx => {
              if (cIdx < run.values[r].length) {
                const liveVal = liveFormulas[r][cIdx] ? liveFormulas[r][cIdx] : liveValues[r][cIdx];
                if (run.values[r][cIdx] !== liveVal) {
                    run.values[r][cIdx] = liveVal; 
                }
              }
            });
          }
        }

        if (run.hasVal && run.hasBg) {
            targetRange.setValues(run.values).setBackgrounds(run.bg);
        } else if (run.hasVal) {
            targetRange.setValues(run.values);
        } else if (run.hasBg) {
            targetRange.setBackgrounds(run.bg);
        }
      }
    };
    try { executeWrite(); } 
    catch(e) { 
      if(DEBUG_LOGS) console.log(`Batch write failed at row ${run.startRow}, retrying both values and backgrounds...`, e);
      try { Utilities.sleep(500); executeWrite(); } 
      catch (retryErr) { throw new Error(`❌ Write failed completely at row ${run.startRow}. Sheet may be out of sync.`); }
    }
  }
}

// ================= DYNAMIC SHEET INITIALIZATION =================
function initializeCoreSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tracker = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  if (tracker.getLastRow() === 0 || tracker.getLastColumn() === 0) {
    const headers = [
      CONFIG.CORE_FIELDS.NAME, "Course Of Interest Code", "Lead Status", "Discord Nickname", "Discord User ID",
      CONFIG.CORE_FIELDS.EMAIL, "Tag", "Cohort Facilitator"
    ];
    
    if (FEATURES.GITHUB_TRACKING) { headers.push("GitHub Username", "Last GitHub Activity", "GitHub Profile"); }
    if (FEATURES.LMS_TRACKING) { headers.push("LMS Last Activity", "Days Since LMS Activity", "Progress", "Current Lesson"); }

    headers.push("Notes");

    if (FEATURES.PROJECT_TRACKING) {
      headers.push("Projects submitted", "Pathway");
      CONFIG.PROJECTS.forEach(p => { 
        headers.push(`${p}_submission_deadline`, `${p}_late_submission_deadline`, `${p}_resubmission_deadline`, `${p}_submission_status`); 
      });
    }

    if (FEATURES.SPECIALISATION) { headers.push("Specialization Programme Name", "Specialization Selection Deadline"); }

    if (FEATURES.L3_MODULE_TRACKING) {
      CONFIG.L3_MODULES.forEach(mod => {
        headers.push(mod, `${mod} notes`);
      });
    }

    headers.push("Additional Learning Needs / LLDD", CONFIG.CORE_FIELDS.ID);

    if (FEATURES.RISK_ENGINE) { headers.push("Auto Risk Reason", "Risk Notes"); }

    tracker.getRange(1,1,1,headers.length).setValues([headers]); 
    tracker.setFrozenRows(1);
  }
  
  if (FEATURES.LMS_TRACKING) {
    let lms = getOrCreateSheet(CONFIG.SHEETS.LMS_RAW);
    if (lms.getLastRow() === 0) lms.getRange("A1").setValue("Paste Cypher export here");
  }

  const sheet1 = ss.getSheetByName("Sheet1");
  if (sheet1) { try { ss.deleteSheet(sheet1); } catch (e) {} }
}

// ================= MENU =================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('📊 Student Tracker')
    .addItem('▶️ Run Full Sync', 'runFullSync')
    .addSeparator()
    .addItem('🔄 Sync CRM Only', 'runCrmSyncOnly')
    .addItem('🎓 Sync Cypher LMS Activity', 'runLmsSyncOnly')
    .addItem('🐙 Update GitHub Activity', 'runGithubSyncOnly')
    .addSeparator()
    .addItem('🧪 Run Setup Check', 'runSetupCheck')
    .addItem('⚙️ Change CRM Report Name', 'resetTrackerConfig')
    .addItem('🔓 Clear Crash Lock', 'resetCrashLock')
    .addItem('🔓 Allow Duplicate CRM File', 'resetCrmFileLock')
    .addItem('🧹 Reset GitHub API Cache', 'resetGithubCache')
    .addToUi();
}

function resetTrackerConfig() {
  PropertiesService.getDocumentProperties().deleteProperty("CRM_FILE_IDENTIFIER");
  PropertiesService.getDocumentProperties().deleteProperty("COURSE_TYPE");
  loadCourseType();
  showProgress(SpreadsheetApp.getActiveSpreadsheet(), "✅ Tracker configuration updated!", "Success", 5);
}

function resetCrashLock() {
  PropertiesService.getScriptProperties().deleteProperty("SYNC_TRANSACTION_ACTIVE");
  showProgress(SpreadsheetApp.getActiveSpreadsheet(), "✅ Crash lock cleared. Please ensure sheet data is clean before syncing.", "Lock Cleared", 5);
}

function resetCrmFileLock() {
  PropertiesService.getDocumentProperties().deleteProperty("LAST_CRM_FILE");
  showProgress(SpreadsheetApp.getActiveSpreadsheet(), "✅ CRM File Lock cleared. You can now re-sync the same CSV file.", "Lock Cleared", 5);
}

function resetGithubCache() {
  PropertiesService.getScriptProperties().deleteProperty("GITHUB_CACHE");
  PropertiesService.getScriptProperties().deleteProperty("GITHUB_ACT_CACHE");
  showProgress(SpreadsheetApp.getActiveSpreadsheet(), "✅ GitHub API Caches cleared. Next sync will fetch fresh data.", "Caches Cleared", 5);
}

function runSetupCheck() {
  try {
    loadCourseType();
    runPreflightChecks(true);
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "✅ System is correctly configured",
      "Setup Check",
      5
    );
  } catch (e) {
    SpreadsheetApp.getUi().alert(e.message);
  }
}

// ================= MASTER ROUTER =================
function setupSyncContext() {
    const props = PropertiesService.getScriptProperties();
    props.setProperty("ACTIVE_SYNC_CONTEXT", "TRUE");
    HEADER_CACHE = {};
    globalAmbiguousDateCount = 0;
    globalUnknownHeaders = 0;
    globalAliasCollisions = 0;
}

function teardownSyncContext() {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty("ACTIVE_SYNC_CONTEXT");
}

function runFullSync(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { showProgress(SpreadsheetApp.getActiveSpreadsheet(), "⏳ Another sync is currently running. Please wait.", "Busy", 5); return; }
  
  checkCrashLock();
  setupSyncContext();
  
  try {
    loadCourseType(); 
    runPreflightChecks(false);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty("SYNC_STAGE"); props.deleteProperty("CRM_META"); props.deleteProperty("SYNC_STATS"); props.deleteProperty("LMS_STATS"); props.deleteProperty("GITHUB_STATS");

    showProgress(ss, "Step 1: Fetching CRM...", "Sync", -1);
    const crmMeta = importCrmFromEmail(); 
    
    initializeCoreSheets(); enforceDateValidation(); normalizeDeadlineDates();
    
    showProgress(ss, "Step 2: Syncing CRM Data...", "Sync", -1);
    const syncStats = syncCrmToTracker();
    
    let lmsStats = null;
    if (FEATURES.LMS_TRACKING) {
      showProgress(ss, "Step 3: Syncing LMS Activity...", "Sync", -1); 
      lmsStats = runLmsSync();
    }
    
    let githubStats = null;
    if (FEATURES.GITHUB_TRACKING) {
      showProgress(ss, "Step 4: GitHub Updates...", "Sync", -1); 
      githubStats = updateGithubActivity();
    }
    
    let riskStats = null;
    if (FEATURES.RISK_ENGINE) {
      showProgress(ss, "Step 5: Evaluating Risk Matrix...", "Sync", -1); 
      riskStats = applyRiskEngine();
    }
    
    if (globalAmbiguousDateCount > 0) console.warn(`⚠️ Encountered ${globalAmbiguousDateCount} ambiguous dates. Safely defaulted to ${CONFIG.EXPECTED_DATE_FORMAT} format based on CONFIG.`);
    const redCount = riskStats ? riskStats.redCount : 0;
    if (syncStats) syncStats.redCount = redCount;
    const validLmsStats = (lmsStats && !lmsStats.missingData) ? lmsStats : null;
    
    updateDashboard(crmMeta, syncStats, validLmsStats, githubStats, redCount);
    showProgress(ss, "✅ Full Sync & Audit Complete", "Done", 5);
    
  } catch (err) {
    console.error("❌ Sync Error:", err.stack || err); showProgress(SpreadsheetApp.getActiveSpreadsheet(), "❌ Sync failed: " + err.message, "Error", 10);
  } finally { 
      teardownSyncContext();
      lock.releaseLock(); 
  }
}

function runCrmSyncOnly(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { showProgress(SpreadsheetApp.getActiveSpreadsheet(), "⏳ Another sync is currently running. Please wait.", "Busy", 5); return; }
  
  checkCrashLock();
  setupSyncContext();
  
  try {
    loadCourseType();
    runPreflightChecks(false);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    showProgress(ss, "Starting CRM Sync...", "Sync", -1);
    
    showProgress(ss, "Step 1: Fetching CRM from Email...", "Sync", -1);
    const crmMeta = importCrmFromEmail();

    initializeCoreSheets(); enforceDateValidation(); normalizeDeadlineDates();
    
    showProgress(ss, "Step 2: Updating Tracker...", "Sync", -1);
    const syncStats = syncCrmToTracker();
    
    let riskStats = null;
    if (FEATURES.RISK_ENGINE) {
      showProgress(ss, "Step 3: Evaluating Risk Matrix...", "Sync", -1); 
      riskStats = applyRiskEngine();
    }
    
    if (globalAmbiguousDateCount > 0) console.warn(`⚠️ Encountered ${globalAmbiguousDateCount} ambiguous dates. Safely defaulted to ${CONFIG.EXPECTED_DATE_FORMAT} format based on CONFIG.`);
    const redCount = riskStats ? riskStats.redCount : 0;
    updateDashboard(crmMeta, syncStats, null, null, redCount);
    showProgress(ss, `✅ CRM Sync Complete! Added: ${syncStats ? syncStats.added : 0}, Updated: ${syncStats ? syncStats.updated : 0}`, "Done", 5);
  } catch (err) {
    console.error("❌ Sync Error:", err.stack || err); showProgress(SpreadsheetApp.getActiveSpreadsheet(), "❌ Sync failed: " + err.message, "Error", -1);
  } finally { 
      teardownSyncContext();
      lock.releaseLock(); 
  }
}

function runLmsSyncOnly(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { showProgress(SpreadsheetApp.getActiveSpreadsheet(), "⏳ Another sync is currently running. Please wait.", "Busy", 5); return; }
  
  checkCrashLock();
  setupSyncContext();
  
  try {
    loadCourseType();
    runPreflightChecks(false);
    if (!FEATURES.LMS_TRACKING) throw new Error("LMS Tracking is disabled for this course type.");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    showProgress(ss, "Starting Cypher LMS Sync...", "Sync", -1);
    initializeCoreSheets();
    
    showProgress(ss, "Step 1: Syncing LMS Activity...", "Sync", -1);
    const lmsStats = runLmsSync();
    
    let riskStats = null;
    if (FEATURES.RISK_ENGINE) {
      showProgress(ss, "Step 2: Evaluating Risk Matrix...", "Sync", -1); 
      riskStats = applyRiskEngine();
    }
    
    if (globalAmbiguousDateCount > 0) console.warn(`⚠️ Encountered ${globalAmbiguousDateCount} ambiguous dates. Safely defaulted to ${CONFIG.EXPECTED_DATE_FORMAT} format based on CONFIG.`);
    const redCount = riskStats ? riskStats.redCount : 0;
    
    if (lmsStats && lmsStats.missingData) {
      SpreadsheetApp.getUi().alert("⚠️ Cypher LMS Data Missing", `The '${CONFIG.SHEETS.LMS_RAW}' tab is empty.\n\nPlease paste your latest Cypher export.`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else if (lmsStats) {
      updateDashboard(null, null, lmsStats, null, redCount);
      showProgress(ss, `✅ LMS Sync Complete! Updated ${lmsStats.updated} students.`, "Done", 5);
    }
  } catch (err) {
    console.error("❌ Sync Error:", err.stack || err); showProgress(SpreadsheetApp.getActiveSpreadsheet(), "❌ Sync failed: " + err.message, "Error", -1);
  } finally { 
      teardownSyncContext();
      lock.releaseLock(); 
  }
}

function runGithubSyncOnly(){
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { showProgress(SpreadsheetApp.getActiveSpreadsheet(), "⏳ Another sync is currently running. Please wait.", "Busy", 5); return; }
  
  checkCrashLock();
  setupSyncContext();
  
  try {
    loadCourseType();
    runPreflightChecks(false);
    if (!FEATURES.GITHUB_TRACKING) throw new Error("GitHub Tracking is disabled for this course type.");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    showProgress(ss, "Starting GitHub Sync...", "Sync", -1);
    initializeCoreSheets();
    
    showProgress(ss, "Step 1: Fetching GitHub Updates...", "Sync", -1);
    const githubStats = updateGithubActivity();
    
    let riskStats = null;
    if (FEATURES.RISK_ENGINE) {
      showProgress(ss, "Step 2: Evaluating Risk Matrix...", "Sync", -1); 
      riskStats = applyRiskEngine();
    }
    
    if (globalAmbiguousDateCount > 0) console.warn(`⚠️ Encountered ${globalAmbiguousDateCount} ambiguous dates. Safely defaulted to ${CONFIG.EXPECTED_DATE_FORMAT} format based on CONFIG.`);
    const redCount = riskStats ? riskStats.redCount : 0;
    
    if (githubStats) {
      updateDashboard(null, null, null, githubStats, redCount);
      showProgress(ss, `✅ GitHub Sync Complete! Found ${githubStats.newFoundCount} new profiles, updated ${githubStats.updated} students.`, "Done", 5);
    }
  } catch (err) {
    console.error("❌ Sync Error:", err.stack || err); showProgress(SpreadsheetApp.getActiveSpreadsheet(), "❌ Sync failed: " + err.message, "Error", -1);
  } finally { 
      teardownSyncContext();
      lock.releaseLock(); 
  }
}

// ================= CRM SYNC ENGINE =================
function importCrmFromEmail() {
  verifySyncContext();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = getOrCreateSheet(CONFIG.SHEETS.CRM_RAW);
  sheet.clearContents();
  
  const fileIdentifier = PropertiesService.getDocumentProperties().getProperty("CRM_FILE_IDENTIFIER") || "student_data";
  
  let threads = GmailApp.search(buildCrmSearchQuery(fileIdentifier, true), 0, 5);
  let latestAttachment = null; 
  let latestMsgDate = 0;

  if (!threads.length) {
      threads = GmailApp.search(buildCrmSearchQuery(fileIdentifier, false), 0, 10);
  }
  
  if (!threads.length) throw new Error(`❌ CRM email not found matching query.`);
  
  const normalizedTarget = fileIdentifier.toLowerCase().replace(/[\s_-]/g, "");

  for (const thread of threads) {
    const messages = thread.getMessages();
    for (const msg of messages) {
      const msgDate = msg.getDate().getTime();
      const found = msg.getAttachments().find(a => { 
        const name = a.getName().toLowerCase(); 
        const normalizedName = name.replace(/[\s_-]/g, "");
        const isCsv = name.endsWith(".csv") || a.getContentType() === "text/csv";
        const validSize = a.getSize() > 500;
        return isCsv && validSize && normalizedName.includes(normalizedTarget) && !name.includes("backup") && !name.includes("copy") && !name.includes("archive"); 
      });
      if (found && msgDate > latestMsgDate) { latestAttachment = found; latestMsgDate = msgDate; }
    }
  }
  if (!latestAttachment) throw new Error(`❌ No valid CRM CSV attachment found containing: '${fileIdentifier}'`);
  
  if (Date.now() - latestMsgDate > MAX_FILE_AGE_MS) {
      throw new Error(`❌ CRM file is too old (older than 7 days). Aborting sync to prevent stale overwrite.`);
  }

  const fileName = latestAttachment.getName();

  let detectedCourseType = PropertiesService.getDocumentProperties().getProperty("COURSE_TYPE") || "L5"; 
  if (/L3/i.test(fileName)) { detectedCourseType = "L3"; } 
  else if (/L5/i.test(fileName)) { detectedCourseType = "L5"; }
  
  PropertiesService.getDocumentProperties().setProperty("COURSE_TYPE", detectedCourseType);
  applyCoursePreset(detectedCourseType);
  buildGlobals();

  const props = PropertiesService.getDocumentProperties(); const lastFile = props.getProperty("LAST_CRM_FILE");
  const fileKey = fileName + "_" + latestMsgDate;
  if (fileKey === lastFile) throw new Error("Duplicate CRM file detected. Sync aborted. Use 'Student Tracker > 🔓 Allow Duplicate CRM File' from the menu to force bypass.");
  props.setProperty("LAST_CRM_FILE", fileKey);
  
  const data = Utilities.parseCsv(latestAttachment.getDataAsString());
  
  let headerIdx = -1; 
  let normHeaders = [];
  for (let i = 0; i < Math.min(data.length, 30); i++) {
    const rowNorm = data[i].map(normHeader);
    if (rowNorm.includes(H.EMAIL) && rowNorm.includes(H.NAME)) { 
        headerIdx = i; normHeaders = rowNorm; break; 
    }
  }
  
  if (headerIdx === -1) {
      throw new Error(`❌ Fetched CSV does not contain 'Email' and 'Name' columns in the first 30 rows. Invalid CRM export format.`);
  }

  let generatedDate = "Not Found"; 
  for (let i = 0; i < headerIdx; i++) {
      const genString = data[i].join(" ");
      const dateMatch = genString.match(/(\d{2}\/\d{2}\/\d{4} \d{2}:\d{2} [AP]M)/);
      if (dateMatch) {
          const [d, m, y, h, min, mer] = dateMatch[0].split(/[\/\s:]/); let hours = parseInt(h);
          if (mer === "PM" && hours < 12) hours += 12; if (mer === "AM" && hours === 12) hours = 0;
          generatedDate = new Date(y, m - 1, d, hours, min);
          break;
      }
  }
  
  function parseUKDateToJS(value) {
    if (!value || typeof value !== "string" || value.trim() === "") return value;
    const parts = value.trim().split(" ")[0].split("/"); if (parts.length !== 3) return value;
    const [day, month, year] = parts.map(Number); const d = new Date(year, month - 1, day); d.setHours(0,0,0,0); return d;
  }
  
  const dateCols = normHeaders.map((hNorm, i) => ({ name: hNorm, index: i })).filter(h => h.name.includes("deadline") || h.name.includes("submission_deadline")).map(h => h.index);
  for (let i = headerIdx + 1; i < data.length; i++) { dateCols.forEach(c => { data[i][c] = parseUKDateToJS(data[i][c]); }); }
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  return { generatedDate, fileName: fileName };
}

function syncCrmToTracker() {
  verifySyncContext();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const crmSheet = getOrCreateSheet(CONFIG.SHEETS.CRM_RAW);
  const trackerSheet = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  
  const raw = crmSheet.getDataRange().getValues();
  let headerIdx = -1; let crmHeadersNorm = [];
  for (let i = 0; i < Math.min(raw.length, 30); i++) {
    const rowNorm = raw[i].map(normHeader);
    if (rowNorm.includes(H.EMAIL) && rowNorm.includes(H.NAME)) { headerIdx = i; crmHeadersNorm = rowNorm; break; }
  }
  if (headerIdx === -1) throw new Error("❌ Could not detect CRM headers properly.");
  const crmDataOnly = raw.slice(headerIdx); 
  
  const tracker = getSheetDataWithFormulas(trackerSheet);
  if (!tracker || tracker.length === 0 || !tracker[0]) throw new Error("❌ Student Tracker sheet is empty or corrupted.");
  
  if (tracker.length > 20 && crmDataOnly.length < (tracker.length * 0.2)) {
      throw new Error(`❌ CRM dataset is suspiciously small (${crmDataOnly.length} rows vs ${tracker.length} tracker rows). Aborting to prevent mass deletion.`);
  }
  
  const crmMap = getHeaderMap(CONFIG.SHEETS.CRM_RAW, crmHeadersNorm, false);
  const CRITICAL_SYNC_FIELDS = new Set([H.EMAIL, H.NAME, H.ID, H.COURSE_CODE]);
  CRITICAL_SYNC_FIELDS.forEach(f => { if (crmMap[f] === undefined) throw new Error(`❌ Critical CRM field missing: ${f}. Schema drift detected.`); });

  const seenEmails = new Set();
  for (let i = 1; i < crmDataOnly.length; i++) {
    const email = normEmail(crmDataOnly[i][crmMap[H.EMAIL]]);
    if (email && seenEmails.has(email)) throw new Error(`❌ Duplicate email in CRM: ${email}. Fix in Zoho before syncing.`);
    if (email) seenEmails.add(email);
  }

  const seenIds = new Set();
  for (let i = 1; i < crmDataOnly.length; i++) {
    const crmStudent = createRowModel(crmDataOnly[i], crmMap);
    const rawId = crmStudent.get(H.ID);
    const id = normId(rawId);
    if (id && seenIds.has(id)) { 
        throw new Error(`❌ Duplicate Record Id in CRM: '${rawId}' (Normalized: '${id}'). Please fix duplicates in Zoho before syncing.`); 
    }
    if (id) seenIds.add(id);
  }

  const originalTrackerSize = tracker.length - 1; 
  
  let bgColors = [];
  const lastRowTracker = trackerSheet.getLastRow(), lastColTracker = trackerSheet.getLastColumn();
  if (lastRowTracker > 1 && lastColTracker > 0) bgColors = trackerSheet.getRange(2, 1, lastRowTracker - 1, lastColTracker).getBackgrounds();
  
  const trackerHeaders = tracker[0];
  const trackerHeadersNorm = trackerHeaders.map(normHeader);
  const headerMap = getHeaderMap(CONFIG.SHEETS.TRACKER, trackerHeadersNorm, true);

  const CRM_COLUMNS = COLUMN_OWNERS.CRM;
  
  const manualColIndexes = [];
  COLUMN_OWNERS.MANUAL.forEach(field => {
      const idx = headerMap[field];
      if (idx !== undefined) manualColIndexes.push(idx);
  });

  CRM_COLUMNS.forEach(hNorm => { if (headerMap[hNorm] === undefined) throw new Error(`❌ Tracker missing CRM column: ${hNorm}. Please restore this column.`); });
  while (bgColors.length < tracker.length - 1) bgColors.push(new Array(trackerHeaders.length).fill("#ffffff"));

  const activeCrmKeys = new Set();
  for (let i = 1; i < crmDataOnly.length; i++) {
     const crmStudent = createRowModel(crmDataOnly[i], crmMap);
     let id = normId(crmStudent.get(H.ID));
     let emailStr = normEmail(crmStudent.get(H.EMAIL));
     if (id) activeCrmKeys.add("ID:" + id);
     if (emailStr) activeCrmKeys.add("EMAIL:" + emailStr);
  }

  let potentialRemovals = 0;
  for (let i = 1; i < tracker.length; i++) {
    const trackerStudent = createRowModel(tracker[i], headerMap);
    const id = normId(trackerStudent.get(H.ID));
    const emailStr = normEmail(trackerStudent.get(H.EMAIL));
    const idKey = id ? "ID:" + id : null;
    const emailKey = emailStr ? "EMAIL:" + emailStr : null;
    
    if (!id && (!emailStr || emailStr === "")) continue; 
    
    let existsInCRM = false;
    if (idKey) existsInCRM = activeCrmKeys.has(idKey);
    else if (emailKey) existsInCRM = activeCrmKeys.has(emailKey);
    
    if (!existsInCRM) potentialRemovals++;
  }
  
  const removalRatio = originalTrackerSize > 0 ? (potentialRemovals / originalTrackerSize) : 0;
  let skipDeletions = false;
  const FORCE_DELETE = PropertiesService.getScriptProperties().getProperty("ALLOW_AUTO_DELETE") === "true";

  if (originalTrackerSize > 10 && removalRatio > 0.30 && !FORCE_DELETE) {
      throw new Error(`❌ CRITICAL: Mass deletion of ${potentialRemovals} students (${Math.round(removalRatio * 100)}%) detected. Hard abort triggered to prevent catastrophic data loss. Set 'ALLOW_AUTO_DELETE' to 'true' in script properties to override.`);
  }

  if (originalTrackerSize > 10 && (potentialRemovals > 20 || removalRatio > 0.15)) {
    try {
      const activeUser = Session.getActiveUser().getEmail();
      if (!activeUser) {
          if (!FORCE_DELETE) skipDeletions = true;
      } else {
          const ui = SpreadsheetApp.getUi();
          const response = ui.alert(
            "⚠️ Mass Deletion Warning",
            `The CRM sync is about to remove ${potentialRemovals} students (${Math.round(removalRatio * 100)}% of your tracker).\n\nThis is completely normal if a cohort just graduated. However, if your CRM CSV export was incomplete or filtered by mistake, you could lose data.\n\nDo you want to PROCEED and permanently delete these ${potentialRemovals} students from the tracker?`,
            ui.ButtonSet.YES_NO
          );
          
          if (response !== ui.Button.YES) {
            skipDeletions = true;
            SpreadsheetApp.getActiveSpreadsheet().toast(`Skipped deleting ${potentialRemovals} students to protect data.`, "Safety Abort", 8);
          }
      }
    } catch (e) {
      if (!FORCE_DELETE) skipDeletions = true;
    }
  }
  
  const existingById = {}, existingByEmail = {};
  const auditLogs = [];
  const logDate = new Date(); 
  
  for (let i = 1; i < tracker.length; i++) {
    const trackerStudent = createRowModel(tracker[i], headerMap);
    const id = normId(trackerStudent.get(H.ID)), emailStr = normEmail(trackerStudent.get(H.EMAIL));
    if (id) existingById[id] = i;
    if (emailStr) existingByEmail[emailStr] = i;
  }
  
  let added = 0, updated = 0, skippedNoId = 0, unchanged = 0;
  let missingFromExportCount = 0; const missingFromExportNames = {};
  const trackSkip = (field) => { missingFromExportCount++; missingFromExportNames[field] = (missingFromExportNames[field] || 0) + 1; };
  
  let anyChange = false; let processedRows = 0; 
  const forceWriteRows = new Set();
  const nextTracker = tracker.map(r => r.slice());
  const nextBgColors = bgColors.map(r => r.slice());
  
  for (let i = 1; i < crmDataOnly.length; i++) {
      shouldStop(); 
      
      const crmStudent = createRowModel(crmDataOnly[i], crmMap);
      let rawId = crmStudent.get(H.ID);
      let id = normId(rawId);
      let emailStr = normEmail(crmStudent.get(H.EMAIL));
      
      if ((!emailStr || emailStr === "") && (!id || id === "")) { skippedNoId++; continue; }
      processedRows++; 
    
    let rowIndex;
    if (id) rowIndex = existingById[id];
    
    if (rowIndex === undefined && emailStr) {
      const possible = existingByEmail[emailStr];
      if (possible !== undefined) {
        const existingRow = createRowModel(nextTracker[possible], headerMap);
        const existingId = normId(existingRow.get(H.ID));
        if (!existingId) { rowIndex = possible; } 
        else if (existingId && existingId !== id) { 
           // Never blindly overwrite an existing ID during identity conflict resolution
           rowIndex = possible;
        } 
        else if (existingId === id) { rowIndex = possible; } 
      }
    }
      
    if (rowIndex !== undefined && emailStr) existingByEmail[emailStr] = rowIndex;
    if (rowIndex !== undefined && id) existingById[id] = rowIndex; 
    
    if (rowIndex !== undefined && rowIndex !== null && !isNaN(rowIndex)) {
      const origRowCopy = nextTracker[rowIndex].slice();
      const origHash = fastHash(origRowCopy);
      const trackerStudent = createRowModel(nextTracker[rowIndex], headerMap);
      
      let criticalMismatch = false;
      const existingRawId = trackerStudent.get(H.ID);
      const existingNormId = normId(existingRawId);
      
      if (isDifferent(existingNormId, id) || isDifferent(existingRawId, id)) {
        if (!existingNormId && id) {
           trackerStudent.set(H.ID, id);
           criticalMismatch = true;
        } else {
           throw new Error(`❌ Identity Conflict: Attempted to overwrite existing ID '${existingNormId}' with new ID '${id}' for Email '${emailStr}'. Fix in Zoho CRM before syncing.`);
        }
      }
      
      const existingEmail = trackerStudent.get(H.EMAIL);
      if (isDifferent(existingEmail, emailStr)) {
        trackerStudent.set(H.EMAIL, emailStr);
        criticalMismatch = true;
      }
      
      CRM_COLUMNS.forEach(hNorm => {
        if (!CRM_COLUMNS.has(hNorm)) return;
        if (hNorm === H.PATHWAY || hNorm === H.PROJ_SUB || hNorm === H.AUTO_RISK_REASON || hNorm === H.RISK_NOTES || hNorm === H.ID || hNorm === H.EMAIL) return;
        const tIdx = headerMap[hNorm];
        if (tIdx === undefined) return;
        if (!crmStudent.has(hNorm)) {
          if (CRITICAL_SYNC_FIELDS.has(hNorm)) throw new Error(`❌ Critical field missing in CRM: ${hNorm}. Schema drift detected.`);
          trackSkip(hNorm); return;
        }
        let val = crmStudent.get(hNorm);
        if (val === undefined || val === null) val = "";
        if (typeof val === "string") val = val.trim();
        
        if (hNorm === H.NAME) {
          const recordId = crmStudent.get(H.ID);
          if (val && recordId) {
            const newFormula = applyCrmProfileLink(val, recordId);
            const currentVal = trackerStudent.get(hNorm);
            if (isDifferent(currentVal, newFormula)) { val = newFormula; } 
            else { val = currentVal; }
          }
        }
        
        if (hNorm.includes(H.SUB_STATUS)) val = mapStatus(val);
        if (hNorm === H.DISC_ID && val && val.toString().trim() !== "") val = applyDiscordLink(val);
        
        trackerStudent.safeSet(hNorm, val);
      }); 
      
      if (FEATURES.PROJECT_TRACKING) {
        let pathwayCalc = 4;
        
        if (FEATURES.SPECIALISATION) {
            const p5Deadline = parseRobustDate(trackerStudent.get(H.P5_DEADLINE));
            const specDeadline = parseRobustDate(trackerStudent.get(H.SPEC_DEADLINE));
            if ((p5Deadline instanceof Date && !isNaN(p5Deadline.getTime())) || (specDeadline instanceof Date && !isNaN(specDeadline.getTime()))) pathwayCalc = 5;
        }
        
        let submittedCountCalc = 0;
        for (const p of CONFIG.PROJECTS) {
           const status = trackerStudent.get(normHeader(`${p}_submission_status`));
           if (status && CONFIG.RISK.SUBMITTED_STATUSES.some(s => status.toString().toLowerCase().trim().includes(s))) submittedCountCalc++;
        }
        
        if (trackerStudent.has(H.PATHWAY)) trackerStudent.safeSet(H.PATHWAY, pathwayCalc);
        if (trackerStudent.has(H.PROJ_SUB)) trackerStudent.safeSet(H.PROJ_SUB, submittedCountCalc);
      }
      
      if (nextTracker[rowIndex]) {
        manualColIndexes.forEach(tIdx => {
          nextTracker[rowIndex][tIdx] = tracker[rowIndex][tIdx];
        });
      }

      const newHash = fastHash(nextTracker[rowIndex]);
      const rowActuallyChanged = origHash !== newHash;
      
      if (rowActuallyChanged || criticalMismatch) {
        updated++;
        logDiff(origRowCopy, nextTracker[rowIndex], headerMap, emailStr || id);
        auditLogs.push([ logDate, trackerStudent.get(H.NAME) || "Unknown", trackerStudent.get(H.COURSE_CODE) || "Unknown", "UPDATED" ]);
        forceWriteRows.add(rowIndex);
        anyChange = true;
      } else {
        unchanged++;
      }

    } else {
      const newRow = new Array(trackerHeaders.length).fill("");
      const newTrackerStudent = createRowModel(newRow, headerMap);
      
      newTrackerStudent.set(H.ID, id);
      newTrackerStudent.set(H.EMAIL, emailStr);
      
      CRM_COLUMNS.forEach(hNorm => {
        if (!CRM_COLUMNS.has(hNorm)) return;
        const tIdx = headerMap[hNorm];
        if (tIdx === undefined) return;
        if (hNorm === H.PATHWAY || hNorm === H.PROJ_SUB || hNorm === H.ID || hNorm === H.EMAIL) return; 
        if (!crmStudent.has(hNorm)) { trackSkip(hNorm); return; }
        
        let val = crmStudent.get(hNorm);
        if (val === undefined || val === null) val = "";
        
        if (hNorm === H.NAME) {
          const recordId = crmStudent.get(H.ID);
          if (val && recordId) { val = applyCrmProfileLink(val, recordId); }
        }
        
        if (hNorm.includes(H.SUB_STATUS)) val = mapStatus(val);
        if (hNorm === H.DISC_ID && val && val.toString().trim() !== "") val = applyDiscordLink(val);
        newTrackerStudent.set(hNorm, val);
      });
      
      if (FEATURES.PROJECT_TRACKING) {
        let pathwayCalc = 4;
        if (FEATURES.SPECIALISATION) {
            const p5Deadline = parseRobustDate(newTrackerStudent.get(H.P5_DEADLINE));
            const specDeadline = parseRobustDate(newTrackerStudent.get(H.SPEC_DEADLINE));
            if ((p5Deadline instanceof Date && !isNaN(p5Deadline.getTime())) || (specDeadline instanceof Date && !isNaN(specDeadline.getTime()))) pathwayCalc = 5;
        }
        
        let submittedCountCalc = 0;
        for (const p of CONFIG.PROJECTS) {
           const status = newTrackerStudent.get(normHeader(`${p}_submission_status`));
           if (status && CONFIG.RISK.SUBMITTED_STATUSES.some(s => status.toString().toLowerCase().trim().includes(s))) submittedCountCalc++;
        }
        
        if (newTrackerStudent.has(H.PATHWAY)) newTrackerStudent.set(H.PATHWAY, pathwayCalc);
        if (newTrackerStudent.has(H.PROJ_SUB)) newTrackerStudent.set(H.PROJ_SUB, submittedCountCalc);
      }
      
      nextTracker.push(newTrackerStudent.raw);
      const newRowIndex = nextTracker.length - 1;
      
      forceWriteRows.add(newRowIndex);
      nextBgColors.push(new Array(trackerHeaders.length).fill("#ffffff"));
      added++;
      anyChange = true;
      auditLogs.push([ logDate, crmStudent.get(H.NAME) || "Unknown", crmStudent.get(H.COURSE_CODE) || "Unknown", "ADDED" ]);
    }
  }

  let removedFromCrm = 0; let removedNoId = 0; const rowsToDelete = [];
  
  for (let i = nextTracker.length - 1; i >= 1; i--) {
    const trackerStudent = createRowModel(nextTracker[i], headerMap);
    const id = normId(trackerStudent.get(H.ID));
    const emailStr = normEmail(trackerStudent.get(H.EMAIL));
    const idKey = id ? "ID:" + id : null;
    const emailKey = emailStr ? "EMAIL:" + emailStr : null;
    
    if (!id && (!emailStr || emailStr === "")) {
      continue; 
    }
    
    let existsInCRM = false;
    if (idKey) existsInCRM = activeCrmKeys.has(idKey);
    else if (emailKey) existsInCRM = activeCrmKeys.has(emailKey);
    
    if (!existsInCRM) {
      auditLogs.push([ logDate, trackerStudent.get(H.NAME) || "Unknown", trackerStudent.get(H.COURSE_CODE) || "Unknown", "REMOVED" ]);
      if (!skipDeletions) rowsToDelete.push({ row: i + 1, id: id, email: emailStr }); 
      removedFromCrm++; anyChange = true;
    }
  }

  const deleteSet = new Set(rowsToDelete.map(t => t.row));
  
  if (!anyChange && forceWriteRows.size === 0 && rowsToDelete.length === 0) {
    return { added, updated, unchanged, removedFromCrm, removedNoId, aborted: false, syncTime: new Date() };
  }

  let rowsToWrite = [];
  for (let i = 1; i < nextTracker.length; i++) {
    const sheetRow = i + 1;
    if (forceWriteRows.has(i) && !deleteSet.has(sheetRow)) {
      const safeBg = nextBgColors[i - 1] ? nextBgColors[i - 1].slice() : new Array(trackerHeaders.length).fill("#ffffff");
      rowsToWrite.push({ row: sheetRow, values: nextTracker[i].slice(), bg: safeBg });
    }
  }

  while (nextBgColors.length < nextTracker.length - 1) { nextBgColors.push(new Array(trackerHeaders.length).fill("#ffffff")); }

  shouldStop(); 
  
  if (rowsToWrite.length > 0 || rowsToDelete.length > 0) {
    executeSecureWrite("CRM_WRITE", () => {
      if (rowsToWrite.length > 0) writeRowsInBatches(trackerSheet, rowsToWrite, headerMap, manualColIndexes);
      if (rowsToDelete.length > 0) {
        try { 
            const idCol = headerMap[H.ID] !== undefined ? headerMap[H.ID] + 1 : null;
            const emailCol = headerMap[H.EMAIL] !== undefined ? headerMap[H.EMAIL] + 1 : null;
            
            rowsToDelete.sort((a,b) => b.row - a.row).forEach(target => {
                if (idCol && emailCol) {
                    const liveId = normId(trackerSheet.getRange(target.row, idCol).getValue());
                    const liveEmail = normEmail(trackerSheet.getRange(target.row, emailCol).getValue());
                    
                    if ((target.id && liveId !== target.id) || (target.email && liveEmail !== target.email)) {
                        console.warn(`⚠️ Race condition averted: Row ${target.row} shifted mid-run. Skipped deletion.`);
                        return;
                    }
                }
                trackerSheet.deleteRow(target.row); 
            }); 
        } 
        catch (e) { throw new Error(`❌ CRITICAL: Row deletion failed. Error: ${e.message}`); }
      }
    });
  }
  
  if (auditLogs.length > 0) {
      let auditSheet = getOrCreateSheet(CONFIG.SHEETS.AUDIT_LOG);
      if (auditSheet.getLastRow() === 0) { auditSheet.appendRow(["Date", "Full Name", "Course Code", "Status"]); auditSheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f3f3"); }
      auditSheet.getRange(auditSheet.getLastRow() + 1, 1, auditLogs.length, 4).setValues(auditLogs);
      auditSheet.getRange(2, 1, auditSheet.getLastRow(), 1).setNumberFormat("dd/MM/yyyy HH:mm:ss");
  }
  return { added, updated, unchanged, removedFromCrm, removedNoId, aborted: false, syncTime: new Date() };
}

// ================= CYPHER LMS SYNC =================
function runLmsSync() {
  verifySyncContext();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trackerSheet = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  const lmsSheet = getOrCreateSheet(CONFIG.SHEETS.LMS_RAW);
  
  const lmsData = lmsSheet.getDataRange().getDisplayValues();
  if (lmsData.length <= 1 || (lmsData.length === 1 && lmsData[0][0].toString().includes("Paste Cypher"))) return { missingData: true };
  
  const lmsHeadersNorm = lmsData[0].map(normHeader);
  const lmsHeaderMap = getHeaderMap(CONFIG.SHEETS.LMS_RAW, lmsHeadersNorm, false);
  
  if (lmsHeaderMap[normHeader("email")] === undefined && lmsHeaderMap[H.ID] === undefined) {
      showProgress(ss, "❌ No valid identifier (Email or Learner ID) found in Cypher data.", "Error", 5);
      return { missingData: true };
  }
  
  const trackerData = getSheetDataWithFormulas(trackerSheet);
  const headers = trackerData[0]; 
  const headersNorm = headers.map(normHeader);
  const headerMap = getHeaderMap(CONFIG.SHEETS.TRACKER, headersNorm, true);
  
  if (headerMap[H.EMAIL] === undefined) throw new Error(`❌ Student Tracker missing '${CONFIG.CORE_FIELDS.EMAIL}' column.`);

  let bgColors = [];
  if (trackerSheet.getLastRow() > 1 && trackerSheet.getLastColumn() > 0) {
    bgColors = trackerSheet.getRange(2, 1, trackerSheet.getLastRow() - 1, trackerSheet.getLastColumn()).getBackgrounds();
  }
  while (bgColors.length < trackerData.length - 1) bgColors.push(new Array(headers.length).fill("#ffffff"));
  
  const forceWriteRows = new Set();
  const lmsMapById = {};
  const lmsMapByEmail = {};
  
  for (let i = 1; i < lmsData.length; i++) {
    const lmsStudent = createRowModel(lmsData[i], lmsHeaderMap);
    const emailStr = normEmail(lmsStudent.get(normHeader("email")));
    const idStr = normId(lmsStudent.get(H.ID));
    if (!emailStr && !idStr) continue;
    
    let d1 = parseRobustDate(lmsStudent.get("last visited at") || lmsStudent.get(normHeader('visited')));
    let d2 = parseRobustDate(lmsStudent.get(normHeader('last login')));
    let best = (d1 instanceof Date && d2 instanceof Date) ? (d1 > d2 ? d1 : d2) : (d1 instanceof Date ? d1 : (d2 instanceof Date ? d2 : null));
    
    const rProgRaw = lmsStudent.get(normHeader('progress')) ?? lmsStudent.get(normHeader('completion'));
    const rProg = rProgRaw !== undefined && rProgRaw !== "" ? rProgRaw : 0;
    const rLesRaw = lmsStudent.get(normHeader('current lesson'));
    const rLes = rLesRaw !== undefined && rLesRaw !== "" ? rLesRaw : ""; 
    let pDate = best ? best : NEVER; 
    
    const payload = { date: pDate, progress: rProg, lesson: rLes };
    
    if (idStr) {
      if (!lmsMapById[idStr]) lmsMapById[idStr] = payload;
      else {
        let ex = lmsMapById[idStr].date;
        if (pDate instanceof Date && (isNever(ex) || !ex || pDate > ex)) lmsMapById[idStr] = payload;
      }
    }
    if (emailStr) {
      if (!lmsMapByEmail[emailStr]) lmsMapByEmail[emailStr] = payload;
      else {
        let ex = lmsMapByEmail[emailStr].date;
        if (pDate instanceof Date && (isNever(ex) || !ex || pDate > ex)) lmsMapByEmail[emailStr] = payload;
      }
    }
  }

  let updatedCount = 0, unchangedCount = 0;
  let processedRows = 0; 
  let anyChange = false;
  
  const unmatched = [["Generated At", new Date()], ["", ""], ["Unmatched Email / ID", "Student Name"]]; 
  const unmatchedSet = new Set();
  const nextTrackerData = trackerData.map(r => r.slice());
  
  const manualColIndexes = [];
  COLUMN_OWNERS.MANUAL.forEach(field => {
      const idx = headerMap[field];
      if (idx !== undefined) manualColIndexes.push(idx);
  });

  for (let i = 1; i < nextTrackerData.length; i++) {
    shouldStop(); 
    const trackerStudent = createRowModel(nextTrackerData[i], headerMap);
    const emailStr = normEmail(trackerStudent.get(H.EMAIL));
    const idStr = normId(trackerStudent.get(H.ID));
    if (!emailStr && !idStr) continue;
    
    processedRows++;
    const origRowCopy = nextTrackerData[i].slice();
    const origHash = fastHash(origRowCopy);
    let data = null;
    
    if (idStr && lmsMapById[idStr]) {
       data = lmsMapById[idStr];
       if (emailStr && lmsMapByEmail[emailStr] && lmsMapById[idStr] !== lmsMapByEmail[emailStr]) {
           if (DEBUG_LOGS) console.warn(`⚠️ LMS Match: ID and Email map to different records for ${emailStr}. Trusting ID.`);
       }
    } else if (emailStr && lmsMapByEmail[emailStr]) {
       data = lmsMapByEmail[emailStr];
       if (idStr && !lmsMapById[idStr]) {
           if (DEBUG_LOGS) console.warn(`⚠️ LMS Match: Matched by Email for ${emailStr}, but Learner ID ${idStr} was not found in export.`);
       }
    }
    
    if (data) {
       if (trackerStudent.has(H.LMS_ACT)) {
         let newDate = isNever(data.date) ? NEVER : data.date;
         trackerStudent.safeSet(H.LMS_ACT, newDate);
       }
       if (trackerStudent.has(H.LMS_PROG)) {
         let progVal = data.progress;
         if (typeof progVal === "string") progVal = progVal.replace(",", ".").replace(/[^\d.]/g, "");
         let num = (progVal !== "" && !isNaN(parseFloat(progVal))) ? parseFloat(progVal) : 0;
         let parsedProg = num > 1 ? num / 100 : num; 
         trackerStudent.safeSet(H.LMS_PROG, parsedProg);
       }
       if (trackerStudent.has(H.LMS_LES)) {
         trackerStudent.safeSet(H.LMS_LES, data.lesson);
       }
    } else {
       const stuName = trackerStudent.has(H.NAME) ? trackerStudent.get(H.NAME) : "Unknown";
       const unmatchKey = idStr ? `ID: ${idStr}` : emailStr;
       if (!unmatchedSet.has(unmatchKey)) {
           unmatchedSet.add(unmatchKey);
           unmatched.push([unmatchKey, stuName]);
       }
       
       if (trackerStudent.has(H.LMS_ACT)) {
         const currentAct = trackerStudent.get(H.LMS_ACT);
         if ((currentAct === undefined || currentAct === null || currentAct === "") && !isNever(currentAct)) {
           trackerStudent.safeSet(H.LMS_ACT, NEVER);
         }
       }
    }

    if (trackerStudent.has(H.LMS_DAYS)) {
      const days = calculateDaysSince(trackerStudent.get(H.LMS_ACT));
      if (days !== null) {
          trackerStudent.safeSet(H.LMS_DAYS, days);
      } else {
          let fallback = isNever(trackerStudent.get(H.LMS_ACT)) ? NEVER : "";
          trackerStudent.safeSet(H.LMS_DAYS, fallback);
      }
    }

    if (nextTrackerData[i]) {
      manualColIndexes.forEach(tIdx => {
        nextTrackerData[i][tIdx] = trackerData[i][tIdx];
      });
    }

    const newHash = fastHash(nextTrackerData[i]);
    const rowActuallyChanged = origHash !== newHash;

    if (rowActuallyChanged) {
      updatedCount++; forceWriteRows.add(i); anyChange = true; logDiff(origRowCopy, nextTrackerData[i], headerMap, emailStr || idStr);
    } else { unchangedCount++; }
  }
  
  if (processedRows === 0 && lmsData.length > 1) {
    throw new Error("❌ CRITICAL: No LMS rows successfully matched. Aborting to prevent blank overwrite.");
  }
  
  if (!anyChange && forceWriteRows.size === 0 && processedRows > 0) {
    return { missingData: false, updated: updatedCount, unchanged: unchangedCount, unmatched: Math.max(0, unmatched.length - 3), aborted: false, syncTime: new Date() };
  }
  
  let rowsToWrite = [];
  for (let i = 1; i < nextTrackerData.length; i++) {
    const sheetRow = i + 1;
    if (forceWriteRows.has(i)) {
      const safeBg = bgColors[i - 1] ? bgColors[i - 1].slice() : new Array(headers.length).fill("#ffffff");
      rowsToWrite.push({ row: sheetRow, values: nextTrackerData[i].slice(), bg: safeBg });
    }
  }
  
  while (bgColors.length < nextTrackerData.length - 1) bgColors.push(new Array(headers.length).fill("#ffffff"));
  
  shouldStop(); 
  
  if (rowsToWrite.length > 0) {
    executeSecureWrite("LMS_WRITE", () => {
      writeRowsInBatches(trackerSheet, rowsToWrite, headerMap, manualColIndexes);
      const progColIdx = headerMap[H.LMS_PROG];
      if (progColIdx !== undefined && trackerData.length > 1) trackerSheet.getRange(2, progColIdx + 1, trackerData.length - 1, 1).setNumberFormat("0.00%");
    });
  }
  
  const MAX_UNMATCHED_RECORDS = 500;
  let unmatchedToWrite = unmatched;
  
  if (unmatchedToWrite.length > MAX_UNMATCHED_RECORDS + 3) {
      unmatchedToWrite = unmatchedToWrite.slice(0, MAX_UNMATCHED_RECORDS + 3);
  }

  if (unmatchedToWrite.length > 3) {
    let unSheet = getOrCreateSheet(CONFIG.SHEETS.UNMATCHED);
    let lastRow = unSheet.getLastRow(); let lastCol = unSheet.getLastColumn();
    let manualDataMap = {}; let manualHeaders = [];

    if (lastCol > 2 && lastRow >= 3) {
      manualHeaders = unSheet.getRange(3, 3, 1, lastCol - 2).getValues()[0];
      if (lastRow > 3) {
        let existingData = unSheet.getRange(4, 1, lastRow - 3, lastCol).getValues();
        existingData.forEach(row => { if (row[0]) manualDataMap[row[0].toString().trim().toLowerCase()] = row.slice(2); });
      }
    }

    for (let i = 3; i < unmatchedToWrite.length; i++) {
      let emailKey = (unmatchedToWrite[i][0] || "").toString().trim().toLowerCase();
      let manualCols = manualDataMap[emailKey] || new Array(manualHeaders.length).fill("");
      unmatchedToWrite[i] = unmatchedToWrite[i].concat(manualCols);
    }
    unmatchedToWrite[2] = unmatchedToWrite[2].concat(manualHeaders); unmatchedToWrite[0] = unmatchedToWrite[0].concat(new Array(manualHeaders.length).fill("")); unmatchedToWrite[1] = unmatchedToWrite[1].concat(new Array(manualHeaders.length).fill(""));

    unSheet.getRange(1, 1, unmatchedToWrite.length, unmatchedToWrite[0].length).clearContent();
    unSheet.getRange(1, 1, unmatchedToWrite.length, unmatchedToWrite[0].length).setValues(unmatchedToWrite);
    unSheet.getRange("A3:B3").setFontWeight("bold").setBackground("#f3f3f3");
    if (manualHeaders.length > 0) unSheet.getRange(3, 3, 1, manualHeaders.length).setFontWeight("bold").setBackground("#f3f3f3");
    unSheet.setColumnWidth(1, 250); unSheet.setColumnWidth(2, 200);
  }
  return { missingData: false, updated: updatedCount, unchanged: unchangedCount, unmatched: Math.max(0, unmatchedToWrite.length - 3), aborted: false, syncTime: new Date() };
}

// ================= GITHUB =================
function updateGithubActivity() {
  verifySyncContext();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateSheet(CONFIG.SHEETS.TRACKER);
  const data = getSheetDataWithFormulas(sheet);
  if (data.length <= 1) return null;
  
  const headers = data[0];
  const headersNorm = headers.map(normHeader);
  const headerMap = getHeaderMap(CONFIG.SHEETS.TRACKER, headersNorm, true);
  
  let bgColors = [];
  if (sheet.getLastRow() > 1 && sheet.getLastColumn() > 0) bgColors = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getBackgrounds();
  while (bgColors.length < data.length - 1) bgColors.push(new Array(headers.length).fill("#ffffff"));
  
  const token = PropertiesService.getScriptProperties().getProperty("GITHUB_TOKEN");
  if (!token) { showProgress(ss, "❌ GITHUB_TOKEN missing!", "Setup Error", 5); return null; }
  const apiHeaders = { Accept: "application/vnd.github.cloak-preview", Authorization: `token ${token}` };
  const props = PropertiesService.getScriptProperties();
  
  const MAX_CACHE_SIZE = 500;
  
  let cacheStr = props.getProperty("GITHUB_CACHE") || "{}"; let cache; try { cache = JSON.parse(cacheStr); } catch(e) { cache = {}; }
  let actCacheStr = props.getProperty("GITHUB_ACT_CACHE") || "{}"; let actCache; try { actCache = JSON.parse(actCacheStr); } catch(e) { actCache = {}; }
  
  enforceCacheLimit(cache, MAX_CACHE_SIZE);
  enforceCacheLimit(actCache, MAX_CACHE_SIZE);
  
  const THIRTY_DAYS = 30 * 24 * 60 * 60 * 1000; const TWELVE_HOURS = 12 * 60 * 60 * 1000; const now = Date.now();
  let searchCount = 0, newFoundCount = 0, updatedCount = 0, unchangedCount = 0; const MAX_SEARCHES_PER_RUN = 10; 
  
  let anyChange = false;
  let processedRows = 0;
  const forceWriteRows = new Set();
  const nextData = data.map(r => r.slice());
  
  const manualColIndexes = [];
  COLUMN_OWNERS.MANUAL.forEach(field => {
      const idx = headerMap[field];
      if (idx !== undefined) manualColIndexes.push(idx);
  });

  for (let i = 1; i < nextData.length; i++) {
    shouldStop();
    
    const trackerStudent = createRowModel(nextData[i], headerMap);
    const emailStr = normEmail(trackerStudent.get(H.EMAIL));
    const idStr = normId(trackerStudent.get(H.ID));
    
    if (!emailStr) continue;
    processedRows++;
    
    const origRowCopy = nextData[i].slice();
    const origHash = fastHash(origRowCopy);

    let currentUname = trackerStudent.get(H.GH_UNAME);
    let cacheKey = idStr || emailStr;
    let cData = cache[cacheKey];
    let cVal = null, cTs = 0;
    
    if (cData && typeof cData === "object" && cData.ts) { 
        cVal = cData.val; cTs = cData.ts; 
        cData.ts = now; 
    } else if (cData) { 
        cVal = cData?.val || cData; cTs = now; cache[cacheKey] = { val: cVal, ts: cTs }; 
    }
    
    if (!currentUname && cVal && cVal !== "NOT_FOUND") {
       trackerStudent.safeSet(H.GH_UNAME, cVal);
       currentUname = cVal;
    }
    
    const isNotFoundExpired = (cVal === "NOT_FOUND") && (now - cTs > THIRTY_DAYS);
    if (!currentUname && (!cVal || isNotFoundExpired) && searchCount < MAX_SEARCHES_PER_RUN) { 
      searchCount++; let delay = 300;
      for (let attempt = 0; attempt < 3; attempt++) {
        try {
          const res = UrlFetchApp.fetch(`https://api.github.com/search/commits?q=author-email:${encodeURIComponent(emailStr)}&per_page=1`, { headers: apiHeaders, muteHttpExceptions: true });
          if (res.getResponseCode() === 200) {
            const found = JSON.parse(res.getContentText()).items?.[0]?.author?.login;
            if (found) {
              trackerStudent.safeSet(H.GH_UNAME, found);
              cache[cacheKey] = { val: found, ts: now }; newFoundCount++;
            } else { cache[cacheKey] = { val: "NOT_FOUND", ts: now }; }
            break;
          } else if (res.getResponseCode() === 403 || res.getResponseCode() === 429) { 
              const rHeaders = res.getHeaders();
              const remaining = rHeaders['X-RateLimit-Remaining'] || rHeaders['x-ratelimit-remaining'];
              const reset = rHeaders['X-RateLimit-Reset'] || rHeaders['x-ratelimit-reset'];
              if (remaining === '0' && reset) {
                  const waitTime = (parseInt(reset) * 1000) - Date.now();
                  if (waitTime > 0 && waitTime < 60000) { Utilities.sleep(waitTime + 1000); } 
                  else { throw new Error("GitHub API Rate limit exhausted. Try again later."); }
              } else {
                  Utilities.sleep(1000 + Math.random() * 1000);
                  delay *= 2;
              }
          } else break;
        } catch(e) { Utilities.sleep(delay); delay *= 2; }
      }
    }
    
    if (nextData[i]) {
      manualColIndexes.forEach(tIdx => {
        nextData[i][tIdx] = data[i][tIdx];
      });
    }

    const newHash = fastHash(nextData[i]);
    const rowActuallyChanged = origHash !== newHash;
    if (rowActuallyChanged) { anyChange = true; forceWriteRows.add(i); }
  }
  
  let requests = [], map = [];
  for (let i = 1; i < nextData.length; i++) {
    const trackerStudent = createRowModel(nextData[i], headerMap);
    let uname = trackerStudent.get(H.GH_UNAME);
    const origRowCopy = nextData[i].slice();
    const origHash = fastHash(origRowCopy);
    
    if (uname && uname.toString().trim() !== "") {
      uname = uname.toString().toLowerCase().replace(/^https?:\/\/github\.com\//, "").split(/[/?#]/)[0].replace("@", "").trim();
      trackerStudent.safeSet(H.GH_UNAME, uname);
      
      let aData = actCache[uname];
      if (aData && typeof aData === "object" && aData.ts && (now - aData.ts < TWELVE_HOURS)) {
        let actVal = aData.val;
        if (actVal !== "No recent activity (90d)" && actVal !== "User Not Found") actVal = new Date(actVal);
        trackerStudent.safeSet(H.GH_ACT, actVal);
        const newLink = `=HYPERLINK("https://github.com/${uname}", "${uname}")`;
        trackerStudent.safeSet(H.GH_PROF, newLink);
        aData.ts = now; 
      } else {
        requests.push({ url: `https://api.github.com/users/${uname}/events/public?per_page=1`, headers: apiHeaders, muteHttpExceptions: true });
        map.push({ rowIdx: i, uname: uname });
      }
    }
    
    if (nextData[i]) {
      manualColIndexes.forEach(tIdx => {
        nextData[i][tIdx] = data[i][tIdx];
      });
    }

    const newHash = fastHash(nextData[i]);
    const rowActuallyChanged = origHash !== newHash;
    
    if (rowActuallyChanged) { updatedCount++; anyChange = true; forceWriteRows.add(i); logDiff(origRowCopy, nextData[i], headerMap, uname || "Unknown"); } 
    else unchangedCount++;
  }
  
  if (requests.length > 0) {
    const chunkSize = 30; let quotaExceeded = false; 
    for (let i = 0; i < requests.length; i += chunkSize) {
      if (quotaExceeded) break;
      shouldStop();
      
      const chunkReqs = requests.slice(i, i + chunkSize), chunkMap = map.slice(i, i + chunkSize);
      try {
        const responses = UrlFetchApp.fetchAll(chunkReqs);
        responses.forEach((res, idx) => {
          const rIdx = chunkMap[idx].rowIdx; const uname = chunkMap[idx].uname; const code = res.getResponseCode();
          const origRowCopy = nextData[rIdx].slice(); 
          const origHash = fastHash(origRowCopy);
          const trackerStudent = createRowModel(nextData[rIdx], headerMap);
          if (code === 200) {
            try {
              const events = JSON.parse(res.getContentText());
              let valToCache = "No recent activity (90d)";
              if (events.length > 0) { 
                const newVal = new Date(events[0].created_at); valToCache = events[0].created_at;
                trackerStudent.safeSet(H.GH_ACT, newVal);
              } else { trackerStudent.safeSet(H.GH_ACT, valToCache); }
              const newFormula = `=HYPERLINK("https://github.com/${uname}", "${uname}")`;
              trackerStudent.safeSet(H.GH_PROF, newFormula);
              actCache[uname] = { val: valToCache, ts: now };
            } catch(e) {}
          } else if (code === 404) {
             trackerStudent.safeSet(H.GH_ACT, "User Not Found");
             actCache[uname] = { val: "User Not Found", ts: now };
          }
          
          if (nextData[rIdx]) {
            manualColIndexes.forEach(tIdx => {
              nextData[rIdx][tIdx] = data[rIdx][tIdx];
            });
          }

          const newHash = fastHash(nextData[rIdx]);
          const rowActuallyChanged = origHash !== newHash;
          
          if (rowActuallyChanged) { updatedCount++; unchangedCount--; anyChange = true; forceWriteRows.add(rIdx); }
        });
      } catch (e) { 
          if (e.message.includes("Bandwidth") || e.message.includes("quota")) quotaExceeded = true; 
          Utilities.sleep(1000 + Math.random() * 1000);
      }
    }
  }
  
  if (processedRows === 0 && data.length > 1) {
    throw new Error("❌ CRITICAL: No GitHub rows successfully matched. Aborting to prevent blank overwrite.");
  }
  
  if (!anyChange && forceWriteRows.size === 0 && processedRows > 0) {
    return { newFoundCount: newFoundCount, updated: updatedCount, unchanged: unchangedCount, aborted: false, syncTime: new Date() };
  }
  
  let rowsToWrite = [];
  for (let i = 1; i < nextData.length; i++) {
    const sheetRow = i + 1;
    if (forceWriteRows.has(i)) { 
      const safeBg = bgColors[i - 1] ? bgColors[i - 1].slice() : new Array(headers.length).fill("#ffffff");
      rowsToWrite.push({ row: sheetRow, values: nextData[i].slice(), bg: safeBg }); 
    }
  }
  
  while (bgColors.length < nextData.length - 1) bgColors.push(new Array(headers.length).fill("#ffffff"));
  
  shouldStop();
  
  if (rowsToWrite.length > 0) {
    executeSecureWrite("GITHUB_WRITE", () => {
      writeRowsInBatches(sheet, rowsToWrite, headerMap, manualColIndexes);
    });
  }
  
  enforceCacheLimit(cache, MAX_CACHE_SIZE);
  enforceCacheLimit(actCache, MAX_CACHE_SIZE);
  
  props.setProperty("GITHUB_CACHE", JSON.stringify(cache));
  props.setProperty("GITHUB_ACT_CACHE", JSON.stringify(actCache));
  return { newFoundCount: newFoundCount, updated: updatedCount, unchanged: unchangedCount, aborted: false, syncTime: new Date() };
}

// ================= DASHBOARD UPDATER =================
function updateDashboard(crm, sync, lms, github, redCount) {
  const dash = getOrCreateSheet(CONFIG.SHEETS.DASHBOARD);
  dash.setColumnWidth(1, 280); dash.setColumnWidth(2, 350);
  dash.getRange("A1:B1").setValues([["AUDIT METRIC", "DETAILS"]]).setFontWeight("bold").setBackground("#f3f3f3");
  
  const lastRow = dash.getLastRow();
  let ex = {};
  
  if (lastRow > 1) {
    dash.getRange(2, 1, lastRow - 1, 2).getValues().forEach(r => { if (r[0]) ex[r[0]] = r[1]; });
  }

  delete ex["Existing Records Actually Updated"];
  delete ex["Rows Checked But Unchanged (CRM)"];
  delete ex["Rows Checked But Unchanged (LMS)"];
  delete ex["Rows Checked But Unchanged (GitHub)"];
  delete ex["Students Removed (Not in CRM)"];
  delete ex["Students Purged (No ID)"];
  delete ex["Students Skipped (No ID & No Email)"];
  delete ex["Ambiguous Dates Defaulted to UK"];
  delete ex["Ambiguous Dates Defaulted to US"];
  delete ex["Unknown Headers Ignored"];
  delete ex["Header Alias Collisions Resolved"];
  
  if (crm) { ex["CRM File Generation Date (from CSV)"] = crm.generatedDate; ex["CRM Source File"] = crm.fileName; }
  if (sync) { 
    ex["New Students Added"] = sync.added; 
    ex["Records Removed / Skipped"] = (sync.removedFromCrm || 0) + (sync.removedNoId || 0) + (sync.skippedNoId || 0);
    ex["Last CRM Sync Time"] = sync.syncTime; 
  }
  if (redCount !== undefined && redCount !== null) { ex["Total At-Risk Students (Red)"] = redCount; }
  if (lms) { 
    ex["Last LMS Sync Time"] = lms.syncTime; 
    ex["LMS Records Updated"] = lms.updated; 
    ex["Unmatched LMS Records (ID missing/mismatch)"] = lms.unmatched;
  }
  if (github) { 
    ex["Last GitHub Sync Time"] = github.syncTime; 
    ex["New GitHub Profiles Found"] = github.newFoundCount; 
    ex["GitHub Records Updated"] = github.updated;
  }
  
  const outMap = [
    "CRM Source File", "CRM File Generation Date (from CSV)", "Last CRM Sync Time",
    "New Students Added", "Records Removed / Skipped", "Total At-Risk Students (Red)"
  ];
  if (FEATURES.LMS_TRACKING) outMap.push("Last LMS Sync Time", "LMS Records Updated", "Unmatched LMS Records (ID missing/mismatch)");
  if (FEATURES.GITHUB_TRACKING) outMap.push("Last GitHub Sync Time", "New GitHub Profiles Found", "GitHub Records Updated");
  
  const outData = [];
  outMap.forEach(k => { if (ex[k] !== undefined) outData.push([k, ex[k]]); });
  
  Object.keys(ex).forEach(k => { if (!outMap.includes(k)) outData.push([k, ex[k]]); });
  
  if (lastRow > 1) {
    dash.getRange(2, 1, lastRow, 2).clearContent();
  }
  
  if (globalUnknownHeaders > 0) outData.push(["Unknown Headers Ignored", globalUnknownHeaders]);
  if (globalAliasCollisions > 0) outData.push(["Header Alias Collisions Resolved", globalAliasCollisions]);
  // Removed the ambiguous dates push here
  
  if (outData.length > 0) {
    dash.getRange(2, 1, outData.length, 2).setValues(outData);
    const formats = outData.map(row => {
      if (row[1] instanceof Date) { return ["dd/MM/yyyy HH:mm:ss"]; } 
      else if (typeof row[1] === "number") { return ["0"]; } 
      else { return ["General"]; }
    });
    dash.getRange(2, 2, outData.length, 1).setNumberFormats(formats);
  }
}
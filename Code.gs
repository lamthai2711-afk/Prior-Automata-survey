/**
 * ==============================================================================
 * GOOGLE APPS SCRIPT: PRE-SESSION SHEET BACKEND (SHEET 1 - 28 COLUMNS)
 * Purpose: Receives and processes Pre-Session Survey data:
 *          - Informed Consent & Demographics (Age, Gender, Standing, Major, English, Prior, Nationality)
 *          - Secondary Matching IDs (Neptun Code, Student Name)
 *          - Baseline Knowledge Pre-Test (Q1 to Q5) with automatic scoring
 *          - Device audit metadata (IP, User Agent, Time on Page)
 * ==============================================================================
 */

var PRE_SESSION_SCHEMA = [
  // 1. Participant Identity & Demographics
  { key: "timestamp",            title: "Timestamp" },
  { key: "participant_id",       title: "Participant ID" },
  { key: "neptun_code",          title: "Neptun Code" },
  { key: "student_name",         title: "Student Full Name" },
  { key: "consent_confirmed",    title: "Consent Confirmed" },
  { key: "age",                  title: "Age" },
  { key: "gender_identity",      title: "Gender Identity" },
  { key: "academic_standing",    title: "Academic Standing" },
  { key: "field_of_study",       title: "Field of Study" },
  { key: "english_proficiency",  title: "English Proficiency" },
  { key: "prior_familiarity",    title: "Prior Familiarity" },

  // 2. Baseline Knowledge Pre-Test Items (Diagnostic Questions 1 to 5)
  { key: "pretest_q1_ans",       title: "Pre Q1 (Power Set) Answer" },
  { key: "pretest_q1_score",     title: "Pre Q1 Score (0-1)" },
  { key: "pretest_q2_ans",       title: "Pre Q2 (Transition Function) Answer" },
  { key: "pretest_q2_score",     title: "Pre Q2 Score (0-1)" },
  { key: "pretest_q3_ans",       title: "Pre Q3 (NFA vs DFA) Answer" },
  { key: "pretest_q3_score",     title: "Pre Q3 Score (0-1)" },
  { key: "pretest_q4_ans",       title: "Pre Q4 (Execution Path) Answer" },
  { key: "pretest_q4_score",     title: "Pre Q4 Score (0-1)" },
  { key: "pretest_q5_ans",       title: "Pre Q5 (E-Closure) Answer" },
  { key: "pretest_q5_score",     title: "Pre Q5 Score (0-1)" },
  { key: "total_pretest_score",  title: "Total Pre-Test Score (0-5)" },

  // 3. Device Integrity, Session Audit Metadata & Nationality
  { key: "time_on_page_seconds", title: "Duration on Page (Seconds)" },
  { key: "ip_address",           title: "IP Address" },
  { key: "client_timezone",      title: "Client Timezone" },
  { key: "screen_resolution",    title: "Screen Resolution" },
  { key: "user_agent",           title: "User Agent (Browser/OS)" },
  { key: "nationality",          title: "Nationality" }
];

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "active",
    service: "Pre-Session Diagnostic Backend (Sheet 1)",
    ready: true,
    total_columns: PRE_SESSION_SCHEMA.length,
    timestamp: new Date()
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  if (!e) {
    return ContentService.createTextOutput("Direct execution without payload is not allowed. Use testDoPost() to test.");
  }

  var lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = {};

    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter || {};
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    // Automated Pre-test Diagnostic Scoring Engine (1 pt each, Total: 0 - 5)
    var q1_ans = String(data.pretest_q1 || "").trim();
    var q2_ans = String(data.pretest_q2 || "").trim();
    var q3_ans = String(data.pretest_q3 || "").trim();
    var q4_ans = String(data.pretest_q4 || "").trim();
    var q5_ans = String(data.pretest_q5 || "").trim();

    var q1_score = (q1_ans === "4") ? 1 : 0;
    var q2_score = (q2_ans === "The state transition function") ? 1 : 0;
    var q3_score = (q3_ans.indexOf("For a given state and input symbol") !== -1) ? 1 : 0;
    var q4_score = (q4_ans === "001" || q4_ans === "1") ? 1 : 0;
    var q5_score = (q5_ans === "{q0, q1, q2}") ? 1 : 0;

    var total_pretest_score = q1_score + q2_score + q3_score + q4_score + q5_score;

    var record = {
      timestamp: new Date(),
      participant_id: (data.participant_id || "").toString().trim() || "NOT_PROVIDED",
      neptun_code: (data.neptun_code || "").toString().trim().toUpperCase(),
      student_name: (data.student_name || "").toString().trim(),
      consent_confirmed: data.consent_confirmed ? "Yes" : "Yes (Confirmed)",
      age: parseInt(data.age, 10) || "",
      gender_identity: data.gender_identity || "",
      academic_standing: data.academic_standing || "",
      field_of_study: data.field_of_study || "",
      english_proficiency: data.english_proficiency || "",
      prior_familiarity: data.prior_familiarity || "",
      nationality: (data.nationality || data.Nationality || data.country || "").toString().trim(),

      // Pre-test questions & scores
      pretest_q1_ans: q1_ans,
      pretest_q1_score: q1_score,
      pretest_q2_ans: q2_ans,
      pretest_q2_score: q2_score,
      pretest_q3_ans: q3_ans,
      pretest_q3_score: q3_score,
      pretest_q4_ans: q4_ans,
      pretest_q4_score: q4_score,
      pretest_q5_ans: q5_ans,
      pretest_q5_score: q5_score,
      total_pretest_score: total_pretest_score,

      // Audit metadata
      time_on_page_seconds: parseFloat(data.time_on_page_seconds) || 0,
      ip_address: data.ip_address || "unknown",
      client_timezone: data.client_timezone || "",
      screen_resolution: data.screen_resolution || "",
      user_agent: data.user_agent || ""
    };

    var alignedRow = getAlignedRowData(sheet, record, PRE_SESSION_SCHEMA, "#1e293b");
    sheet.appendRow(alignedRow);

    return ContentService.createTextOutput(JSON.stringify({
      result: "success",
      participant_id: record.participant_id,
      nationality: record.nationality,
      pretest_score: total_pretest_score
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      result: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function getAlignedRowData(sheet, record, schema, headerBgColor) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    var titles = schema.map(function(col) { return col.title; });
    sheet.appendRow(titles);
    styleHeaderRow(sheet, titles.length, headerBgColor);
    return schema.map(function(col) {
      return (record[col.key] !== undefined && record[col.key] !== null) ? record[col.key] : "";
    });
  }

  var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var normalizedExisting = existingHeaders.map(normalizeHeader);

  schema.forEach(function(col) {
    var normTitle = normalizeHeader(col.title);
    var normKey = normalizeHeader(col.key);
    if (normalizedExisting.indexOf(normTitle) === -1 && normalizedExisting.indexOf(normKey) === -1) {
      existingHeaders.push(col.title);
      normalizedExisting.push(normTitle);
      sheet.getRange(1, existingHeaders.length).setValue(col.title);
    }
  });

  var lookup = {};
  schema.forEach(function(col) {
    lookup[normalizeHeader(col.title)] = record[col.key];
    lookup[normalizeHeader(col.key)] = record[col.key];
  });

  var alignedRow = [];
  for (var i = 0; i < existingHeaders.length; i++) {
    var normHeader = normalizeHeader(existingHeaders[i]);
    if (lookup.hasOwnProperty(normHeader)) {
      var val = lookup[normHeader];
      alignedRow.push((val !== undefined && val !== null) ? val : "");
    } else {
      alignedRow.push("");
    }
  }

  return alignedRow;
}

function styleHeaderRow(sheet, colCount, headerBgColor) {
  var headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setFontWeight("bold");
  headerRange.setBackground(headerBgColor);
  headerRange.setFontColor("#ffffff");
  headerRange.setHorizontalAlignment("center");
  headerRange.setVerticalAlignment("middle");
  headerRange.setWrap(true);
  sheet.setRowHeight(1, 40);
  sheet.setFrozenRows(1);
}

function normalizeHeader(str) {
  return String(str || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function setupPreSessionSheetHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    var titles = PRE_SESSION_SCHEMA.map(function(col) { return col.title; });
    sheet.appendRow(titles);
    styleHeaderRow(sheet, titles.length, "#1e293b");
  } else {
    var existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var normalizedExisting = existingHeaders.map(normalizeHeader);
    PRE_SESSION_SCHEMA.forEach(function(col) {
      var normTitle = normalizeHeader(col.title);
      if (normalizedExisting.indexOf(normTitle) === -1) {
        existingHeaders.push(col.title);
        normalizedExisting.push(normTitle);
        sheet.getRange(1, existingHeaders.length).setValue(col.title);
      }
    });
    styleHeaderRow(sheet, existingHeaders.length, "#1e293b");
  }
  SpreadsheetApp.flush();
  Logger.log("Pre-Session Sheet headers synchronized non-destructively.");
}

function testDoPost() {
  var mockEvent = {
    postData: {
      contents: JSON.stringify({
        participant_id: "SZE-TEST01",
        neptun_code: "XYZ999",
        student_name: "Test Student",
        age: 21,
        gender_identity: "Male",
        nationality: "Vietnam",
        academic_standing: "Bachelor's Degree - 2nd Year",
        field_of_study: "Computer Science / Software Engineering",
        english_proficiency: "Advanced / Fluent",
        prior_familiarity: "Practical",
        pretest_q1: "4",
        pretest_q2: "The state transition function",
        pretest_q3: "For a given state and input symbol, an NFA can transition to a set of states (including ∅), whereas a DFA must transition to exactly one deterministic state.",
        pretest_q4: "001",
        pretest_q5: "{q0, q1, q2}",
        time_on_page_seconds: 120,
        ip_address: "127.0.0.1"
      })
    }
  };
  var result = doPost(mockEvent);
  Logger.log(result.getContent());
}

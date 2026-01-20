/**************************************************************
 * LCM Finance — server.gs (v3 - MATCHES YOUR SPREADSHEET)
 * Spreadsheet ID: 1ZNAMPTjBH2xBKQ5nAlQqbWnn9HaSsCWjf6dlF2RNNCQ
 * 
 * FIXED: Column names to match your actual spreadsheet structure
 **************************************************************/

const SS_ID = "1ZNAMPTjBH2xBKQ5nAlQqbWnn9HaSsCWjf6dlF2RNNCQ";

/* =========================
   SHEET NAMES (matching your spreadsheet)
   ========================= */
const SHEET_PVS        = "PVs";
const SHEET_PAYEES     = "Payees";
const SHEET_LOOKUPS    = "Lookups";
const SHEET_RECURRING  = "Recurring";
const SHEET_FAVES      = "Favourites";
const SHEET_SIGNATURES = "Signatures";
const SHEET_SHARES     = "Shares";
const SHEET_LOGS       = "Logs";
const SHEET_ROLES      = "Roles";
const SHEET_MINISTRY_HEADS = "MinistryHeads";
const DEFAULT_DEPARTMENTS = [
  "Bishop",
  "LCM Mission",
  "LCM Social Concern",
  "LCM Stewardship",
  "LCM Young Adults and Youth",
  "LCM Property",
  "LCM Orang Asli",
  "LCM Communication",
  "LCM Education",
  "Luther Study Centre",
  "LCM HQ Office",
  "Luther Centre Management (BMC)",
  "LCM Payroll"
const DEFAULT_MINISTRIES = [
  "LCM HQ Office",
  "LCM Mission",
  "LCM Education",
  "LCM YAY",
  "Property",
  "Worship",
  "Others"
];

/* =========================
   ROLE EMAIL MAPPINGS
   Add all authorized emails here (lowercase)
   ========================= */
const ROLE_EMAILS = {
  FINANCE_ADMIN: [
    "finance@lcm.org.my"
    // Add more Finance Admin emails as needed
  ],
  BISHOP: [
    "bishop@lcm.org.my"
    // Add actual Bishop email
  ],
  TREASURER: [
    "treasurer@lcm.org.my"
    // Add actual Treasurer email
  ],
  SECRETARY: [
    "secretary@lcm.org.my"
    // Add actual Secretary email
  ]
};

/* =========================
   WEB APP
   ========================= */
function doGet(e) {
  const t = HtmlService.createTemplateFromFile("index");
  t.page  = e?.parameter?.page || "";
  t.pv    = e?.parameter?.pv || "";
  t.token = e?.parameter?.token || "";
  return t.evaluate()
    .setTitle("LCM Finance")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================
   HELPERS
   ========================= */
function ss_() { return SpreadsheetApp.openById(SS_ID); }

function getSheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error("Sheet not found: " + name);
  return sh;
}

function ensureSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (headers && headers.length) {
      sh.appendRow(headers);
    }
  }
  return sh;
}

function getHeaders_(sh) {
  if (sh.getLastColumn() === 0) return [];
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}

function rows_(sh) {
  if (sh.getLastRow() <= 1) return [];
  const data = sh.getDataRange().getValues();
  const headers = data.shift();
  return data.map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function append_(sh, obj) {
  try {
    const headers = getHeaders_(sh);
    if (headers.length === 0) {
      console.error("append_: No headers in sheet " + sh.getName());
      return false;
    }
    const row = headers.map(h => {
      const val = obj[h];
      return val !== undefined ? val : "";
    });
    sh.appendRow(row);
    console.log("append_: Added row to " + sh.getName());
    return true;
  } catch (e) {
    console.error("append_ error:", e);
    return false;
  }
}

function updateRow_(sh, matchCol, matchVal, updates) {
  const headers = getHeaders_(sh);
  const data = sh.getDataRange().getValues();
  const colIdx = headers.indexOf(matchCol);
  if (colIdx === -1) throw new Error("Column not found: " + matchCol);

  for (let i = 1; i < data.length; i++) {
    if (data[i][colIdx] === matchVal) {
      Object.keys(updates).forEach(col => {
        const idx = headers.indexOf(col);
        if (idx !== -1) {
          sh.getRange(i + 1, idx + 1).setValue(updates[col]);
        }
      });
      return true;
    }
  }
  return false;
}

function update_(sh, idx, updates) {
  if (!sh) throw new Error("update_: missing sheet");
  const headers = getHeaders_(sh);
  if (!headers.length) throw new Error("update_: missing headers");
  const rowNumber = idx + 2;
  Object.keys(updates || {}).forEach(key => {
    const colIdx = headers.indexOf(key);
    if (colIdx !== -1) {
      sh.getRange(rowNumber, colIdx + 1).setValue(updates[key]);
    }
  });
  return true;
}

function id_(prefix) { 
  return prefix + "-" + Utilities.getUuid().slice(0, 8).toUpperCase(); 
}

function now_() { return new Date(); }

function jsonParse_(str, fallback) {
  if (!str) return fallback;
  try { return JSON.parse(str); }
  catch (e) { return fallback; }
}

function log_(event, data) {
  try {
    const sh = ss_().getSheetByName(SHEET_LOGS);
    if (sh) {
      sh.appendRow([now_(), event, JSON.stringify(data)]);
    }
  } catch (e) { console.error("log_ error:", e); }
}

function ensurePvColumns_() {
  const sh = ensureSheet_(SHEET_PVS, []);
  const headers = getHeaders_(sh);
  if (!headers.length) return;
  const required = [
    "ministry",
    "ministry_verified",
    "ministry_verified_by",
    "ministry_verified_at",
    "ministry_verified_comment",
    "ministry_verify_bypassed",
    "ministry_bypass_by",
    "ministry_bypass_at",
    "ministry_bypass_reason"
  ];
  const missing = required.filter(col => !headers.includes(col));
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
}

function ensureMinistryHeadsSheet_() {
  const sh = ensureSheet_(SHEET_MINISTRY_HEADS, [
    "Department",
    "Primary_Head_Name",
    "Primary_Head_Email",
    "Secondary_Head_Name",
    "Secondary_Head_Email",
    "Created_At",
    "Updated_At"
  ]);
  const headers = getHeaders_(sh);
  const required = [
    "Department",
    "Primary_Head_Name",
    "Primary_Head_Email",
    "Secondary_Head_Name",
    "Secondary_Head_Email",
    "Created_At",
    "Updated_At"
  ];
  const sh = ensureSheet_(SHEET_MINISTRY_HEADS, ["Ministry", "Head_Email", "Head_Name", "Is_Supplementary", "Created_At", "Updated_At"]);
  const headers = getHeaders_(sh);
  const required = ["Ministry", "Head_Email", "Head_Name", "Is_Supplementary", "Created_At", "Updated_At"];
  const missing = required.filter(h => !headers.includes(h));
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function getCanonicalDepartments_() {
  const ss = ss_();
  const deptSh = ss.getSheetByName("Departments");
  let departments = [];
  if (deptSh) {
    const deptRows = rows_(deptSh);
    departments = deptRows.map(r => (r.Dept_Name || r.dept_name || r.name || "")).filter(Boolean);
  }
  const combined = DEFAULT_DEPARTMENTS.concat(departments);
  return Array.from(new Set(combined));
}

/* =========================
   USER CONTEXT / ROLES
   ========================= */
function getUserContext() {
  let email = "";
  
  // In web app context, getActiveUser() often returns empty
  // getEffectiveUser() returns the script owner's email
  // We need to try both
  try {
    email = Session.getActiveUser().getEmail();
  } catch(e) {}
  
  // If empty, try effective user
  if (!email) {
    try {
      email = Session.getEffectiveUser().getEmail();
    } catch(e) {}
  }
  
  email = (email || "").toLowerCase().trim();
  
  // IMPORTANT: For web apps deployed as "Execute as: Me" with "Anyone" access,
  // getActiveUser() returns empty. In this case, we use getEffectiveUser().
  // If you need per-user tracking, deploy as "Execute as: User accessing the web app"
  
  console.log("getUserContext: email=" + email);

  const props = PropertiesService.getUserProperties();
  const testRole = props.getProperty("TEST_ROLE") || "";

  // Check test mode first
  if (testRole) {
    console.log("getUserContext: Using test role=" + testRole);
    return {
      email: email,
      testRole: testRole,
      isFinanceAdmin: testRole === "ADMIN",
      isBishop: testRole === "BISHOP",
      isTreasurer: testRole === "TREASURER",
      isSecretary: testRole === "SECRETARY",
      isSignatory: ["BISHOP", "TREASURER", "SECRETARY"].includes(testRole),
      signatoryRole: testRole,
      isMinistryHead: false,
      ministries: []
    };
  }

  // Check actual roles - normalize all emails to lowercase
  const financeEmails = ROLE_EMAILS.FINANCE_ADMIN.map(e => (e || "").toLowerCase().trim());
  const bishopEmails = ROLE_EMAILS.BISHOP.map(e => (e || "").toLowerCase().trim());
  const treasurerEmails = ROLE_EMAILS.TREASURER.map(e => (e || "").toLowerCase().trim());
  const secretaryEmails = ROLE_EMAILS.SECRETARY.map(e => (e || "").toLowerCase().trim());

  const isFinanceAdmin = financeEmails.includes(email);
  const isBishop = bishopEmails.includes(email);
  const isTreasurer = treasurerEmails.includes(email);
  const isSecretary = secretaryEmails.includes(email);

  console.log("getUserContext: isFinanceAdmin=" + isFinanceAdmin + ", email=" + email);

  let signatoryRole = "";
  if (isBishop) signatoryRole = "BISHOP";
  else if (isTreasurer) signatoryRole = "TREASURER";
  else if (isSecretary) signatoryRole = "SECRETARY";

  const ministryAssignments = email ? getDepartmentsForVerifier_(email) : [];
  const ministryAssignments = email ? getMinistriesForHead_(email) : [];

  return {
    email: email,
    testRole: "",
    isFinanceAdmin: isFinanceAdmin,
    isBishop: isBishop,
    isTreasurer: isTreasurer,
    isSecretary: isSecretary,
    isSignatory: isBishop || isTreasurer || isSecretary,
    signatoryRole: signatoryRole,
    isMinistryHead: ministryAssignments.length > 0,
    ministries: ministryAssignments
  };
}

function setTestRole(role) {
  PropertiesService.getUserProperties().setProperty("TEST_ROLE", role || "");
  return { ok: true, role: role || "" };
}

/**
 * Debug function - call from web app to see what's happening
 */
function debugWebAppContext() {
  let activeEmail = "", effectiveEmail = "";
  
  try { activeEmail = Session.getActiveUser().getEmail(); } catch(e) { activeEmail = "ERROR: " + e.message; }
  try { effectiveEmail = Session.getEffectiveUser().getEmail(); } catch(e) { effectiveEmail = "ERROR: " + e.message; }
  
  const ctx = getUserContext();
  
  const sh = getSheet_(SHEET_PVS);
  const allPvs = rows_(sh);
  const myPvs = allPvs.filter(r => {
    const pvEmail = String(r.applicant_email || "").toLowerCase().trim();
    return pvEmail === ctx.email;
  });
  
  return {
    activeUserEmail: activeEmail,
    effectiveUserEmail: effectiveEmail,
    finalEmail: ctx.email,
    isFinanceAdmin: ctx.isFinanceAdmin,
    isSignatory: ctx.isSignatory,
    testRole: ctx.testRole,
    totalPVsInSheet: allPvs.length,
    matchingPVs: myPvs.length,
    matchingPVNumbers: myPvs.map(p => p.pv_no)
  };
}

/* =========================
   LOOKUPS
   ========================= */
function getLookups() {
  const ss = ss_();
  ensureMinistryHeadsSheet_();
  ensurePvColumns_();

  // Departments from Departments sheet (with heads)
  let departments = [];
  const deptSh = ss.getSheetByName("Departments");
  if (deptSh) {
    const deptRows = rows_(deptSh);
    departments = deptRows.map(r => ({
      name: r.Dept_Name || r.dept_name || r.name || "",
      head_name: r.Head_Name || r.head_name || "",
      head_email: r.Head_Email || r.head_email || ""
    })).filter(d => d.name);
  }
  if (!departments.length) {
    departments = getCanonicalDepartments_().map(name => ({ name: name, head_name: "", head_email: "" }));
  }

  // From Lookups sheet (for projects, banks, etc.)
  let look = [];
  const lookSh = ss.getSheetByName(SHEET_LOOKUPS);
  if (lookSh) look = rows_(lookSh);

  // Projects from Projects sheet or Lookups
  let projects = [];
  const projSh = ss.getSheetByName("Projects");
  if (projSh) {
    const projRows = rows_(projSh);
    projects = projRows.map(r => r.name || r.project_name || "").filter(Boolean);
  } else {
    projects = look.filter(r => r.category === "projects").map(r => r.value);
  }

  // Payees
  let payees = [];
  const payeeSh = ss.getSheetByName(SHEET_PAYEES);
  if (payeeSh) payees = rows_(payeeSh);

  // Applicants from PVs
  let pvs = [];
  const pvSh = ss.getSheetByName(SHEET_PVS);
  if (pvSh) pvs = rows_(pvSh);

  const departmentsList = getCanonicalDepartments_();
  const ministriesFromLookups = look.filter(r => r.category === "ministry").map(r => r.value).filter(Boolean);
  const ministries = Array.from(new Set(DEFAULT_MINISTRIES.concat(ministriesFromLookups)));

  return {
    departments: departments,
    departmentNames: departmentsList,
    projects: projects,
    banks: look.filter(r => r.category === "banks").map(r => r.value),
    ministries: [],
    ministries: ministries,
    payees: payees.map(p => ({ 
      name: p.name || "", 
      bank_name: p.bank_name || "", 
      bank_acct: p.bank_acct || "" 
    })),
    applicants: [...new Set(pvs.map(p => p.applicant_name).filter(Boolean))]
  };
}

/* =========================
   DEPARTMENT MANAGEMENT
   ========================= */
function getDepartments() {
  const ss = ss_();
  let deptSh = ss.getSheetByName("Departments");
  
  // Create sheet if doesn't exist
  if (!deptSh) {
    deptSh = ss.insertSheet("Departments");
    deptSh.appendRow(["Dept_Name", "Head_Name", "Head_Email", "Created_At", "Updated_At"]);
  }
  
  const depts = rows_(deptSh);
  return {
    ok: true,
    departments: depts.map(r => ({
      name: r.Dept_Name || r.dept_name || "",
      head_name: r.Head_Name || r.head_name || "",
      head_email: r.Head_Email || r.head_email || ""
    })).filter(d => d.name)
  };
}

function saveDepartment(d) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Not authorized" };
    }

    const ss = ss_();
    let deptSh = ss.getSheetByName("Departments");
    if (!deptSh) {
      deptSh = ss.insertSheet("Departments");
      deptSh.appendRow(["Dept_Name", "Head_Name", "Head_Email", "Created_At", "Updated_At"]);
    }

    const depts = rows_(deptSh);
    const existing = depts.find(r => (r.Dept_Name || r.dept_name) === d.name);

    if (existing) {
      updateRow_(deptSh, "Dept_Name", d.name, {
        Head_Name: d.head_name || "",
        Head_Email: d.head_email || "",
        Updated_At: now_()
      });
    } else {
      append_(deptSh, {
        Dept_Name: d.name,
        Head_Name: d.head_name || "",
        Head_Email: d.head_email || "",
        Created_At: now_(),
        Updated_At: now_()
      });
    }

    return { ok: true };
  } catch (e) {
    console.error("saveDepartment error:", e);
    return { ok: false, error: e.message };
  }
}

function deleteDepartment(name) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized" };
    }

    const ss = ss_();
    const deptSh = ss.getSheetByName("Departments");
    if (!deptSh) return { ok: false, error: "Departments sheet not found" };

    const headers = getHeaders_(deptSh);
    const data = deptSh.getDataRange().getValues();
    const nameIdx = headers.indexOf("Dept_Name");
    if (nameIdx === -1) return { ok: false, error: "Dept_Name column not found" };

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === name) {
        deptSh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Department not found" };
  } catch (e) {
    console.error("deleteDepartment error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   PROJECT MANAGEMENT
   ========================= */
function getProjects() {
  const ss = ss_();
  let projSh = ss.getSheetByName("Projects");
  
  if (!projSh) {
    projSh = ss.insertSheet("Projects");
    projSh.appendRow(["name", "created_at", "updated_at"]);
  }
  
  const projs = rows_(projSh);
  return {
    ok: true,
    projects: projs.map(r => r.name || "").filter(Boolean)
  };
}

function saveProject(name) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Not authorized" };
    }

    const ss = ss_();
    let projSh = ss.getSheetByName("Projects");
    if (!projSh) {
      projSh = ss.insertSheet("Projects");
      projSh.appendRow(["name", "created_at", "updated_at"]);
    }

    const projs = rows_(projSh);
    if (projs.some(r => r.name === name)) {
      return { ok: false, error: "Project already exists" };
    }

    append_(projSh, {
      name: name,
      created_at: now_(),
      updated_at: now_()
    });

    return { ok: true };
  } catch (e) {
    console.error("saveProject error:", e);
    return { ok: false, error: e.message };
  }
}

function deleteProject(name) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized" };
    }

    const ss = ss_();
    const projSh = ss.getSheetByName("Projects");
    if (!projSh) return { ok: false, error: "Projects sheet not found" };

    const headers = getHeaders_(projSh);
    const data = projSh.getDataRange().getValues();
    const nameIdx = headers.indexOf("name");
    if (nameIdx === -1) return { ok: false, error: "name column not found" };

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === name) {
        projSh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Project not found" };
  } catch (e) {
    console.error("deleteProject error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   PAYEES
   ========================= */
function savePayee(p) {
  try {
    const sh = getSheet_(SHEET_PAYEES);
    const allRows = rows_(sh);
    const existing = allRows.find(r => r.name === p.name);

    if (existing) {
      // Update
      updateRow_(sh, "name", p.name, { updated_at: now_() });
    } else {
      // Add new - matches your Payees columns
      append_(sh, {
        name: p.name || "",
        dept: p.dept || "",
        bank_name: p.bank_name || "",
        bank_acct: p.bank_acct || "",
        biller_code: p.biller_code || "",
        default_ref_no: p.default_ref_no || "",
        created_at: now_(),
        updated_at: now_()
      });
    }
    return { ok: true };
  } catch (e) {
    console.error("savePayee error:", e);
    return { ok: false, error: e.message };
  }
}

function deletePayee(name) {
  try {
    const sh = getSheet_(SHEET_PAYEES);
    const headers = getHeaders_(sh);
    const data = sh.getDataRange().getValues();
    const nameIdx = headers.indexOf("name");
    if (nameIdx === -1) return { ok: false, error: "name column not found" };

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx] === name) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Payee not found" };
  } catch (e) {
    console.error("deletePayee error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   ATTACHMENTS
   ========================= */
function uploadAttachmentFile(dataUrl, filename, mimeType) {
  try {
    const b64 = dataUrl.split(",")[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(b64),
      mimeType || "application/octet-stream",
      filename || "attachment"
    );
    const f = DriveApp.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { ok: true, url: f.getUrl() };
  } catch (e) {
    console.error("uploadAttachmentFile error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   SUBMIT PV
   Columns in your PVs sheet:
   pv_no, date, created_at, updated_at, status, applicant_name, applicant_email,
   dept, project, payee_name, payment_method, payee_bank_name, payee_bank_acct,
   cheque_no, biller_code, ref_no, purpose, amount, line_items_json,
   attachments_json, sig_applicant_name, sig_applicant_confirm, sig_head_name,
   sig_head_confirm, admin_comment, approvals_json, signed_pdf_url
   ========================= */
function submitPublicPV(d) {
  try {
    const ctx = getUserContext();
    ensurePvColumns_();
    const sh = getSheet_(SHEET_PVS);

    // Generate PV number
    const year = Utilities.formatDate(now_(), Session.getScriptTimeZone(), "yyyy");
    const lastRow = sh.getLastRow();
    const seq = ("0000" + lastRow).slice(-4);
    const pvNo = "PV-" + year + "-" + seq;

    // Get department head info
    const deptResult = getDepartments();
    const departments = deptResult.departments || [];
    const deptInfo = departments.find(dept => dept.name === d.dept) || {};

    // Generate a tracking token for public users
    const trackingToken = Utilities.getUuid().replace(/-/g, "").substring(0, 16);

    // Determine initial status:
    // - If department has a head AND applicant is not the head, status is "PENDING_HEAD"
    // - Otherwise, status is "PENDING" (goes straight to Finance Admin)
    const hasDeptHead = deptInfo.head_email && deptInfo.head_email.trim() !== "";
    const applicantEmail = (d.applicant_email || ctx.email || "").toLowerCase().trim();
    const isApplicantHead = hasDeptHead && deptInfo.head_email.toLowerCase().trim() === applicantEmail;
    const initialStatus = (hasDeptHead && !isApplicantHead) ? "PENDING_HEAD" : "PENDING";

    // Build row matching YOUR column structure
    const ministry = d.ministry || d.dept || "";
    const rowData = {
      pv_no: pvNo,
      date: d.pvDate || "",
      created_at: now_(),
      updated_at: now_(),
      status: initialStatus,
      applicant_name: d.applicant_name || "",
      applicant_email: d.applicant_email || ctx.email || "",
      dept: d.dept || "",
      ministry: ministry,
      dept_head_name: deptInfo.head_name || "",
      dept_head_email: deptInfo.head_email || "",
      project: d.project || "",
      payee_name: d.payee_name || "",
      payment_method: d.payment_method || "",
      payee_bank_name: d.payee_bank_name || "",
      payee_bank_acct: d.payee_bank_acct || "",
      cheque_no: d.cheque_no || "",
      biller_code: d.biller_code || "",
      ref_no: d.ref_no || "",
      purpose: d.purpose || "",
      amount: d.amount || 0,
      line_items_json: JSON.stringify(d.line_items || []),
      attachments_json: JSON.stringify(d.attachments || []),
      sig_applicant_name: d.sig_applicant_name || "",
      sig_applicant_confirm: d.sig_applicant_confirm || "NO",
      sig_head_name: d.sig_head_name || "",
      sig_head_confirm: d.sig_head_confirm || "NO",
      admin_comment: "",
      approvals_json: "[]",
      signed_pdf_url: "",
      tracking_token: trackingToken,
      head_verified: (hasDeptHead && !isApplicantHead) ? "NO" : "N/A",
      head_verified_at: "",
      head_comment: "",
      ministry_verified: "NO",
      ministry_verified_by: "",
      ministry_verified_at: "",
      ministry_verified_comment: "",
      ministry_verify_bypassed: "NO",
      ministry_bypass_by: "",
      ministry_bypass_at: "",
      ministry_bypass_reason: ""
    };

    const success = append_(sh, rowData);
    if (!success) {
      return { ok: false, error: "Failed to append row to PVs sheet" };
    }

    // Log it
    log_("PV submitted", { pv_no: pvNo, by: ctx.email || d.applicant_email, amount: d.amount, status: initialStatus });

    let message = "PV created: " + pvNo;
    if (initialStatus === "PENDING_HEAD") {
      message += ". Awaiting verification by " + deptInfo.head_name + " (" + deptInfo.head_email + ")";
    } else {
      message += ". Sent to Finance Admin for review.";
    }

    return { 
      ok: true, 
      pv_no: pvNo, 
      tracking_token: trackingToken,
      status: initialStatus,
      message: message
    };
  } catch (e) {
    console.error("submitPublicPV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   TRACK PV STATUS (Public - by tracking token or PV number)
   ========================= */
function trackPVStatus(identifier) {
  try {
    if (!identifier) {
      return { ok: false, error: "PV number or tracking token required" };
    }

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    
    // Try to find by tracking token first, then by pv_no
    let pv = allPvs.find(p => p.tracking_token === identifier);
    if (!pv) {
      pv = allPvs.find(p => p.pv_no === identifier);
    }

    if (!pv) {
      return { ok: false, error: "PV not found" };
    }

    // Parse approvals for status display
    const approvals = jsonParse_(pv.approvals_json, []);
    const approvedRoles = approvals.filter(a => a.decision === "APPROVED").map(a => a.role);

    // Build status timeline
    const timeline = [];
    timeline.push({
      step: "Submitted",
      status: "complete",
      date: pv.created_at ? formatDateStr_(pv.created_at) : ""
    });

    const status = String(pv.status || "").toUpperCase();
    const hasDeptHead = pv.dept_head_email && pv.dept_head_email.trim() !== "";

    if (hasDeptHead && pv.head_verified !== "N/A") {
      if (pv.head_verified === "YES") {
        timeline.push({
          step: "Department Head Verified",
          status: "complete",
          date: pv.head_verified_at ? formatDateStr_(pv.head_verified_at) : "",
          by: pv.dept_head_name
        });
      } else if (status === "PENDING_HEAD") {
        timeline.push({
          step: "Awaiting Department Head Verification",
          status: "current",
          by: pv.dept_head_name
        });
      } else if (status === "REJECTED_HEAD") {
        timeline.push({
          step: "Rejected by Department Head",
          status: "rejected",
          comment: pv.head_comment
        });
      }
    }

    if (["PENDING", "REVIEWED", "APPROVED"].includes(status)) {
      if (status === "PENDING") {
        timeline.push({ step: "Awaiting Finance Admin Review", status: "current" });
      } else {
        timeline.push({ step: "Finance Admin Reviewed", status: "complete" });
      }
    }

    const ministryVerified = String(pv.ministry_verified || "").toUpperCase() === "YES";
    const ministryBypassed = String(pv.ministry_verify_bypassed || "").toUpperCase() === "YES";
    const ministryRequired = String(pv.dept || "").trim() !== "";
    const ministryRequired = String(pv.ministry || "").trim() !== "";
    const ministryReady = !ministryRequired || ministryVerified || ministryBypassed;

    if (["REVIEWED", "APPROVED"].includes(status) && ministryRequired) {
      if (ministryVerified) {
        timeline.push({
          step: "Head Verified",
          status: "complete",
          date: pv.ministry_verified_at ? formatDateStr_(pv.ministry_verified_at) : "",
          by: pv.ministry_verified_by || ""
        });
      } else if (ministryBypassed) {
        timeline.push({
          step: "Verification Bypassed",
          status: "complete",
          date: pv.ministry_bypass_at ? formatDateStr_(pv.ministry_bypass_at) : "",
          by: pv.ministry_bypass_by || ""
        });
      } else {
        timeline.push({
          step: "Awaiting Head Verification",
          status: "current"
        });
      }
    }

    if (["REVIEWED", "APPROVED"].includes(status) && ministryReady) {
      const approvalCount = approvedRoles.length;
      if (approvalCount < 2) {
        timeline.push({
          step: "Awaiting Signatory Approval (" + approvalCount + "/2)",
          status: "current",
          details: approvedRoles.length > 0 ? approvedRoles.join(" ✓, ") + " ✓" : ""
        });
      } else {
        timeline.push({
          step: "Fully Approved (2/2)",
          status: "complete",
          details: approvedRoles.join(" ✓, ") + " ✓"
        });
      }
    }

    if (status === "APPROVED") {
      timeline.push({ step: "Payment Voucher Approved", status: "complete" });
    } else if (status === "REJECTED") {
      timeline.push({ step: "Rejected", status: "rejected", comment: pv.admin_comment || pv.signatory_comment });
    }

    return {
      ok: true,
      pv: {
        pv_no: pv.pv_no,
        date: pv.date ? formatDateStr_(pv.date) : "",
        applicant_name: pv.applicant_name,
        dept: pv.dept,
        payee_name: pv.payee_name,
        amount: pv.amount,
        status: pv.status,
        purpose: pv.purpose
      },
      timeline: timeline,
      current_status: status
    };
  } catch (e) {
    console.error("trackPVStatus error:", e);
    return { ok: false, error: e.message };
  }
}

// Helper to format date as dd/mm/yyyy
function formatDateStr_(dateVal) {
  if (!dateVal) return "";
  try {
    var d = new Date(dateVal);
    if (isNaN(d.getTime())) return String(dateVal).substring(0, 10);
    var day = ("0" + d.getDate()).slice(-2);
    var month = ("0" + (d.getMonth() + 1)).slice(-2);
    var year = d.getFullYear();
    return day + "/" + month + "/" + year;
  } catch (e) {
    return String(dateVal).substring(0, 10);
  }
}

/* =========================
   SEARCH PVs (by name, payee, amount)
   ========================= */
function searchPVs(query) {
  try {
    if (!query) {
      return { ok: false, error: "Search query required" };
    }

    query = String(query).toLowerCase().trim();
    
    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    
    // Parse amount range if query looks like a number
    let amountExact = null;
    let amountRange = null;
    const numQuery = parseFloat(query.replace(/[^0-9.-]/g, ""));
    if (!isNaN(numQuery)) {
      amountExact = numQuery;
      // Also search nearest 10s and 100s
      const tens = Math.floor(numQuery / 10) * 10;
      const hundreds = Math.floor(numQuery / 100) * 100;
      amountRange = { exact: numQuery, tens: tens, hundreds: hundreds };
    }

    // Search
    const matches = allPvs.filter(pv => {
      const applicant = String(pv.applicant_name || "").toLowerCase();
      const payee = String(pv.payee_name || "").toLowerCase();
      const purpose = String(pv.purpose || "").toLowerCase();
      const pvNo = String(pv.pv_no || "").toLowerCase();
      const dept = String(pv.dept || "").toLowerCase();
      const amount = parseFloat(pv.amount) || 0;

      // Text match
      if (applicant.includes(query) || payee.includes(query) || 
          purpose.includes(query) || pvNo.includes(query) || dept.includes(query)) {
        return true;
      }

      // Amount match (exact, or within 10s/100s range)
      if (amountRange) {
        if (amount === amountRange.exact) return true;
        if (Math.abs(amount - amountRange.exact) <= 10) return true;
        if (amountRange.hundreds > 0 && Math.abs(amount - amountRange.hundreds) <= 100) return true;
      }

      return false;
    });

    // Limit to 50 results and format
    const results = matches.slice(0, 50).map(pv => ({
      pv_no: pv.pv_no || "",
      date: pv.date ? formatDateStr_(pv.date) : "",
      applicant_name: pv.applicant_name || "",
      payee_name: pv.payee_name || "",
      amount: pv.amount || 0,
      status: pv.status || "",
      dept: pv.dept || ""
    }));

    return { ok: true, pvs: results, total: matches.length };
  } catch (e) {
    console.error("searchPVs error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   DEPARTMENT HEAD VERIFICATION
   ========================= */
function getDepartmentsForVerifier_(email) {
  const sh = ensureMinistryHeadsSheet_();
  const heads = rows_(sh);
  const key = String(email || "").toLowerCase().trim();
  return heads
    .filter(r => {
      const primary = String(r.Primary_Head_Email || r.primary_head_email || r.Head_Email || r.head_email || "").toLowerCase().trim();
      const secondary = String(r.Secondary_Head_Email || r.secondary_head_email || "").toLowerCase().trim();
      return primary === key || secondary === key;
    })
    .map(r => r.Department || r.department || r.Ministry || r.ministry || "")
    .filter(Boolean);
}

function listMinistryHeadAssignments() {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Not authorized" };
    }

    const sh = ensureMinistryHeadsSheet_();
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) {
      const empty = getCanonicalDepartments_().map((dept) => ({
        rowIndex: 0,
        department: dept,
        primary_name: "",
        primary_email: "",
        secondary_name: "",
        secondary_email: ""
      }));
      return { ok: true, assignments: empty };
    }
    const headers = data.shift();

    const idx = {
      Department: headers.indexOf("Department"),
      Legacy_Department: headers.indexOf("Ministry"),
      Primary_Name: headers.indexOf("Primary_Head_Name"),
      Primary_Email: headers.indexOf("Primary_Head_Email"),
      Secondary_Name: headers.indexOf("Secondary_Head_Name"),
      Secondary_Email: headers.indexOf("Secondary_Head_Email"),
      Legacy_Name: headers.indexOf("Head_Name"),
      Legacy_Email: headers.indexOf("Head_Email")
    };

    const rows = data.map((row, i) => ({
      rowIndex: i + 2,
      department: idx.Department >= 0 ? row[idx.Department] : (idx.Legacy_Department >= 0 ? row[idx.Legacy_Department] : ""),
      primary_name: idx.Primary_Name >= 0 ? row[idx.Primary_Name] : (idx.Legacy_Name >= 0 ? row[idx.Legacy_Name] : ""),
      primary_email: idx.Primary_Email >= 0 ? row[idx.Primary_Email] : (idx.Legacy_Email >= 0 ? row[idx.Legacy_Email] : ""),
      secondary_name: idx.Secondary_Name >= 0 ? row[idx.Secondary_Name] : "",
      secondary_email: idx.Secondary_Email >= 0 ? row[idx.Secondary_Email] : ""
    })).filter(r => r.department);

    const deptSet = new Set(rows.map(r => r.department));
    getCanonicalDepartments_().forEach((dept) => {
      if (!deptSet.has(dept)) {
        rows.push({
          rowIndex: 0,
          department: dept,
          primary_name: "",
          primary_email: "",
          secondary_name: "",
          secondary_email: ""
        });
      }
    });

    return { ok: true, assignments: rows };
    if (data.length <= 1) return { ok: true, assignments: [] };
    const headers = data.shift();

    const idx = {
      Ministry: headers.indexOf("Ministry"),
      Head_Email: headers.indexOf("Head_Email"),
      Head_Name: headers.indexOf("Head_Name"),
      Is_Supplementary: headers.indexOf("Is_Supplementary")
    };

    const assignments = data.map((row, i) => ({
      rowIndex: i + 2,
      ministry: idx.Ministry >= 0 ? row[idx.Ministry] : "",
      head_email: idx.Head_Email >= 0 ? row[idx.Head_Email] : "",
      head_name: idx.Head_Name >= 0 ? row[idx.Head_Name] : "",
      is_supplementary: String(idx.Is_Supplementary >= 0 ? row[idx.Is_Supplementary] : "").toUpperCase() === "YES"
    })).filter(r => r.ministry || r.head_email || r.head_name);

    return { ok: true, assignments: assignments };
  } catch (e) {
    console.error("listMinistryHeadAssignments error:", e);
    return { ok: false, error: e.message };
  }
}

function saveMinistryHeadAssignment(data) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Not authorized" };
    }

    const sh = ensureMinistryHeadsSheet_();
    const headers = getHeaders_(sh);
    const rowIndex = data && data.rowIndex ? Number(data.rowIndex) : 0;
    const updates = {
      Department: data.department || "",
      Primary_Head_Name: data.primary_name || "",
      Primary_Head_Email: data.primary_email || "",
      Secondary_Head_Name: data.secondary_name || "",
      Secondary_Head_Email: data.secondary_email || "",
      Ministry: data.ministry || "",
      Head_Email: data.head_email || "",
      Head_Name: data.head_name || "",
      Is_Supplementary: data.is_supplementary ? "YES" : "NO",
      Updated_At: now_()
    };

    if (rowIndex && rowIndex > 1) {
      const idx = rowIndex - 2;
      update_(sh, idx, updates);
    } else {
      const existing = rows_(sh).find(r => String(r.Department || "").trim() === updates.Department);
      if (existing) {
        updateRow_(sh, "Department", updates.Department, updates);
      } else {
        updates.Created_At = now_();
        append_(sh, updates);
      }
      updates.Created_At = now_();
      append_(sh, updates);
    }

    return { ok: true };
  } catch (e) {
    console.error("saveMinistryHeadAssignment error:", e);
    return { ok: false, error: e.message };
  }
}

function deleteMinistryHeadAssignment(rowIndex) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Not authorized" };
    }

    const sh = ensureMinistryHeadsSheet_();
    const idx = Number(rowIndex);
    if (!idx || idx < 2) return { ok: false, error: "Invalid row" };
    sh.deleteRow(idx);
    return { ok: true };
  } catch (e) {
    console.error("deleteMinistryHeadAssignment error:", e);
    return { ok: false, error: e.message };
  }
}

function deptHeadVerifyPV(pvNo, action, comment) {
  try {
    const ctx = getUserContext();
    
    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const pv = allPvs.find(p => p.pv_no === pvNo);

    if (!pv) return { ok: false, error: "PV not found" };

    // Check if user is the department head for this PV
    const deptHeadEmail = String(pv.dept_head_email || "").toLowerCase().trim();
    const userEmail = ctx.email.toLowerCase().trim();
    
    // Allow if user is the dept head OR Finance Admin (can verify on behalf)
    if (deptHeadEmail !== userEmail && !ctx.isFinanceAdmin) {
      return { ok: false, error: "You are not the department head for this PV" };
    }

    if (String(pv.status || "").toUpperCase() !== "PENDING_HEAD") {
      return { ok: false, error: "PV is not pending department head verification" };
    }

    let updates = {
      updated_at: now_(),
      head_comment: comment || ""
    };

    if (action === "VERIFY" || action === "APPROVED") {
      updates.status = "PENDING";
      updates.head_verified = "YES";
      updates.head_verified_at = now_();
      
      updateRow_(sh, "pv_no", pvNo, updates);
      log_("Dept head verified", { pv_no: pvNo, by: ctx.email });
      return { ok: true, status: "PENDING", message: "Verified. PV sent to Finance Admin for review." };
    } else if (action === "REJECT" || action === "REJECTED") {
      updates.status = "REJECTED_HEAD";
      updates.head_verified = "REJECTED";
      
      updateRow_(sh, "pv_no", pvNo, updates);
      log_("Dept head rejected", { pv_no: pvNo, by: ctx.email });
      return { ok: true, status: "REJECTED_HEAD", message: "PV rejected." };
    }

    return { ok: false, error: "Invalid action" };
  } catch (e) {
    console.error("deptHeadVerifyPV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   LIST PVs FOR DEPARTMENT HEAD
   ========================= */
function listPVsForDeptHead() {
  try {
    const ctx = getUserContext();
    const userEmail = ctx.email.toLowerCase().trim();

    const sh = getSheet_(SHEET_PVS);
    let pvs = rows_(sh);

    // Filter PVs where user is the department head
    pvs = pvs.filter(p => {
      const deptHeadEmail = String(p.dept_head_email || "").toLowerCase().trim();
      return deptHeadEmail === userEmail;
    });

    // Separate into pending and verified
    const pending = pvs.filter(p => String(p.status || "").toUpperCase() === "PENDING_HEAD");
    const verified = pvs.filter(p => p.head_verified === "YES");

    const mapPv = pv => {
      const approvals = jsonParse_(pv.approvals_json, []);
      const approvalCount = approvals.filter(a => a.decision === "APPROVED").length;
      return ({
      pv_no: pv.pv_no || "",
      date: pv.date ? String(pv.date) : "",
      applicant_name: pv.applicant_name || "",
      applicant_email: pv.applicant_email || "",
      dept: pv.dept || "",
      payee_name: pv.payee_name || "",
      amount: pv.amount || 0,
      status: pv.status || "",
      purpose: pv.purpose || "",
      head_verified: pv.head_verified || "",
      head_verified_at: pv.head_verified_at ? String(pv.head_verified_at) : "",
      dept_head_name: pv.dept_head_name || "",
      dept_head_email: pv.dept_head_email || "",
      approvals_count: approvalCount + "/2"
      });
    };

    return {
      ok: true,
      pending: pending.map(mapPv),
      verified: verified.map(mapPv)
    };
  } catch (e) {
    console.error("listPVsForDeptHead error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   MINISTRY HEAD VERIFICATION
   ========================= */
function listPVsForMinistryHead() {
  try {
    ensurePvColumns_();
    const ctx = getUserContext();
    const departments = getDepartmentsForVerifier_(ctx.email);
    if (!departments.length) {
      return { ok: false, error: "You are not assigned as a head verifier for any department." };
    }

    const sh = getSheet_(SHEET_PVS);
    let pvs = rows_(sh);

    pvs = pvs.filter(p => departments.includes(p.dept || ""));

    const pending = pvs.filter(p => {
      const verified = String(p.ministry_verified || "").toUpperCase() === "YES";
      const bypassed = String(p.ministry_verify_bypassed || "").toUpperCase() === "YES";
      return !verified && !bypassed;
    });
    const verified = pvs.filter(p => String(p.ministry_verified || "").toUpperCase() === "YES");

    const mapPv = pv => {
      const approvals = jsonParse_(pv.approvals_json, []);
      const approvalCount = approvals.filter(a => a.decision === "APPROVED").length;
      return ({
      pv_no: pv.pv_no || "",
      date: pv.date ? String(pv.date) : "",
      applicant_name: pv.applicant_name || "",
      applicant_email: pv.applicant_email || "",
      dept: pv.dept || "",
      payee_name: pv.payee_name || "",
      amount: pv.amount || 0,
      status: pv.status || "",
      purpose: pv.purpose || "",
      ministry_verified: pv.ministry_verified || "",
      ministry_verified_at: pv.ministry_verified_at ? String(pv.ministry_verified_at) : "",
      ministry_verify_bypassed: pv.ministry_verify_bypassed || "",
      approvals_count: approvalCount + "/2"
      });
    };

    return {
      ok: true,
      pending: pending.map(mapPv),
      verified: verified.map(mapPv)
    };
  } catch (e) {
    console.error("listPVsForMinistryHead error:", e);
    return { ok: false, error: e.message };
  }
}

function ministryHeadVerifyPV(pvNo, action, comment) {
  try {
    ensurePvColumns_();
    const ctx = getUserContext();
    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const pv = allPvs.find(p => p.pv_no === pvNo);

    if (!pv) return { ok: false, error: "PV not found" };

    const userEmail = String(ctx.email || "").toLowerCase().trim();
    const departments = getDepartmentsForVerifier_(ctx.email);
    const dept = pv.dept || "";
    if (!departments.includes(dept)) {
      return { ok: false, error: "You are not assigned as the head verifier for this PV" };
    }

    const updates = {
      updated_at: now_(),
      ministry_verified_comment: comment || ""
    };

    if (action === "VERIFY" || action === "APPROVED") {
      updates.ministry_verified = "YES";
      updates.ministry_verified_by = userEmail;
      updates.ministry_verified_at = now_();
      updates.ministry_verify_bypassed = "NO";
      updates.ministry_bypass_by = "";
      updates.ministry_bypass_at = "";
      updates.ministry_bypass_reason = "";
      updateRow_(sh, "pv_no", pvNo, updates);
      log_("Ministry verified", { pv_no: pvNo, by: ctx.email });
      return { ok: true, status: pv.status || "", message: "Head verification recorded." };
    } else if (action === "REJECT" || action === "REJECTED") {
      updates.ministry_verified = "REJECTED";
      updates.ministry_verified_by = userEmail;
      updates.ministry_verified_at = now_();
      updateRow_(sh, "pv_no", pvNo, updates);
      log_("Ministry rejected", { pv_no: pvNo, by: ctx.email });
      return { ok: true, status: pv.status || "", message: "Head rejection recorded." };
    }

    return { ok: false, error: "Invalid action" };
  } catch (e) {
    console.error("ministryHeadVerifyPV error:", e);
    return { ok: false, error: e.message };
  }
}

function adminBypassMinistryVerification(pvNo, reason) {
  try {
    ensurePvColumns_();
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized as Finance Admin" };
    }

    const updates = {
      ministry_verify_bypassed: "YES",
      ministry_bypass_by: ctx.email,
      ministry_bypass_at: now_(),
      ministry_bypass_reason: reason || "",
      updated_at: now_()
    };

    const sh = getSheet_(SHEET_PVS);
    const updated = updateRow_(sh, "pv_no", pvNo, updates);
    if (!updated) return { ok: false, error: "PV not found" };

    log_("Ministry verification bypassed", { pv_no: pvNo, by: ctx.email, reason: reason || "" });
    return { ok: true };
  } catch (e) {
    console.error("adminBypassMinistryVerification error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   LIST MY PVs
   ========================= */
function listMyPVsForCurrentUser() {
  try {
    let email = "";
    try {
      email = Session.getActiveUser().getEmail();
    } catch(e) {
      email = Session.getEffectiveUser().getEmail();
    }
    email = (email || "").toLowerCase().trim();
    
    console.log("listMyPVsForCurrentUser: Looking for email=" + email);

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    
    console.log("listMyPVsForCurrentUser: Total PVs in sheet=" + allPvs.length);

    const myPvs = allPvs.filter(r => {
      const pvEmail = String(r.applicant_email || "").toLowerCase().trim();
      const match = pvEmail === email;
      if (match) console.log("listMyPVsForCurrentUser: Found match - " + r.pv_no);
      return match;
    });

    console.log("listMyPVsForCurrentUser: Found " + myPvs.length + " PVs");

    // Parse JSON fields and convert dates to strings (dates break serialization!)
    const cleanPvs = myPvs.map(pv => {
      return {
        pv_no: pv.pv_no || "",
        date: pv.date ? String(pv.date) : "",
        status: pv.status || "",
        applicant_name: pv.applicant_name || "",
        applicant_email: pv.applicant_email || "",
        dept: pv.dept || "",
        project: pv.project || "",
        payee_name: pv.payee_name || "",
        payment_method: pv.payment_method || "",
        payee_bank_name: pv.payee_bank_name || "",
        payee_bank_acct: pv.payee_bank_acct || "",
        purpose: pv.purpose || "",
        amount: pv.amount || 0,
        line_items: jsonParse_(pv.line_items_json, []),
        attachments: jsonParse_(pv.attachments_json, []),
        approvals: jsonParse_(pv.approvals_json, []),
        created_at: pv.created_at ? String(pv.created_at) : "",
        signed_pdf_url: pv.signed_pdf_url || ""
      };
    });

    return { ok: true, pvs: cleanPvs };
  } catch (e) {
    console.error("listMyPVsForCurrentUser error:", e);
    return { ok: false, pvs: [], error: e.message };
  }
}

/* =========================
   LIST ALL PVs (Master View)
   ========================= */
function listAllPVs(filter) {
  try {
    console.log("listAllPVs: Starting, filter=" + JSON.stringify(filter));
    
    const ctx = getUserContext();
    console.log("listAllPVs: ctx.isFinanceAdmin=" + ctx.isFinanceAdmin + ", ctx.isSignatory=" + ctx.isSignatory);
    
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      console.log("listAllPVs: Not authorized");
      return { ok: false, error: "Not authorized", pvs: [] };
    }

    const sh = getSheet_(SHEET_PVS);
    let pvs = rows_(sh);
    console.log("listAllPVs: Total rows=" + pvs.length);

    // Apply status filter
    if (filter && filter.statuses && filter.statuses.length) {
      const statuses = filter.statuses.map(s => s.toUpperCase());
      pvs = pvs.filter(p => statuses.includes(String(p.status || "").toUpperCase()));
      console.log("listAllPVs: After filter=" + pvs.length);
    }

    // Sort by created_at descending (safely handle invalid dates)
    pvs.sort((a, b) => {
      const dateA = a.created_at ? new Date(a.created_at).getTime() : 0;
      const dateB = b.created_at ? new Date(b.created_at).getTime() : 0;
      return (dateB || 0) - (dateA || 0);
    });

    // Clean PVs - convert ALL values to safe serializable types
    const cleanPvs = [];
    for (let i = 0; i < pvs.length; i++) {
      try {
        const pv = pvs[i];
        cleanPvs.push({
          pv_no: String(pv.pv_no || ""),
          date: pv.date ? String(pv.date) : "",
          status: String(pv.status || ""),
          applicant_name: String(pv.applicant_name || ""),
          applicant_email: String(pv.applicant_email || ""),
          dept: String(pv.dept || ""),
          project: String(pv.project || ""),
          payee_name: String(pv.payee_name || ""),
          payment_method: String(pv.payment_method || ""),
          purpose: String(pv.purpose || ""),
          amount: Number(pv.amount) || 0,
          line_items: jsonParse_(pv.line_items_json, []),
          attachments: jsonParse_(pv.attachments_json, []),
          approvals: jsonParse_(pv.approvals_json, []),
          created_at: pv.created_at ? String(pv.created_at) : "",
          signed_pdf_url: String(pv.signed_pdf_url || "")
        });
      } catch (rowErr) {
        console.error("listAllPVs: Error processing row " + i + ": " + rowErr);
      }
    }

    console.log("listAllPVs: Returning " + cleanPvs.length + " clean PVs");
    return { ok: true, pvs: cleanPvs };
  } catch (e) {
    console.error("listAllPVs error:", e);
    return { ok: false, error: String(e.message || e), pvs: [] };
  }
}

/* =========================
   GET PV BY NUMBER
   ========================= */
function getPVByNo(pvNo, token) {
  try {
    ensurePvColumns_();
    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const pv = allPvs.find(p => p.pv_no === pvNo);

    if (!pv) return { ok: false, error: "PV not found" };

    // Parse JSON fields
    const lineItems = jsonParse_(pv.line_items_json, []);
    const attachments = jsonParse_(pv.attachments_json, []);
    const approvals = jsonParse_(pv.approvals_json, []);

    // Calculate approvals count
    const approvalCount = approvals.filter(a => a.decision === "APPROVED").length;

    // Clean PV - convert dates to strings
    const cleanPv = {
      pv_no: pv.pv_no || "",
      date: pv.date ? String(pv.date) : "",
      status: pv.status || "",
      applicant_name: pv.applicant_name || "",
      applicant_email: pv.applicant_email || "",
      dept: pv.dept || "",
      ministry: pv.ministry || "",
      project: pv.project || "",
      payee_name: pv.payee_name || "",
      payment_method: pv.payment_method || "",
      payee_bank_name: pv.payee_bank_name || "",
      payee_bank_acct: pv.payee_bank_acct || "",
      biller_code: pv.biller_code || "",
      ref_no: pv.ref_no || "",
      purpose: pv.purpose || "",
      amount: pv.amount || 0,
      line_items: lineItems,
      attachments: attachments,
      approvals: approvals,
      approvals_count: approvalCount + "/2",
      created_at: pv.created_at ? String(pv.created_at) : "",
      signed_pdf_url: pv.signed_pdf_url || "",
      ministry_verified: pv.ministry_verified || "",
      ministry_verified_by: pv.ministry_verified_by || "",
      ministry_verified_at: pv.ministry_verified_at ? String(pv.ministry_verified_at) : "",
      ministry_verified_comment: pv.ministry_verified_comment || "",
      ministry_verify_bypassed: pv.ministry_verify_bypassed || "",
      ministry_bypass_by: pv.ministry_bypass_by || "",
      ministry_bypass_at: pv.ministry_bypass_at ? String(pv.ministry_bypass_at) : "",
      ministry_bypass_reason: pv.ministry_bypass_reason || ""
    };

    return { ok: true, pv: cleanPv };
  } catch (e) {
    console.error("getPVByNo error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   ADMIN FUNCTIONS
   ========================= */
function listPVsForAdmin(filter) {
  try {
    const ctx = getUserContext();
    // Allow Signatories to view (but not act)
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Not authorized", pvs: [] };
    }

    const sh = getSheet_(SHEET_PVS);
    let pvs = rows_(sh);

    // Default to PENDING
    let statuses = ["PENDING"];
    if (filter && filter.statuses && filter.statuses.length) {
      statuses = filter.statuses.map(s => s.toUpperCase());
    }
    pvs = pvs.filter(p => statuses.includes(String(p.status || "").toUpperCase()));

    // Clean PVs - convert dates to strings
    const cleanPvs = pvs.map(pv => ({
      pv_no: pv.pv_no || "",
      date: pv.date ? String(pv.date) : "",
      status: pv.status || "",
      applicant_name: pv.applicant_name || "",
      dept: pv.dept || "",
      payee_name: pv.payee_name || "",
      amount: pv.amount || 0,
      line_items: jsonParse_(pv.line_items_json, []),
      attachments: jsonParse_(pv.attachments_json, [])
    }));

    return { ok: true, pvs: cleanPvs, canAct: ctx.isFinanceAdmin };
  } catch (e) {
    console.error("listPVsForAdmin error:", e);
    return { ok: false, error: e.message, pvs: [] };
  }
}

function adminSendToSignatories(pvNo, comment, visibility, bankAccount) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized as Finance Admin" };
    }

    // Bank account mappings
    const bankMappings = {
      "MAIN_MBB": { bank_name: "Maybank", bank_acct: "Main Account" },
      "BMC_MBB": { bank_name: "Maybank", bank_acct: "BMC Account" },
      "LCM_PBB": { bank_name: "Public Bank", bank_acct: "LCM Account" },
      "LSC_RHB": { bank_name: "RHB", bank_acct: "LSC Account" },
      "FELCMS_RHB": { bank_name: "RHB", bank_acct: "FELCMS Account" }
    };

    // Build update object
    const updateData = {
      status: "REVIEWED",
      admin_comment: comment || "",
      updated_at: now_()
    };

    // Add bank info if selected
    if (bankAccount && bankMappings[bankAccount]) {
      updateData.payer_bank_name = bankMappings[bankAccount].bank_name;
      updateData.payer_bank_acct = bankMappings[bankAccount].bank_acct;
      updateData.payer_account_code = bankAccount;
    }

    // Update status to REVIEWED
    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const pv = allPvs.find(r => r.pv_no === pvNo);
    if (!pv) return { ok: false, error: "PV not found" };
    const currentStatus = String(pv.status || "").toUpperCase();
    if (currentStatus === "APPROVED") {
      return { ok: false, error: "Payout account is locked after approval" };
    }

    const ministryRequired = String(pv.dept || "").trim() !== "";

    const ministryRequired = String(pv.ministry || "").trim() !== "";
    if (!ministryRequired) {
      updateData.ministry_verify_bypassed = "YES";
      updateData.ministry_bypass_by = ctx.email;
      updateData.ministry_bypass_at = now_();
      updateData.ministry_bypass_reason = "No department selected";
      updateData.ministry_bypass_reason = "No ministry selected";
    }

    const updated = updateRow_(sh, "pv_no", pvNo, updateData);

    if (!updated) return { ok: false, error: "PV not found" };

    log_("Admin reviewed", { pv_no: pvNo, by: ctx.email, decision: "REVIEWED", bank: bankAccount || "none" });
    return { ok: true, status: "REVIEWED" };
  } catch (e) {
    console.error("adminSendToSignatories error:", e);
    return { ok: false, error: e.message };
  }
}

function adminRejectPV(pvNo, comment, visibility) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized as Finance Admin" };
    }

    const sh = getSheet_(SHEET_PVS);
    const updated = updateRow_(sh, "pv_no", pvNo, {
      status: "REJECTED",
      admin_comment: comment || "",
      updated_at: now_()
    });

    if (!updated) return { ok: false, error: "PV not found" };

    log_("Admin rejected", { pv_no: pvNo, by: ctx.email });
    return { ok: true, status: "REJECTED" };
  } catch (e) {
    console.error("adminRejectPV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   DELETE PV
   ========================= */
function deletePV(pvNo) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.email) {
      return { ok: false, error: "Not authorized" };
    }

    const sh = getSheet_(SHEET_PVS);
    const headers = getHeaders_(sh);
    const data = sh.getDataRange().getValues();
    const pvNoIdx = headers.indexOf("pv_no");
    
    if (pvNoIdx === -1) return { ok: false, error: "pv_no column not found" };

    // Find and check ownership/permission
    for (let i = 1; i < data.length; i++) {
      if (data[i][pvNoIdx] === pvNo) {
        const emailIdx = headers.indexOf("applicant_email");
        const statusIdx = headers.indexOf("status");
        const rowEmail = String(data[i][emailIdx] || "").toLowerCase();
        const rowStatus = String(data[i][statusIdx] || "").toUpperCase();
        
        // Allow delete if: owner of pending PV, or Finance Admin for pending/reviewed
        const isOwner = rowEmail === ctx.email.toLowerCase();
        const canDelete = (isOwner && rowStatus === "PENDING") || 
                          (ctx.isFinanceAdmin && (rowStatus === "PENDING" || rowStatus === "REVIEWED"));
        
        if (!canDelete) {
          return { ok: false, error: "Cannot delete this PV. Only pending PVs can be deleted by owner, or pending/reviewed by Admin." };
        }
        
        sh.deleteRow(i + 1);
        log_("PV deleted", { pv_no: pvNo, by: ctx.email });
        return { ok: true };
      }
    }
    
    return { ok: false, error: "PV not found" };
  } catch (e) {
    console.error("deletePV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   CLEANUP INVALID PVs - Remove corrupted rows
   ========================= */
function cleanupInvalidPVs() {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized" };
    }

    const sh = getSheet_(SHEET_PVS);
    const headers = getHeaders_(sh);
    const data = sh.getDataRange().getValues();
    const statusIdx = headers.indexOf("status");
    
    if (statusIdx === -1) return { ok: false, error: "status column not found" };

    const validStatuses = ["PENDING", "REVIEWED", "APPROVED", "REJECTED"];
    const rowsToDelete = [];

    // Find invalid rows (go from bottom to avoid index shifting)
    for (let i = data.length - 1; i >= 1; i--) {
      const status = String(data[i][statusIdx] || "").toUpperCase().trim();
      if (!validStatuses.includes(status)) {
        rowsToDelete.push({ row: i + 1, status: status });
      }
    }

    // Delete rows from bottom up
    rowsToDelete.forEach(r => {
      sh.deleteRow(r.row);
    });

    log_("Cleanup invalid PVs", { deleted: rowsToDelete.length, by: ctx.email });
    return { ok: true, deleted: rowsToDelete.length, details: rowsToDelete };
  } catch (e) {
    console.error("cleanupInvalidPVs error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   UNREVIEW PV - Move back to PENDING
   ========================= */
function unreviewPV(pvNo) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized as Finance Admin" };
    }

    const sh = getSheet_(SHEET_PVS);
    const updated = updateRow_(sh, "pv_no", pvNo, {
      status: "PENDING",
      approvals_json: "[]",  // Clear any approvals
      admin_comment: "",
      updated_at: now_()
    });

    if (!updated) return { ok: false, error: "PV not found" };

    log_("PV unreviewed", { pv_no: pvNo, by: ctx.email });
    return { ok: true, status: "PENDING" };
  } catch (e) {
    console.error("unreviewPV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   REVOKE SIGNATORY APPROVAL - Remove one signatory's approval
   ========================= */
function revokeSignatoryApproval(pvNo, roleToRevoke) {
  try {
    const ctx = getUserContext();
    
    // Only the signatory themselves or Finance Admin can revoke
    const sigRole = ctx.signatoryRole || ctx.testRole;
    const isOwnRole = roleToRevoke && sigRole && roleToRevoke.toUpperCase() === sigRole.toUpperCase();
    
    if (!ctx.isFinanceAdmin && !isOwnRole) {
      return { ok: false, error: "Not authorized to revoke this approval" };
    }

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const pv = allPvs.find(p => p.pv_no === pvNo);

    if (!pv) return { ok: false, error: "PV not found" };
    
    const status = String(pv.status || "").toUpperCase();
    if (status !== "REVIEWED" && status !== "APPROVED") {
      return { ok: false, error: "Can only revoke from REVIEWED or APPROVED status" };
    }

    let approvals = jsonParse_(pv.approvals_json, []);
    
    // Remove the approval for this role
    const originalCount = approvals.length;
    approvals = approvals.filter(a => {
      if (!a.role) return true;
      return a.role.toUpperCase() !== roleToRevoke.toUpperCase();
    });
    
    if (approvals.length === originalCount) {
      return { ok: false, error: "No approval found for " + roleToRevoke };
    }

    // Update status based on remaining approvals
    const approvedCount = approvals.filter(a => a.decision === "APPROVED").length;
    const newStatus = approvedCount >= 2 ? "APPROVED" : "REVIEWED";

    const updated = updateRow_(sh, "pv_no", pvNo, {
      approvals_json: JSON.stringify(approvals),
      status: newStatus,
      signed_pdf_url: "", // Clear signed PDF since approvals changed
      updated_at: now_()
    });

    if (!updated) return { ok: false, error: "Update failed" };

    log_("Approval revoked", { pv_no: pvNo, role: roleToRevoke, by: ctx.email });
    return { ok: true, status: newStatus, approvals_count: approvedCount + "/2" };
  } catch (e) {
    console.error("revokeSignatoryApproval error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   UPDATE PV
   ========================= */
function updatePV(d) {
  try {
    const ctx = getUserContext();
    const pvNo = d.pv_no;
    
    if (!pvNo) {
      return { ok: false, error: "No PV number provided" };
    }

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const existing = allPvs.find(p => p.pv_no === pvNo);
    
    if (!existing) {
      return { ok: false, error: "PV not found" };
    }

    // Check permission - owner of pending, or admin for pending/reviewed
    const isOwner = String(existing.applicant_email || "").toLowerCase() === ctx.email.toLowerCase();
    const status = String(existing.status || "").toUpperCase();
    const canEdit = (isOwner && status === "PENDING") || 
                    (ctx.isFinanceAdmin && (status === "PENDING" || status === "REVIEWED"));

    if (!canEdit) {
      return { ok: false, error: "Cannot edit this PV" };
    }

    // Update the PV
    const updates = {
      date: d.pvDate || existing.date,
      applicant_name: d.applicant_name || existing.applicant_name,
      dept: d.dept || existing.dept,
      ministry: d.dept || existing.dept,
      project: d.project || existing.project,
      payee_name: d.payee_name || existing.payee_name,
      payment_method: d.payment_method || existing.payment_method,
      payee_bank_name: d.payee_bank_name || "",
      payee_bank_acct: d.payee_bank_acct || "",
      biller_code: d.biller_code || "",
      ref_no: d.ref_no || "",
      cheque_no: d.cheque_no || "",
      purpose: d.purpose || "",
      amount: d.amount || 0,
      line_items_json: JSON.stringify(d.line_items || []),
      attachments_json: JSON.stringify(d.attachments || []),
      updated_at: now_()
    };

    const updated = updateRow_(sh, "pv_no", pvNo, updates);
    
    if (!updated) {
      return { ok: false, error: "Update failed" };
    }

    log_("PV updated", { pv_no: pvNo, by: ctx.email });
    return { ok: true, pv_no: pvNo };
  } catch (e) {
    console.error("updatePV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   SIGNATORY FUNCTIONS
   ========================= */
function listPVsForSignatory(filter) {
  try {
    ensurePvColumns_();
    const ctx = getUserContext();
    // Allow Finance Admin to view (but not act)
    if (!ctx.isSignatory && !ctx.isFinanceAdmin) {
      return { ok: false, error: "Not authorized", pvs: [] };
    }

    const sh = getSheet_(SHEET_PVS);
    let pvs = rows_(sh);

    // Default to REVIEWED
    let statuses = ["REVIEWED"];
    if (filter && filter.statuses && filter.statuses.length) {
      statuses = filter.statuses.map(s => s.toUpperCase());
    }
    pvs = pvs.filter(p => statuses.includes(String(p.status || "").toUpperCase()));

    // Get current role for checking
    const sigRole = ctx.signatoryRole || ctx.testRole;

    // Clean PVs - convert dates to strings
    const cleanPvs = pvs.map(pv => {
      const approvals = jsonParse_(pv.approvals_json, []);
      const approvedRoles = approvals
        .filter(a => a.decision === "APPROVED")
        .map(a => a.role);
      
      const approvalCount = approvedRoles.length;
      
      // Check if current role has already acted (by role, not email)
      const currentUserApproved = sigRole ? approvals.some(a =>
        a.role && a.role.toUpperCase() === sigRole.toUpperCase()
      ) : false;

      // Build approval display text (e.g., "Bishop ✓, Treasurer ✓")
      const approvalDetails = approvedRoles.map(r => r + " ✓").join(", ");

      return {
        pv_no: pv.pv_no || "",
        date: pv.date ? String(pv.date) : "",
        status: pv.status || "",
        applicant_name: pv.applicant_name || "",
        dept: pv.dept || "",
        ministry: pv.ministry || "",
        payee_name: pv.payee_name || "",
        amount: pv.amount || 0,
        approvals_count: approvalCount + "/2",
        approval_details: approvalDetails,
        currentUserApproved: currentUserApproved,
        ministry_verified: pv.ministry_verified || "",
        ministry_verify_bypassed: pv.ministry_verify_bypassed || "",
        line_items: jsonParse_(pv.line_items_json, []),
        attachments: jsonParse_(pv.attachments_json, [])
      };
    });

    return { ok: true, pvs: cleanPvs, canAct: ctx.isSignatory };
  } catch (e) {
    console.error("listPVsForSignatory error:", e);
    return { ok: false, error: e.message, pvs: [] };
  }
}

function signatoryActOnPV(pvNo, action, comment, visibility) {
  try {
    const ctx = getUserContext();
    if (!ctx.isSignatory) {
      return { ok: false, error: "Not authorized as Signatory" };
    }

    const result = getPVByNo(pvNo);
    if (!result.ok) return result;

    const pv = result.pv;
    if (pv.status !== "REVIEWED") {
      return { ok: false, error: "PV is not in REVIEWED status" };
    }
    const ministryVerified = String(pv.ministry_verified || "").toUpperCase() === "YES";
    const ministryBypassed = String(pv.ministry_verify_bypassed || "").toUpperCase() === "YES";
    const ministryRequired = String(pv.dept || "").trim() !== "";
    if (ministryRequired && !ministryVerified && !ministryBypassed) {
      return { ok: false, error: "Head verification is required before signatory approval" };
    const ministryRequired = String(pv.ministry || "").trim() !== "";
    if (ministryRequired && !ministryVerified && !ministryBypassed) {
      return { ok: false, error: "Ministry verification is required before signatory approval" };
    }

    let approvals = pv.approvals || [];

    // Check if this ROLE has already acted (important for test mode)
    // In test mode, same email can have different roles
    const sigRole = ctx.signatoryRole || ctx.testRole;
    const alreadyActed = approvals.some(a => {
      // Check by role first (for test mode)
      if (a.role && sigRole && a.role.toUpperCase() === sigRole.toUpperCase()) {
        return true;
      }
      return false;
    });
    
    if (alreadyActed) {
      return { ok: false, error: sigRole + " has already acted on this PV" };
    }

    const sh = getSheet_(SHEET_PVS);

    if (action === "APPROVED") {
      approvals.push({
        email: ctx.email,
        role: sigRole,
        decision: "APPROVED",
        comment: comment || "",
        at: now_().toISOString()
      });

      const approvalCount = approvals.filter(a => a.decision === "APPROVED").length;
      const isFinal = approvalCount >= 2;

      const updates = {
        approvals_json: JSON.stringify(approvals),
        updated_at: now_()
      };

      if (isFinal) {
        updates.status = "APPROVED";
        // Generate signed PDF
        try {
          const pdfResult = generateSignedPV(pvNo);
          if (pdfResult.ok && pdfResult.signed_pdf_url) {
            updates.signed_pdf_url = pdfResult.signed_pdf_url;
          }
        } catch (e) { console.error("PDF generation error:", e); }
      }

      updateRow_(sh, "pv_no", pvNo, updates);

      log_("Signatory approved", { pv_no: pvNo, by: ctx.email, role: sigRole, final: isFinal });

      return {
        ok: true,
        approvals_count: approvalCount + "/2",
        final: isFinal,
        status: isFinal ? "APPROVED" : "REVIEWED"
      };

    } else if (action === "REJECTED") {
      approvals.push({
        email: ctx.email,
        role: sigRole,
        decision: "REJECTED",
        comment: comment || "",
        at: now_().toISOString()
      });

      updateRow_(sh, "pv_no", pvNo, {
        status: "REJECTED",
        approvals_json: JSON.stringify(approvals),
        updated_at: now_()
      });

      log_("Signatory rejected", { pv_no: pvNo, by: ctx.email, role: ctx.signatoryRole });

      return { ok: true, status: "REJECTED" };
    }

    return { ok: false, error: "Invalid action" };
  } catch (e) {
    console.error("signatoryActOnPV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   SIGNED PDF GENERATION
   ========================= */
function generateSignedPV(pvNo) {
  try {
    const result = getPVByNo(pvNo);
    if (!result.ok) return result;

    const pv = result.pv;
    const sigResult = getRoleSignatures();
    const signatures = sigResult.signatures || {};

    const html = buildSignedPvHtml_(pv, signatures);

    const blob = Utilities.newBlob(html, "text/html", "pv.html");
    const pdf = blob.getAs("application/pdf");
    pdf.setName(pvNo + "-SIGNED.pdf");

    const file = DriveApp.createFile(pdf);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const pdfUrl = file.getUrl();

    // Update the PV with PDF URL
    const sh = getSheet_(SHEET_PVS);
    updateRow_(sh, "pv_no", pvNo, { signed_pdf_url: pdfUrl });

    return { ok: true, signed_pdf_url: pdfUrl };
  } catch (e) {
    console.error("generateSignedPV error:", e);
    return { ok: false, error: e.message };
  }
}

function buildSignedPvHtml_(pv, signatures) {
  const lineItems = pv.line_items || [];
  const approvals = pv.approvals || [];

  // Format date as dd/mm/yyyy
  function formatDateDDMMYYYY(dateStr) {
    if (!dateStr) return "";
    try {
      var d = new Date(dateStr);
      if (isNaN(d.getTime())) return String(dateStr).substring(0, 10);
      var day = ("0" + d.getDate()).slice(-2);
      var month = ("0" + (d.getMonth() + 1)).slice(-2);
      var year = d.getFullYear();
      return day + "/" + month + "/" + year;
    } catch (e) {
      return String(dateStr).substring(0, 10);
    }
  }

  // Convert number to words (Malaysian Ringgit format)
  function numberToWords(num) {
    const ones = ['', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE', 'TEN', 'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'];
    const tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'];
    
    if (num === 0) return 'ZERO';
    
    function convertHundreds(n) {
      let result = '';
      if (n >= 100) {
        result += ones[Math.floor(n / 100)] + ' HUNDRED ';
        n %= 100;
      }
      if (n >= 20) {
        result += tens[Math.floor(n / 10)] + ' ';
        n %= 10;
      }
      if (n > 0) {
        result += ones[n] + ' ';
      }
      return result;
    }
    
    let result = '';
    const million = Math.floor(num / 1000000);
    const thousand = Math.floor((num % 1000000) / 1000);
    const remainder = Math.floor(num % 1000);
    const cents = Math.round((num % 1) * 100);
    
    if (million > 0) result += convertHundreds(million) + 'MILLION ';
    if (thousand > 0) result += convertHundreds(thousand) + 'THOUSAND ';
    if (remainder > 0) result += convertHundreds(remainder);
    
    result = 'RINGGIT MALAYSIA ' + result.trim();
    if (cents > 0) {
      result += ' AND CENTS ' + convertHundreds(cents).trim();
    }
    result += ' ONLY';
    
    return result;
  }

  // Calculate total
  let total = 0;
  lineItems.forEach(li => {
    total += parseFloat(li.amount) || 0;
  });

  // Build line items
  let paymentForHtml = "";
  lineItems.forEach((li, idx) => {
    const amt = parseFloat(li.amount) || 0;
    const amtStr = amt.toFixed(2);
    paymentForHtml += '<tr>' +
      '<td style="padding:5px 10px;font-size:11px;text-align:center;border-bottom:1px solid #ddd;">' + (idx + 1) + '</td>' +
      '<td style="padding:5px 10px;font-size:11px;border-bottom:1px solid #ddd;">' + (li.description || '') + '</td>' +
      '<td style="padding:5px 10px;font-size:11px;text-align:right;border-bottom:1px solid #ddd;">' + amtStr + '</td>' +
      '</tr>';
  });

  // Format total with commas
  const totalFormatted = total.toLocaleString('en-MY', {minimumFractionDigits: 2, maximumFractionDigits: 2});

  // Payment method - determine which columns to show
  const paymentMethod = pv.payment_method || "Bank transfer";
  let paymentDetailsHtml = "";
  
  // Build payment details row based on method
  let paymentHeaders = '<th style="border:1px solid #000;padding:8px;font-size:10px;background:#f5f5f5;">Date</th>';
  let paymentValues = '<td style="border:1px solid #000;padding:8px;font-size:10px;">' + formatDateDDMMYYYY(pv.date) + '</td>';
  
  if (paymentMethod === "Cheque" && pv.cheque_no) {
    paymentHeaders += '<th style="border:1px solid #000;padding:8px;font-size:10px;background:#f5f5f5;">Cheque No.</th>';
    paymentValues += '<td style="border:1px solid #000;padding:8px;font-size:10px;">' + pv.cheque_no + '</td>';
  }
  
  if (paymentMethod === "JomPAY") {
    paymentHeaders += '<th style="border:1px solid #000;padding:8px;font-size:10px;background:#f5f5f5;">JomPAY</th>';
    paymentValues += '<td style="border:1px solid #000;padding:8px;font-size:10px;">Biller Code: ' + (pv.biller_code || '-') + '<br>Ref No: ' + (pv.ref_no || '-') + '</td>';
  }
  
  if (paymentMethod === "Bank transfer" && pv.payee_bank_name) {
    paymentHeaders += '<th style="border:1px solid #000;padding:8px;font-size:10px;background:#f5f5f5;">Bank</th>';
    paymentValues += '<td style="border:1px solid #000;padding:8px;font-size:10px;">' + pv.payee_bank_name + '<br>' + (pv.payee_bank_acct || '') + '</td>';
  }
  
  if (paymentMethod === "Cash") {
    paymentHeaders += '<th style="border:1px solid #000;padding:8px;font-size:10px;background:#f5f5f5;">Cash</th>';
    paymentValues += '<td style="border:1px solid #000;padding:8px;font-size:10px;">Cash Payment</td>';
  }
  
  // Always show payment amount
  paymentHeaders += '<th style="border:1px solid #000;padding:8px;font-size:10px;background:#f5f5f5;text-align:right;">Payment Amount (RM)</th>';
  paymentValues += '<td style="border:1px solid #000;padding:8px;font-size:10px;text-align:right;">' + totalFormatted + '</td>';

  // Set Bank Payout display
  let setBankDisplay = "";
  if (pv.payer_bank_name) {
    setBankDisplay = pv.payer_bank_name;
    if (pv.payer_account_code) setBankDisplay += " - " + pv.payer_account_code;
  }
  const payFromValue = setBankDisplay || "(Not set)";
  const payFromValue = setBankDisplay || "NOT SPECIFIED";
  const payFromBoxHtml = '<div style="text-align:center;">' +
    '<div style="border:2px solid #000;padding:6px 12px;font-size:12px;font-weight:bold;">PAY FROM ACCOUNT</div>' +
    '<div style="border:1px solid #000;border-top:none;padding:6px 12px;font-size:10px;">' + payFromValue + '</div>' +
    '</div>';
  const statusValue = String(pv.status || "").toUpperCase();
  const statusLabel = statusValue === "APPROVED" ? "APPROVED" :
    (statusValue === "REVIEWED" ? "REVIEWED" :
      (statusValue === "REJECTED" ? "REJECTED" : "PENDING"));
  const statusBoxHtml = '<div style="margin-top:8px;border:1px solid #000;padding:6px 10px;font-size:10px;font-weight:bold;text-align:center;">STATUS: ' + statusLabel + '</div>';

  // Department head verification section
  let deptHeadHtml = "";
  const deptHeadName = pv.dept_head_name || "";
  const deptName = pv.dept || "";
  const headVerified = pv.head_verified || "";
  
  if (deptHeadName || headVerified === "YES") {
    const headSig = signatures["DEPT_HEAD"] || {};
    deptHeadHtml = '<div style="margin-top:30px;">' +
      '<div style="font-weight:bold;font-size:11px;margin-bottom:15px;border-bottom:1px solid #000;padding-bottom:5px;">VERIFIED BY:</div>' +
      '<div style="display:flex;justify-content:flex-start;gap:80px;">' +
      '<div style="text-align:center;width:200px;">' +
      '<div style="height:60px;display:flex;align-items:flex-end;justify-content:center;">' +
      (headSig.file_url ? '<img src="' + headSig.file_url + '" style="max-height:50px;max-width:150px;">' : '') +
      '</div>' +
      '<div style="border-top:1px solid #000;margin-top:5px;padding-top:5px;">' +
      '<div style="font-size:11px;">' + (deptHeadName || 'Department Head') + '</div>' +
      '<div style="font-size:10px;color:#666;">(' + deptName + ')</div>' +
      '<div style="font-size:9px;color:#999;">Head of Department</div>' +
      '</div></div>' +
      '</div></div>';
  }

  const ministryVerified = String(pv.ministry_verified || "").toUpperCase() === "YES";
  const ministryBypassed = String(pv.ministry_verify_bypassed || "").toUpperCase() === "YES";
  const ministryRequired = String(pv.dept || "").trim() !== "";
  const ministryRequired = String(pv.ministry || "").trim() !== "";
  const ministryStatus = ministryVerified ? "VERIFIED" : (ministryBypassed ? "BYPASSED" : "PENDING");
  const ministryBy = ministryVerified ? (pv.ministry_verified_by || "") : (ministryBypassed ? (pv.ministry_bypass_by || "") : "");
  const ministryAt = ministryVerified ? (pv.ministry_verified_at || "") : (ministryBypassed ? (pv.ministry_bypass_at || "") : "");
  const ministryComment = ministryVerified ? (pv.ministry_verified_comment || "") : (ministryBypassed ? (pv.ministry_bypass_reason || "") : "");
  const ministrySigKey = pv.ministry_verified_by ? "MINISTRY_HEAD|" + pv.ministry_verified_by : "MINISTRY_HEAD";
  const ministrySig = signatures[ministrySigKey] || signatures["MINISTRY_HEAD"] || signatures["DEPT_HEAD"] || {};
  const ministrySig = signatures["MINISTRY_HEAD"] || signatures["DEPT_HEAD"] || {};
  const ministryVerifierLabel = ministryVerified ? (ministryBy || ministrySig.name || "Ministry Head") : (ministryBypassed ? "BYPASSED" : "PENDING");

  const ministryHtml = '<div style="margin-top:25px;">' +
    '<div style="font-weight:bold;font-size:11px;margin-bottom:8px;border-bottom:1px solid #000;padding-bottom:5px;">VERIFIED BY (MINISTRY HEAD) - ' + ministryStatus + '</div>' +
    '<div style="display:flex;align-items:flex-start;gap:60px;">' +
    '<div style="text-align:center;width:200px;">' +
    '<div style="height:55px;display:flex;align-items:flex-end;justify-content:center;">' +
    (ministryVerified && ministrySig.file_url ? '<img src="' + ministrySig.file_url + '" style="max-height:50px;max-width:150px;">' : (ministryVerified || ministryBypassed ? '<span style="font-size:18px;color:#16a34a;">✓</span>' : '<span style="color:#999;font-size:10px;">Pending</span>')) +
    '</div>' +
    '<div style="border-top:1px solid #000;margin-top:5px;padding-top:5px;">' +
    '<div style="font-size:11px;">' + ministryVerifierLabel + '</div>' +
    '<div style="font-size:9px;color:#666;">(' + (ministryRequired ? (pv.dept || "") : "No department selected") + ')</div>' +
    '<div style="font-size:9px;color:#666;">(' + (ministryRequired ? (pv.ministry || pv.dept || "") : "No ministry selected") + ')</div>' +
    '</div></div>' +
    '<div style="font-size:10px;line-height:1.4;">' +
    (ministryBy ? '<div><strong>By:</strong> ' + ministryBy + '</div>' : '') +
    (ministryAt ? '<div><strong>At:</strong> ' + formatDateDDMMYYYY(ministryAt) + '</div>' : '') +
    (ministryComment ? '<div><strong>Comment:</strong> ' + ministryComment + '</div>' : '') +
    '</div></div></div>';

  // Build signatures - only the 2 signatories who approved
  let signaturesHtml = "";
  const approvedSignatories = approvals.filter(a => a.decision === "APPROVED");
  
  for (let i = 0; i < 2; i++) {
    const approval = approvedSignatories[i];
    if (approval) {
      const role = approval.role || "";
      const sig = signatures[role] || {};
      const sigName = sig.name || role;
      
      signaturesHtml += '<div style="text-align:center;width:200px;">' +
        '<div style="height:60px;display:flex;align-items:flex-end;justify-content:center;">' +
        (sig.file_url ? '<img src="' + sig.file_url + '" style="max-height:50px;max-width:150px;">' : '') +
        '</div>' +
        '<div style="border-top:1px solid #000;margin-top:5px;padding-top:5px;">' +
        '<div style="font-size:11px;font-weight:bold;">' + sigName + '</div>' +
        '<div style="font-size:10px;color:#666;">' + role + '</div>' +
        '</div></div>';
    }
  }

  // Logo URL
  const logoUrl = "https://www.lutheran.org.my/wp-content/uploads/2018/09/LCM-Logo-120px.png";

  // Build full HTML
  return '<!DOCTYPE html>' +
    '<html><head><meta charset="utf-8">' +
    '<style>' +
    '@page { margin: 12mm; size: A4; }' +
    'body { font-family: Arial, sans-serif; font-size: 11px; margin: 0; padding: 20px 30px; color: #000; line-height: 1.3; }' +
    '</style></head><body>' +
    
    // Header with logo
    '<table style="width:100%;margin-bottom:15px;"><tr>' +
    '<td style="width:80px;vertical-align:top;">' +
    '<img src="' + logoUrl + '" style="width:70px;" onerror="this.style.display=\'none\'">' +
    '</td><td style="vertical-align:top;padding-left:10px;">' +
    '<div style="font-size:16px;font-weight:bold;">LUTHERAN CHURCH IN MALAYSIA <span style="font-size:9px;font-weight:normal;color:#666;">(ROS: PPM-001-10-09031964)</span></div>' +
    '<div style="font-size:10px;margin-top:2px;">Luther Centre, No. 6, Jalan Utara, 46200 Petaling Jaya, Selangor, Malaysia</div>' +
    '<div style="font-size:10px;">Tel: 03-7956 5992 &nbsp;&nbsp; Fax: 03-7957 6953 &nbsp;&nbsp; Whatsapp: 011-10831561 (Finance)</div>' +
    '<div style="font-size:10px;">Email: hq@lcm.org.my or finance@lcm.org.my</div>' +
    '</td></tr></table>' +
    
    // Title with Set Bank
    '<table style="width:100%;margin:15px 0;"><tr>' +
    '<td style="text-align:left;">' +
    '<span style="font-size:18px;font-weight:bold;text-decoration:underline;">PAYMENT VOUCHER</span>' +
    '</td>' +
    '<td style="text-align:right;">' +
    payFromBoxHtml +
    statusBoxHtml +
    '</td></tr></table>' +
    
    // PAY TO and Voucher Info
    '<table style="width:100%;margin-bottom:10px;"><tr>' +
    '<td style="width:70%;vertical-align:top;">' +
    '<table style="width:100%;">' +
    '<tr><td style="font-size:11px;width:180px;vertical-align:top;"><strong>PAY TO:</strong></td>' +
    '<td style="font-size:12px;font-weight:bold;">' + (pv.payee_name || '').toUpperCase() + '</td></tr>' +
    '<tr><td style="font-size:11px;vertical-align:top;"><strong>DEPARTMENT:</strong></td>' +
    '<td style="font-size:11px;">' + (pv.dept || '') + '</td></tr>' +
    (pv.project ? '<tr><td style="font-size:11px;vertical-align:top;"><strong>PROJECT:</strong></td><td style="font-size:11px;">' + pv.project + '</td></tr>' : '') +
    '</table></td>' +
    '<td style="text-align:right;font-size:11px;vertical-align:top;">' +
    'Voucher No : <strong>' + (pv.pv_no || '') + '</strong><br>' +
    'Date : ' + formatDateDDMMYYYY(pv.date) +
    '</td></tr></table>' +
    
    // Pay the sum of
    '<table style="width:100%;margin:15px 0;"><tr>' +
    '<td style="font-size:11px;width:100px;vertical-align:top;"><strong>PAY THE<br>SUM OF:</strong></td>' +
    '<td style="font-size:11px;font-style:italic;text-decoration:underline;">' + numberToWords(total) + '</td>' +
    '</tr></table>' +
    
    // Payment details table (only relevant columns)
    '<table style="border-collapse:collapse;margin:15px 0;width:100%;">' +
    '<tr>' + paymentHeaders + '</tr>' +
    '<tr>' + paymentValues + '</tr>' +
    '</table>' +
    
    // Payment for (line items)
    '<div style="margin-top:20px;">' +
    '<div style="font-size:11px;font-style:italic;margin-bottom:5px;">Payment for</div>' +
    '<table style="width:100%;border-collapse:collapse;">' +
    '<tr style="background:#f5f5f5;">' +
    '<th style="padding:8px;font-size:10px;border:1px solid #000;width:40px;">No</th>' +
    '<th style="padding:8px;font-size:10px;border:1px solid #000;text-align:left;">Description</th>' +
    '<th style="padding:8px;font-size:10px;border:1px solid #000;text-align:right;width:100px;">Amount</th>' +
    '</tr>' +
    paymentForHtml +
    '</table></div>' +
    
    // Total
    '<div style="text-align:right;margin-top:15px;">' +
    '<span style="font-size:12px;"><strong>Total:</strong> &nbsp; <strong>RM ' + totalFormatted + '</strong></span>' +
    '</div>' +
    
    // Department Head Verification (VERIFIED BY)
    deptHeadHtml +

    // Ministry Verification
    ministryHtml +
    
    // Signatories (APPROVED BY)
    '<div style="margin-top:30px;">' +
    '<div style="font-weight:bold;font-size:11px;margin-bottom:15px;border-bottom:1px solid #000;padding-bottom:5px;">APPROVED BY:</div>' +
    '<div style="display:flex;justify-content:space-between;padding:0 50px;">' +
    signaturesHtml +
    '</div></div>' +
    
    // Footer
    '<div style="margin-top:30px;text-align:center;font-size:8px;color:#999;border-top:1px solid #ddd;padding-top:8px;">' +
    'Generated by LCM Finance System on ' + formatDateDDMMYYYY(new Date()) +
    '</div>' +
    
    '</body></html>';
}

/* =========================
   GET PV FOR PREVIEW (with signatures info)
   ========================= */
function getPVForPreview(pvNo) {
  try {
    const ctx = getUserContext();
    if (!pvNo) return { ok: false, error: "PV number required" };

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const pv = allPvs.find(r => r.pv_no === pvNo);
    
    if (!pv) return { ok: false, error: "PV not found" };

    // Get signatures
    const sigResult = getRoleSignatures();
    const signatures = sigResult.signatures || {};

    // Add dept head signature if available
    if (pv.dept_head_name) {
      signatures["DEPT_HEAD"] = {
        name: pv.dept_head_name,
        file_url: "", // Would need to be stored separately
        email: pv.dept_head_email || ""
      };
    }

    // Parse line items
    let lineItems = [];
    try {
      lineItems = typeof pv.line_items === "string" ? JSON.parse(pv.line_items) : (pv.line_items || []);
    } catch (e) {
      lineItems = [];
    }

    // Parse approvals
    let approvals = [];
    try {
      approvals = typeof pv.approvals === "string" ? JSON.parse(pv.approvals) : (pv.approvals || []);
    } catch (e) {
      approvals = [];
    }

    return {
      ok: true,
      pv: {
        pv_no: pv.pv_no,
        date: pv.date,
        applicant_name: pv.applicant_name,
        dept: pv.dept,
        payee_name: pv.payee_name,
        purpose: pv.purpose,
        project: pv.project || "",
        amount: pv.amount,
        status: pv.status,
        payment_method: pv.payment_method,
        payee_bank_name: pv.payee_bank_name,
        payee_bank_acct: pv.payee_bank_acct,
        biller_code: pv.biller_code,
        ref_no: pv.ref_no,
        cheque_no: pv.cheque_no,
        payer_bank_name: pv.payer_bank_name,
        payer_account_code: pv.payer_account_code,
        line_items: lineItems,
        approvals: approvals,
        approvals_count: pv.approvals_count || "0/2",
        dept_head_name: pv.dept_head_name || "",
        dept_head_email: pv.dept_head_email || "",
        head_verified: pv.head_verified || "",
        ministry: pv.ministry || "",
        ministry_verified: pv.ministry_verified || "",
        ministry_verified_by: pv.ministry_verified_by || "",
        ministry_verified_at: pv.ministry_verified_at || "",
        ministry_verified_comment: pv.ministry_verified_comment || "",
        ministry_verify_bypassed: pv.ministry_verify_bypassed || "",
        ministry_bypass_by: pv.ministry_bypass_by || "",
        ministry_bypass_at: pv.ministry_bypass_at || "",
        ministry_bypass_reason: pv.ministry_bypass_reason || ""
      },
      signatures: signatures
    };
  } catch (e) {
    console.error("getPVForPreview error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   CANCEL PV
   ========================= */
function cancelPV(pvNo, reason) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin && !ctx.isSignatory) {
      return { ok: false, error: "Only Finance Admin or Signatories can cancel PVs" };
    }

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const idx = allPvs.findIndex(r => r.pv_no === pvNo);
    
    if (idx < 0) return { ok: false, error: "PV not found" };

    const pv = allPvs[idx];
    const currentStatus = String(pv.status || "").toUpperCase();
    
    if (currentStatus === "CANCELLED") {
      return { ok: false, error: "PV is already cancelled" };
    }

    // Update status to CANCELLED
    const updates = {
      status: "CANCELLED",
      cancelled_at: now_(),
      cancelled_by: ctx.email,
      cancelled_reason: reason || "",
      updated_at: now_()
    };

    update_(sh, idx, updates);
    
    return { ok: true, message: "PV cancelled successfully" };
  } catch (e) {
    console.error("cancelPV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   REVOKE APPROVED PV
   ========================= */
function revokePV(pvNo, reason) {
  try {
    const ctx = getUserContext();
    if (!ctx.isFinanceAdmin) {
      return { ok: false, error: "Only Finance Admin can revoke approved PVs" };
    }

    const sh = getSheet_(SHEET_PVS);
    const allPvs = rows_(sh);
    const idx = allPvs.findIndex(r => r.pv_no === pvNo);
    
    if (idx < 0) return { ok: false, error: "PV not found" };

    const pv = allPvs[idx];
    const currentStatus = String(pv.status || "").toUpperCase();
    
    if (currentStatus !== "APPROVED") {
      return { ok: false, error: "Only APPROVED PVs can be revoked" };
    }

    // Reset status to PENDING and clear approvals
    const updates = {
      status: "PENDING",
      approvals_json: "[]",
      approvals_count: "0/2",
      head_verified: pv.dept_head_name ? "NO" : "N/A",
      head_verified_at: "",
      ministry_verified: "",
      ministry_verified_by: "",
      ministry_verified_at: "",
      ministry_verified_comment: "",
      ministry_verify_bypassed: "NO",
      ministry_bypass_by: "",
      ministry_bypass_at: "",
      ministry_bypass_reason: "",
      revoked_at: now_(),
      revoked_by: ctx.email,
      revoked_reason: reason || "",
      signed_pdf_url: "",
      updated_at: now_()
    };

    update_(sh, idx, updates);
    
    return { ok: true, message: "PV revoked and reset to Pending status" };
  } catch (e) {
    console.error("revokePV error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   SHARE LINKS
   ========================= */
function createShareLink(pvNo) {
  try {
    const token = Utilities.getUuid().replace(/-/g, "");
    const ctx = getUserContext();

    // Save to Shares sheet
    const sh = ss_().getSheetByName(SHEET_SHARES);
    if (sh) {
      append_(sh, {
        token: token,
        pv_no: pvNo,
        created_by: ctx.email,
        created_at: now_(),
        expires_at: ""
      });
    }

    const url = ScriptApp.getService().getUrl() + "?page=claim&pv=" + pvNo + "&token=" + token;
    return { ok: true, url: url, token: token };
  } catch (e) {
    console.error("createShareLink error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   RECURRING
   Columns: id, name, payee_name, dept, project, amount, frequency, next_date, notes, created_by_email, created_at, updated_at
   ========================= */
function saveRecurringExpense(p) {
  try {
    const ctx = getUserContext();
    const sh = getSheet_(SHEET_RECURRING);
    
    const success = append_(sh, {
      id: id_("REC"),
      name: p.name || "",
      payee_name: p.payee_name || "",
      dept: p.dept || "",
      project: p.project || "",
      amount: p.amount || 0,
      frequency: p.frequency || "MONTHLY",
      next_date: p.next_date || "",
      notes: p.notes || "",
      created_by_email: ctx.email,
      created_at: now_(),
      updated_at: now_()
    });

    if (!success) return { ok: false, error: "Failed to save" };
    return { ok: true };
  } catch (e) {
    console.error("saveRecurringExpense error:", e);
    return { ok: false, error: e.message };
  }
}

function listRecurringExpenses() {
  try {
    const sh = ss_().getSheetByName(SHEET_RECURRING);
    if (!sh) return [];
    return rows_(sh);
  } catch (e) {
    console.error("listRecurringExpenses error:", e);
    return [];
  }
}

function generatePVFromRecurring(id) {
  try {
    const sh = ss_().getSheetByName(SHEET_RECURRING);
    if (!sh) return { ok: false, error: "Recurring sheet not found" };

    const r = rows_(sh).find(row => row.id === id);
    if (!r) return { ok: false, error: "Recurring expense not found" };

    return submitPublicPV({
      pvDate: r.next_date || new Date().toISOString().slice(0, 10),
      applicant_name: "",
      dept: r.dept || "",
      project: r.project || "",
      payee_name: r.payee_name || "",
      purpose: r.name || "",
      payment_method: "Bank transfer",
      line_items: [{ date: r.next_date || "", description: r.name || "", amount: r.amount || 0 }],
      attachments: [],
      amount: r.amount || 0,
      sig_applicant_confirm: "YES",
      sig_head_confirm: "YES"
    });
  } catch (e) {
    console.error("generatePVFromRecurring error:", e);
    return { ok: false, error: e.message };
  }
}

function deleteRecurringExpense(id) {
  try {
    const sh = getSheet_(SHEET_RECURRING);
    const headers = getHeaders_(sh);
    const data = sh.getDataRange().getValues();
    const idIdx = headers.indexOf("id");
    if (idIdx === -1) return { ok: false, error: "id column not found" };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === id) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Not found" };
  } catch (e) {
    console.error("deleteRecurringExpense error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   FAVOURITES
   Columns: id, name, owner_email, payee, dept, method, amount, payload_json, created_at, updated_at
   ========================= */
function saveFavouriteTemplate(f) {
  try {
    const ctx = getUserContext();
    const sh = getSheet_(SHEET_FAVES);
    const payload = f.payload || {};

    const success = append_(sh, {
      id: id_("FAV"),
      name: f.name || "",
      owner_email: ctx.email,
      payee: payload.payee_name || "",
      dept: payload.dept || "",
      method: payload.payment_method || "",
      amount: payload.amount || 0,
      payload_json: JSON.stringify(payload),
      created_at: now_(),
      updated_at: now_()
    });

    if (!success) return { ok: false, error: "Failed to save" };
    return { ok: true };
  } catch (e) {
    console.error("saveFavouriteTemplate error:", e);
    return { ok: false, error: e.message };
  }
}

function listFavourites() {
  try {
    const email = getUserContext().email.toLowerCase();
    const sh = ss_().getSheetByName(SHEET_FAVES);
    if (!sh) return [];

    return rows_(sh)
      .filter(r => String(r.owner_email || "").toLowerCase() === email)
      .map(r => ({
        id: r.id,
        name: r.name,
        payee: r.payee || "",
        dept: r.dept || "",
        method: r.method || "",
        amount: r.amount || 0
      }));
  } catch (e) {
    console.error("listFavourites error:", e);
    return [];
  }
}

function getFavouriteById(id) {
  try {
    const sh = ss_().getSheetByName(SHEET_FAVES);
    if (!sh) return { ok: false, error: "Favourites sheet not found" };

    const r = rows_(sh).find(row => row.id === id);
    if (!r) return { ok: false, error: "Favourite not found" };

    return {
      ok: true,
      favourite: {
        id: r.id,
        name: r.name,
        payload: jsonParse_(r.payload_json, {})
      }
    };
  } catch (e) {
    console.error("getFavouriteById error:", e);
    return { ok: false, error: e.message };
  }
}

function deleteFavourite(id) {
  try {
    const sh = getSheet_(SHEET_FAVES);
    const headers = getHeaders_(sh);
    const data = sh.getDataRange().getValues();
    const idIdx = headers.indexOf("id");
    if (idIdx === -1) return { ok: false, error: "id column not found" };

    for (let i = 1; i < data.length; i++) {
      if (data[i][idIdx] === id) {
        sh.deleteRow(i + 1);
        return { ok: true };
      }
    }
    return { ok: false, error: "Not found" };
  } catch (e) {
    console.error("deleteFavourite error:", e);
    return { ok: false, error: e.message };
  }
}

/* =========================
   SIGNATURES
   Columns in your sheet: email, role, file_id, file_url, updated_at
   ========================= */
function ensureSignatureSheet_() {
  const sh = ensureSheet_(SHEET_SIGNATURES, ["email", "role", "file_id", "file_url", "name", "updated_at"]);
  const headers = getHeaders_(sh);
  const required = ["email", "role", "file_id", "file_url", "name", "updated_at"];
  const missing = required.filter(h => !headers.includes(h));
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function saveRoleSignature(role, name, dataUrl) {
  try {
    const ctx = getUserContext();
    const allowed = {
      FINANCE_ADMIN: ctx.isFinanceAdmin,
      BISHOP: ctx.isBishop,
      TREASURER: ctx.isTreasurer,
      SECRETARY: ctx.isSecretary,
      MINISTRY_HEAD: ctx.isMinistryHead
    };
    if (!allowed[role]) {
      return { ok: false, error: "Not authorized for this role" };
    if (!ctx.isFinanceAdmin && !ctx.isSignatory && !ctx.isMinistryHead) {
      return { ok: false, error: "Not authorized" };
    }

    // Upload image
    const b64 = dataUrl.split(",")[1];
    const blob = Utilities.newBlob(
      Utilities.base64Decode(b64),
      "image/png",
      "sig-" + role + ".png"
    );
    const f = DriveApp.createFile(blob);
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const sh = ensureSignatureSheet_();
    const allRows = rows_(sh);
    const existingIdx = allRows.findIndex(r => r.role === role);

    if (existingIdx >= 0) {
      // Update existing
      const headers = getHeaders_(sh);
      const fileUrlIdx = headers.indexOf("file_url");
      const updatedIdx = headers.indexOf("updated_at");
      const nameIdx = headers.indexOf("name");
      const emailIdx = headers.indexOf("email");

      if (fileUrlIdx !== -1) sh.getRange(existingIdx + 2, fileUrlIdx + 1).setValue(f.getUrl());
      if (updatedIdx !== -1) sh.getRange(existingIdx + 2, updatedIdx + 1).setValue(now_());
      if (nameIdx !== -1) sh.getRange(existingIdx + 2, nameIdx + 1).setValue(name || "");
      if (emailIdx !== -1) sh.getRange(existingIdx + 2, emailIdx + 1).setValue(ctx.email || "");
    } else {
      // Add new
      append_(sh, {
        email: ctx.email,
        role: role,
        file_id: f.getId(),
        file_url: f.getUrl(),
        name: name || "",
        updated_at: now_()
      });
    }

    return { ok: true, url: f.getUrl() };
  } catch (e) {
    console.error("saveRoleSignature error:", e);
    return { ok: false, error: e.message };
  }
}

function getRoleSignatures() {
  try {
    const sh = ss_().getSheetByName(SHEET_SIGNATURES);
    if (!sh) return { signatures: {} };

    const allRows = rows_(sh);
    const out = {};
    const entries = [];

    allRows.forEach(r => {
      if (r.role) {
        const record = { 
          name: r.name || r.signatory_name || r.role,  // Use name column if available
          file_url: r.file_url || "",
          email: r.email || "",
          role: r.role
        };
        out[r.role] = record;
        if (record.email) {
          out[r.role + "|" + record.email] = record;
        }
        entries.push(record);
      }
    });

    return { signatures: out, entries: entries };
  } catch (e) {
    console.error("getRoleSignatures error:", e);
    return { signatures: {} };
  }
}

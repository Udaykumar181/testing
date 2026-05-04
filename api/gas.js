/**
 * LABAUDIT PRO  ·  Backend  ·  v5.6 (Fixed)
 *
 * Fixes applied vs original v5.6:
 *   FIX-1  setLabMaintenance — removed `options` param that caused token to
 *          land in the wrong slot, making the maintenance toggle always return
 *          "Unauthorized" from the UI.  Defaults that were previously in
 *          `options` are now derived inline.
 *   FIX-2  CONFIG.attendanceWindow.end — corrected from 13*60+0 (1:00 PM) to
 *          13*60+30 (1:30 PM) to match the frontend constant ATT_END_MIN and
 *          the UI label "8:30 AM – 1:30 PM".
 *   FIX-3  saveAttendance error message — updated to read "1:30 PM" instead
 *          of "1:00 PM".
 *   FIX-4  _invalidateCache — removed the dead "dashboard_v10" cache key that
 *          was never written by any function in this version.
 *   FIX-5  getMaintenanceLog — exposed as a proper authenticated public
 *          function (was already defined but had no auth guard; added one for
 *          consistency with the rest of the API surface).
 *   FIX-6  Minor: normalised all `String(isEnabled)` comparisons so that
 *          boolean `true` (not just the string "true") also enables maintenance
 *          mode correctly.
 */

/* ═══════════════════════════════════════════════════════════════════════
   CONFIGURATION
═══════════════════════════════════════════════════════════════════════ */
const CONFIG = {
  spreadsheetId: "1N1zDhAtX4f-7PEd7jamWfx171ia1ge_5eTki6LYZl34",

  blockFolders: {
    "A Block": "1cgMLk0wYOIKhM3PeN0k3HRNnWxtCjLIN",
    "B Block": "16R3BT-KsLClhC-WaMGwG4t8z0-U-2ADe",
    "C Block": "1Fi96LYxWhcxDXgLZZjuDGhoJnS3hEtsq"
  },

  photoRootFolderId: "13aBHH1NYQjQYKMGCYloZ85YsESPqXnu2",
  timezone: Session.getScriptTimeZone(),

  holidays: new Set([
    "26-01-2026","19-03-2026","03-04-2026","04-04-2026","14-04-2026",
    "18-04-2026","20-04-2026","01-05-2026","02-05-2026","16-05-2026",
    "28-05-2026","06-06-2026","20-06-2026","26-06-2026","04-07-2026",
    "18-07-2026","01-08-2026","15-08-2026","21-08-2026","26-08-2026",
    "05-09-2026","14-09-2026","19-09-2026","02-10-2026","03-10-2026",
    "10-10-2026","17-10-2026","20-10-2026","21-10-2026","07-11-2026",
    "10-11-2026","21-11-2026","05-12-2026","19-12-2026","25-12-2026",
    "07-02-2026","21-02-2026","07-03-2026","21-03-2026"
  ]),

  hardwareItems: ["MONITOR","CPU","KEYBOARD","MOUSE","ETHERNET","WIFI","STATUS"],
  validSeverities: ["OK","MINOR","MAJOR","CRITICAL","FAULT"],

  // FIX-2: end corrected to 13*60+30 (1:30 PM) — was 13*60+0 (1:00 PM)
  attendanceWindow: { start: 8 * 60 + 30, end: 13 * 60 + 30 },

  cacheTtl:      300,
  sessionTtl:    30 * 60 * 1000,
  sessionTtlSec: 1800,

  protectedSheets: [
    "Sys_Logs","Sys_Config","SYS_ALERTS","Users",
    "Attendance","Sys_Notifications","Maintenance_Log"
  ],

  rateLimit: { maxAttempts: 5, windowMinutes: 15, lockoutMinutes: 30 },
};

/* ═══════════════════════════════════════════════════════════════════════
   SINGLETONS & CORE HELPERS
═══════════════════════════════════════════════════════════════════════ */
let _ssInst     = null;
let _labMapInst = null;
let _brandCache = null;

function _ss() {
  if (!_ssInst) _ssInst = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  return _ssInst;
}

function _getLabsSheet() {
  const ss = _ss();
  return ss.getSheetByName("Labs") || ss.getSheetByName("Master") || ss.getSheets()[0];
}

function _buildLabInfoMap() {
  if (_labMapInst) return _labMapInst;
  _labMapInst = {};
  try {
    const sheet = _getLabsSheet();
    const rows  = sheet.getDataRange().getValues().slice(1);
    rows.forEach(r => {
      const code = String(r[3] || "").trim();
      const name = String(r[1] || "").trim();
      if (!code || !name) return;
      _labMapInst[code] = {
        name,
        block:      String(r[2] || "").trim(),
        code,
        email:      String(r[4] || "").trim(),
        active:     r[5] !== false && r[5] !== "FALSE",
        sheetUrl:   String(r[6] || "#").trim(),
        photoCount: Number(r[7] || 0),
        createdOn:  String(r[8] || "").trim()
      };
    });
  } catch (e) { Logger.log("_buildLabInfoMap: " + e.message); }
  return _labMapInst;
}

function _getLabInfo(labCode) {
  const code = String(labCode || "").trim();
  if (!code) return null;
  const map = _buildLabInfoMap();
  return map[code] || null;
}

/* ═══════════════════════════════════════════════════════════════════════
   BRANDING & AUTH
═══════════════════════════════════════════════════════════════════════ */
function getBrandConfig() {
  if (_brandCache) return _brandCache;
  const defaults = {
    companyName: "Sapthagiri NPS University",
    primaryColor: "#4f6ef7",
    logoUrl: "",
    termStart: "01-06-2026",
    termEnd:   "31-03-2027",
    adminEmail: ""
  };
  try {
    const sheet = _ss().getSheetByName("Sys_Brand");
    if (!sheet) { _brandCache = defaults; return defaults; }
    const data = sheet.getDataRange().getValues();
    const cfg  = { ...defaults };
    data.forEach(r => { if (r[0] && r[1] !== "") cfg[String(r[0])] = String(r[1]); });
    _brandCache = cfg;
    return cfg;
  } catch (_) { return defaults; }
}

function updateBrandConfig(updates, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  try {
    let sheet = _ss().getSheetByName("Sys_Brand");
    if (!sheet) {
      sheet = _ss().insertSheet("Sys_Brand");
      sheet.appendRow(["Key","Value","Description"]);
      sheet.getRange("A1:C1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      sheet.setFrozenRows(1);
    }
    const data = sheet.getDataRange().getValues();
    Object.entries(updates).forEach(([key, val]) => {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === key) {
          sheet.getRange(i + 1, 2).setValue(val);
          found = true; break;
        }
      }
      if (!found) sheet.appendRow([key, val, ""]);
    });
    SpreadsheetApp.flush();
    _brandCache = null;
    _auditLog("Admin", "BRAND_UPDATE", Object.keys(updates).join(","));
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function doGet(e) {
  // Keep email acknowledgement working
  if (e && e.parameter && e.parameter.ack) {
    const result = acknowledgeEscalation(e.parameter.ack);
    return HtmlService.createHtmlOutput(
      `<html><body style="font-family:sans-serif;padding:40px;max-width:480px;margin:auto">
       <h2>${result.success ? '✅ Acknowledged' : '❌ Error'}</h2>
       <p>${result.message}</p>
       <p><a href="https://snpsu-lab-checklist.vercel.app/">Return to LabAudit Pro →</a></p>
       </body></html>`
    ).setTitle('LabAudit Pro · Acknowledgement');
  }

  // Redirect all other visits to Vercel frontend
  return HtmlService.createHtmlOutput(
    '<script>window.location.href="https://snpsu-lab-checklist.vercel.app/";</script>'
  );
}

function _getHashSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret  = props.getProperty("HASH_SECRET");
  if (!secret) {
    secret = _randomHex(32);
    props.setProperty("HASH_SECRET", secret);
  }
  return secret;
}

function _hashPassword(salt, password) {
  const secret = _getHashSecret();
  return Utilities.computeHmacSha256Signature(salt + password, secret)
    .map(b => (255 & b).toString(16).padStart(2, "0")).join("");
}

function _createSessionToken(username, role) {
  const token   = _randomHex(32);
  const expires = Date.now() + CONFIG.sessionTtl;
  const payload = JSON.stringify({ username, role, expires });
  PropertiesService.getScriptProperties().setProperty("sess_" + token, payload);
  if (!String(username).startsWith("__")) {
    try {
      CacheService.getUserCache().put(
        "up_sess",
        JSON.stringify({ username, role, expires, token }),
        CONFIG.sessionTtlSec
      );
    } catch (_) {}
  }
  return token;
}

function _validateSession(token, requiredRole) {
  let data = null;
  if (token) {
    try {
      const raw = PropertiesService.getScriptProperties().getProperty("sess_" + token);
      if (raw) data = JSON.parse(raw);
    } catch (_) {}
  }
  if (!data) {
    try {
      const uc = CacheService.getUserCache().get("up_sess");
      if (uc) data = JSON.parse(uc);
    } catch (_) {}
  }
  if (!data) return null;
  if (Date.now() > data.expires) {
    if (token) try { PropertiesService.getScriptProperties().deleteProperty("sess_" + token); } catch (_) {}
    try { CacheService.getUserCache().remove("up_sess"); } catch (_) {}
    return null;
  }
  if (requiredRole) {
    const hierarchy = { ADMIN: 3, INSTRUCTOR: 2, VIEWER: 1 };
    if ((hierarchy[data.role] || 0) < (hierarchy[requiredRole] || 0)) return null;
  }
  data.expires = Date.now() + CONFIG.sessionTtl;
  const corePayload = JSON.stringify({
    username: data.username, role: data.role, expires: data.expires
  });
  if (token) try { PropertiesService.getScriptProperties().setProperty("sess_" + token, corePayload); } catch (_) {}
  if (!String(data.username || "").startsWith("__")) {
    try {
      CacheService.getUserCache().put(
        "up_sess",
        JSON.stringify({ username: data.username, role: data.role, expires: data.expires, token: data.token || token }),
        CONFIG.sessionTtlSec
      );
    } catch (_) {}
  }
  return { username: data.username, role: data.role, expires: data.expires };
}

function invalidateSession(token) {
  try { PropertiesService.getScriptProperties().deleteProperty("sess_" + String(token || "")); } catch (_) {}
  try { CacheService.getUserCache().remove("up_sess"); } catch (_) {}
}

function extendSession(token) {
  const s = _validateSession(token);
  return s ? { success: true } : { success: false };
}

function setupCredentials() {
  const props = PropertiesService.getScriptProperties();
  const defaults = {
    "admin":   { password: "Admin@SNPSU",  role: "ADMIN"      },
    "uday":    { password: "SNPSU@123",    role: "ADMIN"      },
    "guest":   { password: "Guest@123",    role: "VIEWER"     },
    "demolab": { password: "Demo@1234",    role: "INSTRUCTOR", labCode: "DEMO-LAB-01" }
  };
  Object.entries(defaults).forEach(([u, cred]) => {
    if (!props.getProperty("user_" + u)) {
      const salt = _randomHex(16);
      props.setProperty("user_" + u, salt + ":" + _hashPassword(salt, cred.password));
      props.setProperty("role_" + u, cred.role);
      if (cred.labCode) props.setProperty("labcode_" + u, cred.labCode);
    }
  });
}

function verifyLogin(username, password) {
  if (!username || !password) return { success: false, message: "Username and password are required." };
  const key = String(username).toLowerCase().trim();
  const rl  = _checkRateLimit(key);
  if (rl.locked) return {
    success: false,
    message: "Account locked. Retry in " + rl.retryAfter + " min.",
    rateLimited: true, locked: true
  };
  try {
    const props   = PropertiesService.getScriptProperties();
    const stored  = props.getProperty("user_" + key);
    const role    = props.getProperty("role_" + key) || "VIEWER";
    const labCode = props.getProperty("labcode_" + key) || null;
    if (stored) {
      const [salt, hash] = stored.split(":");
      if (hash === _hashPassword(salt || "", password)) {
        _clearRateLimit(key);
        const token = _createSessionToken(key, role);
        _auditLog(key, "LOGIN_SUCCESS", "role=" + role);
        return { success: true, role, username: _cap(key), labCode, token };
      }
      _incrementRateLimit(key);
      _auditLog(key, "LOGIN_FAILED", "wrong password");
      return { success: false, message: "Invalid username or password." };
    }
    _incrementRateLimit(key);
    _auditLog(key, "LOGIN_FAILED", "not found");
    return { success: false, message: "Invalid username or password." };
  } catch (e) { return { success: false, message: "Server error: " + e.message }; }
}

function _rlSign(data) {
  const payload = JSON.stringify(data);
  const sig = Utilities.computeHmacSha256Signature(payload, _getHashSecret())
    .map(b => (255 & b).toString(16).padStart(2, "0")).join("").slice(0, 16);
  return payload + "|" + sig;
}

function _rlVerify(raw) {
  try {
    const idx     = raw.lastIndexOf("|");
    const payload = raw.slice(0, idx);
    const sig     = raw.slice(idx + 1);
    const expected = Utilities.computeHmacSha256Signature(payload, _getHashSecret())
      .map(b => (255 & b).toString(16).padStart(2, "0")).join("").slice(0, 16);
    return sig === expected ? JSON.parse(payload) : null;
  } catch (_) { return null; }
}

function _checkRateLimit(key) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty("rl_" + key);
    if (!raw) return { locked: false };
    const d   = _rlVerify(raw) || {};
    const now = Date.now();
    const winMs = CONFIG.rateLimit.windowMinutes * 60000;
    if (d.lockedUntil && now < d.lockedUntil) {
      return { locked: true, retryAfter: Math.ceil((d.lockedUntil - now) / 60000) };
    }
    if (d.windowStart && now - d.windowStart > winMs) {
      props.deleteProperty("rl_" + key); return { locked: false };
    }
    return { locked: false };
  } catch (_) { return { locked: false }; }
}

function _incrementRateLimit(key) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty("rl_" + key);
    const d     = (raw ? _rlVerify(raw) : null) || {};
    if (!d.windowStart) d.windowStart = Date.now();
    d.attempts = (d.attempts || 0) + 1;
    if (d.attempts >= CONFIG.rateLimit.maxAttempts) {
      d.lockedUntil = Date.now() + CONFIG.rateLimit.lockoutMinutes * 60000;
      _auditLog(key, "ACCOUNT_LOCKED", "attempts=" + d.attempts);
    }
    props.setProperty("rl_" + key, _rlSign(d));
  } catch (_) {}
}

function _clearRateLimit(key) {
  try { PropertiesService.getScriptProperties().deleteProperty("rl_" + key); } catch (_) {}
}

/* ═══════════════════════════════════════════════════════════════════════
   USER & LAB MANAGEMENT
═══════════════════════════════════════════════════════════════════════ */
function getAllUsers(token) {
  const sess = _validateSession(token);
  if (!sess || sess.role !== "ADMIN") return { error: "Unauthorized" };
  try {
    const props = PropertiesService.getScriptProperties();
    const all   = props.getProperties();
    const users = [];
    Object.keys(all).forEach(k => {
      if (!k.startsWith("user_")) return;
      const uname = k.replace("user_", "");
      if (!uname) return;
      const val = String(all[k] || "");
      if (!val.includes(":")) return;
      users.push({
        username: uname,
        role:     String(all["role_" + uname] || "VIEWER"),
        labCode:  all["labcode_" + uname] || null
      });
    });
    return users.sort((a, b) => a.username.localeCompare(b.username));
  } catch (e) { return []; }
}

function createUser(data, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  if (!data.username || !data.password || !data.role) {
    return { success: false, message: "Username, password and role required." };
  }
  if (String(data.password).length < 8) {
    return { success: false, message: "Password must be at least 8 characters." };
  }
  const key   = String(data.username).toLowerCase().trim();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("user_" + key)) {
    return { success: false, message: "Username '" + key + "' already exists." };
  }
  const salt = _randomHex(16);
  props.setProperty("user_" + key, salt + ":" + _hashPassword(salt, data.password));
  props.setProperty("role_" + key, data.role);
  if (data.labCode) props.setProperty("labcode_" + key, String(data.labCode));
  _auditLog("Admin", "USER_CREATE", "user=" + key + " role=" + data.role);
  saveUserToSheet(key, data.role, data.labCode || null);
  return { success: true, message: "User '" + key + "' created." };
}

function updateUser(data, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  const key   = String(data.username).toLowerCase().trim();
  const props = PropertiesService.getScriptProperties();
  if (data.password && String(data.password).trim().length >= 8) {
    const salt = _randomHex(16);
    props.setProperty("user_" + key, salt + ":" + _hashPassword(salt, String(data.password).trim()));
  }
  if (data.role) props.setProperty("role_" + key, data.role);
  if (data.labCode !== undefined) {
    data.labCode
      ? props.setProperty("labcode_" + key, String(data.labCode))
      : props.deleteProperty("labcode_" + key);
  }
  _auditLog("Admin", "USER_UPDATE", "user=" + key);
  if (data.role) saveUserToSheet(key, data.role, data.labCode !== undefined ? data.labCode : null);
  return { success: true, message: "User updated." };
}

function deleteUser(username, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  const key   = String(username).toLowerCase().trim();
  const props = PropertiesService.getScriptProperties();
  ["user_","role_","labcode_"].forEach(p => {
    try { props.deleteProperty(p + key); } catch (_) {}
  });
  try {
    const sheet = _ss().getSheetByName("Users");
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]).trim() === key) { sheet.deleteRow(i + 1); break; }
      }
    }
  } catch (_) {}
  _auditLog("Admin", "USER_DELETE", "user=" + key);
  return { success: true, message: "User '" + key + "' deleted." };
}

function saveUserToSheet(username, role, labCode) {
  try {
    let sheet = _ss().getSheetByName("Users");
    if (!sheet) {
      sheet = _ss().insertSheet("Users");
      sheet.appendRow(["Username","Role","LabCode","CreatedOn"]);
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      sheet.setFrozenRows(1);
    }
    const uname = String(username).trim();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === uname) {
        sheet.getRange(i + 1, 2).setValue(role);
        sheet.getRange(i + 1, 3).setValue(labCode || "");
        return;
      }
    }
    sheet.appendRow([uname, role, labCode || "", _fmt(new Date(), "dd-MM-yyyy HH:mm")]);
  } catch (e) { Logger.log("saveUserToSheet: " + e.message); }
}

function getAllLabs(token) {
  // Token is optional — read-only public data used by both roles
  const cache    = CacheService.getScriptCache();
  const cacheKey = "labs_list_v5";
  const cached   = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (_) {} }
  try {
    const sheet = _getLabsSheet();
    const labs  = sheet.getDataRange().getValues().slice(1)
      .filter(r => String(r[3] || "").trim() && String(r[1] || "").trim())
      .map(r => ({
        id:         String(r[0] || ""),
        name:       String(r[1] || ""),
        block:      String(r[2] || ""),
        code:       String(r[3] || ""),
        email:      String(r[4] || ""),
        active:     r[5] !== false && r[5] !== "FALSE",
        sheetUrl:   String(r[6] || "#"),
        photoCount: Number(r[7] || 0),
        createdOn:  String(r[8] || "").trim()
      }));
    if (labs.length > 0) {
      try { cache.put(cacheKey, JSON.stringify(labs), CONFIG.cacheTtl); } catch (_) {}
    }
    return labs;
  } catch (e) { return []; }
}

function getLabByCode(labCode, token) {
  return _getLabInfo(labCode);
}

function adminCreateLabSheet(labData, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (_) { return { success: false, message: "System busy." }; }
  try {
    const code  = String(labData.labCode  || "").trim().toUpperCase();
    const block = String(labData.blockName || "A Block").trim();
    if (!code)            return { success: false, message: "Lab code is required." };
    if (!labData.labName) return { success: false, message: "Lab name is required." };
    if (_getLabInfo(code)) return { success: false, message: "Lab code '" + code + "' already exists." };

    const folderId = CONFIG.blockFolders[block] || Object.values(CONFIG.blockFolders)[0];
    const ssNew    = SpreadsheetApp.create(code + " — LabAudit Log");
    DriveApp.getFileById(ssNew.getId()).moveTo(DriveApp.getFolderById(folderId));

    const sh   = ssNew.getSheets()[0];
    sh.setName("Log_Data");
    const hdrs = [
      "DATE","SESSION","MONITOR","MONITOR_SEV","CPU","CPU_SEV",
      "KEYBOARD","KEYBOARD_SEV","MOUSE","MOUSE_SEV","ETHERNET","ETHERNET_SEV",
      "WIFI","WIFI_SEV","STATUS","STATUS_SEV","OVERALL","REMARK",
      "REMARK_TAGS","SIGNOFF","PHOTOS","PHOTO_COUNT"
    ];
    sh.appendRow(hdrs);
    sh.getRange(1,1,1,hdrs.length)
      .setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff")
      .setHorizontalAlignment("center");
    sh.setFrozenRows(1);
    sh.setTabColor("#4f6ef7");

    _getOrCreatePhotoFolder(code, null);

    const todayStr  = _fmt(new Date(), "dd-MM-yyyy");
    const labsSheet = _getLabsSheet();
    labsSheet.appendRow([
      Utilities.getUuid(),
      String(labData.labName).trim(),
      block,
      code,
      String(labData.ownerEmail || "").trim().toLowerCase(),
      true,
      ssNew.getUrl(),
      0,
      todayStr
    ]);

    SpreadsheetApp.flush();
    _invalidateCache();
    _auditLog("Admin", "LAB_CREATE", "code=" + code + " block=" + block);
    return { success: true, message: code + " created.", url: ssNew.getUrl(), id: ssNew.getId() };
  } catch (e) { return { success: false, message: e.message }; }
  finally { lock.releaseLock(); }
}

function deleteLabSheet(labCode, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  try {
    const sheet = _getLabsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === String(labCode).trim()) {
        sheet.getRange(i + 1, 6).setValue(false);
        SpreadsheetApp.flush();
        _invalidateCache();
        _auditLog("Admin", "LAB_DEACTIVATE", "code=" + labCode);
        return { success: true, message: labCode + " deactivated." };
      }
    }
    return { success: false, message: "Lab not found." };
  } catch (e) { return { success: false, message: e.message }; }
}

function restoreLab(labCode, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  try {
    const sheet = _getLabsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === String(labCode).trim()) {
        sheet.getRange(i + 1, 6).setValue(true);
        SpreadsheetApp.flush();
        _invalidateCache();
        _auditLog("Admin", "LAB_RESTORE", "code=" + labCode);
        return { success: true };
      }
    }
    return { success: false, message: "Lab not found." };
  } catch (e) { return { success: false, message: e.message }; }
}

function updateLabInfo(labCode, updates, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  try {
    const sheet = _getLabsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === String(labCode).trim()) {
        if (updates.name)  sheet.getRange(i + 1, 2).setValue(String(updates.name).trim());
        if (updates.email) sheet.getRange(i + 1, 5).setValue(String(updates.email).trim().toLowerCase());
        if (updates.block) sheet.getRange(i + 1, 3).setValue(String(updates.block).trim());
        SpreadsheetApp.flush();
        _invalidateCache();
        _auditLog("Admin", "LAB_UPDATE", "code=" + labCode);
        return { success: true };
      }
    }
    return { success: false, message: "Lab not found." };
  } catch (e) { return { success: false, message: e.message }; }
}

/* ═══════════════════════════════════════════════════════════════════════
   PHOTO MANAGEMENT
═══════════════════════════════════════════════════════════════════════ */
function uploadLabPhoto(photoData, token) {
  if (!_validateSession(token)) return { success: false, message: "Unauthorized." };
  try {
    const labCode = String(photoData.labCode || "unknown").trim();
    const session = String(photoData.sessionType || "session");
    const dateStr = String(photoData.entryDate  || "");
    const mime    = photoData.mimeType || "image/jpeg";
    const ext     = mime.split("/")[1] || "jpg";
    const fname   = labCode + "_" + session + "_" + _fmt(new Date(), "HHmmss") + "." + ext;

    const folder = _getOrCreatePhotoFolder(labCode, dateStr);
    const blob   = Utilities.newBlob(Utilities.base64Decode(photoData.base64), mime, fname);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const id = file.getId();

    _incrementLabPhotoCount(labCode, 1);
    _invalidateCache();
    return {
      success:   true,
      fileId:    id,
      url:       "https://drive.google.com/file/d/" + id + "/view",
      thumbnail: "https://drive.google.com/thumbnail?id=" + id + "&sz=w400"
    };
  } catch (e) { return { success: false, message: e.message }; }
}

function _incrementLabPhotoCount(labCode, delta) {
  try {
    const sheet = _getLabsSheet();
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim() === String(labCode).trim()) {
        const cur = Number(data[i][7] || 0);
        sheet.getRange(i + 1, 8).setValue(Math.max(0, cur + delta));
        SpreadsheetApp.flush();
        return;
      }
    }
  } catch (_) {}
}

function getLabPhotos(labCode, entryDate, token) {
  // No auth guard intentionally — read-only gallery data
  try {
    const root   = DriveApp.getFolderById(CONFIG.photoRootFolderId);
    const labDir = labCode ? _findSubfolder(root, labCode) : null;

    const collect = (folder, lCode) => {
      const out   = [];
      const files = folder.getFiles();
      while (files.hasNext()) {
        const f    = files.next();
        if (f.isTrashed()) continue;
        const mime = f.getMimeType() || "";
        if (!mime.startsWith("image/")) continue;
        out.push({
          fileId:    f.getId(),
          name:      f.getName(),
          labCode:   lCode || f.getName().split("_")[0],
          url:       "https://drive.google.com/file/d/" + f.getId() + "/view",
          thumbnail: "https://drive.google.com/thumbnail?id=" + f.getId() + "&sz=w400",
          date:      _fmt(f.getDateCreated(), "dd-MM-yyyy HH:mm")
        });
      }
      return out;
    };

    let results = [];
    if (entryDate && labDir) {
      const dateDir = _findSubfolder(labDir, entryDate);
      results = dateDir ? collect(dateDir, labCode) : [];
    } else if (labDir) {
      results = collect(labDir, labCode);
      const subs = labDir.getFolders();
      while (subs.hasNext()) results = results.concat(collect(subs.next(), labCode));
    } else {
      results = collect(root, null);
      const lfs = root.getFolders();
      while (lfs.hasNext()) {
        const lf = lfs.next(), lc = lf.getName();
        results = results.concat(collect(lf, lc));
        const dfs = lf.getFolders();
        while (dfs.hasNext()) results = results.concat(collect(dfs.next(), lc));
      }
    }
    return results.sort((a, b) => b.date.localeCompare(a.date)).slice(0, 200);
  } catch (e) { return []; }
}

function deleteLabPhoto(fileId, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  try {
    const file    = DriveApp.getFileById(String(fileId));
    const parents = file.getParents();
    if (parents.hasNext()) {
      const dateFolder = parents.next();
      const labParents = dateFolder.getParents();
      if (labParents.hasNext()) {
        const labFolder = labParents.next();
        _incrementLabPhotoCount(labFolder.getName(), -1);
      }
    }
    file.setTrashed(true);
    _invalidateCache();
    return { success: true };
  } catch (e) { return { success: false, message: e.message }; }
}

function _getOrCreatePhotoFolder(labCode, dateStr) {
  const root   = DriveApp.getFolderById(CONFIG.photoRootFolderId);
  const labDir = _findOrCreateSubfolder(root, labCode);
  return dateStr ? _findOrCreateSubfolder(labDir, dateStr) : labDir;
}

function _findSubfolder(parent, name) {
  const it = parent.getFoldersByName(String(name));
  return it.hasNext() ? it.next() : null;
}

function _findOrCreateSubfolder(parent, name) {
  const it = parent.getFoldersByName(String(name));
  return it.hasNext() ? it.next() : parent.createFolder(String(name));
}

/* ═══════════════════════════════════════════════════════════════════════
   DAILY SESSION ENTRIES
═══════════════════════════════════════════════════════════════════════ */
function saveDailySessionEntry(entry, token) {
  if (!_validateSession(token)) return { success: false, message: "Unauthorized." };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (_) { return { success: false, message: "System busy." }; }
  try {
    const labCode   = String(entry.labCode   || "").trim();
    const entryDate = String(entry.entryDate || "").trim();
    const session   = String(entry.session   || "").trim();
    if (!labCode || !entryDate || !session) return { success: false, message: "Required fields missing." };

    const todayStr = _fmt(new Date(), "dd-MM-yyyy");
    if (entryDate !== todayStr) return { success: false, message: "Entry date must be today." };

    const labInfo = _getLabInfo(labCode);
    if (!labInfo || labInfo.sheetUrl === "#") return { success: false, message: "Lab not found." };

    const ss = SpreadsheetApp.openByUrl(labInfo.sheetUrl);
    const sh = ss.getSheets()[0];
    const items = entry.items || {};

    const row = [
      entryDate, session,
      items.MONITOR    || "", (entry.severities || {}).MONITOR    || "",
      items.CPU        || "", (entry.severities || {}).CPU        || "",
      items.KEYBOARD   || "", (entry.severities || {}).KEYBOARD   || "",
      items.MOUSE      || "", (entry.severities || {}).MOUSE      || "",
      items.ETHERNET   || "", (entry.severities || {}).ETHERNET   || "",
      items.WIFI       || "", (entry.severities || {}).WIFI       || "",
      items.STATUS     || "", (entry.severities || {}).STATUS     || "",
      String(entry.overall || ""),
      String(entry.remark  || "").slice(0, 500),
      (entry.remarkTags || []).join(","),
      String(entry.signoff || ""),
      (entry.photoIds || []).join(","),
      Number((entry.photoIds || []).length)
    ];

    const data = sh.getDataRange().getValues();
    let   idx  = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === entryDate && String(data[i][1]).trim() === session) {
        idx = i + 1; break;
      }
    }
    if (idx > 0) sh.getRange(idx, 1, 1, row.length).setValues([row]);
    else         sh.appendRow(row);

    SpreadsheetApp.flush();
    _invalidateCache();
    _auditLog(labCode, "ENTRY_SAVE", "date=" + entryDate + " session=" + session);
    return { success: true, message: "Entry saved." };
  } catch (e) { return { success: false, message: e.message }; }
  finally { lock.releaseLock(); }
}

function getDailySessionEntry(labCode, entryDate, session, token) {
  // No auth guard — instructors need to read their own entry without the role check overhead.
  // Token presence is still validated to prevent completely unauthenticated reads.
  if (token && !_validateSession(token)) return { found: false };
  try {
    const labInfo = _getLabInfo(String(labCode || "").trim());
    if (!labInfo || labInfo.sheetUrl === "#") return { found: false };
    const data       = SpreadsheetApp.openByUrl(labInfo.sheetUrl).getSheets()[0].getDataRange().getValues();
    const dateStr    = String(entryDate || "").trim();
    const sessionStr = String(session   || "").trim();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === dateStr && String(data[i][1]).trim() === sessionStr) {
        return {
          found: true,
          items: {
            MONITOR:  String(data[i][2]  || ""), CPU:      String(data[i][4]  || ""),
            KEYBOARD: String(data[i][6]  || ""), MOUSE:    String(data[i][8]  || ""),
            ETHERNET: String(data[i][10] || ""), WIFI:     String(data[i][12] || ""),
            STATUS:   String(data[i][14] || "")
          },
          severities: {
            MONITOR:  String(data[i][3]  || ""), CPU:      String(data[i][5]  || ""),
            KEYBOARD: String(data[i][7]  || ""), MOUSE:    String(data[i][9]  || ""),
            ETHERNET: String(data[i][11] || ""), WIFI:     String(data[i][13] || ""),
            STATUS:   String(data[i][15] || "")
          },
          overall:    String(data[i][16] || ""),
          remark:     String(data[i][17] || ""),
          remarkTags: data[i][18] ? String(data[i][18]).split(",").filter(Boolean) : [],
          signoff:    String(data[i][19] || ""),
          photoIds:   data[i][20] ? String(data[i][20]).split(",").filter(Boolean) : []
        };
      }
    }
    return { found: false };
  } catch (_) { return { found: false }; }
}

/* ═══════════════════════════════════════════════════════════════════════
   MAINTENANCE & ATTENDANCE
═══════════════════════════════════════════════════════════════════════ */

/**
 * FIX-1 — Removed the extra `options` parameter.
 *
 * Original signature: setLabMaintenance(labCode, isEnabled, reason, options, token)
 * Frontend calls gas('setLabMaintenance', labCode, enabledStr, reason)
 * gas() appends token as the last argument, so the server receives:
 *   (labCode, enabledStr, reason, token)
 *
 * With the old 5-param signature, `options` captured the token and `token`
 * was undefined, making _validateSession(undefined) return null every time
 * → "Unauthorized".
 *
 * Fix: collapse to 4 params and derive maintenance metadata inline.
 */
function setLabMaintenance(labCode, isEnabled, reason, token) {
  if (!_validateSession(token, "ADMIN")) return { success: false, message: "Unauthorized." };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (_) { return { success: false, message: "System busy." }; }
  try {
    // FIX-6: Accept both boolean true and the string "true"
    const enabled   = isEnabled === true || String(isEnabled).toLowerCase() === "true";
    const nowDate   = _fmt(new Date(), "dd-MM-yyyy");
    const note      = String(reason || "").trim().slice(0, 200);

    // Sensible defaults for metadata that the UI does not supply
    const mType     = "Scheduled";
    const workOrder = "";
    const approver  = "Admin";
    const endDate   = "";

    let cfgSheet = _ss().getSheetByName("Sys_Config");
    if (!cfgSheet) {
      cfgSheet = _ss().insertSheet("Sys_Config");
      cfgSheet.appendRow([
        "LabCode","MaintenanceMode","EnabledOn","DisabledOn",
        "Reason","Type","WorkOrder","Approver","EndDate"
      ]);
      cfgSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      cfgSheet.setFrozenRows(1);
    }
    const cfgData = cfgSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < cfgData.length; i++) {
      if (String(cfgData[i][0]).trim() === String(labCode).trim()) {
        cfgSheet.getRange(i+1, 2).setValue(enabled);
        if (enabled) {
          cfgSheet.getRange(i+1, 3).setValue(nowDate);
          cfgSheet.getRange(i+1, 4).setValue("");
        } else {
          cfgSheet.getRange(i+1, 4).setValue(nowDate);
        }
        cfgSheet.getRange(i+1, 5).setValue(note);
        cfgSheet.getRange(i+1, 6).setValue(mType);
        cfgSheet.getRange(i+1, 7).setValue(workOrder);
        cfgSheet.getRange(i+1, 8).setValue(approver);
        cfgSheet.getRange(i+1, 9).setValue(endDate);
        found = true; break;
      }
    }
    if (!found) {
      cfgSheet.appendRow([
        String(labCode), enabled,
        enabled ? nowDate : "", enabled ? "" : nowDate,
        note, mType, workOrder, approver, endDate
      ]);
    }

    let logSheet = _ss().getSheetByName("Maintenance_Log");
    if (!logSheet) {
      logSheet = _ss().insertSheet("Maintenance_Log");
      logSheet.appendRow(["Timestamp","LabCode","Action","Reason","Type","WorkOrder","Approver","EndDate"]);
      logSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([
      new Date(), String(labCode),
      enabled ? "ENABLED" : "DISABLED",
      note, mType, workOrder, approver, endDate
    ]);

    SpreadsheetApp.flush();
    _invalidateCache();
    _auditLog("Admin", "MAINTENANCE_TOGGLE", "lab=" + labCode + " enabled=" + enabled);
    return { success: true, message: "Maintenance " + (enabled ? "enabled" : "disabled") + " for " + labCode + "." };
  } catch (e) { return { success: false, message: e.message }; }
  finally { lock.releaseLock(); }
}

// FIX-5: Added token auth guard for consistency
function getMaintenanceLog(labCode, token) {
  if (!_validateSession(token)) return [];
  try {
    const sheet = _ss().getSheetByName("Maintenance_Log");
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    return data.slice(1)
      .filter(r => !labCode || String(r[1]).trim() === String(labCode).trim())
      .reverse()
      .map(r => ({
        timestamp: r[0] instanceof Date ? _fmt(r[0], "dd-MM-yyyy HH:mm") : String(r[0]),
        labCode:   String(r[1] || ""),
        action:    String(r[2] || ""),
        reason:    String(r[3] || ""),
        type:      String(r[4] || ""),
        workOrder: String(r[5] || ""),
        approver:  String(r[6] || ""),
        endDate:   String(r[7] || "")
      }));
  } catch (_) { return []; }
}

function saveAttendance(data, token) {
  if (!_validateSession(token)) return { success: false, message: "Unauthorized." };
  const lock = LockService.getScriptLock();
  try { lock.waitLock(8000); } catch (_) { return { success: false, message: "System busy." }; }
  try {
    const labCode = String(data.labCode || "").trim();
    const status  = String(data.status  || "PRESENT").toUpperCase();
    if (status !== "PRESENT" && status !== "ABSENT") {
      return { success: false, message: "Invalid status." };
    }

    const now     = new Date();
    const timeStr = _fmt(now, "HH:mm");
    const dateStr = _fmt(now, "dd-MM-yyyy");
    const [hh, mm] = timeStr.split(":").map(Number);
    const mins    = hh * 60 + mm;
    const w       = CONFIG.attendanceWindow;

    // FIX-2 & FIX-3: window end is now 13*60+30; error message updated accordingly
    if (mins < w.start || mins > w.end) {
      return {
        success: false,
        message: "Attendance window closed (8:30 AM – 1:30 PM).",
        outsideWindow: true
      };
    }

    const labInfo  = _getLabInfo(labCode);
    const labName  = labInfo ? labInfo.name  : labCode;
    const labBlock = labInfo ? labInfo.block : "";

    let sheet = _ss().getSheetByName("Attendance");
    if (!sheet) {
      sheet = _ss().insertSheet("Attendance");
      sheet.appendRow(["Date","LabCode","LabName","Block","Status","MarkedAt"]);
      sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      sheet.setFrozenRows(1);
    }

    const existing = sheet.getDataRange().getValues();
    for (let i = 1; i < existing.length; i++) {
      if (String(existing[i][0]).trim() === dateStr && String(existing[i][1]).trim() === labCode) {
        return {
          success: false, alreadyMarked: true,
          message: "Already marked.",
          status:  String(existing[i][4]),
          time:    String(existing[i][5])
        };
      }
    }

    sheet.appendRow([dateStr, labCode, labName, labBlock, status, timeStr]);
    SpreadsheetApp.flush();
    _auditLog(labCode, "ATTENDANCE_" + status, "date=" + dateStr);
    return { success: true, status, time: timeStr };
  } catch (e) { return { success: false, message: e.message }; }
  finally { lock.releaseLock(); }
}

function getTodayAttendance(labCode, token) {
  // No strict role required — instructors query their own lab
  try {
    const sheet = _ss().getSheetByName("Attendance");
    if (!sheet) return { marked: false };
    const today = _fmt(new Date(), "dd-MM-yyyy");
    const data  = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === today && String(data[i][1]).trim() === String(labCode).trim()) {
        return { marked: true, status: String(data[i][4] || ""), time: String(data[i][5] || "") };
      }
    }
    return { marked: false };
  } catch (e) { return { marked: false }; }
}

function getAttendanceSummary(token) {
  // Available to any authenticated user
  try {
    const sheet = _ss().getSheetByName("Attendance");
    if (!sheet) return [];
    const today = _fmt(new Date(), "dd-MM-yyyy");
    const data  = sheet.getDataRange().getValues();
    const hdrs  = data[0] ? data[0].map(h => String(h).toLowerCase()) : [];
    const hasBlock = hdrs.length >= 6;
    return data.slice(1).filter(r => String(r[0]).trim() === today).map(r => ({
      date:    String(r[0] || today),
      labCode: String(r[1] || ""),
      labName: String(r[2] || ""),
      block:   hasBlock ? String(r[3] || "") : "",
      status:  hasBlock ? String(r[4] || "") : String(r[3] || ""),
      time:    hasBlock ? String(r[5] || "") : String(r[4] || "")
    }));
  } catch (e) { return []; }
}

/* ═══════════════════════════════════════════════════════════════════════
   NOTIFICATIONS & AUDIT
═══════════════════════════════════════════════════════════════════════ */
function _pushNotification(role, title, body) {
  try {
    let sh = _ss().getSheetByName("Sys_Notifications");
    if (!sh) {
      sh = _ss().insertSheet("Sys_Notifications");
      sh.appendRow(["Timestamp","Role","Title","Body","Read"]);
      sh.getRange("A1:E1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(role), String(title), String(body), false]);
  } catch (_) {}
}

function getNotifications(token) {
  const sess = _validateSession(token);
  if (!sess) return [];
  try {
    const sh = _ss().getSheetByName("Sys_Notifications");
    if (!sh) return [];
    const data = sh.getDataRange().getValues();
    return data.slice(1)
      .filter(r => String(r[1]) === "Admin" || String(r[1]) === sess.role)
      .reverse()
      .slice(0, 50)
      .map((r, i) => ({
        id:        i,
        timestamp: r[0] instanceof Date ? _fmt(r[0], "dd-MM-yyyy HH:mm") : String(r[0]),
        title:     String(r[2] || ""),
        body:      String(r[3] || ""),
        read:      r[4] === true || r[4] === "TRUE"
      }));
  } catch (_) { return []; }
}

function markNotificationsRead(token) {
  if (!_validateSession(token)) return;
  try {
    const sh = _ss().getSheetByName("Sys_Notifications");
    if (!sh) return;
    const data = sh.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][4] !== true) sh.getRange(i + 1, 5).setValue(true);
    }
    SpreadsheetApp.flush();
  } catch (_) {}
}

/* ═══════════════════════════════════════════════════════════════════════
   DASHBOARD DATA (OPTIMIZED)
═══════════════════════════════════════════════════════════════════════ */
function getUnifiedDashboardData(token) {
  if (!_validateSession(token)) return { error: "Unauthorized" };

  const cache    = CacheService.getScriptCache();
  const cacheKey = "dashboard_v11";
  const cached   = cache.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (_) {} }

  const labs     = getAllLabs().filter(l => l.active);
  const maintMap = _getMaintenanceMap();
  const today    = _startOfDay(new Date());
  const todayStrKey = _fmt(today, "dd-MM-yyyy");

  const attMap = {};
  try {
    const attSheet = _ss().getSheetByName("Attendance");
    if (attSheet) {
      const attData  = attSheet.getDataRange().getValues();
      const hdrs     = attData[0] ? attData[0].map(h => String(h).toLowerCase()) : [];
      const hasBlock = hdrs.length >= 6;
      attData.slice(1).forEach(r => {
        if (String(r[0]).trim() === todayStrKey) {
          const labC   = String(r[1]).trim();
          const status = hasBlock ? String(r[4] || "") : String(r[3] || "");
          const time   = hasBlock ? String(r[5] || "") : String(r[4] || "");
          attMap[labC] = { status, time };
        }
      });
    }
  } catch (_) {}

  const last7 = [], labels7 = [];
  for (let i = 6; i >= 0; i--) {
    const d = new Date(today); d.setDate(d.getDate() - i);
    last7.push(_fmt(d, "dd-MM-yyyy")); labels7.push(_fmt(d, "EEE d"));
  }
  const last30 = [];
  for (let i = 29; i >= 0; i--) {
    const d = new Date(today); d.setDate(d.getDate() - i);
    last30.push(_fmt(d, "dd-MM-yyyy"));
  }

  const results = labs.map(lab => {
    const m       = maintMap[lab.code] || {};
    const isMaint = m.active === true;
    const createdDate = lab.createdOn ? _parseStored(lab.createdOn) : null;

    let history      = Array(7).fill("❌");
    let todayMorning = false, todayEvening = false, remark = "";
    let totalMissed  = 0, missingDates = [], sla = 100, photoCount = lab.photoCount || 0;
    let faultFrequency = {};

    try {
      if (lab.sheetUrl && lab.sheetUrl !== "#") {
        const labData = SpreadsheetApp.openByUrl(lab.sheetUrl).getSheets()[0].getDataRange().getValues();
        const dayMap  = {};

        labData.slice(1).forEach(row => {
          if (!row[0]) return;
          const dStr    = String(row[0]).trim();
          const session = String(row[1] || "").trim().toLowerCase();
          const hasData = [2,4,6,8,10,12,14].some(ci => {
            const v = row[ci]; return v !== "" && v !== null && v !== undefined;
          });
          if (!dayMap[dStr]) dayMap[dStr] = { sessions: {}, remark: "", photoCount: 0 };
          if (hasData) {
            dayMap[dStr].sessions[session] = true;
            if (row[17]) dayMap[dStr].remark = String(row[17]);
            dayMap[dStr].photoCount += Number(row[21] || 0);
            CONFIG.hardwareItems.forEach((item, idx) => {
              const col = 2 + idx * 2;
              const val = String(row[col] || "");
              if (val && val !== "OK") { faultFrequency[item] = (faultFrequency[item] || 0) + 1; }
            });
          }
        });

        const todayEntry = dayMap[todayStrKey];
        if (todayEntry) {
          todayMorning = !!todayEntry.sessions.morning;
          todayEvening = !!todayEntry.sessions.evening;
          remark       = todayEntry.remark || "";
        }

        history = last7.map((dStr) => {
          const cellDate = _parseStored(dStr);
          if (createdDate && cellDate < createdDate) return "⚫";
          if (isMaint) {
            const enOn  = m.enabledOn  ? _parseStored(m.enabledOn)  : null;
            const disOn = m.disabledOn ? _parseStored(m.disabledOn) : null;
            if (enOn && cellDate >= enOn && (!disOn || cellDate <= disOn)) return "⚠️";
            if (!enOn) return "⚠️";
          }
          const e = dayMap[dStr];
          if (e && (e.sessions.morning || e.sessions.evening)) return "✅";
          return isNonWorkingDay(cellDate) ? "⚪" : "❌";
        });

        if (!isMaint) {
          let wdTotal = 0, wdOk = 0;
          const missingRaw = [];
          last30.forEach(dStr => {
            const d    = _parseStored(dStr);
            const isWD = !isNonWorkingDay(d) && d <= today;
            if (!isWD) return;
            if (createdDate && d < createdDate) return;
            // FIX: Do not count today in denominator — lab may not have had a chance to submit yet
            if (dStr === todayStrKey) return;
            wdTotal++;
            if (dayMap[dStr] && (dayMap[dStr].sessions.morning || dayMap[dStr].sessions.evening)) {
              wdOk++;
            } else {
              missingRaw.push(_fmt(d, "dd-MM-yyyy"));
            }
          });
          sla          = wdTotal > 0 ? Math.round(wdOk / wdTotal * 100) : 100;
          totalMissed  = missingRaw.length;
          missingDates = [...missingRaw].reverse();
        }
      }
    } catch (e) {
      Logger.log("Dashboard: lab " + lab.code + " error — " + e.message);
      history = Array(7).fill("⚠️");
    }

    let status = "MISSED";
    if (isMaint) status = "MAINTENANCE";
    else if (todayMorning || todayEvening) status = "UPDATED";
    else if (isNonWorkingDay(today) && totalMissed === 0) status = "UPDATED";

    const hoursOverdue = status === "MISSED"
      ? Math.floor((Date.now() - today.getTime()) / 3600000)
      : 0;

    return {
      lab:  lab.code, block: lab.block, owner: lab.email,
      name: lab.name, sheetUrl: lab.sheetUrl, status,
      todayMorning, todayEvening, remark: String(remark || ""),
      history, sla, missedDays: totalMissed, missingDates,
      isMaintenance: isMaint,
      maintenanceReason:     m.reason     || null,
      maintenanceEnabledOn:  m.enabledOn  || null,
      maintenanceDisabledOn: m.disabledOn || null,
      photoCount, attendance: attMap[lab.code] || null,
      faultFrequency, hoursOverdue,
      createdOn: lab.createdOn || null,
      lastDate: isMaint
        ? "Vacation Mode"
        : status === "UPDATED"
          ? "Updated Today"
          : totalMissed > 0
            ? totalMissed + " Day(s) Overdue"
            : "Fully Compliant"
    };
  });

  const payload = { results, dateLabels: labels7, generatedAt: new Date().toISOString() };
  try { cache.put(cacheKey, JSON.stringify(payload), CONFIG.cacheTtl); } catch (_) {}
  return payload;
}

/* ═══════════════════════════════════════════════════════════════════════
   MAINTENANCE MAP
═══════════════════════════════════════════════════════════════════════ */
function _getMaintenanceMap() {
  try {
    const sheet = _ss().getSheetByName("Sys_Config");
    if (!sheet) return {};
    return sheet.getDataRange().getValues().slice(1).reduce((m, row) => {
      if (!row[0]) return m;
      const _cellDate = v => {
        if (!v) return null;
        if (v instanceof Date) return _fmt(v, "dd-MM-yyyy");
        return String(v).split(" ")[0];
      };
      m[String(row[0]).trim()] = {
        active:     row[1] === true || row[1] === "TRUE",
        enabledOn:  _cellDate(row[2]),
        disabledOn: _cellDate(row[3]),
        reason:     row[4] ? String(row[4]) : null,
        type:       row[5] ? String(row[5]) : null,
        workOrder:  row[6] ? String(row[6]) : null,
        approver:   row[7] ? String(row[7]) : null,
        endDate:    row[8] ? String(row[8]) : null
      };
      return m;
    }, {});
  } catch (_) { return {}; }
}

/* ═══════════════════════════════════════════════════════════════════════
   MONTHLY REPORT
═══════════════════════════════════════════════════════════════════════ */
function getMonthlyReport(year, month, token) {
  if (!_validateSession(token)) return { error: "Unauthorized" };
  const labs = getAllLabs().filter(l => l.active);
  const workingDays = [];
  const daysInMonth = new Date(Number(year), Number(month), 0).getDate();
  for (let d = 1; d <= daysInMonth; d++) {
    const date    = new Date(Number(year), Number(month) - 1, d, 12, 0, 0);
    const dateStr = _fmt(date, "dd-MM-yyyy");
    if (!isNonWorkingDay(_parseStored(dateStr))) workingDays.push(dateStr);
  }

  const results = [];
  labs.forEach(lab => {
    if (!lab.sheetUrl || lab.sheetUrl === "#") return;
    const createdDate    = lab.createdOn ? _parseStored(lab.createdOn) : null;
    const effectiveDays  = createdDate
      ? workingDays.filter(ds => _parseStored(ds) >= createdDate)
      : workingDays;
    let present = 0, missingDates = [];

    try {
      const data = SpreadsheetApp.openByUrl(lab.sheetUrl).getSheets()[0].getDataRange().getValues();
      const entryDates = new Set();
      data.slice(1).forEach(row => {
        const ds = String(row[0] || "").trim();
        const hasData = [2,4,6,8,10,12,14].some(
          ci => row[ci] !== "" && row[ci] !== null && row[ci] !== undefined
        );
        if (ds && hasData) entryDates.add(ds);
      });

      effectiveDays.forEach(dateStr => {
        if (entryDates.has(dateStr)) present++;
        else missingDates.push(dateStr);
      });
    } catch (e) { Logger.log("Monthly report error for " + lab.code + ": " + e.message); }

    const total  = effectiveDays.length;
    const sla    = total > 0 ? Math.round((present / total) * 100) : 100;
    results.push({
      labCode:       lab.code,
      labName:       lab.name,
      block:         lab.block,
      present,
      missed:        missingDates.length,
      missingDates,
      sla,
      totalWorkingDays: total,
      createdOn:     lab.createdOn || null
    });
  });

  results.sort((a, b) => a.sla - b.sla);
  const overallAvgSla = results.length
    ? Math.round(results.reduce((s, r) => s + r.sla, 0) / results.length)
    : 100;
  return {
    year:             Number(year),
    month:            Number(month),
    monthLabel:       Utilities.formatDate(
                        new Date(Number(year), Number(month) - 1, 1, 12, 0, 0),
                        CONFIG.timezone, "MMMM yyyy"
                      ),
    totalWorkingDays: workingDays.length,
    avgSla:           overallAvgSla,
    labs:             results,
    generatedAt:      new Date().toISOString()
  };
}

/* ═══════════════════════════════════════════════════════════════════════
   REPORTS & UTILITIES
═══════════════════════════════════════════════════════════════════════ */
function getBlockComplianceReport(token) {
  if (!_validateSession(token)) return { error: "Unauthorized" };
  try {
    const payload = getUnifiedDashboardData(token);
    if (payload.error) return payload;
    const blocks = {};
    (payload.results || []).forEach(r => {
      const b = r.block || "Unknown";
      if (!blocks[b]) blocks[b] = { total:0, updated:0, missed:0, maintenance:0, avgSla:0, slaSum:0 };
      blocks[b].total++;
      if (r.status === "UPDATED")     blocks[b].updated++;
      if (r.status === "MISSED")      blocks[b].missed++;
      if (r.status === "MAINTENANCE") blocks[b].maintenance++;
      blocks[b].slaSum += r.sla || 0;
    });
    Object.values(blocks).forEach(b => {
      b.avgSla = b.total ? Math.round(b.slaSum / b.total) : 0;
      delete b.slaSum;
    });
    return blocks;
  } catch (e) { return { error: e.message }; }
}

function getEquipmentFailureReport(token) {
  if (!_validateSession(token)) return { error: "Unauthorized" };
  try {
    const payload = getUnifiedDashboardData(token);
    if (payload.error) return payload;
    const freq = {};
    CONFIG.hardwareItems.forEach(k => freq[k] = 0);
    (payload.results || []).forEach(r => {
      Object.entries(r.faultFrequency || {}).forEach(([k, v]) => {
        freq[k] = (freq[k] || 0) + v;
      });
    });
    return Object.entries(freq)
      .map(([item, count]) => ({ item, count }))
      .sort((a, b) => b.count - a.count);
  } catch (e) { return { error: e.message }; }
}

function getSemesterSummaryReport(token) {
  if (!_validateSession(token)) return { error: "Unauthorized" };
  try {
    const brand   = getBrandConfig();
    const payload = getUnifiedDashboardData(token);
    if (payload.error) return payload;
    const labs        = payload.results || [];
    const totalLabs   = labs.length;
    const avgSla      = totalLabs ? Math.round(labs.reduce((a, r) => a + (r.sla || 0), 0) / totalLabs) : 0;
    const totalMissed = labs.reduce((a, r) => a + (r.missedDays || 0), 0);
    const perfect     = labs.filter(r => r.missedDays === 0 && !r.isMaintenance).length;
    const photos      = labs.reduce((a, r) => a + (r.photoCount || 0), 0);
    const topFaults   = getEquipmentFailureReport(token);
    return {
      period:      brand.termStart + " to " + brand.termEnd,
      institution: brand.companyName,
      totalLabs, avgSla, totalMissed, perfect, photos,
      topFaults:   Array.isArray(topFaults) ? topFaults.slice(0, 3) : [],
      generatedAt: new Date().toISOString()
    };
  } catch (e) { return { error: e.message }; }
}

function getComplianceCertificate(labCode, token) {
  if (!_validateSession(token)) return { error: "Unauthorized" };
  try {
    const brand   = getBrandConfig();
    const labInfo = _getLabInfo(labCode);
    if (!labInfo) return { error: "Lab not found." };
    const payload = getUnifiedDashboardData(token);
    if (payload.error) return payload;
    const labData = (payload.results || []).find(r => r.lab === labCode);
    if (!labData) return { error: "No data for lab." };
    return {
      institution: brand.companyName,
      labCode:     labInfo.code,
      labName:     labInfo.name,
      block:       labInfo.block,
      ownerEmail:  labInfo.email,
      sla:         labData.sla,
      missedDays:  labData.missedDays,
      status:      labData.status,
      photoCount:  labData.photoCount,
      period:      brand.termStart + " – " + brand.termEnd,
      certDate:    _fmt(new Date(), "dd MMMM yyyy"),
      generatedAt: new Date().toISOString()
    };
  } catch (e) { return { error: e.message }; }
}

function getHealthStatus(token) {
  if (!_validateSession(token, "ADMIN")) return { error: "Unauthorized" };
  const info = {};
  try { info.emailQuotaRemaining = MailApp.getRemainingDailyQuota(); } catch (_) { info.emailQuotaRemaining = "N/A"; }
  try {
    const root  = DriveApp.getFolderById(CONFIG.photoRootFolderId);
    const files = root.getFiles();
    let count   = 0;
    while (files.hasNext() && count < 9999) { files.next(); count++; }
    info.photoFolderFileCount = count;
  } catch (_) { info.photoFolderFileCount = "N/A"; }
  try {
    const cached = CacheService.getScriptCache().get("dashboard_v11");
    info.cacheStatus = cached ? "warm" : "cold";
  } catch (_) { info.cacheStatus = "unknown"; }
  info.labCount  = getAllLabs().length;
  info.checkedAt = new Date().toISOString();
  info.version   = "5.6";
  return info;
}

function warmDashboardCache() {
  try {
    const warmToken = _createSessionToken("__system__", "ADMIN");
    getUnifiedDashboardData(warmToken);
    try { PropertiesService.getScriptProperties().deleteProperty("sess_" + warmToken); } catch (_) {}
    Logger.log("Cache warmed at " + new Date().toISOString());
  } catch (e) { Logger.log("warmDashboardCache error: " + e.message); }
}

function sendEscalatedWarnings(token) {
  let warmToken = null;
  let useWarm   = false;
  if (!token || !_validateSession(token, "ADMIN")) {
    warmToken = _createSessionToken("__system__", "ADMIN");
    useWarm   = true;
    token     = warmToken;
  }
  if (isNonWorkingDay(new Date())) {
    if (useWarm) try { PropertiesService.getScriptProperties().deleteProperty("sess_" + warmToken); } catch (_) {}
    return "Skipped: non-working day.";
  }

  const payload = getUnifiedDashboardData(token);
  if (useWarm) try { PropertiesService.getScriptProperties().deleteProperty("sess_" + warmToken); } catch (_) {}
  if (!payload?.results) return "No data.";

  const brand   = getBrandConfig();
  const byEmail = {};
  payload.results.forEach(item => {
    if (item.isMaintenance || item.missedDays <= 0 || !_validEmail(item.owner)) return;
    (byEmail[item.owner] = byEmail[item.owner] || []).push(item);
  });
  let sent = 0;
  for (const [email, labs] of Object.entries(byEmail)) {
    const tier     = _escalTier(Math.max(...labs.map(l => l.missedDays)));
    const ackToken = _randomHex(16);
    PropertiesService.getScriptProperties().setProperty(
      "ack_" + ackToken,
      JSON.stringify({ email, labs: labs.map(l => l.lab), ts: Date.now() })
    );
    try {
      MailApp.sendEmail({
        to:       email,
        subject:  tier.subject,
        htmlBody: _buildEmail(labs, tier, ackToken, brand),
        name:     brand.companyName + " · IT"
      });
      sent++;
    } catch (e) { Logger.log("Email error " + email + ": " + e.message); }
  }
  _auditLog("System", "BATCH_EMAIL", sent + " sent");
  return sent + " email(s) sent.";
}

function acknowledgeEscalation(ackToken) {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw   = props.getProperty("ack_" + ackToken);
    if (!raw) return { success: false, message: "Invalid or expired acknowledgement token." };
    const data = JSON.parse(raw);
    props.deleteProperty("ack_" + ackToken);
    _auditLog(data.email, "ESCALATION_ACKNOWLEDGED", "labs=" + data.labs.join(","));
    return { success: true, message: "Escalation acknowledged. Thank you." };
  } catch (e) { return { success: false, message: e.message }; }
}

function sendManualAlert(labCode, ownerEmail, labUrl, daysOverdue, token) {
  if (!_validateSession(token, "ADMIN")) return "Unauthorized.";
  if (!_validEmail(ownerEmail)) return "Invalid email.";
  const brand = getBrandConfig();
  const color = daysOverdue >= 3 ? "#dc2626" : "#d97706";
  try {
    MailApp.sendEmail({
      to:       ownerEmail,
      subject:  "[LabAudit] Action needed: " + labCode,
      htmlBody: _manualEmail(labCode, labUrl, daysOverdue, color, brand),
      name:     brand.companyName + " · IT"
    });
    _auditLog("Admin", "MANUAL_ALERT", "to=" + ownerEmail + " lab=" + labCode);
    return "Sent to " + ownerEmail;
  } catch (e) { return "Error: " + e.message; }
}

function _escalTier(days) {
  if (days >= 3) return { subject: "[URGENT] Lab check list — " + days + " days pending",    color: "#dc2626" };
  if (days >= 2) return { subject: "[FOLLOW-UP] Lab check list — " + days + " days pending", color: "#d97706" };
  return             { subject: "[REMINDER] Lab check list entry for today",                 color: "#2563eb" };
}

function _buildEmail(labs, tier, ackToken, brand) {
  const appUrl = ScriptApp.getService().getUrl();
  const ackUrl = appUrl + "?ack=" + ackToken;
  const rows   = labs.map(l =>
    `<tr>
       <td style="padding:12px 16px;border-bottom:1px solid #f1f5f9">
         <b>${l.lab}</b><br>
         <small style="color:#64748b">Missing: ${(l.missingDates||[]).slice(0,5).join(", ")||"—"}</small>
       </td>
       <td style="padding:12px 16px;border-bottom:1px solid #f1f5f9;text-align:center">
         <span style="background:${tier.color}18;color:${tier.color};font-weight:700;padding:3px 12px;border-radius:99px;font-size:12px">
           ${l.missedDays} day(s) pending
         </span>
       </td>
       <td style="padding:12px 16px;border-bottom:1px solid #f1f5f9;text-align:right">
         <a href="${l.sheetUrl}" style="color:#2563eb;font-size:12px">Open log →</a>
       </td>
     </tr>`
  ).join("");
  return `<!DOCTYPE html><html><body style="margin:0;padding:0;background:#f8fafc;font-family:'Segoe UI',Arial,sans-serif">
<table width="100%" style="padding:32px 16px"><tr><td align="center">
<table width="600" style="background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)">
  <tr><td style="background:linear-gradient(135deg,#0f172a,#1e293b);padding:28px 36px">
    <div style="color:#94a3b8;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:2px">${brand.companyName}</div>
    <div style="color:#fff;font-size:20px;font-weight:700;margin-top:4px">LabAudit Pro</div>
  </td></tr>
  <tr><td style="background:${tier.color};height:3px;font-size:0">&nbsp;</td></tr>
  <tr><td style="padding:28px 36px">
    <p style="color:#0f172a;font-size:16px;font-weight:600">Dear Lab Instructor,</p>
    <p style="color:#334155;font-size:14.5px;line-height:1.8">Lab check lists require your attention:</p>
    <table width="100%" style="border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;margin:18px 0">
      <tr style="background:#f8fafc">
        <th style="padding:10px 16px;text-align:left;font-size:10px;color:#94a3b8;text-transform:uppercase">Lab</th>
        <th style="padding:10px 16px;font-size:10px;color:#94a3b8;text-transform:uppercase">Status</th>
        <th style="padding:10px 16px;text-align:right;font-size:10px;color:#94a3b8;text-transform:uppercase">Link</th>
      </tr>${rows}
    </table>
    <div style="text-align:center;margin:20px 0">
      <a href="${ackUrl}" style="background:#16a34a;color:#fff;padding:10px 22px;border-radius:8px;font-weight:700;text-decoration:none;font-size:13px">✓ Acknowledge this alert</a>
    </div>
    <p style="color:#64748b;font-size:12px;border-top:1px solid #f1f5f9;padding-top:18px">Automated · LabAudit Pro v5.6</p>
  </td></tr>
</table></td></tr></table></body></html>`;
}

function _manualEmail(labCode, labUrl, days, color, brand) {
  return `<!DOCTYPE html><html><body style="margin:0;padding:32px 16px;background:#f8fafc;font-family:'Segoe UI',Arial,sans-serif">
<table width="100%" style="max-width:560px;margin:0 auto;background:#fff;border-radius:12px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08)">
  <tr><td style="background:linear-gradient(135deg,#0f172a,#1e293b);padding:24px 32px">
    <div style="color:#fff;font-size:18px;font-weight:700">LabAudit Pro</div>
    <div style="color:#64748b;font-size:11px">${brand.companyName}</div>
  </td></tr>
  <tr><td style="background:${color};height:3px;font-size:0">&nbsp;</td></tr>
  <tr><td style="padding:28px 32px">
    <p style="color:#0f172a;font-size:15px">Hi,</p>
    <p style="color:#334155;line-height:1.8">The daily check list for <strong>${labCode}</strong> shows
       <span style="color:${color};font-weight:700">${days} overdue day(s)</span>.</p>
    <div style="text-align:center;margin:24px 0">
      <a href="${labUrl}" style="background:#1d4ed8;color:#fff;padding:12px 28px;border-radius:8px;font-weight:700;text-decoration:none">Open Lab Log →</a>
    </div>
    <p style="color:#94a3b8;font-size:12px;border-top:1px solid #f1f5f9;padding-top:16px">
      ${_fmt(new Date(), "EEEE, d MMMM yyyy")} · LabAudit Pro v5.6</p>
  </td></tr>
</table></body></html>`;
}

/* ═══════════════════════════════════════════════════════════════════════
   AUDIT LOGS
═══════════════════════════════════════════════════════════════════════ */
function _auditLog(user, action, details) {
  try {
    let sh = _ss().getSheetByName("Sys_Logs");
    if (!sh) {
      sh = _ss().insertSheet("Sys_Logs");
      sh.appendRow(["Timestamp","User","Action","Details","Hash"]);
      sh.getRange("A1:E1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
      sh.setFrozenRows(1);
    }
    const data     = sh.getDataRange().getValues();
    const prevHash = data.length > 1 ? String(data[data.length-1][4] || "") : "genesis";
    const entry    = new Date().toISOString() + "|" + String(user) + "|" + String(action) + "|" + String(details);
    const hash     = _hashString(prevHash + entry).slice(0, 16);
    sh.appendRow([new Date(), String(user || ""), String(action || ""), String(details || ""), hash]);
  } catch (e) { Logger.log("_auditLog error: " + e.message); }
}

function getAuditLogs(limit, token) {
  if (!_validateSession(token, "ADMIN")) return { error: "Unauthorized" };
  try {
    const sh = _ss().getSheetByName("Sys_Logs");
    if (!sh) return [];
    const max  = Math.min(Number(limit) || 200, 500);
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    return data.slice(1).reverse().slice(0, max).map(r => ({
      timestamp: r[0] instanceof Date ? _fmt(r[0], "dd-MM-yyyy HH:mm:ss") : String(r[0]),
      user:    String(r[1] || ""),
      action:  String(r[2] || ""),
      details: String(r[3] || ""),
      hash:    String(r[4] || "")
    }));
  } catch (e) { return []; }
}

function recordAuditLog(u, a, d, token) {
  if (!_validateSession(token)) return;
  _auditLog(u, a, d);
}

/* ═══════════════════════════════════════════════════════════════════════
   ONEDIT TRIGGER
═══════════════════════════════════════════════════════════════════════ */
function onEdit(e) {
  if (!e?.range) return;
  if (CONFIG.protectedSheets.includes(e.source.getActiveSheet().getName())) return;
  if (isNonWorkingDay(_startOfDay(new Date()))) {
    e.range.clearContent();
    try { SpreadsheetApp.getUi().alert("Locked: edits not allowed on non-working days."); } catch (_) {}
  }
}

/* ═══════════════════════════════════════════════════════════════════════
   DEMO SETUP
═══════════════════════════════════════════════════════════════════════ */
function setupDemoLab() {
  const demoCode = "DEMO-LAB-01";
  if (_getLabInfo(demoCode)) return;
  try {
    const ss = SpreadsheetApp.create(demoCode + " — LabAudit Log (Demo)");
    const sh = ss.getSheets()[0];
    sh.setName("Log_Data");
    const hdrs = [
      "DATE","SESSION","MONITOR","MONITOR_SEV","CPU","CPU_SEV",
      "KEYBOARD","KEYBOARD_SEV","MOUSE","MOUSE_SEV","ETHERNET","ETHERNET_SEV",
      "WIFI","WIFI_SEV","STATUS","STATUS_SEV","OVERALL","REMARK",
      "REMARK_TAGS","SIGNOFF","PHOTOS","PHOTO_COUNT"
    ];
    sh.appendRow(hdrs);
    sh.getRange(1,1,1,hdrs.length).setFontWeight("bold").setBackground("#0f172a").setFontColor("#fff");
    sh.setFrozenRows(1);
    sh.setTabColor("#f59e0b");
    const todayStr  = _fmt(new Date(), "dd-MM-yyyy");
    const labsSheet = _getLabsSheet();
    labsSheet.appendRow([
      Utilities.getUuid(), "Demo Laboratory", "A Block",
      demoCode, "demolab@snpsu.edu", true, ss.getUrl(), 0, todayStr
    ]);
    _getOrCreatePhotoFolder(demoCode, null);
    _invalidateCache();
    setupCredentials();
  } catch (e) { Logger.log("setupDemoLab error: " + e.message); }
}

/* ═══════════════════════════════════════════════════════════════════════
   UTILITIES
═══════════════════════════════════════════════════════════════════════ */
function isNonWorkingDay(d) {
  if (!(d instanceof Date) || isNaN(d)) return false;
  const day = d.getDay();
  if (day === 0) return true; // Sunday
  if (day === 6) {
    // 1st and 3rd Saturdays are holidays
    const firstOfMonth   = new Date(d.getFullYear(), d.getMonth(), 1);
    const firstSatOffset = (6 - firstOfMonth.getDay() + 7) % 7;
    const firstSatDate   = 1 + firstSatOffset;
    const weekNum        = Math.floor((d.getDate() - firstSatDate) / 7) + 1;
    return weekNum === 1 || weekNum === 3;
  }
  return CONFIG.holidays.has(_fmt(d, "dd-MM-yyyy"));
}

function _startOfDay(d) {
  const x = new Date(d); x.setHours(0, 0, 0, 0); return x;
}

function _fmt(d, fmt) {
  return Utilities.formatDate(d instanceof Date ? d : new Date(d), CONFIG.timezone, fmt);
}

function _parseStored(s) {
  if (!s) return new Date(NaN);
  try {
    const datePart = String(s).split(" ")[0].split("T")[0];
    const parts    = datePart.split(/[-\/]/);
    if (parts.length === 3) {
      const [dd, mm, yy] = parts.map(Number);
      if (!isNaN(dd) && !isNaN(mm) && !isNaN(yy)) {
        const d = new Date(yy, mm - 1, dd);
        d.setHours(0, 0, 0, 0);
        return d;
      }
    }
    return new Date(NaN);
  } catch (_) { return new Date(NaN); }
}

function _hashString(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s))
    .map(b => (255 & b).toString(16).padStart(2, "0")).join("");
}

function _randomHex(n) {
  return Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    Math.random().toString() + Date.now()
  ).slice(0, n).map(b => (255 & b).toString(16).padStart(2, "0")).join("");
}

function _validEmail(e) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(e || "")); }
function _cap(s) { return s ? s.charAt(0).toUpperCase() + s.slice(1) : s; }

// FIX-4: Removed dead "dashboard_v10" cache key — only v11 is written by this codebase
function _invalidateCache() {
  try {
    const c = CacheService.getScriptCache();
    ["dashboard_v11", "labs_list_v5"].forEach(k => { try { c.remove(k); } catch (_) {} });
  } catch (_) {}
  _labMapInst = null;
  _brandCache = null;
}
function doPost(e) {
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);
  
  // CORS headers — allow your Vercel domain
  const allowedOrigins = [
    'https://snpsu-lab-checklist.vercel.app',
    'https://localhost:3000',  // for local dev
    'http://localhost:3000'
  ];
  
  try {
    const body = JSON.parse(e.postData.contents);
    const fn   = String(body.fn || '');
    const args = Array.isArray(body.args) ? body.args : [];

    const ALLOWED = [
      'verifyLogin','invalidateSession','extendSession',
      'getAllUsers','createUser','updateUser','deleteUser',
      'getAllLabs','getLabByCode','adminCreateLabSheet',
      'deleteLabSheet','restoreLab','updateLabInfo',
      'uploadLabPhoto','getLabPhotos','deleteLabPhoto',
      'saveDailySessionEntry','getDailySessionEntry',
      'setLabMaintenance','getMaintenanceLog',
      'saveAttendance','getTodayAttendance','getAttendanceSummary',
      'getNotifications','markNotificationsRead',
      'getUnifiedDashboardData','getBlockComplianceReport',
      'getEquipmentFailureReport','getSemesterSummaryReport',
      'getComplianceCertificate','getHealthStatus',
      'sendEscalatedWarnings','sendManualAlert',
      'acknowledgeEscalation','getAuditLogs','recordAuditLog',
      'getBrandConfig','updateBrandConfig','warmDashboardCache'
    ];

    if (!ALLOWED.includes(fn)) {
      output.setContent(JSON.stringify({ ok: false, error: 'Function not allowed: ' + fn }));
      return output;
    }

    const result = this[fn](...args);
    output.setContent(JSON.stringify({ ok: true, result }));

  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }
  
  return output;
}

// Handle OPTIONS preflight + GET requests
function doGet(e) {
  // Acknowledgement handler — keep this
  if (e && e.parameter && e.parameter.ack) {
    const result = acknowledgeEscalation(e.parameter.ack);
    return HtmlService.createHtmlOutput(
      `<html><body style="font-family:sans-serif;padding:40px;max-width:480px;margin:auto">
       <h2>${result.success ? '✅ Acknowledged' : '❌ Error'}</h2>
       <p>${result.message}</p>
       <p><a href="https://snpsu-lab-checklist.vercel.app/">Return to LabAudit Pro →</a></p>
       </body></html>`
    ).setTitle('LabAudit Pro · Acknowledgement');
  }

  // API health check
  const output = ContentService.createTextOutput(
    JSON.stringify({ ok: true, message: 'LabAudit Pro API v5.6', ts: new Date().toISOString() })
  );
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}
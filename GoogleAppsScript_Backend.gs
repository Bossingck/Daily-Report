// ================================================================
//  INOC Field Service — Google Apps Script Backend
//  HEAD Deployment — แก้ code กด Ctrl+S ไม่ต้อง Deploy ใหม่
//
//  วิธีใส่ API Key (ทำครั้งเดียว):
//  Project Settings → Script Properties → Add property
//  Key: ANTHROPIC_API_KEY
//  Value: sk-ant-api03-xxxxxxxxxxxxxxxx
// ================================================================

const CONFIG = {
  SHEET_ID     : "1Pt0ZOFsI_b6mLd2XlN680V_NNhWfH6_xoKc5RLzyHLk",
  SHEET_ARRIVE : "1st_Arrive_Log",
  SHEET_PDT    : "PDT_Log",
  SHEET_AUDIT  : "Audit_Trail",
  SHEET_BACKUP : "Backup_Before_Update",
  TIMEZONE     : "Asia/Bangkok",
  AI_MODEL     : "claude-sonnet-4-20250514",
  AI_MAX_TOKENS: 1000,
};

// ── Headers ──────────────────────────────────────────────────────
const ARRIVE_HEADERS = [
  "Saved At","File Name","Record Date","Team","Skill","Province",
  "Zone (FM)","Supervisor","Manager","FM Office","Clock In",
  "Last Work","Last Action","WO Today","Saved By"
];
const PDT_HEADERS = [
  "Saved At","File Name","Record Date","Team","Skill","Province",
  "Zone (FM)","Supervisor","Manager","FM Office","Clock In",
  "Man Hour","WRK","Saved By"
];
const AUDIT_HEADERS = [
  "Timestamp","Action","Sheet","Record Date","File Name",
  "Records Written","Records Deleted","Operator","Note"
];
const BACKUP_HEADERS = [
  "Backed Up At","Original Sheet","Record Date","File Name","Backed Up By",
  "Col1","Col2","Col3","Col4","Col5","Col6","Col7","Col8","Col9","Col10",
  "Col11","Col12","Col13","Col14","Col15"
];

// ================================================================
//  doGet — Health / Read / AI Analyze
// ================================================================
function doGet(e) {
  try {
    const p      = e.parameter || {};
    const action = p.action || "health";

    if (action === "health") {
      const hasKey = !!_getApiKey();
      return _ok({
        message  : "INOC GAS Backend Online ✅",
        version  : "HEAD (auto-update)",
        sheet_id : CONFIG.SHEET_ID.substring(0, 8) + "...",
        ai_ready : hasKey,
        ts       : _now(),
      });
    }
    if (action === "read") {
      const name = p.sheet === "pdt" ? CONFIG.SHEET_PDT : CONFIG.SHEET_ARRIVE;
      return _ok({ data: _readSheet(name) });
    }
    if (action === "audit")  return _ok({ data: _readSheet(CONFIG.SHEET_AUDIT) });
    if (action === "backup") return _ok({ data: _readSheet(CONFIG.SHEET_BACKUP) });
    if (action === "dates") {
      const name  = p.sheet === "pdt" ? CONFIG.SHEET_PDT : CONFIG.SHEET_ARRIVE;
      const rows  = _readSheet(name);
      const dates = [...new Set(rows.map(r => String(r["Record Date"]).substring(0,10)))].sort();
      return _ok({ dates });
    }
    if (action === "restore") {
      return _restore(p.date, p.sheet, p.operator || "Web App");
    }

    return _err("Unknown action: " + action);
  } catch (err) {
    return _err(err.message);
  }
}

// ================================================================
//  doPost — Save / Update / AI Analyze
// ================================================================
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action || "save";

    // ── Route: AI Analyze ──────────────────────────────────────
    if (action === "ai_analyze") {
      return _aiAnalyze(body.payload || {});
    }

    // ── Route: Save / Update ───────────────────────────────────
    const type       = body.type   || "1st_arrive";
    const rows       = body.data   || [];
    const meta       = body.meta   || {};
    const filename   = meta.filename   || "unknown";
    const recordDate = meta.recordDate || (rows[0] && rows[0].date) || _today();
    const operator   = meta.operator   || "Web App";

    const sheetName = type === "pdt" ? CONFIG.SHEET_PDT : CONFIG.SHEET_ARRIVE;
    const headers   = type === "pdt" ? PDT_HEADERS      : ARRIVE_HEADERS;

    _ensureSheet(sheetName, headers);
    _ensureSheet(CONFIG.SHEET_AUDIT,  AUDIT_HEADERS);
    _ensureSheet(CONFIG.SHEET_BACKUP, BACKUP_HEADERS);

    const ss       = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh       = ss.getSheetByName(sheetName);
    const backedUp = _backupByDate(ss, sh, sheetName, recordDate, filename, operator);
    const deleted  = _deleteByDate(sh, recordDate, 3);

    if (!rows.length) return _err("ไม่มีข้อมูลส่งมา");
    const written  = _appendRows(sh, rows, type, filename, recordDate, operator);
    _audit(ss, action, sheetName, recordDate, filename, written, deleted, backedUp, operator);

    return _ok({
      written, deleted, backedUp,
      message: `${action === "update" ? "Update" : "Save"} สำเร็จ ${written} รายการ | Backup ${backedUp} แถว | ลบเก่า ${deleted} แถว`,
    });

  } catch (err) {
    return _err(err.message);
  }
}

// ================================================================
//  AI Analyze — เรียก Claude API ผ่าน GAS (API Key ปลอดภัยใน Properties)
// ================================================================
function _aiAnalyze(payload) {
  const apiKey = _getApiKey();
  if (!apiKey) {
    return _err("ยังไม่ได้ตั้งค่า API Key — ไปที่ Project Settings → Script Properties → เพิ่ม ANTHROPIC_API_KEY");
  }

  // สร้าง prompt จาก payload ที่ HTML ส่งมา
  const { totalRows, days, avgPerDay, zoneCounts, topSups, streaks, dayPattern, zoneTrend } = payload;

  const prompt = `คุณคือ INOC AI Analyst ระบบ Field Service Management
วิเคราะห์ข้อมูล 1st Arrive (ทีมที่ไม่มี First Arrive ก่อน 10:00 น.) ต่อไปนี้:

📊 ภาพรวม:
- รายการสายทั้งหมด: ${totalRows || 0} รายการ
- ช่วงเวลา: ${days || 0} วัน (เฉลี่ย ${avgPerDay || 0} รายการ/วัน)

🗺 จำนวนสายตาม Zone:
${Object.entries(zoneCounts || {}).map(([z,c]) => `- ${z}: ${c} รายการ`).join("\n") || "ไม่มีข้อมูล"}

🏆 Top Supervisor ที่สายมาก:
${(topSups || []).slice(0,8).map((s,i) => `${i+1}. ${s.name}: ${s.count} ครั้ง`).join("\n") || "ไม่มีข้อมูล"}

🔥 Streak สายติดต่อกัน (≥2 วัน):
${(streaks || []).slice(0,6).map(s => `- ${s.name}: ${s.maxStreak} วันติดกัน`).join("\n") || "ไม่พบ"}

⚠️ Day Pattern (สายซ้ำวันเดิม):
${(dayPattern || []).slice(0,5).map(p => `- ${p.name} สายวัน${p.day} ${p.count} ครั้ง`).join("\n") || "ไม่พบ"}

📈 Zone Trend (สัปดาห์นี้ vs ก่อน):
${(zoneTrend || []).map(z => `- ${z.zone}: ${z.changePct > 0 ? "+" : ""}${z.changePct}%`).join("\n") || "ไม่มีข้อมูล"}

วิเคราะห์เป็นภาษาไทย ตอบเฉพาะ JSON array ไม่มี markdown backtick:
[{"icon":"...","text":"..."},...]
4 ประเด็น: 1) pattern/แนวโน้มน่ากังวล 2) Zone ที่มีปัญหา 3) บุคคลที่ต้องติดตามด่วน 4) คำแนะนำเชิงปฏิบัติสำหรับ Manager`;

  try {
    const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method      : "post",
      contentType : "application/json",
      headers     : {
        "x-api-key"        : apiKey,
        "anthropic-version": "2023-06-01",
        "content-type"     : "application/json",
      },
      payload: JSON.stringify({
        model      : CONFIG.AI_MODEL,
        max_tokens : CONFIG.AI_MAX_TOKENS,
        messages   : [{ role: "user", content: prompt }],
      }),
      muteHttpExceptions: true,
    });

    const status = response.getResponseCode();
    const result = JSON.parse(response.getContentText());

    if (status !== 200) {
      return _err(`Claude API error ${status}: ${result.error?.message || "unknown"}`);
    }

    const text = (result.content || []).map(c => c.text || "").join("");
    let insights = [];
    try {
      insights = JSON.parse(text.replace(/```json|```/g, "").trim());
    } catch(_) {
      insights = [{ icon: "📊", text }];
    }

    return _ok({ insights });

  } catch (err) {
    return _err("เรียก Claude API ไม่ได้: " + err.message);
  }
}

// ── Get API Key from Script Properties (ปลอดภัย ไม่อยู่ใน code) ─
function _getApiKey() {
  try {
    return PropertiesService.getScriptProperties().getProperty("ANTHROPIC_API_KEY") || "";
  } catch(_) {
    return "";
  }
}

// ================================================================
//  Restore — คืนข้อมูลจาก Backup
// ================================================================
function _restore(dateStr, sheetType, operator) {
  try {
    if (!dateStr) return _err("กรุณาระบุวันที่");
    const ss        = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const backupSh  = ss.getSheetByName(CONFIG.SHEET_BACKUP);
    const sheetName = sheetType === "pdt" ? CONFIG.SHEET_PDT : CONFIG.SHEET_ARRIVE;
    const targetSh  = ss.getSheetByName(sheetName);
    if (!backupSh || backupSh.getLastRow() < 2) return _err("ไม่มีข้อมูลใน Backup");

    const target  = dateStr.substring(0, 10);
    const allVals = backupSh.getRange(2, 1, backupSh.getLastRow()-1, 20).getValues();
    const toRestore = allVals.filter(r =>
      String(r[2]).substring(0,10) === target && String(r[1]) === sheetName
    );
    if (!toRestore.length) return _err(`ไม่พบ Backup วันที่ ${dateStr}`);

    const deleted = _deleteByDate(targetSh, dateStr, 3);
    const batch   = toRestore.map(r => r.slice(5, 20));
    if (batch.length) {
      targetSh.getRange(targetSh.getLastRow()+1, 1, batch.length, batch[0].length).setValues(batch);
    }
    _audit(ss, "RESTORE", sheetName, dateStr, "backup", batch.length, deleted, toRestore.length, operator);
    return _ok({ restored: batch.length, message: `Restore สำเร็จ ${batch.length} แถว` });
  } catch(err) {
    return _err("Restore failed: " + err.message);
  }
}

// ================================================================
//  Internal Helpers
// ================================================================
function _backupByDate(ss, sheet, sheetName, dateStr, filename, operator) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const target  = String(dateStr).substring(0, 10);
  const allVals = sheet.getRange(2, 1, last-1, sheet.getLastColumn()).getValues();
  const toBack  = allVals.filter(r => String(r[2]).substring(0,10) === target);
  if (!toBack.length) return 0;
  const backupSh = ss.getSheetByName(CONFIG.SHEET_BACKUP);
  const backedAt = _now();
  const batch    = toBack.map(r => [backedAt, sheetName, target, filename, operator, ...r.slice(0, 15)]);
  const startRow = backupSh.getLastRow() + 1;
  backupSh.getRange(startRow, 1, batch.length, batch[0].length).setValues(batch);
  backupSh.getRange(startRow, 1, batch.length, 5).setBackground("#FFF8E7").setFontColor("#5C4000");
  return toBack.length;
}

function _appendRows(sh, rows, type, filename, recordDate, operator) {
  const savedAt = _now();
  const batch = rows.map(r => {
    const date = r.date || recordDate;
    return type === "pdt"
      ? [savedAt, filename, date, r.team||"", r.skill||"", r.province||"", r.fmprefix||r.zone||"",
         r.sup||"", r.manager||"", r.woid||"", r.timeclock||"", r.manhour||"", r.wrk||"", operator]
      : [savedAt, filename, date, r.team||"", r.skill||"", r.province||"", r.fmprefix||r.zone||"",
         r.sup||"", r.manager||"", r.woid||"", r.timeclock||"", r.lastwork||"", r.lastaction||"",
         r.wo_today||"", operator];
  });
  if (batch.length) {
    sh.getRange(sh.getLastRow()+1, 1, batch.length, batch[0].length).setValues(batch);
  }
  return batch.length;
}

function _deleteByDate(sheet, dateStr, dateCol) {
  const last = sheet.getLastRow();
  if (last < 2) return 0;
  const target = String(dateStr).substring(0, 10);
  const vals   = sheet.getRange(2, dateCol, last-1, 1).getValues();
  let deleted  = 0;
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]).substring(0,10) === target) {
      sheet.deleteRow(i + 2);
      deleted++;
    }
  }
  return deleted;
}

function _ensureSheet(name, headers) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let sh   = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setBackground("#1a1e3c").setFontColor("#ffffff").setFontWeight("bold");
    sh.setColumnWidths(1, 2, 185);
    sh.setColumnWidth(3, 110);
  }
  return sh;
}

function _readSheet(name) {
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const hdrs = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const vals = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  return vals.map(row => {
    const obj = {};
    hdrs.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

function _audit(ss, action, sheetName, recordDate, filename, written, deleted, backedUp, operator) {
  const sh   = ss.getSheetByName(CONFIG.SHEET_AUDIT);
  const note = action === "ai_analyze" ? "AI analysis via GAS proxy"
             : backedUp > 0 ? `Backup ${backedUp} แถวก่อน override`
             : action === "RESTORE" ? "Restore จาก Backup" : "";
  sh.appendRow([_now(), action.toUpperCase(), sheetName, recordDate, filename, written, deleted, operator, note]);
}

function _now()   { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss"); }
function _today() { return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd"); }
function _ok(obj) {
  return ContentService.createTextOutput(JSON.stringify({ status:"ok", ...obj }))
    .setMimeType(ContentService.MimeType.JSON);
}
function _err(msg) {
  return ContentService.createTextOutput(JSON.stringify({ status:"error", message:msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

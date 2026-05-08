const SPREADSHEET_ID = "1MfLXPjhzkgqHiovQAUDglWKx2C6FzAJge9wJ6aOBjmk";
const SHEET_NAME = "submissions";
const DRIVE_FOLDER_ID = "13Wdf68VbudCZyRa4S5xaUfOxfjElgA5C";

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : "{}");

    const studentName = safeText_(payload?.student?.name);
    const studentClass = safeText_(payload?.student?.class);
    const lkpdSlug = safeText_(payload?.lkpd?.slug) || "lkpd";
    const lkpdTitle = safeText_(payload?.lkpd?.title) || "LKPD";

    if (!studentName || !studentClass) return json_({ ok: false, error: "Nama dan Kelas wajib." });

    const uiReason = safeText_(payload?.tasks?.ui?.reason);
    const uxReason = safeText_(payload?.tasks?.ux?.reason);
    const uiShot = payload?.tasks?.ui?.screenshot;
    const uxShot = payload?.tasks?.ux?.screenshot;

    if (!uiReason || !uxReason) return json_({ ok: false, error: "Alasan UI dan UX wajib." });
    if (!uiShot || !uxShot) return json_({ ok: false, error: "Screenshot UI dan UX wajib." });

    const submissionId = Utilities.getUuid();
    const now = new Date();

    const uiUrl = saveScreenshot_(submissionId, lkpdSlug, studentName, studentClass, "UI", uiShot);
    const uxUrl = saveScreenshot_(submissionId, lkpdSlug, studentName, studentClass, "UX", uxShot);

    const sheet = ensureSheet_();
    const answersJson = JSON.stringify(payload?.answers || {});
    const ua = safeText_(payload?.client?.userAgent);

    sheet.appendRow([
      submissionId,
      now.toISOString(),
      lkpdSlug,
      lkpdTitle,
      studentName,
      studentClass,
      uiUrl,
      uiReason,
      uxUrl,
      uxReason,
      answersJson,
      ua
    ]);

    return json_({ ok: true, submissionId, uiUrl, uxUrl });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function ensureSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      "submission_id",
      "created_at",
      "lkpd_slug",
      "lkpd_title",
      "student_name",
      "student_class",
      "ui_screenshot_url",
      "ui_reason",
      "ux_screenshot_url",
      "ux_reason",
      "answers_json",
      "user_agent"
    ]);
  }
  return sh;
}

function saveScreenshot_(submissionId, lkpdSlug, studentName, studentClass, kind, shot) {
  const base64 = safeText_(shot?.base64);
  const mimeType = safeText_(shot?.mimeType) || "image/jpeg";
  const filenameIn = safeText_(shot?.filename) || (kind.toLowerCase() + ".jpg");

  if (!base64) throw new Error("Screenshot kosong: " + kind);

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType, filenameIn);

  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const safeName = fileSafe_(studentName);
  const safeClass = fileSafe_(studentClass);
  const safeSlug = fileSafe_(lkpdSlug);

  const outName = `${safeSlug}__${safeName}__${safeClass}__${kind}__${submissionId}.jpg`;
  const file = folder.createFile(blob).setName(outName);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function safeText_(v) {
  return String(v == null ? "" : v).trim().replace(/\s+/g, " ").slice(0, 20000);
}

function fileSafe_(s) {
  return safeText_(s).replace(/[^\w\- ]+/g, "").trim().replace(/\s+/g, "_").slice(0, 80) || "x";
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
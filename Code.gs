/*************************************************
 * Adaptive Intelligence 2026
 * Registration + QR + Email + Check-in
 * Includes Web App endpoints for scanner frontend
 *************************************************/

const CONFIG = {
  spreadsheetId: '1OX-uDDWuC9Pzv80S1M9S-p6s_VDd-4NoMx15gyykozE',
  sheetName: 'การตอบแบบฟอร์ม 1',

  apiName: 'Adaptive Intelligence 2026 Check-in API',
  apiVersion: '1.0.0',
  lockTimeoutMs: 30000,

  eventName: 'Adaptive Intelligence 2026',
  eventDateText: 'วันอังคารที่ 28 เมษายน พ.ศ. 2569',
  eventTimeText: 'ลงทะเบียน 15:00 - 15:30 น.',
  eventLocationText: 'fabcafe Bangkok อาคารไปรษณีย์กลาง Back building 3rd floor, ถ.เจริญกรุง แขวงบางรัก เขตบางรัก กรุงเทพมหานคร 10500',
  eventStartIso: '2026-04-28T15:00:00+07:00',
  eventEndIso: '2026-04-28T15:30:00+07:00',
  eventCalendarDescription: 'กรุณานำ QR Code นี้มาแสดงที่จุดลงทะเบียนเพื่อเช็กอินหน้างาน',

  contactName: '______',
  contactEmail: '____@gmail.com',
  contactPhone: '094-351-9265',

  registrationPrefix: 'ADAP26',
  qrSize: 240,
  subjectPrefix: '[ยืนยันการลงทะเบียน]',
  frontendBaseUrl: 'https://kaweewatworr-cyber.github.io/smes-github-pages-scanner/',
  checkinWebAppUrl: '',
  defaultCheckedInBy: 'Staff'
};

/*************************************************
 * WEB APP API
 *************************************************/
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || '';

  if (action === 'status') {
    return jsonOutput(buildStatusPayload_());
  }

  if (action === 'lookup') {
    const registrationId = String((e && e.parameter && e.parameter.registrationId) || '').trim();
    return jsonOutput(findByRegistrationId(registrationId));
  }

  return jsonOutput(
    Object.assign({}, buildStatusPayload_(), {
      message: 'Use POST action=checkin or action=lookup'
    })
  );
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const action = String(data.action || '').trim();

    if (action === 'status') {
      return jsonOutput(buildStatusPayload_());
    }

    if (action === 'lookup') {
      const registrationId = String(data.registrationId || '').trim();
      return jsonOutput(findByRegistrationId(registrationId));
    }

    if (action === 'checkin') {
      const registrationId = String(data.registrationId || '').trim();
      const staffName = String(data.staffName || '').trim();
      return jsonOutput(markCheckInByRegistrationId(registrationId, staffName));
    }

    return jsonOutput({
      success: false,
      message: 'Unknown action'
    });
  } catch (err) {
    return jsonOutput({
      success: false,
      message: err.message || String(err)
    });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/*************************************************
 * SETUP
 *************************************************/
function setupSystem() {
  ensureBackofficeColumns();
  installFormSubmitTrigger_();
  Logger.log('Setup และ Trigger เรียบร้อย');
}

function createOnFormSubmitTrigger() {
  installFormSubmitTrigger_();
  Logger.log('สร้าง Trigger onFormSubmit เรียบร้อย');
}

function installFormSubmitTrigger_() {
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('onFormSubmit')
    .forSpreadsheet(spreadsheet)
    .onFormSubmit()
    .create();
}

function ensureBackofficeColumns() {
  const sheet = getResponseSheet();
  const headers = getHeaders_(sheet);
  const requiredColumns = getRequiredBackofficeColumns_();

  const missing = requiredColumns.filter(h => !headers.includes(h));

  if (missing.length > 0) {
    const startCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
  }

  Logger.log('เพิ่มคอลัมน์หลังบ้านเรียบร้อย');
}

/*************************************************
 * TRIGGER
 *************************************************/
function onFormSubmit(e) {
  try {
    const row = e && e.range ? e.range.getRow() : getLatestDataRow_();
    processRow_(row, true);
    Logger.log('Processed row: ' + row);
  } catch (err) {
    Logger.log('onFormSubmit error: ' + err.message);
    throw err;
  }
}

/*************************************************
 * MAIN PROCESS
 *************************************************/
function processRow_(row, sendEmail) {
  ensureBackofficeColumns();
  const sheet = getResponseSheet();
  const targetRow = resolveTargetRow_(sheet, row);
  const headers = getHeaders_(sheet);
  const colMap = getColumnMap_(headers);
  const rowValues = getRowValues_(sheet, targetRow);

  let registrationId = getValueByAliases_(rowValues, colMap, ['Registration ID']);
  if (!registrationId) {
    registrationId = generateRegistrationId_(targetRow);
    setValue_(sheet, targetRow, colMap, 'Registration ID', registrationId);
  }

  const qrCodeUrl = buildQrCodeUrl_(registrationId);
  setValue_(sheet, targetRow, colMap, 'QR Code URL', qrCodeUrl);

  if (!getValueByAliases_(rowValues, colMap, ['Email Status'])) {
    setValue_(sheet, targetRow, colMap, 'Email Status', 'Pending');
  }

  if (!getValueByAliases_(rowValues, colMap, ['Check-in Status'])) {
    setValue_(sheet, targetRow, colMap, 'Check-in Status', 'Not Checked-in');
  }

  if (!getValueByAliases_(rowValues, colMap, ['Created At'])) {
    setValue_(sheet, targetRow, colMap, 'Created At', new Date());
  }

  if (!getValueByAliases_(rowValues, colMap, ['Attendee Type'])) {
    setValue_(sheet, targetRow, colMap, 'Attendee Type', 'Registered');
  }

  const email = getPrimaryEmail_(rowValues, colMap);
  if (sendEmail && email) {
    sendConfirmationEmailByRow(targetRow);
  } else if (sendEmail && !email) {
    setValue_(sheet, targetRow, colMap, 'Email Status', 'Error');
    setValue_(sheet, targetRow, colMap, 'Notes', 'ไม่พบอีเมลผู้ลงทะเบียน');
  }
}

/*************************************************
 * EMAIL
 *************************************************/
function sendConfirmationEmailByRow(row) {
  const sheet = getResponseSheet();
  const targetRow = resolveTargetRow_(sheet, row);
  const headers = getHeaders_(sheet);
  const colMap = getColumnMap_(headers);
  const rowValues = getRowValues_(sheet, targetRow);

  const fullName = getValueByAliases_(rowValues, colMap, ['ชื่อ-นามสกุล']);
  const email = getPrimaryEmail_(rowValues, colMap);
  const registrationId = getValueByAliases_(rowValues, colMap, ['Registration ID']);
  const qrCodeBlob = getQrCodeBlob_(registrationId);
  const calendarPackage = buildCalendarPackage_(registrationId, fullName);
  const checkinUrl = getConfiguredCheckinUrl_();

  if (!email || !registrationId) {
    throw new Error('ไม่พบอีเมลหรือ Registration ID');
  }

  const subject = `${CONFIG.subjectPrefix} ${CONFIG.eventName}`;
  const hasCheckinLink = Boolean(checkinUrl);
  const calendarPlainText = calendarPackage.googleCalendarUrl
    ? `เพิ่มลงปฏิทิน: ${calendarPackage.googleCalendarUrl}\n\n`
    : '';
  const calendarHtml = calendarPackage.googleCalendarUrl
    ? `
      <div style="text-align:center;margin:20px 0 28px;">
        <a href="${escapeHtml_(calendarPackage.googleCalendarUrl)}" style="display:inline-block;background:#0f3d91;color:#ffffff;text-decoration:none;padding:12px 20px;border-radius:10px;font-weight:700;">
          เพิ่มลงปฏิทิน
        </a>
      </div>
    `
    : '';
  const checkinPlainText = hasCheckinLink
    ? `ลิงก์ระบบเช็กอินวันงาน: ${checkinUrl}\n\n`
    : '';
  const checkinHtml = hasCheckinLink
    ? `
      <p style="margin:0 0 18px;text-align:center;font-size:14px;color:#555;">
        หากต้องการเปิดระบบเช็กอินโดยตรง:
        <a href="${escapeHtml_(checkinUrl)}">${escapeHtml_(checkinUrl)}</a>
      </p>
    `
    : '';
  const contactPlainText = buildContactPlainText_();
  const contactHtml = buildContactHtml_();

  const plainBody =
    `เรียน ${fullName || 'ผู้ลงทะเบียน'}\n\n` +
    `ยืนยันการลงทะเบียนเข้าร่วมงาน ${CONFIG.eventName}\n\n` +
    `Registration ID: ${registrationId}\n` +
    `วันจัดงาน: ${CONFIG.eventDateText}\n` +
    `เวลา: ${CONFIG.eventTimeText}\n` +
    `สถานที่: ${CONFIG.eventLocationText}\n\n` +
    `ระบบได้แนบ QR Code สำหรับเช็กอินไว้ในอีเมลนี้แล้ว\n` +
    `กรุณาแสดง QR Code นี้ที่จุดลงทะเบียนหน้างาน\n\n` +
    checkinPlainText +
    calendarPlainText +
    `หากมีข้อสงสัย กรุณาติดต่อ\n${contactPlainText}`;

  const htmlBody = `
    <div style="font-family:Arial,'Noto Sans Thai',sans-serif;color:#222;line-height:1.6;max-width:640px;margin:0 auto;padding:24px;">
      <h2 style="margin:0 0 16px;color:#0f3d91;">ยืนยันการลงทะเบียนเข้าร่วมงาน</h2>

      <p>เรียน ${escapeHtml_(fullName || 'ผู้ลงทะเบียน')}</p>

      <p>การลงทะเบียนสำหรับงาน <strong>${escapeHtml_(CONFIG.eventName)}</strong> เสร็จสมบูรณ์แล้ว</p>

      <div style="background:#f7f9fc;border:1px solid #e3e8f2;border-radius:12px;padding:16px;margin:16px 0;">
        <div><strong>Registration ID:</strong> ${escapeHtml_(registrationId || '-')}</div>
        <div><strong>วันจัดงาน:</strong> ${escapeHtml_(CONFIG.eventDateText)}</div>
        <div><strong>เวลา:</strong> ${escapeHtml_(CONFIG.eventTimeText)}</div>
        <div><strong>สถานที่:</strong> ${escapeHtml_(CONFIG.eventLocationText)}</div>
      </div>

      <div style="text-align:center;margin:24px 0;">
        <p style="margin-bottom:12px;"><strong>QR Code สำหรับเช็กอินหน้างาน</strong></p>
        <img src="cid:qrCodeImage" alt="QR Code" width="${CONFIG.qrSize}" height="${CONFIG.qrSize}" style="display:block;margin:0 auto;" />
        <p style="margin-top:12px;font-size:14px;color:#555;">กรุณาแสดง QR Code นี้ที่จุดลงทะเบียน</p>
      </div>

      ${checkinHtml}
      ${calendarHtml}

      <div style="margin-top:20px;font-size:14px;color:#555;">
        <div style="margin-bottom:6px;">หากมีข้อสงสัย กรุณาติดต่อ</div>
        ${contactHtml}
      </div>
    </div>
  `;

  try {
    const attachments = [qrCodeBlob.copyBlob()];
    if (calendarPackage.icsBlob) {
      attachments.push(calendarPackage.icsBlob);
    }

    GmailApp.sendEmail(email, subject, plainBody, {
      htmlBody: htmlBody,
      name: CONFIG.contactName,
      inlineImages: {
        qrCodeImage: qrCodeBlob
      },
      attachments: attachments
    });

    setValue_(sheet, targetRow, colMap, 'Email Status', 'Sent');
    setValue_(sheet, targetRow, colMap, 'Email Sent At', new Date());
    setValue_(sheet, targetRow, colMap, 'Notes', '');
  } catch (err) {
    setValue_(sheet, targetRow, colMap, 'Email Status', 'Error');
    setValue_(sheet, targetRow, colMap, 'Notes', 'Email Error: ' + (err.message || err));
    throw err;
  }
}

/*************************************************
 * BULK PROCESS
 *************************************************/
function processExistingRowsWithoutEmail() {
  const sheet = getResponseSheet();
  const lastRow = sheet.getLastRow();

  for (let row = 2; row <= lastRow; row++) {
    const values = getRowValues_(sheet, row);
    const hasData = values.some(v => String(v || '').trim() !== '');
    if (!hasData) continue;
    processRow_(row, false);
    Utilities.sleep(100);
  }

  Logger.log('ประมวลผลข้อมูลเก่าเรียบร้อย');
}

function resendPendingEmails() {
  const sheet = getResponseSheet();
  const headers = getHeaders_(sheet);
  const colMap = getColumnMap_(headers);
  const lastRow = sheet.getLastRow();

  let sent = 0;
  let skipped = 0;
  let errors = 0;

  for (let row = 2; row <= lastRow; row++) {
    const rowValues = getRowValues_(sheet, row);
    const email = getPrimaryEmail_(rowValues, colMap);
    const emailStatus = getValueByAliases_(rowValues, colMap, ['Email Status']);
    const registrationId = getValueByAliases_(rowValues, colMap, ['Registration ID']);

    if (!email) {
      setValue_(sheet, row, colMap, 'Email Status', 'Error');
      setValue_(sheet, row, colMap, 'Notes', 'ไม่พบอีเมลผู้ลงทะเบียน');
      skipped++;
      continue;
    }

    if (!registrationId) {
      processRow_(row, false);
    }

    if (!emailStatus || emailStatus === 'Pending' || emailStatus === 'Error') {
      try {
        sendConfirmationEmailByRow(row);
        sent++;
        Utilities.sleep(500);
      } catch (err) {
        errors++;
        setValue_(sheet, row, colMap, 'Email Status', 'Error');
        setValue_(sheet, row, colMap, 'Notes', 'Email Error: ' + err.message);
      }
    } else {
      skipped++;
    }
  }

  Logger.log(`ส่งสำเร็จ ${sent} | ข้าม ${skipped} | Error ${errors}`);
}

function sendTestEmailToLatestRow() {
  const sheet = getResponseSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('ยังไม่มีข้อมูลสำหรับทดสอบส่งเมล');
  }

  processRow_(lastRow, false);
  sendConfirmationEmailByRow(lastRow);
  Logger.log('ส่งเมลทดสอบให้แถวล่าสุดเรียบร้อย: ' + lastRow);
}

/*************************************************
 * CHECK-IN
 *************************************************/
function findByRegistrationId(registrationId) {
  const target = normalizeText_(registrationId);
  if (!target) {
    return { success: false, message: 'กรุณากรอก Registration ID' };
  }

  const sheet = getResponseSheet();
  const data = sheet.getDataRange().getValues();
  const headers = getHeaders_(sheet);
  const colMap = getColumnMap_(headers);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const regId = normalizeText_(getValueByAliases_(row, colMap, ['Registration ID']));

    if (regId === target) {
      return {
        success: true,
        fullName: getValueByAliases_(row, colMap, ['ชื่อ-นามสกุล']),
        registrationId: getValueByAliases_(row, colMap, ['Registration ID']),
        organization: getOrganization_(row, colMap),
        status: getValueByAliases_(row, colMap, ['Check-in Status']) || 'Not Checked-in',
        checkedInAt: getValueByAliases_(row, colMap, ['Check-in Time']) || ''
      };
    }
  }

  return { success: false, message: 'ไม่พบข้อมูลผู้ลงทะเบียน' };
}

function markCheckInByRegistrationId(registrationId, staffName) {
  const target = normalizeText_(registrationId);
  if (!target) {
    return { success: false, message: 'กรุณากรอก Registration ID' };
  }

  return withScriptLock_(function () {
    const sheet = getResponseSheet();
    const data = sheet.getDataRange().getValues();
    const headers = getHeaders_(sheet);
    const colMap = getColumnMap_(headers);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const regId = normalizeText_(getValueByAliases_(row, colMap, ['Registration ID']));

      if (regId === target) {
        const fullName = getValueByAliases_(row, colMap, ['ชื่อ-นามสกุล']);
        const currentStatus = getValueByAliases_(row, colMap, ['Check-in Status']);
        const checkedInAt = getValueByAliases_(row, colMap, ['Check-in Time']);
        const organization = getOrganization_(row, colMap);

        if (currentStatus === 'Checked-in') {
          return {
            success: false,
            message: 'ผู้เข้าร่วมท่านนี้เช็กอินแล้ว',
            fullName,
            registrationId: getValueByAliases_(row, colMap, ['Registration ID']),
            organization,
            checkedInAt
          };
        }

        const rowNumber = i + 1;
        const now = new Date();

        setValue_(sheet, rowNumber, colMap, 'Check-in Status', 'Checked-in');
        setValue_(sheet, rowNumber, colMap, 'Check-in Time', now);
        setValue_(sheet, rowNumber, colMap, 'Checked-in By', staffName || CONFIG.defaultCheckedInBy);

        return {
          success: true,
          message: 'เช็กอินสำเร็จ',
          fullName,
          registrationId: getValueByAliases_(row, colMap, ['Registration ID']),
          organization,
          checkedInAt: now
        };
      }
    }

    return { success: false, message: 'ไม่พบข้อมูลผู้ลงทะเบียน' };
  });
}

function withScriptLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.lockTimeoutMs);

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

/*************************************************
 * HELPERS
 *************************************************/
function getResponseSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  const sheet = ss.getSheetByName(CONFIG.sheetName);
  if (!sheet) throw new Error('ไม่พบชีตชื่อ: ' + CONFIG.sheetName);
  return sheet;
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
}

function getRowValues_(sheet, row) {
  return sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function resolveTargetRow_(sheet, row) {
  const numericRow = Number(row);

  if (numericRow >= 2) {
    return numericRow;
  }

  return getLatestDataRow_(sheet);
}

function getLatestDataRow_(sheet) {
  const targetSheet = sheet || getResponseSheet();
  const lastRow = targetSheet.getLastRow();

  if (lastRow < 2) {
    throw new Error('ยังไม่มีข้อมูลผู้ลงทะเบียนสำหรับดำเนินการ');
  }

  return lastRow;
}

function getColumnMap_(headers) {
  const map = {};
  headers.forEach((header, index) => {
    map[String(header).trim()] = index + 1;
  });
  return map;
}

function getValueByAliases_(rowValues, colMap, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const colIndex = colMap[aliases[i]];
    if (colIndex) {
      return rowValues[colIndex - 1];
    }
  }
  return '';
}

function getPrimaryEmail_(rowValues, colMap) {
  const directEmail = (
    getValueByAliases_(rowValues, colMap, ['อีเมล']) ||
    getValueByAliases_(rowValues, colMap, ['ที่อยู่อีเมล']) ||
    getValueByAliases_(rowValues, colMap, ['อีเมลสำหรับติดต่อ']) ||
    getValueByAliases_(rowValues, colMap, ['Email Address']) ||
    getValueByAliases_(rowValues, colMap, ['Email']) ||
    ''
  );

  if (directEmail) {
    return directEmail;
  }

  const fallbackHeader = Object.keys(colMap).find(header => {
    const normalizedHeader = normalizeText_(header);
    return (
      normalizedHeader.includes('email') ||
      normalizedHeader.includes('e-mail') ||
      normalizedHeader.includes('อีเมล')
    );
  });

  if (!fallbackHeader) {
    return '';
  }

  return rowValues[colMap[fallbackHeader] - 1] || '';
}

function getOrganization_(rowValues, colMap) {
  return getValueByAliases_(rowValues, colMap, [
    'องค์กร / บริษัท / ชื่อธุรกิจ / สถาบัน',
    'องค์กร / บริษัท / ชื่อธุรกิจ',
    'ชื่อองค์กร',
    'Organization'
  ]);
}

function buildStatusPayload_() {
  return {
    ok: true,
    apiName: CONFIG.apiName,
    apiVersion: CONFIG.apiVersion,
    eventName: CONFIG.eventName,
    eventDateText: CONFIG.eventDateText,
    eventTimeText: CONFIG.eventTimeText,
    eventLocationText: CONFIG.eventLocationText,
    defaultStaffName: CONFIG.defaultCheckedInBy,
    scannerBaseUrl: getScannerBaseUrl_(),
    scannerUrl: getConfiguredCheckinUrl_(),
    frontendConfigured: hasConfiguredFrontendBaseUrl_() || hasConfiguredCheckinUrl_(),
    message: 'API ready',
    actions: ['status', 'checkin', 'lookup']
  };
}

function hasConfiguredCheckinUrl_() {
  const value = normalizeConfigUrl_(CONFIG.checkinWebAppUrl);
  return Boolean(value) && value !== 'PUT_YOUR_CHECKIN_WEBAPP_URL_HERE';
}

function hasConfiguredFrontendBaseUrl_() {
  const value = getScannerBaseUrl_();
  return Boolean(value) && value !== 'PUT_YOUR_FRONTEND_BASE_URL_HERE';
}

function getConfiguredCheckinUrl_() {
  if (hasConfiguredCheckinUrl_()) {
    return normalizeConfigUrl_(CONFIG.checkinWebAppUrl);
  }

  return buildPublicScannerUrl_();
}

function buildPublicScannerUrl_(staffName) {
  const baseUrl = getScannerBaseUrl_();
  const apiUrl = getCurrentApiWebAppUrl_();

  if (!baseUrl || !apiUrl) {
    return '';
  }

  const params = { apiUrl: apiUrl };

  if (staffName) {
    params.staffName = staffName;
  }

  return appendQueryParams_(baseUrl, params);
}

function getScannerBaseUrl_() {
  return normalizeConfigUrl_(CONFIG.frontendBaseUrl);
}

function getCurrentApiWebAppUrl_() {
  return normalizeConfigUrl_(ScriptApp.getService().getUrl());
}

function normalizeConfigUrl_(value) {
  return String(value || '').trim();
}

function appendQueryParams_(baseUrl, params) {
  const hashIndex = baseUrl.indexOf('#');
  const hashPart = hashIndex >= 0 ? baseUrl.slice(hashIndex) : '';
  const rawBaseUrl = hashIndex >= 0 ? baseUrl.slice(0, hashIndex) : baseUrl;
  const pairs = [];

  Object.keys(params).forEach(key => {
    const value = normalizeConfigUrl_(params[key]);
    if (value) {
      pairs.push(`${encodeURIComponent(key)}=${encodeURIComponent(value)}`);
    }
  });

  if (!pairs.length) {
    return baseUrl;
  }

  const separator = rawBaseUrl.indexOf('?') >= 0 ? '&' : '?';
  return `${rawBaseUrl}${separator}${pairs.join('&')}${hashPart}`;
}

function buildContactPlainText_() {
  const lines = [CONFIG.contactName || 'ทีมงานผู้จัดงาน'];

  if (CONFIG.contactEmail) {
    lines.push(`อีเมล: ${CONFIG.contactEmail}`);
  }

  if (CONFIG.contactPhone) {
    lines.push(`โทร: ${CONFIG.contactPhone}`);
  }

  return lines.join('\n');
}

function buildContactHtml_() {
  const lines = [];

  if (CONFIG.contactName) {
    lines.push(`<div>${escapeHtml_(CONFIG.contactName)}</div>`);
  }

  if (CONFIG.contactEmail) {
    lines.push(`<div>อีเมล: ${escapeHtml_(CONFIG.contactEmail)}</div>`);
  }

  if (CONFIG.contactPhone) {
    lines.push(`<div>โทร: ${escapeHtml_(CONFIG.contactPhone)}</div>`);
  }

  return lines.join('');
}

function setValue_(sheet, row, colMap, headerName, value) {
  let colIndex = colMap[headerName];

  if (!colIndex && isBackofficeColumn_(headerName)) {
    colIndex = ensureColumnExists_(sheet, colMap, headerName);
  }

  if (!colIndex) {
    throw new Error('ไม่พบคอลัมน์: ' + headerName);
  }

  sheet.getRange(row, colIndex).setValue(value);
}

function getRequiredBackofficeColumns_() {
  return [
    'Registration ID',
    'Email Status',
    'QR Code URL',
    'Check-in Status',
    'Check-in Time',
    'Checked-in By',
    'Notes',
    'Created At',
    'Email Sent At',
    'Attendee Type'
  ];
}

function isBackofficeColumn_(headerName) {
  return getRequiredBackofficeColumns_().includes(headerName);
}

function ensureColumnExists_(sheet, colMap, headerName) {
  const freshHeaders = getHeaders_(sheet);
  const freshColMap = getColumnMap_(freshHeaders);

  if (freshColMap[headerName]) {
    Object.assign(colMap, freshColMap);
    return freshColMap[headerName];
  }

  const newColumnIndex = sheet.getLastColumn() + 1;
  sheet.getRange(1, newColumnIndex).setValue(headerName);
  colMap[headerName] = newColumnIndex;
  return newColumnIndex;
}

function generateRegistrationId_(row) {
  return CONFIG.registrationPrefix + String(row - 1).padStart(4, '0');
}

function buildQrCodeUrl_(registrationId) {
  return `https://quickchart.io/qr?text=${encodeURIComponent(registrationId)}&size=${CONFIG.qrSize}`;
}

function buildCalendarPackage_(registrationId, fullName) {
  const start = parseIsoDate_(CONFIG.eventStartIso);
  const end = parseIsoDate_(CONFIG.eventEndIso);

  if (!start || !end || end.getTime() <= start.getTime()) {
    return {
      googleCalendarUrl: '',
      icsBlob: null
    };
  }

  const title = CONFIG.eventName;
  const description =
    `${CONFIG.eventCalendarDescription}\nRegistration ID: ${registrationId || '-'}\nชื่อผู้ลงทะเบียน: ${fullName || '-'}`;
  const googleCalendarUrl =
    'https://calendar.google.com/calendar/render?action=TEMPLATE' +
    '&text=' + encodeURIComponent(title) +
    '&dates=' + encodeURIComponent(formatGoogleCalendarDate_(start) + '/' + formatGoogleCalendarDate_(end)) +
    '&location=' + encodeURIComponent(CONFIG.eventLocationText || '') +
    '&details=' + encodeURIComponent(description);

  return {
    googleCalendarUrl: googleCalendarUrl,
    icsBlob: buildIcsBlob_(title, start, end, description, CONFIG.eventLocationText)
  };
}

function getQrCodeBlob_(registrationId) {
  if (!registrationId) {
    throw new Error('ไม่พบ Registration ID สำหรับสร้าง QR Code');
  }

  const qrCodeUrl = buildQrCodeUrl_(registrationId);
  const response = UrlFetchApp.fetch(qrCodeUrl, {
    followRedirects: true,
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();

  if (responseCode < 200 || responseCode >= 300) {
    throw new Error('สร้าง QR Code ไม่สำเร็จ (HTTP ' + responseCode + ')');
  }

  return response.getBlob().setName(`qr-${registrationId}.png`);
}

function buildIcsBlob_(title, start, end, description, location) {
  const lines = [
    'BEGIN:VCALENDAR',
    'VERSION:2.0',
    'PRODID:-//SMEs QR Scanner//TH',
    'CALSCALE:GREGORIAN',
    'METHOD:PUBLISH',
    'BEGIN:VEVENT',
    'UID:' + Utilities.getUuid(),
    'DTSTAMP:' + formatGoogleCalendarDate_(new Date()),
    'DTSTART:' + formatGoogleCalendarDate_(start),
    'DTEND:' + formatGoogleCalendarDate_(end),
    'SUMMARY:' + escapeIcsText_(title),
    'DESCRIPTION:' + escapeIcsText_(description),
    'LOCATION:' + escapeIcsText_(location || ''),
    'END:VEVENT',
    'END:VCALENDAR'
  ];

  return Utilities.newBlob(lines.join('\r\n'), 'text/calendar', 'event.ics');
}

function parseIsoDate_(value) {
  if (!value) {
    return null;
  }

  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatGoogleCalendarDate_(date) {
  return Utilities.formatDate(date, 'UTC', "yyyyMMdd'T'HHmmss'Z'");
}

function escapeIcsText_(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/\r?\n/g, '\\n')
    .replace(/,/g, '\\,')
    .replace(/;/g, '\\;');
}

function normalizeText_(value) {
  return String(value || '').trim().toLowerCase();
}

function escapeHtml_(text) {
  if (text == null) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

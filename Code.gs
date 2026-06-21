const SPREADSHEET_ID = '1IgfzKJT4_ohsvP6FZYaj8K_2eMgHBSKrybLtUevM67o';
const SHEET_NAME = 'QR_DB';
const PHOTO_FOLDER_NAME = 'SAEGIM_CHILD_PHOTOS';

// ─── 진입점 ───────────────────────────────────────────────

function doGet(e) {
  const action = e?.parameter?.action;

  if (action) {
    return handleApi_(e.parameter);
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.code = e?.parameter?.code || '';
  template.lang = e?.parameter?.lang || 'ko';
  return template.evaluate()
    .setTitle('새김 QR')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  const params = e?.postData?.contents
    ? JSON.parse(e.postData.contents)
    : {};
  return handleApi_(params);
}

function handleApi_(params) {
  let result;

  try {
    switch (params.action) {
      case 'getQrData':
        result = getQrData(params.code);
        break;
      case 'verifyPassword':
        result = verifyPassword(params.code, params.password);
        break;
      case 'changeStatus':
        result = changeStatus(params.code, params.password, params.status);
        break;
      case 'registerSoldQr':
        result = registerSoldQr(
          typeof params.form === 'string' ? JSON.parse(params.form) : params
        );
        break;
      case 'updateCustomerData':
        result = updateCustomerData(
          typeof params.form === 'string' ? JSON.parse(params.form) : params
        );
        break;
      case 'changeLostMode':
        result = changeLostMode(params.code, params.password, params.lost_mode);
        break;
      case 'sendOwnerEmailAlert':
        result = sendOwnerEmailAlert(params);
        break;
      case 'savePushToken':
        result = savePushToken(params);
        break;
      case 'sendOwnerPushAlert':
        result = sendOwnerPushAlert(params);
        break;
      default:
        result = { success: false, message: '알 수 없는 action: ' + params.action };
    }
  } catch (err) {
    result = { success: false, message: String(err) };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────

function getSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('QR_DB 시트를 찾을 수 없습니다.');
  return sheet;
}

function getHeaderMap_() {
  const sheet = getSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => {
    if (h) map[String(h).trim()] = i + 1;
  });
  return map;
}

function findQrRow_(code) {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]).trim() === String(code).trim()) {
      return i + 1;
    }
  }
  return -1;
}

function getRowObject_(row) {
  const sheet = getSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((key, i) => {
    if (key) obj[String(key).trim()] = values[i];
  });
  return obj;
}

// ─── 제품별 시트 동기화 헬퍼 ──────────────────────────────

function getHeaderMapFromSheet_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) {
    if (h) map[String(h).trim()] = i + 1;
  });
  return map;
}

function getProductTypeFromCode_(code) {
  return String(code || '').trim().split('-')[0].toUpperCase();
}

function getProductSheetByCode_(code) {
  const productType = getProductTypeFromCode_(code);
  if (!productType) return null;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return ss.getSheetByName(productType);
}

function findQrRowInSheet_(sheet, code) {
  if (!sheet) return -1;
  const headers = getHeaderMapFromSheet_(sheet);
  const codeCol = headers.code || 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, codeCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === String(code || '').trim()) {
      return i + 2;
    }
  }
  return -1;
}

function syncQrDbRowToProductSheet_(code) {
  code = String(code || '').trim();
  if (!code) return false;

  const masterRow = findQrRow_(code);
  if (masterRow === -1) return false;

  const productSheet = getProductSheetByCode_(code);
  if (!productSheet) {
    console.warn('[제품별 시트 동기화 건너뜀] 시트 없음: ' + code);
    return false;
  }

  const masterData = getRowObject_(masterRow);
  const productHeaders = getHeaderMapFromSheet_(productSheet);

  let productRow = findQrRowInSheet_(productSheet, code);
  if (productRow === -1) {
    productRow = productSheet.getLastRow() + 1;
  }

  const lastCol = productSheet.getLastColumn();
  let rowValues = new Array(lastCol).fill('');

  if (productRow <= productSheet.getLastRow()) {
    rowValues = productSheet.getRange(productRow, 1, 1, lastCol).getValues()[0];
  }

  Object.keys(masterData).forEach(function(key) {
    const col = productHeaders[key];
    if (col) rowValues[col - 1] = masterData[key];
  });

  productSheet.getRange(productRow, 1, 1, lastCol).setValues([rowValues]);
  return true;
}

function updateProductSheetField_(code, key, value) {
  code = String(code || '').trim();
  if (!code || !key) return false;

  const productSheet = getProductSheetByCode_(code);
  if (!productSheet) return false;

  const row = findQrRowInSheet_(productSheet, code);
  if (row === -1) return false;

  const headers = getHeaderMapFromSheet_(productSheet);
  const col = headers[key];
  if (!col) return false;

  productSheet.getRange(row, col).setValue(value);
  return true;
}

function getQrData(code) {
  try {
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };

    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const data = getRowObject_(row);

    if (data.status === '사용중' || data.status === '분실') {
      data.scan_count = increaseScanCount_(row);
    }

    return {
      success: true,
      code: data.code || '',
      product_code: String(data.code || '').split('-')[0],
      product: data.product || '',
      owner: data.owner || '',
      status: data.status || '미등록',

      child_photo: data.child_photo || '',
      child_name: data.child_name || '',
      guardian_name: data.guardian_name || '',
      guardian_phone: data.guardian_phone || '',
      child_allergy: data.child_allergy || '',
      blood_type: data.blood_type || '',
      child_note: data.child_note || '',
      child_message: data.child_message || '',

      bmt_photo1: data.bmt_photo1 || '',
      bmt_photo2: data.bmt_photo2 || '',
      bmt_photo3: data.bmt_photo3 || '',
      bmt_photo4: data.bmt_photo4 || '',
      bmt_photo5: data.bmt_photo5 || '',
      bmt_travel_memo: data.bmt_travel_memo || '',
      bmt_places: data.bmt_places || '',
      bmt_voice: data.bmt_voice || '',
      bmt_visit_date: data.bmt_visit_date || '',
      bmt_photo_fit: data.bmt_photo_fit || 'single',
      bmt_photo_position: data.bmt_photo_position || 'center',
      // ── BMT 분실물 모드 v2 ──
      lost_mode: data.lost_mode || '',
      lost_contact_type: data.lost_contact_type || '',
      lost_contact_label: data.lost_contact_label || '',
      lost_contact_url: data.lost_contact_url || '',
      owner_email: data.owner_email || '',
      push_token: data.push_token || '',
      lost_message_ko: data.lost_message_ko || '',
      lost_message_en: data.lost_message_en || '',
      lost_message_ja: data.lost_message_ja || '',
      lost_message_zh: data.lost_message_zh || '',

      scan_count: data.scan_count || 0,
      links: buildLinks_(data)
    };

  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function buildLinks_(data) {
  const links = [];
  for (let i = 1; i <= 5; i++) {
    const type = data['link' + i + '_type'] || '';
    const label = data['link' + i + '_label'] || '';
    const url = data['link' + i + '_url'] || '';
    if (!url) continue;
    links.push({
      type: String(type).trim(),
      label: String(label || type || '연결하기').trim(),
      url: normalizeUrl_(type, url)
    });
  }
  return links;
}

function normalizeUrl_(type, url) {
  const value = String(url || '').trim();
  const t = String(type || '').trim();
  if (!value) return '';
  if (t === '전화' || t === '전화하기') {
    return value.startsWith('tel:')
      ? value
      : 'tel:' + value.replace(/[^0-9+]/g, '');
  }
  if (t === '이메일' || t === '메일') {
    return value.startsWith('mailto:')
      ? value
      : 'mailto:' + value;
  }
  return value;
}

function registerSoldQr(form) {
  try {
    const sheet = getSheet_();
    const headers = getHeaderMap_();
    const row = findQrRow_(form.code);

    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const data = getRowObject_(row);
    if (data.status !== '판매완료') {
      return { success: false, message: '등록 가능한 상태가 아닙니다.' };
    }
    if (!String(form.password || '').trim()) {
      return { success: false, message: '수정 비밀번호를 입력해주세요.' };
    }

    saveCustomerFields_(sheet, headers, row, form);
    saveLinks_(sheet, headers, row, form.links);
    setByHeader_(sheet, headers, row, 'password', form.password || '');
    setByHeader_(sheet, headers, row, 'status', '사용중');
    setByHeader_(sheet, headers, row, 'registered_at', nowText_());
    setByHeader_(sheet, headers, row, 'updated_at', nowText_());

    syncQrDbRowToProductSheet_(form.code);

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function verifyPassword(code, password) {
  try {
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const data = getRowObject_(row);
    const input = String(password || '').trim();
    const customerPw = String(data.password || '').trim();
    const adminPw = String(data.admin_password || '').trim();

    if (!input) return { success: false, message: '비밀번호를 입력해주세요.' };
    if (input === customerPw || input === adminPw) return { success: true };
    return { success: false, message: '비밀번호가 일치하지 않습니다.' };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function updateCustomerData(form) {
  try {
    const sheet = getSheet_();
    const headers = getHeaderMap_();
    const row = findQrRow_(form.code);

    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const pwCheck = verifyPassword(form.code, form.password);
    if (!pwCheck.success) return pwCheck;

    saveCustomerFields_(sheet, headers, row, form);
    saveLinks_(sheet, headers, row, form.links);
    setByHeader_(sheet, headers, row, 'updated_at', nowText_());

    syncQrDbRowToProductSheet_(form.code);

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function changeStatus(code, password, newStatus) {
  try {
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const pwCheck = verifyPassword(code, password);
    if (!pwCheck.success) return pwCheck;

    const allowed = ['사용중', '분실'];
    if (!allowed.includes(newStatus)) {
      return { success: false, message: '변경할 수 없는 상태입니다.' };
    }

    const sheet = getSheet_();
    const headers = getHeaderMap_();
    setByHeader_(sheet, headers, row, 'status', newStatus);
    setByHeader_(sheet, headers, row, 'updated_at', nowText_());

    syncQrDbRowToProductSheet_(code);

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function saveCustomerFields_(sheet, headers, row, form) {
  setByHeader_(sheet, headers, row, 'owner', form.owner || '');

  setByHeader_(sheet, headers, row, 'child_name', form.child_name || '');
  setByHeader_(sheet, headers, row, 'guardian_name', form.guardian_name || '');
  setByHeader_(sheet, headers, row, 'guardian_phone', form.guardian_phone || '');
  setByHeader_(sheet, headers, row, 'child_allergy', form.child_allergy || '');
  setByHeader_(sheet, headers, row, 'blood_type', form.blood_type || '');
  setByHeader_(sheet, headers, row, 'child_note', form.child_note || '');
  setByHeader_(sheet, headers, row, 'child_message', form.child_message || '');

  setByHeader_(sheet, headers, row, 'bmt_photo1', form.bmt_photo1 || '');
  setByHeader_(sheet, headers, row, 'bmt_photo2', form.bmt_photo2 || '');
  setByHeader_(sheet, headers, row, 'bmt_photo3', form.bmt_photo3 || '');
  setByHeader_(sheet, headers, row, 'bmt_photo4', form.bmt_photo4 || '');
  setByHeader_(sheet, headers, row, 'bmt_photo5', form.bmt_photo5 || '');
  setByHeader_(sheet, headers, row, 'bmt_travel_memo', form.bmt_travel_memo || '');
  setByHeader_(sheet, headers, row, 'bmt_places', form.bmt_places || '');
  setByHeader_(sheet, headers, row, 'bmt_voice', form.bmt_voice || '');
  setByHeader_(sheet, headers, row, 'bmt_visit_date', form.bmt_visit_date || '');
  setByHeader_(sheet, headers, row, 'bmt_photo_fit', form.bmt_photo_fit || 'single');
  setByHeader_(sheet, headers, row, 'bmt_photo_position', form.bmt_photo_position || 'center');
  // ── BMT 분실물 모드 v2 ──
  setByHeader_(sheet, headers, row, 'lost_mode', form.lost_mode || '');
  setByHeader_(sheet, headers, row, 'lost_contact_type', form.lost_contact_type || '');
  setByHeader_(sheet, headers, row, 'lost_contact_label', form.lost_contact_label || '');
  setByHeader_(sheet, headers, row, 'lost_contact_url', form.lost_contact_url || '');
  setByHeader_(sheet, headers, row, 'owner_email', form.owner_email || '');
  setByHeader_(sheet, headers, row, 'push_token', form.push_token || '');
  setByHeader_(sheet, headers, row, 'lost_message_ko', form.lost_message_ko || '');
  setByHeader_(sheet, headers, row, 'lost_message_en', form.lost_message_en || '');
  setByHeader_(sheet, headers, row, 'lost_message_ja', form.lost_message_ja || '');
  setByHeader_(sheet, headers, row, 'lost_message_zh', form.lost_message_zh || '');
}

function saveLinks_(sheet, headers, row, links) {
  const parsed = parseLinks_(links);
  for (let i = 1; i <= 5; i++) {
    const link = parsed[i - 1] || {};
    setByHeader_(sheet, headers, row, 'link' + i + '_type', link.type || '');
    setByHeader_(sheet, headers, row, 'link' + i + '_label', link.label || '');
    setByHeader_(sheet, headers, row, 'link' + i + '_url', link.url || '');
  }
}

function parseLinks_(links) {
  let arr = [];
  if (Array.isArray(links)) {
    arr = links;
  } else if (typeof links === 'string' && links.trim()) {
    arr = JSON.parse(links);
  }
  return arr
    .filter(link => link && String(link.url || '').trim())
    .slice(0, 5)
    .map(link => ({
      type: String(link.type || '').trim(),
      label: String(link.label || link.type || '연결하기').trim(),
      url: String(link.url || '').trim()
    }));
}

function saveChildPhoto_(code, dataUrl) {
  const match = String(dataUrl || '').match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error('아이 사진 형식이 올바르지 않습니다.');

  const contentType = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  const ext = contentType.includes('png') ? 'png' : 'jpg';
  const fileName = String(code || 'child') + '_child_photo_' + Date.now() + '.' + ext;

  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const folder = getPhotoFolder_();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w600';
}

function getPhotoFolder_() {
  const folders = DriveApp.getFoldersByName(PHOTO_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PHOTO_FOLDER_NAME);
}

function setByHeader_(sheet, headers, row, key, value) {
  if (!headers[key]) return;
  sheet.getRange(row, headers[key]).setValue(value);
}

function increaseScanCount_(row) {
  const sheet = getSheet_();
  const headers = getHeaderMap_();

  if (!headers.scan_count) return 0;

  const codeCol = headers.code || 1;
  const code = String(sheet.getRange(row, codeCol).getValue() || '').trim();

  const cell = sheet.getRange(row, headers.scan_count);
  const current = Number(cell.getValue()) || 0;
  const next = current + 1;

  cell.setValue(next);

  if (code) {
    updateProductSheetField_(code, 'scan_count', next);
  }

  return next;
}

function nowText_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}

function changeLostMode(code, password, lostMode) {
  try {
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const pwCheck = verifyPassword(code, password);
    if (!pwCheck.success) return pwCheck;

    const sheet = getSheet_();
    const headers = getHeaderMap_();

    setByHeader_(sheet, headers, row, 'lost_mode', lostMode === 'on' ? 'on' : '');
    setByHeader_(sheet, headers, row, 'updated_at', nowText_());

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function sendOwnerEmailAlert(params) {
  try {
    const code = String(params.code || '').trim();
    const nickname = String(params.nickname || '').trim() || '발견자';
    const message = String(params.message || '').trim();

    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };
    if (!message) return { success: false, message: '메시지가 없습니다.' };

    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const data = getRowObject_(row);
    const ownerEmail = String(data.owner_email || '').trim();

    if (!ownerEmail) return { success: false, message: '알림 이메일이 등록되어 있지 않습니다.' };

    const pageUrl = 'https://saegim-memory.web.app/?code=' + encodeURIComponent(code);
    const subject = '[새김] ' + code + ' 새 메시지가 도착했습니다.';
    const body =
      '새김 부산 메모리 태그에 새로운 메시지가 도착했습니다.\n\n' +
      '태그번호: ' + code + '\n' +
      '닉네임: ' + nickname + '\n\n' +
      '메시지:\n' + message + '\n\n' +
      '확인하기:\n' + pageUrl + '\n\n' +
      'SAEGIM · MEMORY QR';

    MailApp.sendEmail({ to: ownerEmail, subject: subject, body: body, name: 'SAEGIM' });

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function savePushToken(params) {
  try {
    const code = String(params.code || '').trim();
    const token = String(params.token || '').trim();
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };
    if (!token) return { success: false, message: '토큰이 없습니다.' };
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };
    const sheet = getSheet_();
    const headers = getHeaderMap_();
    setByHeader_(sheet, headers, row, 'push_token', token);

    syncQrDbRowToProductSheet_(code);

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function sendOwnerPushAlert(params) {
  try {
    const code = String(params.code || '').trim();
    const nickname = String(params.nickname || '').trim() || '발견자';
    const message = String(params.message || '').trim();
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };
    if (!message) return { success: false, message: '메시지가 없습니다.' };
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };
    const data = getRowObject_(row);
    const playerId = String(data.push_token || '').trim();
    if (!playerId) return { success: false, message: '푸시 토큰이 없습니다.' };
    const pageUrl = 'https://saegim-memory.web.app/?code=' + encodeURIComponent(code);
    const payload = JSON.stringify({
      app_id: '4454afde-61d3-4c7a-9f09-84de382a029d',
      include_subscription_ids: [playerId],
      headings: { en: '[새김] ' + code + ' 새 메시지', ko: '[새김] ' + code + ' 새 메시지' },
      contents: { en: nickname + ': ' + message, ko: nickname + ': ' + message },
      url: pageUrl
    });
    const response = UrlFetchApp.fetch('https://api.onesignal.com/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Key os_v2_app_irkk7xtb2nghvhyjqtpdqkqctul2cyqw76buolfthjqff6znmsqemlbos3nmy7pob6ma2bk7zptx76cm655dyaccxtyvxuoyl5sglia' },
      payload: payload,
      muteHttpExceptions: true
    });
    const result = JSON.parse(response.getContentText());
    if (result.errors) return { success: false, message: JSON.stringify(result.errors) };
    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

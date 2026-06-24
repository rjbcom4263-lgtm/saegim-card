const SPREADSHEET_ID = '1IgfzKJT4_ohsvP6FZYaj8K_2eMgHBSKrybLtUevM67o';
const SHEET_NAME = 'QR_DB';

// ─── 진입점 ───────────────────────────────────────────────

function doGet(e) {
  const action = e?.parameter?.action;
  if (action) {
    return handleApi_(e.parameter);
  }
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, message: 'action 파라미터가 필요합니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
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
      case 'resetPasswordByEmail':
        result = resetPasswordByEmail(params.code, params.email);
        break;
      case 'getOwnerEmailHint':
        result = getOwnerEmailHint(params.code);
        break;
      case 'changePassword':
        result = changePassword(params.code, params.password, params.new_password);
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
      case 'recordScan':
        result = recordScan(params.code);
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
  const targetCode = String(code || '').trim().toUpperCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toUpperCase() === targetCode) {
      return i + 1;
    }
  }
  return -1;
}

// 한국 전화번호처럼 앞자리 0이 필요한 필드 목록
const PHONE_FIELDS_ = ['guardian_phone', 'link1_url', 'link2_url', 'link3_url', 'link4_url', 'link5_url'];

function fixPhone_(key, val) {
  if (!PHONE_FIELDS_.includes(key)) return val;
  const s = String(val ?? '');
  // 순수 숫자 9~10자리면 앞에 0 붙이기 (010으로 시작해야 할 번호)
  if (/^\d{9,10}$/.test(s)) return '0' + s;
  return s;
}

function getRowObject_(row) {
  const sheet = getSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((key, i) => {
    if (!key) return;
    const k = String(key).trim();
    obj[k] = fixPhone_(k, values[i]);
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

// ─────────────────────────────────────────────────────────

function getQrData(code) {
  try {
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };

    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const data = getRowObject_(row);

    if (data.status === '사용중' || data.status === '분실') {
      data.scan_count = increaseScanCount_(row);
      syncQrDbRowToProductSheet_(code);
      backupCustomerQrToFirestoreV1_(code);
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

      lost_mode: data.lost_mode || '',
      lost_contact_type: data.lost_contact_type || '',
      lost_contact_label: data.lost_contact_label || '',
      lost_contact_url: data.lost_contact_url || '',
      has_owner_email: !!String(data.owner_email || '').trim(),
      has_push_token: !!String(data.push_token || '').trim(),
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
    backupCustomerQrToFirestoreV1_(form.code);

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
    if (input === customerPw || input === adminPw) return { success: true, owner_email: String(data.owner_email || '').trim() };
    return { success: false, message: '비밀번호가 일치하지 않습니다.' };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function changePassword(code, currentPassword, newPassword) {
  try {
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };

    const pwCheck = verifyPassword(code, currentPassword);
    if (!pwCheck.success) return { success: false, message: '현재 비밀번호가 일치하지 않습니다.' };

    const newPw = String(newPassword || '').trim();
    if (!newPw || newPw.length < 4) return { success: false, message: '새 비밀번호는 4자 이상이어야 합니다.' };

    const sheet = getSheet_();
    const headers = getHeaderMap_();
    setByHeader_(sheet, headers, row, 'password', newPw);
    setByHeader_(sheet, headers, row, 'updated_at', nowText_());

    backupCustomerQrToFirestoreV1_(code);
    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function getOwnerEmailHint(code) {
  try {
    if (!code) return { success: false };
    const row = findQrRow_(code);
    if (row === -1) return { success: false };
    const data = getRowObject_(row);
    const email = String(data.owner_email || '').trim();
    if (!email) return { success: true, has_email: false };
    const p = email.split('@');
    const local = p[0];
    const masked = local.length > 2 ? local[0] + '***' + local[local.length-1] : local[0] + '***';
    return { success: true, has_email: true, email_masked: masked + '@' + p[1] };
  } catch (err) {
    return { success: false };
  }
}

function resetPasswordByEmail(code, email) {
  try {
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };
    const data = getRowObject_(row);
    const registered = String(data.owner_email || '').trim().toLowerCase();
    if (!registered) return { success: false, message: '등록된 이메일이 없습니다.\n정보 수정 화면에서 이메일을 먼저 등록해주세요.' };
    // email이 빈 문자열이면 등록된 이메일로 직접 전송 (마스킹 힌트 표시 후 전송 경우)
    if (email && String(email).trim() && registered !== String(email || '').trim().toLowerCase()) {
      return { success: false, message: '등록된 이메일과 일치하지 않습니다.' };
    }
    const pw = String(data.password || '').trim();
    if (!pw) return { success: false, message: '비밀번호 정보가 없습니다. 관리자에게 문의하세요.' };
    const subject = '[새김] ' + code + ' 비밀번호 안내';
    const body = '새김 안심태그 비밀번호 안내입니다.\n\n태그번호: ' + code + '\n비밀번호: ' + pw + '\n\n보안을 위해 확인 후 비밀번호를 변경해주세요.';
    MailApp.sendEmail({ to: registered, subject: subject, body: body });
    return { success: true };
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
    backupCustomerQrToFirestoreV1_(form.code);

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
    backupCustomerQrToFirestoreV1_(code);

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
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

    syncQrDbRowToProductSheet_(code);
    backupCustomerQrToFirestoreV1_(code);

    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

function saveCustomerFields_(sheet, headers, row, form) {
  setByHeader_(sheet, headers, row, 'owner', form.owner || '');

  // CD 필드
  setByHeader_(sheet, headers, row, 'child_photo', form.child_photo || '');
  setByHeader_(sheet, headers, row, 'child_name', form.child_name || '');
  setByHeader_(sheet, headers, row, 'guardian_name', form.guardian_name || '');
  setPhoneByHeader_(sheet, headers, row, 'guardian_phone', form.guardian_phone || '');
  setByHeader_(sheet, headers, row, 'child_allergy', form.child_allergy || '');
  setByHeader_(sheet, headers, row, 'blood_type', form.blood_type || '');
  setByHeader_(sheet, headers, row, 'child_note', form.child_note || '');
  setByHeader_(sheet, headers, row, 'child_message', form.child_message || '');

  // BMT 필드
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

  // WT 새김 바램태그 필드
  setIfProvided_(sheet, headers, row, 'wt_birth_date', form);
  setIfProvided_(sheet, headers, row, 'wt_theme', form);
  setIfProvided_(sheet, headers, row, 'wt_last_message_id', form);
  setIfProvided_(sheet, headers, row, 'wt_lang', form);

  // 분실 모드 공통
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
    setPhoneByHeader_(sheet, headers, row, 'link' + i + '_url', link.url || '');
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

function setByHeader_(sheet, headers, row, key, value) {
  if (!headers[key]) return;
  sheet.getRange(row, headers[key]).setValue(value);
}

// 전화번호처럼 앞자리 0이 있는 값은 텍스트 형식으로 강제 저장
function setPhoneByHeader_(sheet, headers, row, key, value) {
  if (!headers[key]) return;
  const strVal = String(value || '');
  const cell = sheet.getRange(row, headers[key]);
  // 숫자로만 된 값은 수식으로 강제 텍스트 저장 (="010...")
  if (strVal && /^\d+$/.test(strVal)) {
    cell.setFormula('="' + strVal + '"');
  } else {
    cell.setValue(strVal);
  }
}

function setIfProvided_(sheet, headers, row, key, form) {
  if (!form) return;
  if (Object.prototype.hasOwnProperty.call(form, key)) {
    setByHeader_(sheet, headers, row, key, form[key] || '');
  }
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

// ─── Firestore V1 백업 (고객용) ───────────────────────────────────────────────

const FIRESTORE_PROJECT_ID = 'saegim-memory';
const FIRESTORE_PUBLIC_COLLECTION = 'qr_cards';
const FIRESTORE_PRIVATE_COLLECTION = 'qr_card_private';

const FIRESTORE_PUBLIC_FIELDS = [
  'code', 'product', 'product_type', 'owner', 'status', 'qr_file',
  'link1_type', 'link1_label', 'link1_url',
  'link2_type', 'link2_label', 'link2_url',
  'link3_type', 'link3_label', 'link3_url',
  'link4_type', 'link4_label', 'link4_url',
  'link5_type', 'link5_label', 'link5_url',
  'child_name', 'guardian_name', 'guardian_phone',
  'child_allergy', 'blood_type', 'child_note', 'child_message',
  'scan_count', 'sold_at', 'registered_at', 'updated_at',
  'bmt_photo1', 'bmt_photo2', 'bmt_photo3', 'bmt_photo4', 'bmt_photo5',
  'bmt_travel_memo', 'bmt_places', 'bmt_voice', 'bmt_visit_date',
  'bmt_photo_fit', 'bmt_photo_position',
  'lost_mode', 'lost_contact_type', 'lost_contact_label', 'lost_contact_url',
  'lost_message_ko', 'lost_message_en', 'lost_message_ja', 'lost_message_zh',
  'gt_garden_name', 'gt_growth_point', 'gt_stage', 'gt_slots', 'gt_inventory',
  'wt_birth_date', 'wt_theme', 'wt_last_message_id', 'wt_lang',
  'owner_email',
  'gm_default_area', 'gm_enabled_categories',
  'firestore_backup_at'
];

const FIRESTORE_PRIVATE_FIELDS = [
  'code', 'password', 'admin_password', 'owner_email', 'push_token', 'memo',
  'firestore_backup_at'
];

function getFirestoreToken_() {
  const props = PropertiesService.getScriptProperties();
  const clientEmail = props.getProperty('FIREBASE_CLIENT_EMAIL');
  const rawPrivateKey = props.getProperty('FIREBASE_PRIVATE_KEY');
  if (!clientEmail || !rawPrivateKey) {
    throw new Error('Script Properties에 FIREBASE_CLIENT_EMAIL / FIREBASE_PRIVATE_KEY가 없습니다.');
  }
  const privateKey = rawPrivateKey.replace(/\\n/g, '\n');
  const now = Math.floor(Date.now() / 1000);
  const toB64 = function(s) {
    return Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');
  };
  const header = toB64(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload = toB64(JSON.stringify({
    iss: clientEmail,
    scope: 'https://www.googleapis.com/auth/datastore',
    aud: 'https://oauth2.googleapis.com/token',
    iat: now,
    exp: now + 3600
  }));
  const sigInput = header + '.' + payload;
  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(sigInput, privateKey)
  ).replace(/=+$/, '');
  const assertion = sigInput + '.' + signature;
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: 'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + encodeURIComponent(assertion),
    muteHttpExceptions: true
  });
  const json = JSON.parse(res.getContentText());
  if (!json.access_token) {
    throw new Error('Firestore 토큰 발급 실패: ' + res.getContentText());
  }
  return json.access_token;
}

function pickFirestoreFields_(rowData, fields) {
  const obj = {};
  fields.forEach(function(key) {
    obj[key] = rowData[key] === undefined ? '' : rowData[key];
  });
  return obj;
}

function toFirestoreFields_(data) {
  const fields = {};
  Object.keys(data).forEach(function(key) {
    const val = data[key];
    if (val === undefined || val === null || val === '') {
      fields[key] = { nullValue: null };
      return;
    }
    if (key === 'scan_count' || key === 'gt_growth_point' || key === 'gt_stage') {
      fields[key] = { integerValue: String(Number(val) || 0) };
      return;
    }
    fields[key] = { stringValue: String(val) };
  });
  return fields;
}

function writeFirestoreDoc_(token, collectionName, docId, data) {
  const url =
    'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID +
    '/databases/(default)/documents/' +
    encodeURIComponent(collectionName) + '/' +
    encodeURIComponent(docId);
  const res = UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify({ fields: toFirestoreFields_(data) }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(
      collectionName + '/' + docId + ' 저장 실패 (' +
      res.getResponseCode() + '): ' + res.getContentText()
    );
  }
}

function backupCustomerQrToFirestoreV1_(code) {
  try {
    code = String(code || '').trim().toUpperCase();
    if (!code) return false;
    const row = findQrRow_(code);
    if (row === -1) return false;
    const rowData = getRowObject_(row);
    const backupAt = nowText_();
    rowData.code = code;
    rowData.firestore_backup_at = backupAt;
    const publicData = pickFirestoreFields_(rowData, FIRESTORE_PUBLIC_FIELDS);
    const privateData = pickFirestoreFields_(rowData, FIRESTORE_PRIVATE_FIELDS);
    const token = getFirestoreToken_();
    writeFirestoreDoc_(token, FIRESTORE_PUBLIC_COLLECTION, code, publicData);
    writeFirestoreDoc_(token, FIRESTORE_PRIVATE_COLLECTION, code, privateData);
    const sheet = getSheet_();
    const headers = getHeaderMap_();
    const targetRow = findQrRow_(code);
    if (targetRow !== -1) {
      setByHeader_(sheet, headers, targetRow, 'firestore_backup_at', backupAt);
    }
    syncQrDbRowToProductSheet_(code);
    return true;
  } catch (e) {
    console.error('[고객용 Firestore V1 백업 실패] ' + code + ': ' + e.message);
    return false;
  }
}

// ─── savePushToken ────────────────────────────────────────────────────────────

function savePushToken(params) {
  try {
    const code = String(params.code || '').trim().toUpperCase();
    const token = String(params.token || '').trim();
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };
    if (!token) return { success: false, message: '토큰이 없습니다.' };
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };
    const sheet = getSheet_();
    const headers = getHeaderMap_();
    setByHeader_(sheet, headers, row, 'push_token', token);
    setByHeader_(sheet, headers, row, 'updated_at', nowText_());
    syncQrDbRowToProductSheet_(code);
    backupCustomerQrToFirestoreV1_(code);
    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ─── recordScan ──────────────────────────────────────────────────────────────

function recordScan(code) {
  try {
    code = String(code || '').trim();
    if (!code) return { success: false, message: 'QR 코드가 없습니다.' };
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };
    const data = getRowObject_(row);
    const status = String(data.status || '').trim();
    if (status !== '사용중' && status !== '분실') {
      return { success: true, counted: false, scan_count: Number(data.scan_count || 0) };
    }
    const next = increaseScanCount_(row);
    syncQrDbRowToProductSheet_(code);
    backupCustomerQrToFirestoreV1_(code);
    return { success: true, counted: true, scan_count: next };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}

// ─── OneSignal ───────────────────────────────────────────────────────────────
function getOneSignalConfig_() {
  const props = PropertiesService.getScriptProperties();
  const appId = String(props.getProperty('ONESIGNAL_APP_ID') || '').trim();
  const restApiKey = String(props.getProperty('ONESIGNAL_REST_API_KEY') || '').trim();
  if (!appId) throw new Error('Script Properties에 ONESIGNAL_APP_ID가 없습니다.');
  if (!restApiKey) throw new Error('Script Properties에 ONESIGNAL_REST_API_KEY가 없습니다.');
  return { app_id: appId, rest_api_key: restApiKey };
}

function checkOneSignalSettings() {
  const config = getOneSignalConfig_();
  Logger.log('ONESIGNAL_APP_ID: ' + config.app_id);
  Logger.log('ONESIGNAL_REST_API_KEY 확인 완료');
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
    const productCode = code.split('-')[0].toUpperCase();
    const productNames = { BMT: '부산 메모리 태그', TT: '길동무 태그', CD: '어린이 안심 태그', GT: '가든 태그' };
    const productName = productNames[productCode] || '새김 태그';
    const pageUrl = 'https://saegim-memory.web.app/' + productCode.toLowerCase() + '.html?code=' + encodeURIComponent(code);
    const subject = '[새김] ' + code + ' 새 메시지가 도착했습니다.';
    const body =
      '새김 ' + productName + '에 새로운 메시지가 도착했습니다.\n\n' +
      '태그번호: ' + code + '\n닉네임: ' + nickname + '\n\n메시지:\n' + message +
      '\n\n확인하기: ' + pageUrl;
    MailApp.sendEmail({ to: ownerEmail, subject: subject, body: body });
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
    const row = findQrRow_(code);
    if (row === -1) return { success: false, message: 'QR 정보를 찾을 수 없습니다.' };
    const data = getRowObject_(row);
    const pushToken = String(data.push_token || '').trim();
    if (!pushToken) return { success: false, message: '푸시 토큰이 없습니다.' };
    const config = getOneSignalConfig_();
    const payload = {
      app_id: config.app_id,
      include_subscription_ids: [pushToken],
      headings: { en: '[새김] ' + code + ' 새 메시지', ko: '[새김] ' + code + ' 새 메시지' },
      contents: { en: nickname + ': ' + message, ko: nickname + ': ' + message }
    };
    UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Basic ' + config.rest_api_key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    return { success: true };
  } catch (err) {
    return { success: false, message: String(err) };
  }
}


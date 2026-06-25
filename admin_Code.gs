const APP_TITLE = '새김 QR 통합 관리자';

const MASTER_SHEET_NAME = 'QR_DB';
const PRODUCTS_SHEET_NAME = 'PRODUCTS';
const QR_FOLDER_ID = '1gSlkEIZT-5AuffHwAeEATYMczf4lmlBR';
const SPREADSHEET_ID = '1IgfzKJT4_ohsvP6FZYaj8K_2eMgHBSKrybLtUevM67o';
const CARD_BASE_URL = 'https://saegim-memory.web.app';

// ✅ 활성 제품 타입 (여기서만 관리)
// TT = 길동무 태그
const ACTIVE_PRODUCT_TYPES = ['CD', 'BMT', 'GT', 'WT', 'TT'];

const QR_HEADERS = [
  'code',
  'product',
  'owner',
  'status',
  'qr_file',
  'password',
  'admin_password',

  'link1_type', 'link1_label', 'link1_url',
  'link2_type', 'link2_label', 'link2_url',
  'link3_type', 'link3_label', 'link3_url',
  'link4_type', 'link4_label', 'link4_url',
  'link5_type', 'link5_label', 'link5_url',

  'child_name',
  'guardian_name',
  'guardian_phone',
  'child_allergy',
  'blood_type',
  'child_note',
  'child_message',

  'scan_count',
  'sold_at',
  'registered_at',
  'updated_at',
  'memo',

  'bmt_photo1',
  'bmt_photo2',
  'bmt_photo3',
  'bmt_photo4',
  'bmt_photo5',

  'bmt_travel_memo',
  'bmt_places',
  'bmt_voice',
  'bmt_visit_date',
  'bmt_photo_fit',
  'bmt_photo_position',

  'lost_mode',
  'lost_contact_type',
  'lost_contact_label',
  'lost_contact_url',

  'owner_email',
  'push_token',
  'product_type',

  // GT 새김 가든 태그
  'gt_garden_name',
  'gt_growth_point',
  'gt_stage',
  'gt_slots',
  'gt_inventory',

  // WT 새김 바램태그
  'wt_birth_date',
  'wt_theme',
  'wt_last_message_id',
  'wt_lang',

  // TT 길동무 태그 (필드명 gm_ 유지 — 시트 구조 변경 금지)
  'gm_default_area',
  'gm_enabled_categories',

  // 분실 다국어 메시지
  'lost_message_ko',
  'lost_message_en',
  'lost_message_ja',
  'lost_message_zh',

  // Firestore 백업
  'firestore_backup_at'
];

// ─────────────────────────────────────────────
// Entry
// ─────────────────────────────────────────────

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─────────────────────────────────────────────
// Products
// ─────────────────────────────────────────────

function getProducts() {
  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers     = values[0].map(h => String(h).trim());
  const typeCol     = headers.indexOf('product_type');
  const nameCol     = headers.indexOf('product_name');
  const activeCol   = headers.indexOf('is_active');
  const featuresCol = headers.indexOf('features');

  return values.slice(1)
    .filter(row => {
      const t = String(row[typeCol] || '').trim().toUpperCase();
      return t && ACTIVE_PRODUCT_TYPES.includes(t);
    })
    .map(row => ({
      product_type: String(row[typeCol] || '').trim().toUpperCase(),
      product_name: String(row[nameCol] || '').trim(),
      is_active:    String(row[activeCol] || '').trim().toUpperCase() === 'TRUE',
      features:     String(row[featuresCol] || '').trim()
    }));
}

function saveProduct(form) {
  const productType = String(form.product_type || '').trim().toUpperCase();
  const productName = String(form.product_name || '').trim();
  const features    = String(form.features    || '').trim();

  if (!productType) throw new Error('제품 코드를 입력해주세요.');
  if (!productName) throw new Error('제품명을 입력해주세요.');

  if (!/^[A-Z0-9]{2,6}$/.test(productType)) {
    throw new Error('제품 코드는 영문/숫자 2~6자리만 가능합니다.');
  }

  if (!ACTIVE_PRODUCT_TYPES.includes(productType)) {
    throw new Error(
      '허용되지 않은 제품 코드입니다: ' + productType +
      '\n허용 목록: ' + ACTIVE_PRODUCT_TYPES.join(', ')
    );
  }

  const sheet   = getProductsSheet_();
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const typeCol     = headers.indexOf('product_type');
  const nameCol     = headers.indexOf('product_name');
  const activeCol   = headers.indexOf('is_active');
  const featuresCol = headers.indexOf('features');

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol] || '').trim().toUpperCase() === productType) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow < 0) targetRow = sheet.getLastRow() + 1;

  sheet.getRange(targetRow, typeCol     + 1).setValue(productType);
  sheet.getRange(targetRow, nameCol     + 1).setValue(productName);
  sheet.getRange(targetRow, activeCol   + 1).setValue('TRUE');
  sheet.getRange(targetRow, featuresCol + 1).setValue(features);

  ensureProductSheet_(productType);

  return {
    success:  true,
    message:  productType + ' · ' + productName + ' 저장 완료',
    products: getProducts()
  };
}

// ─────────────────────────────────────────────
// Sheet helpers
// ─────────────────────────────────────────────

function getProductsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(PRODUCTS_SHEET_NAME);

  const phHeaders = ['product_type', 'product_name', 'is_active', 'features'];
  const firstRow  = sheet.getRange(1, 1, 1, phHeaders.length).getValues()[0];
  if (String(firstRow[0] || '').trim() !== 'product_type') {
    sheet.getRange(1, 1, 1, phHeaders.length).setValues([phHeaders]);
  }
  return sheet;
}

function ensureProductsHeaders_() {
  const sheet = getProductsSheet_();
  const requiredHeaders = [
    'product_type', 'product_name', 'is_active', 'features',
    'is_hidden', 'archived_at', 'updated_at'
  ];
  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(h => String(h || '').trim());
  const missingHeaders = requiredHeaders.filter(h => currentHeaders.indexOf(h) < 0);
  if (missingHeaders.length > 0) {
    sheet.getRange(1, lastCol + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }
  return sheet;
}

function ensureSheetHeaders_(sheet) {
  const firstCell = String(sheet.getRange(1, 1).getValue() || '').trim();

  if (!firstCell) {
    sheet.getRange(1, 1, 1, QR_HEADERS.length).setValues([QR_HEADERS]);
    return sheet;
  }

  if (firstCell !== 'code') {
    throw new Error(sheet.getName() + ' 시트의 A1이 code가 아닙니다. 헤더를 확인해주세요.');
  }

  const lastCol = sheet.getLastColumn();
  const currentHeaders = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map(h => String(h || '').trim());

  const missingHeaders = QR_HEADERS.filter(h => currentHeaders.indexOf(h) < 0);

  if (missingHeaders.length > 0) {
    sheet
      .getRange(1, lastCol + 1, 1, missingHeaders.length)
      .setValues([missingHeaders]);
  }

  return sheet;
}

function ensureProductSheet_(productType) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(productType);
  if (!sheet) sheet = ss.insertSheet(productType);
  return ensureSheetHeaders_(sheet);
}

function ensureMasterSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(MASTER_SHEET_NAME);
  return ensureSheetHeaders_(sheet);
}

function getMasterSheet_() {
  return ensureMasterSheet_();
}

// ─────────────────────────────────────────────
// QR Inventory
// ─────────────────────────────────────────────

function makeQrRow_(code, productType, productName, qrFile) {
  const row = new Array(QR_HEADERS.length).fill('');

  const set = (key, val) => {
    const i = QR_HEADERS.indexOf(key);
    if (i >= 0) row[i] = val;
  };

  set('code',               code);
  set('product_type',       productType);
  set('product',            productName);
  set('status',             '미등록');
  set('qr_file',            qrFile);
  set('admin_password',     generateAdminPassword_());
  set('scan_count',         0);
  set('updated_at',         nowText_());
  set('bmt_photo_fit',      'single');
  set('bmt_photo_position', 'center');
  set('lost_mode',          '');
  set('firestore_backup_at', '');

  return row;
}

// ✅ QR_DB에서 코드 중복 검사
function codeExistsInMaster_(code) {
  const sheet  = getMasterSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toUpperCase() === String(code || '').trim().toUpperCase()) {
      return true;
    }
  }
  return false;
}

function createQrInventory(form) {
  const productType = String(form.product_type || '').trim().toUpperCase();
  const quantity    = Number(form.quantity || 0);

  if (!productType) throw new Error('제품을 선택해주세요.');
  if (!quantity || quantity < 1) throw new Error('생성 수량을 확인해주세요.');

  if (!ACTIVE_PRODUCT_TYPES.includes(productType)) {
    throw new Error('허용되지 않은 제품 타입입니다: ' + productType);
  }

  const products = getProducts();
  const product  = products.find(p => p.product_type === productType);
  if (!product) throw new Error('등록되지 않은 제품입니다: ' + productType);

  const masterSheet  = ensureMasterSheet_();
  const productSheet = ensureProductSheet_(productType);
  const folder       = DriveApp.getFolderById(QR_FOLDER_ID);

  const startNumber  = getNextNumber_(productSheet, productType);
  const rows         = [];
  const createdCodes = [];

  let nextNumber = startNumber;

  for (let i = 0; i < quantity; i++) {
    // ✅ 중복 코드 건너뜀
    let code = makeCode_(productType, nextNumber);
    while (codeExistsInMaster_(code)) {
      nextNumber++;
      code = makeCode_(productType, nextNumber);
    }

    const qrUrl  = CARD_BASE_URL + '/?code=' + encodeURIComponent(code);
    const qrFile = saveQrImage_(folder, code, qrUrl);

    rows.push(makeQrRow_(code, productType, product.product_name, qrFile));
    createdCodes.push(code);
    nextNumber++;
  }

  productSheet
    .getRange(productSheet.getLastRow() + 1, 1, rows.length, QR_HEADERS.length)
    .setValues(rows);

  masterSheet
    .getRange(masterSheet.getLastRow() + 1, 1, rows.length, QR_HEADERS.length)
    .setValues(rows);

  createdCodes.forEach(function(c) { backupOneQrCodeToFirestoreV1_(c); });

  return {
    success:   true,
    message:   createdCodes[0] + (createdCodes.length > 1 ? ' ~ ' + createdCodes[createdCodes.length - 1] : '') + ' 생성 완료',
    firstCode: createdCodes[0],
    lastCode:  createdCodes[createdCodes.length - 1],
    count:     createdCodes.length,
    folderUrl: 'https://drive.google.com/drive/folders/' + QR_FOLDER_ID
  };
}

// ─────────────────────────────────────────────
// Inventory & Dashboard
// ─────────────────────────────────────────────

function getInventoryItems() {
  const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
  const products = getProducts();
  const items    = [];

  products.forEach(function(product) {
    const sheet = ss.getSheetByName(product.product_type);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    const headers    = values[0].map(h => String(h).trim());
    const codeCol    = headers.indexOf('code');
    const productCol = headers.indexOf('product');
    const ownerCol   = headers.indexOf('owner');
    const statusCol  = headers.indexOf('status');
    const scanCol    = headers.indexOf('scan_count');
    const updatedCol = headers.indexOf('updated_at');

    values.slice(1).forEach(function(row) {
      if (!row[codeCol]) return;
      items.push({
        code:         String(row[codeCol]    || '').trim(),
        product_type: product.product_type,
        product_name: String(row[productCol] || product.product_name).trim(),
        status:       String(row[statusCol]  || '').trim(),
        owner_name:   String(row[ownerCol]   || '').trim(),
        scan_count:   Number(row[scanCol]    || 0),
        created_at:   String(row[updatedCol] || '')
      });
    });
  });

  return items;
}

function updateQrStatus(code, status) {
  code   = String(code   || '').trim().toUpperCase();
  status = String(status || '').trim();

  if (!code)   throw new Error('코드가 없습니다.');
  if (!status) throw new Error('상태값이 없습니다.');

  const productType = code.split('-')[0];

  if (!ACTIVE_PRODUCT_TYPES.includes(productType)) {
    throw new Error('현재 운영 제품이 아닙니다: ' + productType);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const productSheet = ss.getSheetByName(productType);
  if (!productSheet) throw new Error(productType + ' 시트를 찾을 수 없습니다.');

  const okProduct   = updateStatusInSheet_(productSheet, code, status);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const okMaster    = masterSheet ? updateStatusInSheet_(masterSheet, code, status) : false;

  if (okProduct && okMaster) {
    backupOneQrCodeToFirestoreV1_(code);
  }

  return {
    success:       okProduct,
    message:       code + ' 상태를 [' + status + ']로 변경했습니다.',
    backupUpdated: okMaster
  };
}

function updateStatusInSheet_(sheet, code, status) {
  const values  = sheet.getDataRange().getValues();
  if (values.length < 2) return false;

  const headers       = values[0].map(h => String(h).trim());
  const codeCol       = headers.indexOf('code');
  const statusCol     = headers.indexOf('status');
  const updatedCol    = headers.indexOf('updated_at');
  const soldCol       = headers.indexOf('sold_at');
  const registeredCol = headers.indexOf('registered_at');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][codeCol] || '').trim().toUpperCase() !== code) continue;

    sheet.getRange(i + 1, statusCol + 1).setValue(status);
    if (updatedCol    >= 0) sheet.getRange(i + 1, updatedCol    + 1).setValue(nowText_());
    if (status === '판매완료' && soldCol       >= 0) sheet.getRange(i + 1, soldCol       + 1).setValue(nowText_());
    if (status === '사용중'   && registeredCol >= 0) sheet.getRange(i + 1, registeredCol + 1).setValue(nowText_());
    return true;
  }
  return false;
}

function getDashboardData() {
  const items   = getInventoryItems();
  const summary = { total: items.length, using: 0, sold: 0, unused: 0, stopped: 0, lost: 0 };
  const productMap = {};

  items.forEach(function(item) {
    if      (item.status === '사용중')   summary.using++;
    else if (item.status === '판매완료') summary.sold++;
    else if (item.status === '중지')     summary.stopped++;
    else if (item.status === '분실')     summary.lost++;
    else                                 summary.unused++;

    const key = item.product_type + '|' + item.product_name;
    if (!productMap[key]) {
      productMap[key] = { product_type: item.product_type, product_name: item.product_name, count: 0 };
    }
    productMap[key].count++;
  });

  return {
    summary:  summary,
    products: Object.values(productMap),
    recent:   items.slice(-5).reverse()
  };
}

function getSettingsData() {
  return {
    spreadsheet_id:       SPREADSHEET_ID,
    master_sheet_name:    MASTER_SHEET_NAME,
    products_sheet_name:  PRODUCTS_SHEET_NAME,
    qr_folder_id:         QR_FOLDER_ID,
    card_base_url:        CARD_BASE_URL,
    active_product_types: ACTIVE_PRODUCT_TYPES
  };
}

// ─────────────────────────────────────────────
// Utilities
// ─────────────────────────────────────────────

// ✅ 제품 시트 + QR_DB 둘 다 스캔해서 다음 번호 계산 (중복 방지)
function getNextNumber_(sheet, productType) {
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);

  let max = 0;

  function scanSheet_(targetSheet) {
    if (!targetSheet) return;
    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return;
    const values = targetSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
    values.forEach(function(code) {
      const match = String(code || '').trim().match(new RegExp('^' + productType + '-(\\d+)$'));
      if (match) max = Math.max(max, Number(match[1]));
    });
  }

  scanSheet_(sheet);
  scanSheet_(masterSheet);

  return max + 1;
}

function makeCode_(productType, number) {
  return productType + '-' + String(number).padStart(4, '0');
}

function saveQrImage_(folder, code, qrUrl) {
  const apiUrl =
    'https://quickchart.io/qr?' +
    'text=' + encodeURIComponent(qrUrl) +
    '&size=800&margin=2&ecLevel=H&format=png';

  const response = UrlFetchApp.fetch(apiUrl, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) throw new Error(code + ' QR 이미지 생성 실패');

  const blob = response.getBlob().setName(code + '.png');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function generateAdminPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let result = '';
  for (let i = 0; i < 8; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
}

function nowText_() {
  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}

// ─────────────────────────────────────────────
// Firestore 백업
// ─────────────────────────────────────────────

const FIRESTORE_PROJECT_ID = 'saegim-memory';
const FIRESTORE_COLLECTION = 'qr_cards';

const FIRESTORE_BACKUP_FIELDS = [
  'code', 'product', 'product_type', 'owner', 'status', 'qr_file',
  'scan_count', 'sold_at', 'registered_at', 'updated_at', 'memo',
  'owner_email', 'lost_mode', 'lost_contact_type', 'lost_contact_label',
  'lost_contact_url', 'firestore_backup_at'
];

function getFirestoreToken_() {
  const props         = PropertiesService.getScriptProperties();
  const clientEmail   = props.getProperty('FIREBASE_CLIENT_EMAIL');
  const rawPrivateKey = props.getProperty('FIREBASE_PRIVATE_KEY');

  if (!clientEmail || !rawPrivateKey) {
    throw new Error('Script Properties에 FIREBASE_CLIENT_EMAIL / FIREBASE_PRIVATE_KEY가 없습니다.');
  }

  const privateKey = rawPrivateKey.replace(/\\n/g, '\n');
  const now        = Math.floor(Date.now() / 1000);
  const toB64      = s => Utilities.base64EncodeWebSafe(s).replace(/=+$/, '');
  const header     = toB64(JSON.stringify({ alg: 'RS256', typ: 'JWT' }));
  const payload    = toB64(JSON.stringify({
    iss:   clientEmail,
    scope: 'https://www.googleapis.com/auth/datastore',
    aud:   'https://oauth2.googleapis.com/token',
    iat:   now,
    exp:   now + 3600
  }));

  const sigInput  = header + '.' + payload;
  const signature = Utilities.base64EncodeWebSafe(
    Utilities.computeRsaSha256Signature(sigInput, privateKey)
  ).replace(/=+$/, '');

  const assertion = sigInput + '.' + signature;
  const res = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
    method:      'post',
    contentType: 'application/x-www-form-urlencoded',
    payload:     'grant_type=urn%3Aietf%3Aparams%3Aoauth%3Agrant-type%3Ajwt-bearer&assertion=' + encodeURIComponent(assertion),
    muteHttpExceptions: true
  });

  const json = JSON.parse(res.getContentText());
  if (!json.access_token) throw new Error('토큰 발급 실패: ' + res.getContentText());
  return json.access_token;
}

function backupQrToFirestore_(rowData) {
  const code = String(rowData['code'] || '').trim();
  if (!code) throw new Error('code 값이 없습니다.');

  const backupAt = nowText_();
  rowData['firestore_backup_at'] = backupAt;

  const fields = {};
  FIRESTORE_BACKUP_FIELDS.forEach(function(key) {
    const val = rowData[key];
    if (val === undefined || val === null || val === '') {
      fields[key] = { nullValue: null };
    } else if (key === 'scan_count') {
      fields[key] = { integerValue: String(Number(val) || 0) };
    } else {
      fields[key] = { stringValue: String(val) };
    }
  });

  const token = getFirestoreToken_();
  const url   = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID
              + '/databases/(default)/documents/' + FIRESTORE_COLLECTION
              + '/' + encodeURIComponent(code);

  const res = UrlFetchApp.fetch(url, {
    method:      'patch',
    contentType: 'application/json',
    headers:     { Authorization: 'Bearer ' + token },
    payload:     JSON.stringify({ fields: fields }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error('Firestore 저장 실패 (' + res.getResponseCode() + '): ' + res.getContentText());
  }
  return backupAt;
}

function updateFirestoreBackupAt_(code, backupAt) {
  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const codeCol   = headers.indexOf('code');
  const backupCol = headers.indexOf('firestore_backup_at');
  if (codeCol < 0 || backupCol < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codeCol]).trim() === code) {
      sheet.getRange(i + 1, backupCol + 1).setValue(backupAt);
      return;
    }
  }
}

function tryBackupToFirestore_(code) {
  try {
    const sheet   = getMasterSheet_();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).trim());
    const codeCol = headers.indexOf('code');

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][codeCol]).trim() !== code) continue;
      const rowData = {};
      headers.forEach(function(h, j) { rowData[h] = data[i][j]; });
      const backupAt = backupQrToFirestore_(rowData);
      updateFirestoreBackupAt_(code, backupAt);
      return;
    }
  } catch (e) {
    console.error('[Firestore 백업 실패] ' + code + ': ' + e.message);
  }
}

function testFirestoreBackup() {
  const code = 'CD-0023';
  tryBackupToFirestore_(code);
  Logger.log('백업 완료: ' + code);
}

function testFirestoreBackupStrict() {
  const code = 'CD-0023';

  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const codeCol = headers.indexOf('code');

  if (codeCol < 0) throw new Error('code 헤더를 찾을 수 없습니다.');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codeCol]).trim() !== code) continue;

    const rowData = {};
    headers.forEach(function(h, j) { rowData[h] = data[i][j]; });

    const backupAt = backupQrToFirestore_(rowData);
    updateFirestoreBackupAt_(code, backupAt);

    Logger.log('Firestore 백업 성공: ' + code + ' / ' + backupAt);
    return;
  }

  throw new Error(code + ' 행을 QR_DB에서 찾을 수 없습니다.');
}

// ─────────────────────────────────────────────
// 제품군 보관 / 복구
// ─────────────────────────────────────────────

function archiveProduct(productType) {
  productType = String(productType || '').trim().toUpperCase();
  if (!productType) throw new Error('제품 코드가 없습니다.');
  if (!/^[A-Z0-9]{2,6}$/.test(productType)) throw new Error('제품 코드는 영문/숫자 2~6자리만 가능합니다.');
  if (ACTIVE_PRODUCT_TYPES.includes(productType)) throw new Error('운영 제품은 보관할 수 없습니다: ' + productType);

  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet       = ss.getSheetByName(productType);
  const archiveName = 'ARCHIVE_' + productType;

  if (!sheet) throw new Error(productType + ' 시트를 찾을 수 없습니다.');
  if (ss.getSheetByName(archiveName)) throw new Error(archiveName + ' 시트가 이미 있습니다.');

  sheet.setName(archiveName);
  sheet.hideSheet();
  markProductArchived_(productType);

  return { success: true, message: productType + ' -> ' + archiveName + ' 보관 완료' };
}

function markProductArchived_(productType) {
  const sheet   = ensureProductsHeaders_();
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());

  const typeCol     = headers.indexOf('product_type');
  const activeCol   = headers.indexOf('is_active');
  const hiddenCol   = headers.indexOf('is_hidden');
  const archivedCol = headers.indexOf('archived_at');
  const updatedCol  = headers.indexOf('updated_at');

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol] || '').trim().toUpperCase() === productType) {
      targetRow = i + 1;
      break;
    }
  }
  if (targetRow < 0) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, typeCol + 1).setValue(productType);
  }

  if (activeCol   >= 0) sheet.getRange(targetRow, activeCol   + 1).setValue('FALSE');
  if (hiddenCol   >= 0) sheet.getRange(targetRow, hiddenCol   + 1).setValue('TRUE');
  if (archivedCol >= 0) sheet.getRange(targetRow, archivedCol + 1).setValue(nowText_());
  if (updatedCol  >= 0) sheet.getRange(targetRow, updatedCol  + 1).setValue(nowText_());
}

function restoreArchivedProduct(productType) {
  productType = String(productType || '').trim().toUpperCase();
  if (!productType) throw new Error('제품 코드가 없습니다.');
  if (ACTIVE_PRODUCT_TYPES.includes(productType)) throw new Error('운영 제품은 복구 대상이 아닙니다: ' + productType);

  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const archiveName = 'ARCHIVE_' + productType;
  const sheet       = ss.getSheetByName(archiveName);

  if (!sheet) throw new Error(archiveName + ' 시트를 찾을 수 없습니다.');
  if (ss.getSheetByName(productType)) throw new Error(productType + ' 시트가 이미 있습니다.');

  sheet.showSheet();
  sheet.setName(productType);
  markProductRestored_(productType);

  return { success: true, message: archiveName + ' -> ' + productType + ' 복구 완료' };
}

function markProductRestored_(productType) {
  const sheet   = ensureProductsHeaders_();
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());

  const typeCol    = headers.indexOf('product_type');
  const activeCol  = headers.indexOf('is_active');
  const hiddenCol  = headers.indexOf('is_hidden');
  const updatedCol = headers.indexOf('updated_at');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol] || '').trim().toUpperCase() !== productType) continue;
    if (activeCol  >= 0) sheet.getRange(i + 1, activeCol  + 1).setValue('FALSE');
    if (hiddenCol  >= 0) sheet.getRange(i + 1, hiddenCol  + 1).setValue('FALSE');
    if (updatedCol >= 0) sheet.getRange(i + 1, updatedCol + 1).setValue(nowText_());
    return;
  }
}

function archiveOldProductSheets() {
  const targets = ['BT', 'PT', 'BM', 'AA', 'TH', 'CE', 'TEST'];
  const results = [];

  targets.forEach(function(productType) {
    try {
      const res = archiveProduct(productType);
      results.push(res.message);
    } catch (e) {
      results.push(productType + ': ' + e.message);
    }
  });

  Logger.log(results.join('\n'));
  return { success: true, message: results.join('\n') };
}

function getArchivedProducts() {
  const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const result = [];

  sheets.forEach(function(sheet) {
    const name = sheet.getName();
    if (!name.startsWith('ARCHIVE_')) return;
    result.push({
      product_type: name.replace('ARCHIVE_', ''),
      sheet_name:   name,
      can_restore:  true
    });
  });

  return result;
}

function archiveOldProductSheetsFromUi() {
  return archiveOldProductSheets();
}

function restoreArchivedProductFromUi(productType) {
  productType = String(productType || '').trim().toUpperCase();
  if (!productType) throw new Error('복구할 제품 코드가 없습니다.');
  return restoreArchivedProduct(productType);
}

function restoreTTProductDirect() {
  const ss          = SpreadsheetApp.openById(SPREADSHEET_ID);
  const archiveName = 'ARCHIVE_TT';
  const productType = 'TT';

  const sheet = ss.getSheetByName(archiveName);
  if (!sheet) throw new Error(archiveName + ' 시트를 찾을 수 없습니다.');
  if (ss.getSheetByName(productType)) throw new Error(productType + ' 시트가 이미 있습니다.');

  sheet.showSheet();
  sheet.setName(productType);
  markProductRestored_(productType);

  Logger.log(archiveName + ' -> ' + productType + ' 복구 완료');
  return { success: true, message: archiveName + ' -> ' + productType + ' 복구 완료' };
}

// ─────────────────────────────────────────────
// 초기화 (Apps Script 에디터에서 직접 실행)
// ─────────────────────────────────────────────

function initializeActiveSheets() {
  ensureMasterSheet_();

  ACTIVE_PRODUCT_TYPES.forEach(function(productType) {
    ensureProductSheet_(productType);
  });

  return {
    success: true,
    message: 'QR_DB 및 활성 제품 시트 헤더 정리 완료: ' + ACTIVE_PRODUCT_TYPES.join(', ')
  };
}

// ─────────────────────────────────────────────
// QR_DB → Firestore 전체 이관
// Apps Script 에디터에서 직접 실행
// ─────────────────────────────────────────────

// 이관 제외 필드 (비밀번호류)
const MIGRATE_EXCLUDE_FIELDS = ['password', 'admin_password'];

// QR_DB 전체 행을 Firestore qr_cards 컬렉션에 upsert
// 실행 전 반드시 CD-0025 같은 중복 행 정리 필요
function migrateQrDbToFirestore() {
  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());
  const codeIdx = headers.indexOf('code');

  if (codeIdx < 0) throw new Error('QR_DB에 code 헤더가 없습니다.');

  const token    = getFirestoreToken_();
  const results  = { success: 0, skipped: 0, errors: [] };

  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeIdx] || '').trim();
    if (!code) { results.skipped++; continue; }

    try {
      const rowData = {};
      headers.forEach(function(h, j) {
        if (h && MIGRATE_EXCLUDE_FIELDS.indexOf(h) < 0) {
          rowData[h] = data[i][j];
        }
      });

      const backupAt = nowText_();
      rowData['firestore_backup_at'] = backupAt;

      // Firestore REST fields 변환
      const fields = {};
      Object.keys(rowData).forEach(function(key) {
        var val = rowData[key];
        if (val === undefined || val === null || val === '') {
          fields[key] = { nullValue: null };
        } else if (key === 'scan_count') {
          fields[key] = { integerValue: String(Number(val) || 0) };
        } else {
          fields[key] = { stringValue: String(val) };
        }
      });

      const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID
                + '/databases/(default)/documents/' + FIRESTORE_COLLECTION
                + '/' + encodeURIComponent(code);

      const res = UrlFetchApp.fetch(url, {
        method:      'patch',
        contentType: 'application/json',
        headers:     { Authorization: 'Bearer ' + token },
        payload:     JSON.stringify({ fields: fields }),
        muteHttpExceptions: true
      });

      if (res.getResponseCode() === 200) {
        // firestore_backup_at 시트에도 기록
        const backupColIdx = headers.indexOf('firestore_backup_at');
        if (backupColIdx >= 0) {
          sheet.getRange(i + 1, backupColIdx + 1).setValue(backupAt);
        }
        results.success++;
      } else {
        results.errors.push(code + ': HTTP ' + res.getResponseCode());
      }

      // 토큰은 1시간 유효, 대량 이관 시 속도 제한 대응
      Utilities.sleep(200);

    } catch (e) {
      results.errors.push(code + ': ' + e.message);
    }
  }

  const summary = '이관 완료: ' + results.success + '건'
    + ' / 스킵: ' + results.skipped + '건'
    + (results.errors.length ? ' / 오류: ' + results.errors.join(', ') : '');

  Logger.log(summary);
  return { success: true, message: summary, details: results };
}

// ─────────────────────────────────────────────
// V1: qr_cards(공개) + qr_card_private(비공개) 분리 이관
// ─────────────────────────────────────────────

const FIRESTORE_PUBLIC_COLLECTION  = 'qr_cards';
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
  'gm_default_area', 'gm_enabled_categories',
  'firestore_backup_at'
];

const FIRESTORE_PRIVATE_FIELDS = [
  'code', 'password', 'admin_password', 'owner_email', 'push_token', 'memo',
  'firestore_backup_at'
];

// 중복 코드 검사
function findDuplicateQrCodes() {
  const result = collectDuplicateQrCodes_();
  Logger.log(JSON.stringify(result, null, 2));
  return { success: true, duplicate_count: result.length, duplicates: result };
}

function collectDuplicateQrCodes_() {
  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(function(h) { return String(h || '').trim(); });
  const codeCol = headers.indexOf('code');
  if (codeCol < 0) throw new Error('QR_DB에서 code 헤더를 찾을 수 없습니다.');

  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeCol] || '').trim().toUpperCase();
    if (!code) continue;
    if (!seen[code]) { seen[code] = [i + 1]; } else { seen[code].push(i + 1); }
  }

  return Object.keys(seen)
    .filter(function(c) { return seen[c].length > 1; })
    .map(function(c) { return { code: c, rows: seen[c] }; });
}

// qr_cards(공개) + qr_card_private(비공개) 분리 이관
function migrateQrDbToFirestoreV1() {
  const duplicates = collectDuplicateQrCodes_();
  if (duplicates.length > 0) {
    throw new Error(
      '중복 코드 정리 먼저: ' +
      duplicates.map(function(d) { return d.code + ' rows=' + d.rows.join(','); }).join(' / ')
    );
  }

  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  if (data.length < 2) return { success: true, message: '이관할 데이터가 없습니다.' };

  const headers = data[0].map(function(h) { return String(h || '').trim(); });
  const codeCol = headers.indexOf('code');
  if (codeCol < 0) throw new Error('QR_DB에서 code 헤더를 찾을 수 없습니다.');

  // 토큰 1회 발급 — 행마다 재발급하지 않음
  const token = getFirestoreToken_();

  let successCount = 0;
  let skipCount = 0;
  const failed = [];

  for (let i = 1; i < data.length; i++) {
    const code = String(data[i][codeCol] || '').trim().toUpperCase();
    if (!code) { skipCount++; continue; }

    const rowData = {};
    headers.forEach(function(h, j) { if (h) rowData[h] = data[i][j]; });

    try {
      migrateOneQrRowV1_(token, rowData);
      successCount++;
    } catch (e) {
      failed.push(code + ': ' + e.message);
    }

    Utilities.sleep(150); // 속도 제한 대응
  }

  const message = '분리 이관 완료: 성공 ' + successCount
    + '건 / 스킵 ' + skipCount
    + '건 / 실패 ' + failed.length + '건';
  Logger.log(message);
  if (failed.length) Logger.log(failed.join('\n'));

  return { success: failed.length === 0, message: message, failed: failed };
}

function migrateOneQrRowV1_(token, rowData) {
  const code = String(rowData.code || '').trim().toUpperCase();
  if (!code) throw new Error('code 없음');

  const backupAt = nowText_();
  rowData.firestore_backup_at = backupAt;

  const publicData  = pickFields_(rowData, FIRESTORE_PUBLIC_FIELDS);
  const privateData = pickFields_(rowData, FIRESTORE_PRIVATE_FIELDS);

  writeFirestoreDoc_(token, FIRESTORE_PUBLIC_COLLECTION,  code, publicData);
  writeFirestoreDoc_(token, FIRESTORE_PRIVATE_COLLECTION, code, privateData);

  updateFirestoreBackupAt_(code, backupAt);
}

function pickFields_(rowData, fields) {
  const obj = {};
  fields.forEach(function(key) { obj[key] = rowData[key] === undefined ? '' : rowData[key]; });
  return obj;
}

function writeFirestoreDoc_(token, collectionName, docId, data) {
  const url = 'https://firestore.googleapis.com/v1/projects/' + FIRESTORE_PROJECT_ID
    + '/databases/(default)/documents/'
    + encodeURIComponent(collectionName) + '/' + encodeURIComponent(docId);

  const res = UrlFetchApp.fetch(url, {
    method:      'patch',
    contentType: 'application/json',
    headers:     { Authorization: 'Bearer ' + token },
    payload:     JSON.stringify({ fields: toFirestoreFields_(data) }),
    muteHttpExceptions: true
  });

  if (res.getResponseCode() !== 200) {
    throw new Error(collectionName + '/' + docId + ' 저장 실패 (' + res.getResponseCode() + '): ' + res.getContentText());
  }
}

function toFirestoreFields_(data) {
  const fields = {};
  Object.keys(data).forEach(function(key) {
    var val = data[key];
    if (val === undefined || val === null || val === '') {
      fields[key] = { nullValue: null };
    } else if (key === 'scan_count' || key === 'gt_growth_point' || key === 'gt_stage') {
      fields[key] = { integerValue: String(Number(val) || 0) };
    } else {
      fields[key] = { stringValue: String(val) };
    }
  });
  return fields;
}

// ─────────────────────────────────────────────
// QR 생성 시 Firestore V1 분리 백업 (공개 + 비공개)
// ─────────────────────────────────────────────

function backupOneQrCodeToFirestoreV1_(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) return;

  try {
    const sheet   = getMasterSheet_();
    const data    = sheet.getDataRange().getValues();
    const headers = data[0].map(function(h) { return String(h || '').trim(); });
    const codeCol = headers.indexOf('code');
    if (codeCol < 0) return;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][codeCol] || '').trim().toUpperCase() !== code) continue;
      const rowData = {};
      headers.forEach(function(h, j) { if (h) rowData[h] = data[i][j]; });
      const token = getFirestoreToken_();
      migrateOneQrRowV1_(token, rowData);
      return;
    }
  } catch (e) {
    console.error('[Firestore V1 백업 실패] ' + code + ': ' + e.message);
  }
}

// ─────────────────────────────────────────────
// Admin REST API — doPost
// Firebase /admin 화면에서 POST 요청을 받아 처리
// ─────────────────────────────────────────────

function doPost(e) {
  try {
    const payload = parseAdminPostPayload_(e);
    const action  = String(payload.action || '').trim();

    assertAdminApiKey_(payload.admin_key);

    if (action === 'createQrInventoryFromAdmin') {
      const form = payload.form || payload;
      const result = createQrInventory(form);
      return adminJsonResponse_(result);

    } else if (action === 'updateQrStatusFromAdmin') {
      const result = updateQrStatusFromAdmin(payload.code, payload.status);
      return adminJsonResponse_(result);

    } else if (action === 'regenerateQrImageFromAdmin') {
      const result = regenerateQrImageFromAdmin(payload.code);
      return adminJsonResponse_(result);

    } else if (action === 'updateQrQuickFieldsFromAdmin') {
      const result = updateQrQuickFieldsFromAdmin(payload.code, payload.fields || {});
      return adminJsonResponse_(result);

    } else {
      return adminJsonResponse_({ success: false, message: '알 수 없는 action: ' + action });
    }
  } catch (err) {
    return adminJsonResponse_({ success: false, message: err.message });
  }
}

function parseAdminPostPayload_(e) {
  try {
    const body = e && e.postData ? e.postData.contents : '{}';
    return JSON.parse(body);
  } catch (err) {
    throw new Error('POST body 파싱 실패: ' + err.message);
  }
}

function assertAdminApiKey_(key) {
  const props    = PropertiesService.getScriptProperties();
  const expected = props.getProperty('ADMIN_ACCESS_KEY');
  if (!expected) throw new Error('ADMIN_ACCESS_KEY가 Script Properties에 없습니다.');
  if (String(key || '').trim() !== expected) throw new Error('관리자 키가 올바르지 않습니다.');
}

function adminJsonResponse_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function updateQrStatusFromAdmin(code, status) {
  code   = String(code   || '').trim().toUpperCase();
  status = String(status || '').trim();

  if (!code)   throw new Error('코드가 없습니다.');
  if (!status) throw new Error('상태값이 없습니다.');

  const allowed = ['미등록', '판매완료', '사용중', '중지', '분실'];
  if (allowed.indexOf(status) < 0) {
    throw new Error('허용되지 않은 상태값입니다: ' + status);
  }

  return updateQrStatus(code, status);
}

function regenerateQrImageFromAdmin(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) throw new Error('QR 코드가 없습니다.');

  const sheet = adminGetMasterSheetForRegen_();
  const headers = adminGetHeaderMapForRegen_(sheet);
  const row = adminFindQrRowForRegen_(sheet, code);
  if (row === -1) throw new Error(code + ' 행을 QR_DB에서 찾을 수 없습니다.');

  const pageUrl = 'https://saegim-memory.web.app/?code=' + encodeURIComponent(code);
  const file = adminCreateQrPngFileForRegen_(code, pageUrl);
  const fileUrl = file.getUrl();
  const now = adminNowTextForRegen_();

  adminSetByHeaderForRegen_(sheet, headers, row, 'qr_file', fileUrl);
  adminSetByHeaderForRegen_(sheet, headers, row, 'updated_at', now);
  adminUpdateProductSheetQrFileForRegen_(code, fileUrl, now);

  if (typeof backupOneQrCodeToFirestoreV1_ === 'function') {
    backupOneQrCodeToFirestoreV1_(code);
  }

  return {
    success: true,
    message: code + ' QR 이미지를 재생성했습니다.',
    code: code,
    qr_file: fileUrl
  };
}

function adminGetMasterSheetForRegen_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('QR_DB');
  if (!sheet) throw new Error('QR_DB 시트를 찾을 수 없습니다.');
  return sheet;
}

function adminGetHeaderMapForRegen_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

function adminFindQrRowForRegen_(sheet, code) {
  const headers = adminGetHeaderMapForRegen_(sheet);
  const codeCol = headers.code || 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, codeCol, lastRow - 1, 1).getValues();
  const target = String(code || '').trim().toUpperCase();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toUpperCase() === target) return i + 2;
  }
  return -1;
}

function adminSetByHeaderForRegen_(sheet, headers, row, key, value) {
  const col = headers[key];
  if (!col) return;
  sheet.getRange(row, col).setValue(value);
}

function adminGetQrImageFolderForRegen_() {
  const props = PropertiesService.getScriptProperties();
  const folderId = String(props.getProperty('QR_FOLDER_ID') || '').trim();
  if (folderId) return DriveApp.getFolderById(folderId);
  const name = 'SAEGIM_QR_IMAGES';
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function adminCreateQrPngFileForRegen_(code, pageUrl) {
  const folder = adminGetQrImageFolderForRegen_();
  const qrUrl = 'https://quickchart.io/qr?text=' + encodeURIComponent(pageUrl) +
    '&size=1000&margin=2&ecLevel=H&format=png';
  const blob = UrlFetchApp.fetch(qrUrl).getBlob().setName(code + '.png');
  const existing = folder.getFilesByName(code + '.png');
  while (existing.hasNext()) existing.next().setTrashed(true);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file;
}

function adminUpdateProductSheetQrFileForRegen_(code, fileUrl, now) {
  const productType = String(code || '').split('-')[0].toUpperCase();
  if (!productType) return false;
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(productType);
  if (!sheet) return false;
  const headers = adminGetHeaderMapForRegen_(sheet);
  const row = adminFindQrRowForRegen_(sheet, code);
  if (row === -1) return false;
  adminSetByHeaderForRegen_(sheet, headers, row, 'qr_file', fileUrl);
  adminSetByHeaderForRegen_(sheet, headers, row, 'updated_at', now);
  return true;
}

function adminNowTextForRegen_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// 에디터에서 직접 실행 — ADMIN_ACCESS_KEY 설정 여부 확인
function checkAdminApiKeySetting() {
  const props = PropertiesService.getScriptProperties();
  const key   = props.getProperty('ADMIN_ACCESS_KEY');
  if (!key) {
    Logger.log('❌ ADMIN_ACCESS_KEY 없음 — Script Properties에 추가 필요');
  } else {
    Logger.log('✅ ADMIN_ACCESS_KEY 설정됨 (길이: ' + key.length + ')');
  }
}

// ─────────────────────────────────────────────
// 단일 코드 이관 테스트 (에디터에서 바로 실행)
function testMigrateSingle() {
  const code    = 'CD-0024';
  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h || '').trim());
  const codeIdx = headers.indexOf('code');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codeIdx] || '').trim() !== code) continue;

    const rowData = {};
    headers.forEach(function(h, j) {
      if (h && MIGRATE_EXCLUDE_FIELDS.indexOf(h) < 0) {
        rowData[h] = data[i][j];
      }
    });

    const backupAt = backupQrToFirestore_(rowData);
    updateFirestoreBackupAt_(code, backupAt);
    Logger.log('이관 성공: ' + code + ' / ' + backupAt);
    return;
  }

  throw new Error(code + ' 을 QR_DB에서 찾을 수 없습니다.');
}

// ── 빠른 수정 ────────────────────────────────────────────
function updateQrQuickFieldsFromAdmin(code, fields) {
  code = String(code || '').trim().toUpperCase();
  if (!code) throw new Error('QR 코드가 없습니다.');
  if (!fields || typeof fields !== 'object') throw new Error('수정할 fields 값이 없습니다.');

  const allowedFields = [
    'owner', 'memo', 'lost_mode', 'lost_contact_type', 'lost_contact_label', 'lost_contact_url',
    'child_name', 'guardian_name', 'guardian_phone', 'child_allergy', 'blood_type', 'child_note', 'child_message',
    'bmt_travel_memo', 'bmt_places', 'bmt_visit_date', 'bmt_voice',
    'wt_lang', 'wt_theme', 'wt_birth_date',
    'gt_garden_name'
  ];

  const sheet = adminQuickGetMasterSheet_();
  const headers = adminQuickGetHeaderMap_(sheet);
  const row = adminQuickFindQrRow_(sheet, code);
  if (row === -1) throw new Error(code + ' 행을 QR_DB에서 찾을 수 없습니다.');

  let changedCount = 0;
  allowedFields.forEach(function(key) {
    if (!Object.prototype.hasOwnProperty.call(fields, key)) return;
    if (!headers[key]) return;
    let value = fields[key];
    if (value === undefined || value === null) value = '';
    sheet.getRange(row, headers[key]).setValue(value);
    changedCount++;
  });

  if (!changedCount) return { success: false, message: '수정할 수 있는 필드가 없습니다.' };

  if (headers.updated_at) {
    sheet.getRange(row, headers.updated_at).setValue(adminQuickNowText_());
  }

  adminQuickSyncQrDbRowToProductSheet_(code);

  if (typeof backupOneQrCodeToFirestoreV1_ === 'function') {
    backupOneQrCodeToFirestoreV1_(code);
  }

  return { success: true, message: code + ' 빠른 수정 완료', changed_count: changedCount };
}

function adminQuickGetMasterSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('QR_DB');
  if (!sheet) throw new Error('QR_DB 시트를 찾을 수 없습니다.');
  return sheet;
}

function adminQuickGetHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach(function(h, i) { if (h) map[String(h).trim()] = i + 1; });
  return map;
}

function adminQuickFindQrRow_(sheet, code) {
  const headers = adminQuickGetHeaderMap_(sheet);
  const codeCol = headers.code || 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const values = sheet.getRange(2, codeCol, lastRow - 1, 1).getValues();
  const target = String(code || '').trim().toUpperCase();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toUpperCase() === target) return i + 2;
  }
  return -1;
}

function adminQuickGetRowObject_(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach(function(h, i) { if (h) obj[String(h).trim()] = values[i]; });
  return obj;
}

function adminQuickSyncQrDbRowToProductSheet_(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) return false;

  const masterSheet = adminQuickGetMasterSheet_();
  const masterRow = adminQuickFindQrRow_(masterSheet, code);
  if (masterRow === -1) return false;

  const productType = code.split('-')[0];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const productSheet = ss.getSheetByName(productType);
  if (!productSheet) return false;

  const masterData = adminQuickGetRowObject_(masterSheet, masterRow);
  const productHeaders = adminQuickGetHeaderMap_(productSheet);
  let productRow = adminQuickFindQrRow_(productSheet, code);
  if (productRow === -1) productRow = productSheet.getLastRow() + 1;

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

function adminQuickNowText_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

const APP_TITLE = '새김 QR 통합 관리자';

const MASTER_SHEET_NAME = 'QR_DB';
const PRODUCTS_SHEET_NAME = 'PRODUCTS';
const QR_FOLDER_ID = '1gSlkEIZT-5AuffHwAeEATYMczf4lmlBR';
const SPREADSHEET_ID = '1IgfzKJT4_ohsvP6FZYaj8K_2eMgHBSKrybLtUevM67o';
const CARD_BASE_URL = 'https://saegim-memory.web.app';

// ✅ 활성 제품 타입 (여기서만 관리)
// GM = 길동무태그 (기존 TT에서 변경 시 tt.html → gm.html 리네임 필요)
const ACTIVE_PRODUCT_TYPES = ['CD', 'BMT', 'GT', 'WT', 'GM'];

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
  'bmt_visit_date',       // ← bmt_voice 바로 다음
  'bmt_photo_fit',
  'bmt_photo_position',

  'lost_mode',
  'lost_contact_type',
  'lost_contact_label',
  'lost_contact_url',

  'owner_email',
  'push_token',           // ← 알림 확장용 유지
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

  // GM 길동무 태그
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

  const headers = values[0].map(h => String(h).trim());
  const typeCol   = headers.indexOf('product_type');
  const nameCol   = headers.indexOf('product_name');
  const activeCol = headers.indexOf('is_active');
  const featuresCol = headers.indexOf('features');

  return values.slice(1)
    .filter(row => {
      const t = String(row[typeCol] || '').trim().toUpperCase();
      // ✅ ACTIVE_PRODUCT_TYPES에 있는 것만 노출
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

  // ✅ ACTIVE_PRODUCT_TYPES 외 코드 차단
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

// ✅ 공통 헤더 보장 함수 (기존 시트의 누락 헤더도 자동 추가)
function ensureSheetHeaders_(sheet) {
  const firstCell = String(sheet.getRange(1, 1).getValue() || '').trim();

  // 빈 시트: 전체 헤더 작성
  if (!firstCell) {
    sheet.getRange(1, 1, 1, QR_HEADERS.length).setValues([QR_HEADERS]);
    return sheet;
  }

  // A1이 'code'가 아니면 구조 이상
  if (firstCell !== 'code') {
    throw new Error(sheet.getName() + ' 시트의 A1이 code가 아닙니다. 헤더를 확인해주세요.');
  }

  // 기존 헤더와 비교해 없는 것만 맨 뒤에 추가
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
  return ensureSheetHeaders_(sheet);   // ✅ 공통 함수 사용
}

function ensureMasterSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(MASTER_SHEET_NAME);
  return ensureSheetHeaders_(sheet);   // ✅ 공통 함수 사용
}

function getMasterSheet_() {
  return ensureMasterSheet_();
}

// ─────────────────────────────────────────────
// QR Inventory
// ─────────────────────────────────────────────

// ✅ 행 데이터 생성 헬퍼 (헤더 순서 변경에 자동 대응)
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

function createQrInventory(form) {
  const productType = String(form.product_type || '').trim().toUpperCase();
  const quantity    = Number(form.quantity || 0);

  if (!productType) throw new Error('제품을 선택해주세요.');
  if (!quantity || quantity < 1) throw new Error('생성 수량을 확인해주세요.');

  // ✅ 활성 제품만 허용
  if (!ACTIVE_PRODUCT_TYPES.includes(productType)) {
    throw new Error('허용되지 않은 제품 타입입니다: ' + productType);
  }

  const products = getProducts();
  const product  = products.find(p => p.product_type === productType);
  if (!product) throw new Error('등록되지 않은 제품입니다: ' + productType);

  const masterSheet  = ensureMasterSheet_();
  const productSheet = ensureProductSheet_(productType);
  const folder       = DriveApp.getFolderById(QR_FOLDER_ID);

  const startNumber = getNextNumber_(productSheet, productType);
  const rows        = [];
  const createdCodes = [];

  for (let i = 0; i < quantity; i++) {
    const code   = makeCode_(productType, startNumber + i);
    const qrUrl  = CARD_BASE_URL + '/?code=' + encodeURIComponent(code);
    const qrFile = saveQrImage_(folder, code, qrUrl);

    rows.push(makeQrRow_(code, productType, product.product_name, qrFile));   // ✅ 헬퍼 사용
    createdCodes.push(code);
  }

  productSheet
    .getRange(productSheet.getLastRow() + 1, 1, rows.length, QR_HEADERS.length)
    .setValues(rows);

  masterSheet
    .getRange(masterSheet.getLastRow() + 1, 1, rows.length, QR_HEADERS.length)
    .setValues(rows);

  // ✅ 시트 저장 완료 후 Firestore 백업
  createdCodes.forEach(function(c) { tryBackupToFirestore_(c); });

  return {
    success:    true,
    message:    createdCodes[0] + (createdCodes.length > 1 ? ' ~ ' + createdCodes[createdCodes.length - 1] : '') + ' 생성 완료',
    firstCode:  createdCodes[0],
    lastCode:   createdCodes[createdCodes.length - 1],
    count:      createdCodes.length,
    folderUrl:  'https://drive.google.com/drive/folders/' + QR_FOLDER_ID
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

    const headers     = values[0].map(h => String(h).trim());
    const codeCol     = headers.indexOf('code');
    const productCol  = headers.indexOf('product');
    const ownerCol    = headers.indexOf('owner');
    const statusCol   = headers.indexOf('status');
    const scanCol     = headers.indexOf('scan_count');
    const updatedCol  = headers.indexOf('updated_at');

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

  // ✅ 운영 제품만 상태 변경 허용
  if (!ACTIVE_PRODUCT_TYPES.includes(productType)) {
    throw new Error('현재 운영 제품이 아닙니다: ' + productType);
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const productSheet = ss.getSheetByName(productType);
  if (!productSheet) throw new Error(productType + ' 시트를 찾을 수 없습니다.');

  const okProduct = updateStatusInSheet_(productSheet, code, status);
  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  const okMaster    = masterSheet ? updateStatusInSheet_(masterSheet, code, status) : false;

  // ✅ 시트 저장 완료 후 Firestore 백업 (QR_DB까지 성공한 경우에만)
  if (okProduct && okMaster) tryBackupToFirestore_(code);

  return {
    success:       okProduct,
    message:       code + ' 상태를 [' + status + ']로 변경했습니다.',
    backupUpdated: okMaster
  };
}

function updateStatusInSheet_(sheet, code, status) {
  const values  = sheet.getDataRange().getValues();
  if (values.length < 2) return false;

  const headers      = values[0].map(h => String(h).trim());
  const codeCol      = headers.indexOf('code');
  const statusCol    = headers.indexOf('status');
  const updatedCol   = headers.indexOf('updated_at');
  const soldCol      = headers.indexOf('sold_at');
  const registeredCol = headers.indexOf('registered_at');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][codeCol] || '').trim().toUpperCase() !== code) continue;

    sheet.getRange(i + 1, statusCol + 1).setValue(status);
    if (updatedCol    >= 0) sheet.getRange(i + 1, updatedCol    + 1).setValue(nowText_());
    if (status === '판매완료'  && soldCol      >= 0) sheet.getRange(i + 1, soldCol      + 1).setValue(nowText_());
    if (status === '사용중'    && registeredCol >= 0) sheet.getRange(i + 1, registeredCol + 1).setValue(nowText_());
    return true;
  }
  return false;
}

function getDashboardData() {
  const items = getInventoryItems();
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
    spreadsheet_id:      SPREADSHEET_ID,
    master_sheet_name:   MASTER_SHEET_NAME,
    products_sheet_name: PRODUCTS_SHEET_NAME,
    qr_folder_id:        QR_FOLDER_ID,
    card_base_url:       CARD_BASE_URL,
    active_product_types: ACTIVE_PRODUCT_TYPES   // ✅ 설정 화면에서 확인 가능
  };
}

// ─────────────────────────────────────────────
// Utilities
// ─────────────────────────────────────────────

function getNextNumber_(sheet, productType) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const codes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let max = 0;
  codes.forEach(code => {
    const match = String(code || '').match(new RegExp('^' + productType + '-(\\d+)$'));
    if (match) max = Math.max(max, Number(match[1]));
  });
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

const FIRESTORE_PROJECT_ID  = 'saegim-memory';
const FIRESTORE_COLLECTION  = 'qr_cards';

// 백업 대상 필드 (password / admin_password 제외)
const FIRESTORE_BACKUP_FIELDS = [
  'code', 'product', 'product_type', 'owner', 'status', 'qr_file',
  'scan_count', 'sold_at', 'registered_at', 'updated_at', 'memo',
  'owner_email', 'lost_mode', 'lost_contact_type', 'lost_contact_label',
  'lost_contact_url', 'firestore_backup_at'
];

// 서비스 계정 JWT → access_token
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

// rowData: 헤더-값 매핑 객체  ex) { code: 'CD-0023', status: '판매완료', ... }
// 반환: firestore_backup_at 값 (성공 시)
function backupQrToFirestore_(rowData) {
  const code = String(rowData['code'] || '').trim();
  if (!code) throw new Error('code 값이 없습니다.');

  const backupAt = nowText_();
  rowData['firestore_backup_at'] = backupAt;

  // Firestore REST API fields 형식으로 변환
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

// 시트에서 해당 code 행의 firestore_backup_at 컬럼 업데이트
function updateFirestoreBackupAt_(code, backupAt) {
  const sheet   = getMasterSheet_();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const codeCol    = headers.indexOf('code');
  const backupCol  = headers.indexOf('firestore_backup_at');
  if (codeCol < 0 || backupCol < 0) return;

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][codeCol]).trim() === code) {
      sheet.getRange(i + 1, backupCol + 1).setValue(backupAt);
      return;
    }
  }
}

// 시트에서 code로 rowData 빌드 후 Firestore 백업 + backup_at 시트 기록
// try-catch 래핑: Firestore 실패가 시트 저장에 영향 없도록
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
    // Firestore 실패는 로그만 남기고 무시 (시트 저장은 이미 완료)
    console.error('[Firestore 백업 실패] ' + code + ': ' + e.message);
  }
}

// 테스트용 — Apps Script 에디터에서 직접 실행
function testFirestoreBackup() {
  const code = 'CD-0023';
  tryBackupToFirestore_(code);
  Logger.log('백업 완료: ' + code);
}

// ─────────────────────────────────────────────
// 초기화 (Apps Script 에디터에서 직접 실행)
// ─────────────────────────────────────────────

// 기존 시트 포함 전체 헤더 일괄 정리
// 실행 방법: Apps Script 에디터 > 함수 선택 > initializeActiveSheets > 실행
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

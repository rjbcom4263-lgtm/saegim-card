const APP_TITLE = '새김 QR 통합 관리자';

const MASTER_SHEET_NAME = 'QR_DB';
const PRODUCTS_SHEET_NAME = 'PRODUCTS';
const QR_FOLDER_ID = '1gSlkEIZT-5AuffHwAeEATYMczf4lmlBR';
const SPREADSHEET_ID = '1IgfzKJT4_ohsvP6FZYaj8K_2eMgHBSKrybLtUevM67o';

const CARD_BASE_URL = 'https://saegim-memory.web.app';

// ── 컬럼 헤더 (스프레드시트 열 순서) ────────────────────────
const QR_HEADERS = [
  // 기본 정보
  'code',
  'product',
  'owner',
  'status',
  'qr_file',
  'password',
  'admin_password',

  // 링크 (최대 5개)
  'link1_type', 'link1_label', 'link1_url',
  'link2_type', 'link2_label', 'link2_url',
  'link3_type', 'link3_label', 'link3_url',
  'link4_type', 'link4_label', 'link4_url',
  'link5_type', 'link5_label', 'link5_url',

  // CD 어린이 안심 태그 필드
  'child_photo',
  'child_name',
  'guardian_name',
  'guardian_phone',
  'child_allergy',
  'blood_type',
  'child_note',
  'child_message',

  // 메타
  'scan_count',
  'sold_at',
  'registered_at',
  'updated_at',
  'memo',

  // BMT 부산 메모리 태그 필드
  'bmt_photo1', 'bmt_photo2', 'bmt_photo3', 'bmt_photo4', 'bmt_photo5',
  'bmt_travel_memo',
  'bmt_places',
  'bmt_voice',
  'bmt_visit_date',
  'bmt_photo_fit',
  'bmt_photo_position',

  // BMT/TT 분실 모드 공통 필드
  'lost_mode',
  'lost_contact_type',
  'lost_contact_label',
  'lost_contact_url',
  'owner_email',
  'push_token',
  'lost_message_ko',
  'lost_message_en',
  'lost_message_ja',
  'lost_message_zh',
];

// ─── 진입점 ───────────────────────────────────────────────

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── 제품 관리 ───────────────────────────────────────────

function getProducts(includeInactive) {
  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();

  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const typeCol = headers.indexOf('product_type');
  const nameCol = headers.indexOf('product_name');
  const activeCol = headers.indexOf('is_active');
  const featuresCol = headers.indexOf('features');

  return values.slice(1)
    .filter(row => {
      const type = String(row[typeCol] || '').trim();
      if (!type) return false;
      if (!includeInactive) {
        return String(row[activeCol] || '').trim().toUpperCase() === 'TRUE';
      }
      return true;
    })
    .map(row => ({
      product_type: String(row[typeCol] || '').trim().toUpperCase(),
      product_name: String(row[nameCol] || '').trim(),
      is_active: String(row[activeCol] || '').trim().toUpperCase() === 'TRUE',
      features: String(row[featuresCol] || '').trim()
    }));
}

function saveProduct(form) {
  const productType = String(form.product_type || '').trim().toUpperCase();
  const productName = String(form.product_name || '').trim();
  const features = String(form.features || '').trim();

  if (!productType) throw new Error('제품 코드를 입력해주세요.');
  if (!productName) throw new Error('제품명을 입력해주세요.');
  if (!/^[A-Z0-9]{2,6}$/.test(productType)) {
    throw new Error('제품 코드는 영문/숫자 2~6자리만 가능합니다.');
  }

  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const typeCol = headers.indexOf('product_type');
  const nameCol = headers.indexOf('product_name');
  const activeCol = headers.indexOf('is_active');
  const featuresCol = headers.indexOf('features');

  let targetRow = -1;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol] || '').trim().toUpperCase() === productType) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow < 0) targetRow = sheet.getLastRow() + 1;

  sheet.getRange(targetRow, typeCol + 1).setValue(productType);
  sheet.getRange(targetRow, nameCol + 1).setValue(productName);
  sheet.getRange(targetRow, activeCol + 1).setValue('TRUE');
  sheet.getRange(targetRow, featuresCol + 1).setValue(features);

  ensureProductSheet_(productType);

  return {
    success: true,
    message: productType + ' · ' + productName + ' 저장 완료',
    products: getProducts(true)
  };
}

// 제품 비활성화 (is_active = FALSE) — 재고 목록에서 제외됨
function deactivateProduct(productType) {
  productType = String(productType || '').trim().toUpperCase();
  if (!productType) throw new Error('제품 코드가 없습니다.');

  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());

  const typeCol = headers.indexOf('product_type');
  const activeCol = headers.indexOf('is_active');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol] || '').trim().toUpperCase() === productType) {
      sheet.getRange(i + 1, activeCol + 1).setValue('FALSE');
      return { success: true, message: productType + ' 비활성화 완료', products: getProducts(true) };
    }
  }

  throw new Error(productType + ' 제품을 찾을 수 없습니다.');
}

// ─── 시트 헬퍼 ───────────────────────────────────────────

function getProductsSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PRODUCTS_SHEET_NAME);

  if (!sheet) sheet = ss.insertSheet(PRODUCTS_SHEET_NAME);

  const headers = ['product_type', 'product_name', 'is_active', 'features'];
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];

  if (String(firstRow[0] || '').trim() !== 'product_type') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function ensureProductSheet_(productType) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(productType);

  if (!sheet) sheet = ss.insertSheet(productType);

  const firstRow = sheet.getRange(1, 1, 1, QR_HEADERS.length).getValues()[0];

  if (String(firstRow[0] || '').trim() !== 'code') {
    sheet.getRange(1, 1, 1, QR_HEADERS.length).setValues([QR_HEADERS]);
  } else {
    // 기존 시트에 누락된 컬럼 추가
    addMissingColumns_(sheet);
  }

  return sheet;
}

function ensureMasterSheet_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(MASTER_SHEET_NAME);

  if (!sheet) sheet = ss.insertSheet(MASTER_SHEET_NAME);

  const firstRow = sheet.getRange(1, 1, 1, QR_HEADERS.length).getValues()[0];

  if (String(firstRow[0] || '').trim() !== 'code') {
    sheet.getRange(1, 1, 1, QR_HEADERS.length).setValues([QR_HEADERS]);
  } else {
    addMissingColumns_(sheet);
  }

  return sheet;
}

// 기존 시트에 QR_HEADERS에 있지만 빠진 컬럼을 오른쪽에 추가
function addMissingColumns_(sheet) {
  const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(h => String(h).trim());

  const missing = QR_HEADERS.filter(h => !existingHeaders.includes(h));
  if (missing.length === 0) return;

  const startCol = sheet.getLastColumn() + 1;
  sheet.getRange(1, startCol, 1, missing.length).setValues([missing]);
}

// ─── QR 재고 생성 ─────────────────────────────────────────

function createQrInventory(form) {
  const productType = String(form.product_type || '').trim().toUpperCase();
  const quantity = Number(form.quantity || 0);

  if (!productType) throw new Error('제품을 선택해주세요.');
  if (!quantity || quantity < 1) throw new Error('생성 수량을 확인해주세요.');

  // 활성 제품인지 확인
  const products = getProducts();
  const product = products.find(p => p.product_type === productType);
  if (!product) throw new Error('등록되지 않은 제품입니다. 먼저 제품을 등록해주세요.');

  const masterSheet = ensureMasterSheet_();
  const productSheet = ensureProductSheet_(productType);
  const folder = DriveApp.getFolderById(QR_FOLDER_ID);

  const startNumber = getNextNumber_(productSheet, productType);
  const rows = [];
  const createdCodes = [];

  for (let i = 0; i < quantity; i++) {
    const code = makeCode_(productType, startNumber + i);
    const qrUrl = CARD_BASE_URL + '/?code=' + encodeURIComponent(code);
    const qrFile = saveQrImage_(folder, code, qrUrl);
    const now = nowText_();

    // QR_HEADERS 순서에 맞게 정확히 매핑
    const row = QR_HEADERS.map(key => {
      switch (key) {
        case 'code': return code;
        case 'product': return product.product_name;
        case 'status': return '미등록';
        case 'qr_file': return qrFile;
        case 'admin_password': return generateAdminPassword_();
        case 'scan_count': return 0;
        case 'updated_at': return now;
        case 'bmt_photo_fit': return 'single';
        case 'bmt_photo_position': return 'center';
        default: return '';
      }
    });

    rows.push(row);
    createdCodes.push(code);
  }

  productSheet
    .getRange(productSheet.getLastRow() + 1, 1, rows.length, QR_HEADERS.length)
    .setValues(rows);

  masterSheet
    .getRange(masterSheet.getLastRow() + 1, 1, rows.length, QR_HEADERS.length)
    .setValues(rows);

  return {
    success: true,
    message: createdCodes[0] + ' ~ ' + createdCodes[createdCodes.length - 1] + ' 생성 완료',
    firstCode: createdCodes[0],
    lastCode: createdCodes[createdCodes.length - 1],
    count: createdCodes.length,
    folderUrl: 'https://drive.google.com/drive/folders/' + QR_FOLDER_ID
  };
}

// ─── 재고 조회 ────────────────────────────────────────────

function getInventoryItems() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  // 비활성 포함 전체 제품 시트에서 조회
  const products = getProducts(true);
  const items = [];

  products.forEach(function(product) {
    const sheet = ss.getSheetByName(product.product_type);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return;

    const headers = values[0].map(h => String(h).trim());
    const codeCol = headers.indexOf('code');
    const productCol = headers.indexOf('product');
    const ownerCol = headers.indexOf('owner');
    const statusCol = headers.indexOf('status');
    const scanCol = headers.indexOf('scan_count');
    const updatedCol = headers.indexOf('updated_at');

    values.slice(1).forEach(function(row) {
      const code = String(row[codeCol] || '').trim();
      if (!code) return;

      items.push({
        code: code,
        product_type: product.product_type,
        product_name: String(row[productCol] || product.product_name).trim(),
        is_active: product.is_active,
        status: String(row[statusCol] || '').trim(),
        owner_name: String(row[ownerCol] || '').trim(),
        scan_count: Number(row[scanCol] || 0),
        created_at: String(row[updatedCol] || '')
      });
    });
  });

  return items;
}

// ─── 상태 변경 ────────────────────────────────────────────

function updateQrStatus(code, status) {
  code = String(code || '').trim().toUpperCase();
  status = String(status || '').trim();

  if (!code) throw new Error('코드가 없습니다.');
  if (!status) throw new Error('상태값이 없습니다.');

  const allowed = ['미등록', '판매완료', '사용중', '분실', '중지'];
  if (!allowed.includes(status)) throw new Error('올바르지 않은 상태입니다: ' + status);

  const productType = code.split('-')[0];
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  const productSheet = ss.getSheetByName(productType);
  if (!productSheet) throw new Error(productType + ' 시트를 찾을 수 없습니다.');

  const okProduct = updateStatusInSheet_(productSheet, code, status);

  const masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  let okMaster = false;
  if (masterSheet) okMaster = updateStatusInSheet_(masterSheet, code, status);

  if (!okProduct) throw new Error(code + ' 코드를 찾을 수 없습니다.');

  return {
    success: true,
    message: code + ' 상태를 [' + status + ']로 변경했습니다.',
    backupUpdated: okMaster
  };
}

function updateStatusInSheet_(sheet, code, status) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return false;

  const headers = values[0].map(h => String(h).trim());
  const codeCol = headers.indexOf('code');
  const statusCol = headers.indexOf('status');
  const updatedCol = headers.indexOf('updated_at');
  const soldCol = headers.indexOf('sold_at');
  const registeredCol = headers.indexOf('registered_at');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][codeCol] || '').trim().toUpperCase() !== code) continue;

    sheet.getRange(i + 1, statusCol + 1).setValue(status);
    if (updatedCol >= 0) sheet.getRange(i + 1, updatedCol + 1).setValue(nowText_());
    if (status === '판매완료' && soldCol >= 0) sheet.getRange(i + 1, soldCol + 1).setValue(nowText_());
    if (status === '사용중' && registeredCol >= 0) sheet.getRange(i + 1, registeredCol + 1).setValue(nowText_());

    return true;
  }

  return false;
}

// ─── 대시보드 ────────────────────────────────────────────

function getDashboardData() {
  const items = getInventoryItems();

  const summary = { total: 0, using: 0, sold: 0, unused: 0, stopped: 0, lost: 0 };
  const productMap = {};

  // 활성 제품 항목만 집계
  items.filter(item => item.is_active).forEach(function(item) {
    summary.total++;
    switch (item.status) {
      case '사용중': summary.using++; break;
      case '판매완료': summary.sold++; break;
      case '중지': summary.stopped++; break;
      case '분실': summary.lost++; break;
      default: summary.unused++; // '미등록' 포함
    }

    const key = item.product_type + '|' + item.product_name;
    if (!productMap[key]) {
      productMap[key] = { product_type: item.product_type, product_name: item.product_name, count: 0 };
    }
    productMap[key].count++;
  });

  const recent = items
    .filter(item => item.is_active && item.status !== '미등록')
    .sort((a, b) => b.created_at > a.created_at ? 1 : -1)
    .slice(0, 5);

  return {
    summary: summary,
    products: Object.values(productMap),
    recent: recent
  };
}

function getSettingsData() {
  return {
    spreadsheet_id: SPREADSHEET_ID,
    master_sheet_name: MASTER_SHEET_NAME,
    products_sheet_name: PRODUCTS_SHEET_NAME,
    qr_folder_id: QR_FOLDER_ID,
    card_base_url: CARD_BASE_URL
  };
}

// ─── 유틸 ────────────────────────────────────────────────

function getNextNumber_(sheet, productType) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const codes = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  let max = 0;
  const re = new RegExp('^' + productType + '-(\\d+)$');

  codes.forEach(code => {
    const match = String(code || '').match(re);
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

  if (response.getResponseCode() !== 200) {
    throw new Error(code + ' QR 이미지 생성 실패');
  }

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

// ─────────────────────────────────────────────────────────
// 보관 제품 관리 (비활성 제품 아카이브)
// ─────────────────────────────────────────────────────────

function getArchivedProducts() {
  const sheet = getProductsSheet_();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(h => String(h).trim());
  const typeCol = headers.indexOf('product_type');
  const nameCol = headers.indexOf('product_name');
  const activeCol = headers.indexOf('is_active');
  const featuresCol = headers.indexOf('features');

  return values.slice(1)
    .filter(row => {
      const type = String(row[typeCol] || '').trim();
      const active = String(row[activeCol] || '').trim().toUpperCase();
      return type && active !== 'TRUE';
    })
    .map(row => ({
      product_type: String(row[typeCol] || '').trim().toUpperCase(),
      product_name: String(row[nameCol] || '').trim(),
      is_active: false,
      features: String(row[featuresCol] || '').trim()
    }));
}

function archiveOldProductSheetsFromUi() {
  const products = getProducts(true);
  const inactive = products.filter(p => !p.is_active);
  if (inactive.length === 0) return { success: true, message: '보관할 비활성 제품이 없습니다.' };

  const archived = [];
  inactive.forEach(function(p) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(p.product_type);
    if (!sheet) return;
    // 시트 이름에 _ARCHIVED 접미사 추가
    const newName = p.product_type + '_ARCHIVED';
    const existing = ss.getSheetByName(newName);
    if (!existing) {
      sheet.setName(newName);
      archived.push(p.product_type);
    }
  });

  return {
    success: true,
    message: archived.length > 0
      ? archived.join(', ') + ' 시트 보관 완료'
      : '이미 보관 처리된 제품입니다.',
    archived: archived
  };
}

function restoreArchivedProductFromUi(productType) {
  productType = String(productType || '').trim().toUpperCase();
  if (!productType) throw new Error('제품 코드가 없습니다.');

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const archivedName = productType + '_ARCHIVED';
  const sheet = ss.getSheetByName(archivedName);

  if (!sheet) throw new Error(archivedName + ' 시트를 찾을 수 없습니다.');

  sheet.setName(productType);

  // PRODUCTS 시트에서 is_active를 TRUE로 복원
  const prodSheet = getProductsSheet_();
  const values = prodSheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const typeCol = headers.indexOf('product_type');
  const activeCol = headers.indexOf('is_active');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][typeCol] || '').trim().toUpperCase() === productType) {
      prodSheet.getRange(i + 1, activeCol + 1).setValue('TRUE');
      break;
    }
  }

  return { success: true, message: productType + ' 복원 완료' };
}

// ─────────────────────────────────────────────────────────
// Firebase Admin API Router
// Firebase /admin 화면에서 fetch()로 호출
// Content-Type: text/plain 사용
// ─────────────────────────────────────────────────────────

function doPost(e) {
  try {
    const payload = parsePostPayload_(e);
    const action = String(payload.action || '').trim();

    if (!action) throw new Error('action 값이 없습니다.');

    // verifyAdminKey 외 모든 액션은 키 검증 필수
    assertAdminAccess_(payload.admin_key);

    let result;

    if (action === 'verifyAdminKey') {
      result = { success: true, message: '관리자 인증 성공' };

    } else if (action === 'getProducts') {
      result = { success: true, data: getProducts(true) };

    } else if (action === 'getDashboardData') {
      result = { success: true, data: getDashboardData() };

    } else if (action === 'getInventoryItems') {
      result = { success: true, data: getInventoryItems() };

    } else if (action === 'getSettingsData') {
      result = { success: true, data: getSettingsData() };

    } else if (action === 'createQrInventory') {
      result = createQrInventory(payload.form || {});

    } else if (action === 'updateQrStatus') {
      result = updateQrStatus(payload.code, payload.status);

    } else if (action === 'saveProduct') {
      result = saveProduct(payload.form || {});

    } else if (action === 'deactivateProduct') {
      result = deactivateProduct(payload.product_type);

    } else if (action === 'getArchivedProducts') {
      result = { success: true, data: getArchivedProducts() };

    } else if (action === 'archiveOldProductSheets') {
      result = archiveOldProductSheetsFromUi();

    } else if (action === 'restoreArchivedProduct') {
      result = restoreArchivedProductFromUi(payload.product_type);

    } else {
      throw new Error('알 수 없는 action입니다: ' + action);
    }

    return jsonResponse_(result);

  } catch (err) {
    return jsonResponse_({
      success: false,
      message: err && err.message ? err.message : String(err)
    });
  }
}

function parsePostPayload_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  const text = String(e.postData.contents || '').trim();
  if (!text) return {};
  try {
    return JSON.parse(text);
  } catch (err) {
    throw new Error('요청 JSON 파싱 실패: ' + err.message);
  }
}

function assertAdminAccess_(adminKey) {
  const props = PropertiesService.getScriptProperties();
  const savedKey = props.getProperty('ADMIN_ACCESS_KEY');
  if (!savedKey) throw new Error('Script Properties에 ADMIN_ACCESS_KEY가 없습니다.');
  if (String(adminKey || '') !== String(savedKey)) throw new Error('관리자 인증 실패');
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// Script Properties 설정 확인용 (Apps Script에서 직접 실행)
function checkAdminAccessKeySetting() {
  const props = PropertiesService.getScriptProperties();
  const savedKey = props.getProperty('ADMIN_ACCESS_KEY');
  if (!savedKey) throw new Error('ADMIN_ACCESS_KEY가 설정되어 있지 않습니다.');
  Logger.log('ADMIN_ACCESS_KEY 설정 확인 완료');
}

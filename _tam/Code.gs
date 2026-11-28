/**
 * HỆ THỐNG MƯỢN THIẾT BỊ DẠY HỌC – TRƯỜNG THCS NGHĨA LỘC
 * Ứng dụng được phát triển bởi Thầy Đặng Cao Cường
 * Backend: Google Apps Script + Google Sheets (miễn phí)
 * Phiên bản: DEMO 0.1.0
 *
 * CÀI ĐẶT: chạy hàm setup() một lần, sau đó Triển khai > Ứng dụng web
 * (Thực thi với tư cách: Tôi; Ai có quyền truy cập: Bất kỳ ai).
 */

var VERSION = 'DEMO 0.7.2';
var TZ = 'Asia/Ho_Chi_Minh';
var SESSION_DAYS = 7;
var MAX_WRONG = 5;
var DEFAULT_PASS = '123456';

var SITES = { DC: 'Điểm chính', PH1: 'Phân hiệu 1', PH2: 'Phân hiệu 2' };
var TO_LIST = ['KHTN', 'KHXH', 'Năng khiếu'];

var COLS = {
  USERS: ['username', 'hoTen', 'to', 'diemTruong', 'vaiTro', 'passHash', 'salt', 'doiMK', 'sai', 'khoa', 'an', 'taoLuc', 'monDay', 'yeuThich', 'chuKy', 'hsXong'],
  SESS: ['token', 'username', 'hetHan'],
  CATALOG: ['id', 'ma', 'ten', 'loai', 'mon', 'khoi', 'kho', 'soLuong', 'donVi', 'diemTruong', 'ghiChu', 'an', 'capNhat'],
  WEEKS: ['id', 'username', 'hoTen', 'to', 'tuanBD', 'soTuan', 'trangThai', 'soThietBi', 'guiLuc', 'lbg'],
  ITEMS: ['id', 'weekId', 'username', 'hoTen', 'to', 'diemTruong', 'ngay', 'buoi', 'tiet', 'lop', 'mon', 'tietPPCT',
    'tenBai', 'catalogId', 'maTB', 'tenTB', 'soLuong', 'hinhThuc', 'gioDuKien', 'trangThai', 'gioGiao', 'nguoiGiao',
    'gioTra', 'nguoiNhan', 'tinhTrang', 'lyDo', 'taoLuc'],
  SETTINGS: ['key', 'value'],
  RULES: ['id', 'username', 'mon', 'khoi', 'tenTB', 'maTB', 'soLuong', 'hinhThuc', 'taoLuc'],
  PHIEU: ['id', 'weekId', 'username', 'hoTen', 'to', 'diemTruong', 'tuanBD', 'soTuan', 'gvKyLuc', 'trangThai', 'nvUsername', 'nvHoTen', 'nvKyLuc', 'lyDo', 'soThietBi']
};

var DEFAULT_SETTINGS = {
  TEN_TRUONG: 'TRƯỜNG THCS NGHĨA LỘC',
  CAP_TREN: 'UBND XÃ NGHĨA LỘC',
  GIO_SANG: '07:00',
  GIO_CHIEU: '14:00',
  PHUT_TIET: '45',
  PHUT_NGHI: '5',
  RA_CHOI_SAU_TIET: '2',
  PHUT_RA_CHOI: '5',
  GIO_MUON_SANG: '06:50',
  GIO_TRA_SANG: '11:35',
  GIO_MUON_CHIEU: '13:50',
  GIO_TRA_CHIEU: '17:35',
  NGAY_TUAN1: '2026-09-07'
};

/* ============================ CÀI ĐẶT ============================ */

function setup() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss;
  if (id) {
    ss = SpreadsheetApp.openById(id);
  } else {
    // Ưu tiên bảng tính đang chứa script (mở từ Tiện ích mở rộng > Apps Script)
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
    if (!ss) ss = SpreadsheetApp.create('MuonTB_DEMO_DuLieu');
    props.setProperty('SHEET_ID', ss.getId());
  }
  Object.keys(COLS).forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), Math.max(COLS[name].length, sh.getMaxColumns())).setNumberFormat('@');
    sh.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  var def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  // Cài đặt mặc định
  var cur = readSettings_();
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
    if (cur[k] === undefined) appendRow_('SETTINGS', { key: k, value: DEFAULT_SETTINGS[k] });
  });

  // Tài khoản quản trị + dữ liệu mẫu DEMO
  if (!findUser_('dangcaocuong')) {
    createUser_({ hoTen: 'Đặng Cao Cường', to: 'KHXH', diemTruong: 'PH1,PH2', vaiTro: 'GV,ADMIN' });
  }
  if (!findUser_('nvtbdiemchinh')) {
    createUser_({ username: 'nvtbdiemchinh', hoTen: 'NV thiết bị Điểm chính', to: '', diemTruong: 'DC', vaiTro: 'NVTB' });
    createUser_({ username: 'nvtbphanhieu1', hoTen: 'NV thiết bị Phân hiệu 1', to: '', diemTruong: 'PH1', vaiTro: 'NVTB' });
    createUser_({ username: 'nvtbphanhieu2', hoTen: 'NV thiết bị Phân hiệu 2', to: '', diemTruong: 'PH2', vaiTro: 'NVTB' });
    createUser_({ hoTen: 'Nguyễn Văn Toán', to: 'KHTN', diemTruong: 'DC,PH1', vaiTro: 'GV' });
    createUser_({ hoTen: 'Trần Thị Hóa', to: 'KHTN', diemTruong: 'PH2', vaiTro: 'GV' });
    createUser_({ hoTen: 'Lê Thị Nhạc', to: 'Năng khiếu', diemTruong: 'DC,PH1,PH2', vaiTro: 'GV' });
  }
  if (readAll_('CATALOG').length === 0) seedCatalog_();
  return 'Đã cài đặt xong. ID dữ liệu: ' + ss.getId();
}

function seedCatalog_() {
  var base = [
    ['Máy chiếu (projector)', 'Thiết bị', 'Dùng chung', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung', 2, 'Chiếc'],
    ['Tivi thông minh di động', 'Thiết bị', 'Dùng chung', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung', 1, 'Chiếc'],
    ['Loa kéo', 'Thiết bị', 'Dùng chung', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung', 1, 'Chiếc'],
    ['Phòng Tin học', 'Phòng học bộ môn', 'Tin học', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Phòng bộ môn', 1, 'Phòng'],
    ['Phòng thực hành KHTN', 'Phòng học bộ môn', 'Khoa học tự nhiên', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Phòng bộ môn', 1, 'Phòng'],
    ['Bộ thước vẽ bảng dạy và học', 'Dụng cụ', 'Toán học', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung KHTN', 3, 'Bộ'],
    ['Bộ mô hình hình khối không gian', 'Dụng cụ', 'Toán học', 'Khối 8, Khối 9', 'Kho dùng chung KHTN', 2, 'Bộ'],
    ['Kính hiển vi quang học', 'Thiết bị', 'Khoa học tự nhiên', 'Khối 6, Khối 7', 'Kho dùng chung KHTN', 6, 'Chiếc'],
    ['Bộ dụng cụ đo điện (ampe kế, vôn kế)', 'Dụng cụ', 'Khoa học tự nhiên', 'Khối 8, Khối 9', 'Kho dùng chung KHTN', 6, 'Bộ'],
    ['Đèn cồn thí nghiệm', 'Dụng cụ', 'Khoa học tự nhiên', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung KHTN', 10, 'Chiếc'],
    ['Dung dịch HCl loãng', 'Hóa chất', 'Khoa học tự nhiên', 'Khối 8, Khối 9', 'Kho dùng chung KHTN', 5, 'Lọ'],
    ['Mô hình cơ thể người', 'Thiết bị', 'Khoa học tự nhiên', 'Khối 8', 'Kho dùng chung KHTN', 1, 'Bộ'],
    ['Lược đồ Việt Nam treo tường', 'Tranh ảnh', 'Lịch sử và Địa lí', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung KHXH', 2, 'Tờ'],
    ['Quả địa cầu', 'Thiết bị', 'Lịch sử và Địa lí', 'Khối 6', 'Kho dùng chung KHXH', 4, 'Chiếc'],
    ['Bộ tranh minh họa Ngữ văn 6', 'Tranh ảnh', 'Ngữ văn', 'Khối 6', 'Kho dùng chung KHXH', 1, 'Bộ'],
    ['Bộ tranh minh họa Ngữ văn 8', 'Tranh ảnh', 'Ngữ văn', 'Khối 8', 'Kho dùng chung KHXH', 1, 'Bộ'],
    ['Bộ tranh Giáo dục công dân', 'Tranh ảnh', 'Giáo dục công dân', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho dùng chung KHXH', 1, 'Bộ'],
    ['Đàn phím điện tử (organ)', 'Thiết bị', 'Âm nhạc', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho Năng khiếu', 1, 'Chiếc'],
    ['Bộ giá vẽ', 'Dụng cụ', 'Mĩ thuật', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho Năng khiếu', 10, 'Chiếc'],
    ['Bóng đá', 'Dụng cụ', 'Giáo dục thể chất', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho Năng khiếu', 10, 'Quả'],
    ['Đồng hồ bấm giây', 'Dụng cụ', 'Giáo dục thể chất', 'Khối 6, Khối 7, Khối 8, Khối 9', 'Kho Năng khiếu', 2, 'Chiếc'],
    ['Bộ thiết bị lắp ráp robot', 'Thiết bị', 'Tin học', 'Khối 8, Khối 9', 'Kho dùng chung KHTN', 3, 'Bộ']
  ];
  var n = 1;
  Object.keys(SITES).forEach(function (site) {
    var rows = base.map(function (b) {
      var code = 'TB' + ('000' + (n++)).slice(-4);
      return { ma: code, ten: b[0], loai: b[1], mon: b[2], khoi: b[3], kho: b[4], soLuong: b[5], donVi: b[6], ghiChu: 'Dữ liệu mẫu DEMO' };
    });
    importCatalogRows_(site, rows, true);
  });
}

/* ============================ API ============================ */

function doGet() {
  return json_({ ok: true, app: 'Mượn thiết bị dạy học – THCS Nghĩa Lộc', version: VERSION });
}

function doPost(e) {
  var res;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    res = route_(req);
  } catch (err) {
    res = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return json_(res);
}

function route_(req) {
  var a = req.action;
  if (a === 'login') return login_(req.username, req.password);
  if (a === 'ping') return { ok: true, version: VERSION };

  var me = auth_(req.token);
  var R = {
    logout: function () { return logout_(req.token); },
    bootstrap: function () { return bootstrap_(me); },
    changePass: function () { return changePass_(me, req.oldPass, req.newPass); },
    // Giáo viên
    myWeek: function () { return myWeek_(me, req.tuanBD); },
    myWeeks: function () { return myWeeks_(me); },
    hints: function () { return hints_(me, req.keys); },
    saveProfile: function () { return saveProfile_(me, req.profile); },
    saveSignature: function () { return saveSignature_(me, req.dataUrl); },
    mySignature: function () { var u = findUser_(me.username); return { ok: true, dataUrl: u.chuKy || '' }; },
    phieuDetail: function () { return phieuDetail_(me, req.id); },
    phieuApprove: function () { return phieuApprove_(me, req.ids, req.op, req.lyDo); },
    phieuBundle: function () { return phieuBundle_(me, req.tuanBD, req.site); },
    saveRule: function () { need_(me, 'GV'); return saveRule_(me, req.rule); },
    deleteRule: function () { return deleteRule_(me, req.id); },
    submitWeek: function () { need_(me, 'GV'); return submitWeek_(me, req); },
    noBorrow: function () { need_(me, 'GV'); return noBorrow_(me, req); },
    updateLessons: function () { need_(me, 'GV'); return updateLessons_(me, req); },
    // NV thiết bị / Admin
    siteDay: function () { return siteDay_(me, req.site, req.ngay); },
    setItems: function () { return setItems_(me, req.ids, req.op, req.tinhTrang, req.lyDo); },
    gvReturn: function () { return gvReturn_(me, req.ids, req.tinhTrang, req.lyDo); },
    weekStatus: function () { return weekStatus_(me, req.tuanBD, req.site); },
    importCatalog: function () { return importCatalog_(me, req.site, req.rows, req.mode); },
    saveCatalogItem: function () { return saveCatalogItem_(me, req.item); },
    report: function () { return report_(me, req); },
    // Admin
    listUsers: function () { need_(me, 'ADMIN'); return { ok: true, users: listUsers_() }; },
    importUsers: function () { need_(me, 'ADMIN'); return importUsers_(req.rows); },
    saveUser: function () { need_(me, 'ADMIN'); return saveUser_(req.user); },
    resetPass: function () { need_(me, 'ADMIN'); return resetPass_(req.username); },
    saveSettings: function () { need_(me, 'ADMIN'); return saveSettings_(req.settings); }
  };
  if (!R[a]) throw new Error('Yêu cầu không hợp lệ: ' + a);
  return R[a]();
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================ SHEET HELPERS ============================ */

var _ss = null;
function ss_() {
  if (_ss) return _ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Chưa chạy setup()');
  _ss = SpreadsheetApp.openById(id);
  return _ss;
}
function sh_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh && COLS[name]) {
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, sh.getMaxRows(), COLS[name].length).setNumberFormat('@');
    sh.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function str_(v) {
  if (v === null || v === undefined) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

function readAll_(name) {
  var sh = sh_(name);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var cols = COLS[name];
  var vals = sh.getRange(2, 1, last - 1, cols.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var o = { _row: i + 2 };
    for (var j = 0; j < cols.length; j++) o[cols[j]] = str_(vals[i][j]);
    out.push(o);
  }
  return out;
}

function toRow_(name, o) {
  return COLS[name].map(function (c) { var v = o[c]; return (v === null || v === undefined) ? '' : String(v); });
}

function appendRow_(name, o) { appendRows_(name, [o]); }

function appendRows_(name, list) {
  if (!list.length) return;
  var sh = sh_(name);
  var start = sh.getLastRow() + 1;
  var rng = sh.getRange(start, 1, list.length, COLS[name].length);
  rng.setNumberFormat('@');
  rng.setValues(list.map(function (o) { return toRow_(name, o); }));
}

function updateRow_(name, o) {
  sh_(name).getRange(o._row, 1, 1, COLS[name].length).setValues([toRow_(name, o)]);
}

/** Ghi lại nhiều dòng đã sửa (gom theo khối liên tiếp). */
function updateRows_(name, list) {
  list.sort(function (a, b) { return a._row - b._row; });
  var i = 0;
  while (i < list.length) {
    var j = i;
    while (j + 1 < list.length && list[j + 1]._row === list[j]._row + 1) j++;
    var chunk = list.slice(i, j + 1);
    sh_(name).getRange(chunk[0]._row, 1, chunk.length, COLS[name].length)
      .setValues(chunk.map(function (o) { return toRow_(name, o); }));
    i = j + 1;
  }
}

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function uid_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm'); }
function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }

/* ============================ SETTINGS ============================ */

function readSettings_() {
  var o = {};
  readAll_('SETTINGS').forEach(function (r) { o[r.key] = r.value; });
  return o;
}
function settings_() {
  var o = readSettings_();
  Object.keys(DEFAULT_SETTINGS).forEach(function (k) { if (o[k] === undefined || o[k] === '') o[k] = DEFAULT_SETTINGS[k]; });
  return o;
}
function saveSettings_(s) {
  return withLock_(function () {
    var rows = readAll_('SETTINGS');
    Object.keys(DEFAULT_SETTINGS).forEach(function (k) {
      if (s[k] === undefined) return;
      var r = rows.filter(function (x) { return x.key === k; })[0];
      if (r) { r.value = String(s[k]); updateRow_('SETTINGS', r); }
      else appendRow_('SETTINGS', { key: k, value: String(s[k]) });
    });
    return { ok: true, settings: settings_() };
  });
}

/** Giờ dự kiến của một tiết, ví dụ "07:00-07:45". */
function periodTime_(st, buoi, tiet) {
  var start = (buoi === 'Chiều' ? st.GIO_CHIEU : st.GIO_SANG).split(':');
  var m = Number(start[0]) * 60 + Number(start[1] || 0);
  var len = Number(st.PHUT_TIET), gap = Number(st.PHUT_NGHI);
  var after = Number(st.RA_CHOI_SAU_TIET), rc = Number(st.PHUT_RA_CHOI);
  for (var t = 1; t < tiet; t++) m += len + (t === after ? rc : gap);
  function f(x) { return ('0' + Math.floor(x / 60)).slice(-2) + ':' + ('0' + (x % 60)).slice(-2); }
  return f(m) + '-' + f(m + len);
}

/* ============================ USERS & AUTH ============================ */

function noAccent_(s) {
  return String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd').replace(/Đ/g, 'D');
}
function makeUsername_(hoTen) {
  return noAccent_(hoTen).toLowerCase().replace(/[^a-z0-9]/g, '');
}
function hash_(pass, salt) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pass, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}
function findUser_(username) {
  username = String(username || '').toLowerCase().trim();
  return readAll_('USERS').filter(function (u) { return u.username === username && u.an !== '1'; })[0] || null;
}
function normSites_(s) {
  var out = [];
  String(s || '').split(/[,;\s]+/).forEach(function (x) {
    var k = noAccent_(x).toUpperCase().replace(/[^A-Z0-9]/g, '');
    if (!k) return;
    if (k === 'DC' || k === 'DIEMCHINH' || k === 'CHINH') k = 'DC';
    else if (k === 'PH1' || k === '1' || k === 'PHANHIEU1') k = 'PH1';
    else if (k === 'PH2' || k === '2' || k === 'PHANHIEU2') k = 'PH2';
    else return;
    if (out.indexOf(k) < 0) out.push(k);
  });
  return out.join(',');
}
function normRoles_(s) {
  var out = [];
  String(s || 'GV').toUpperCase().split(/[,;\s]+/).forEach(function (x) {
    if (x === 'GIAOVIEN') x = 'GV';
    if (x === 'NVTHIETBI' || x === 'THIETBI') x = 'NVTB';
    if (x === 'QUANTRI') x = 'ADMIN';
    if (['GV', 'NVTB', 'ADMIN'].indexOf(x) >= 0 && out.indexOf(x) < 0) out.push(x);
  });
  return out.length ? out.join(',') : 'GV';
}
function normTo_(s) {
  var k = noAccent_(s).toUpperCase().replace(/[^A-Z]/g, '');
  if (k.indexOf('KHTN') >= 0 || k.indexOf('TUNHIEN') >= 0) return 'KHTN';
  if (k.indexOf('KHXH') >= 0 || k.indexOf('XAHOI') >= 0) return 'KHXH';
  if (k.indexOf('NANGKHIEU') >= 0 || k === 'NK') return 'Năng khiếu';
  return String(s || '').trim();
}

function createUser_(u, existing) {
  var all = existing || readAll_('USERS');
  var base = u.username ? String(u.username).toLowerCase() : makeUsername_(u.hoTen);
  var name = base, n = 2;
  while (all.some(function (x) { return x.username === name; })) name = base + (n++);
  var salt = uid_();
  var o = {
    username: name, hoTen: String(u.hoTen).trim(), to: normTo_(u.to), diemTruong: normSites_(u.diemTruong),
    vaiTro: normRoles_(u.vaiTro), passHash: hash_(DEFAULT_PASS, salt), salt: salt, doiMK: '1', sai: '0', khoa: '0',
    an: '0', taoLuc: now_()
  };
  appendRow_('USERS', o);
  all.push(o);
  return o;
}

function publicUser_(u) {
  return { username: u.username, hoTen: u.hoTen, to: u.to, diemTruong: u.diemTruong, vaiTro: u.vaiTro,
    doiMK: u.doiMK === '1', khoa: u.khoa === '1', monDay: u.monDay || '', yeuThich: parseJson_(u.yeuThich, []), coKy: !!u.chuKy,
    hsXong: u.hsXong === '1' || !!u.monDay || normRoles_(u.vaiTro) !== 'GV' };
}
function parseJson_(s, d) { try { return s ? JSON.parse(s) : d; } catch (e) { return d; } }

function login_(username, password) {
  return withLock_(function () {
    var u = findUser_(username);
    if (!u) return { ok: false, error: 'Sai tên đăng nhập hoặc mật khẩu.' };
    if (u.khoa === '1') return { ok: false, error: 'Tài khoản đã bị khóa do nhập sai nhiều lần. Liên hệ quản trị để mở khóa.' };
    if (hash_(String(password || ''), u.salt) !== u.passHash) {
      u.sai = String(Number(u.sai || 0) + 1);
      if (Number(u.sai) >= MAX_WRONG) u.khoa = '1';
      updateRow_('USERS', u);
      var left = MAX_WRONG - Number(u.sai);
      return { ok: false, error: left > 0 ? 'Sai mật khẩu. Còn ' + left + ' lần thử.' : 'Tài khoản đã bị khóa.' };
    }
    u.sai = '0';
    updateRow_('USERS', u);
    var token = Utilities.getUuid().replace(/-/g, '') + uid_();
    appendRow_('SESS', { token: token, username: u.username, hetHan: String(Date.now() + SESSION_DAYS * 864e5) });
    return { ok: true, token: token, user: publicUser_(u) };
  });
}

function auth_(token) {
  if (!token) throw new Error('AUTH: Vui lòng đăng nhập.');
  var s = readAll_('SESS').filter(function (x) { return x.token === token; })[0];
  if (!s || Number(s.hetHan) < Date.now()) throw new Error('AUTH: Phiên đăng nhập đã hết hạn, vui lòng đăng nhập lại.');
  var u = findUser_(s.username);
  if (!u || u.khoa === '1') throw new Error('AUTH: Tài khoản không còn hiệu lực.');
  // Gia hạn: 7 ngày kể từ lần dùng gần nhất (chỉ ghi khi còn < 6 ngày để đỡ ghi nhiều)
  if (Number(s.hetHan) - Date.now() < (SESSION_DAYS - 1) * 864e5) {
    s.hetHan = String(Date.now() + SESSION_DAYS * 864e5);
    updateRow_('SESS', s);
  }
  u.roles = u.vaiTro.split(',');
  u.sites = u.diemTruong ? u.diemTruong.split(',') : [];
  return u;
}

function has_(me, role) { return me.roles.indexOf(role) >= 0; }
function need_(me, role) {
  if (!has_(me, role) && !has_(me, 'ADMIN')) throw new Error('Bạn không có quyền thực hiện chức năng này.');
}
function canSite_(me, site) {
  return has_(me, 'ADMIN') || (has_(me, 'NVTB') && me.sites.indexOf(site) >= 0);
}

function logout_(token) {
  return withLock_(function () {
    var s = readAll_('SESS').filter(function (x) { return x.token === token; })[0];
    if (s) { s.hetHan = '0'; updateRow_('SESS', s); }
    return { ok: true };
  });
}

function changePass_(me, oldPass, newPass) {
  if (!newPass || String(newPass).length < 6) return { ok: false, error: 'Mật khẩu mới phải có ít nhất 6 ký tự.' };
  return withLock_(function () {
    var u = findUser_(me.username);
    if (hash_(String(oldPass || ''), u.salt) !== u.passHash) return { ok: false, error: 'Mật khẩu cũ không đúng.' };
    u.salt = uid_();
    u.passHash = hash_(String(newPass), u.salt);
    u.doiMK = '0';
    updateRow_('USERS', u);
    return { ok: true };
  });
}

function listUsers_() {
  return readAll_('USERS').filter(function (u) { return u.an !== '1'; }).map(publicUser_);
}

function importUsers_(rows) {
  return withLock_(function () {
    var all = readAll_('USERS');
    var added = [], updated = 0;
    (rows || []).forEach(function (r) {
      if (!r.hoTen || !String(r.hoTen).trim()) return;
      var exist = all.filter(function (u) { return u.an !== '1' && u.hoTen.toLowerCase() === String(r.hoTen).trim().toLowerCase(); })[0];
      if (exist) {
        if (r.to) exist.to = normTo_(r.to);
        if (r.diemTruong) exist.diemTruong = normSites_(r.diemTruong);
        if (r.vaiTro) exist.vaiTro = normRoles_(r.vaiTro);
        updateRow_('USERS', exist);
        updated++;
      } else {
        var o = createUser_(r, all);
        added.push({ hoTen: o.hoTen, username: o.username });
      }
    });
    return { ok: true, added: added, updated: updated, users: listUsers_() };
  });
}

function saveUser_(x) {
  return withLock_(function () {
    if (!x.username) {
      var o = createUser_(x);
      return { ok: true, created: publicUser_(o), users: listUsers_() };
    }
    var u = findUser_(x.username);
    if (!u) return { ok: false, error: 'Không tìm thấy tài khoản.' };
    if (x.xoa) { u.an = '1'; }
    else {
      if (x.hoTen) u.hoTen = String(x.hoTen).trim();
      if (x.to !== undefined) u.to = normTo_(x.to);
      if (x.diemTruong !== undefined) u.diemTruong = normSites_(x.diemTruong);
      if (x.vaiTro !== undefined) u.vaiTro = normRoles_(x.vaiTro);
      if (x.moKhoa) { u.khoa = '0'; u.sai = '0'; }
    }
    updateRow_('USERS', u);
    return { ok: true, users: listUsers_() };
  });
}

function resetPass_(username) {
  return withLock_(function () {
    var u = findUser_(username);
    if (!u) return { ok: false, error: 'Không tìm thấy tài khoản.' };
    u.salt = uid_();
    u.passHash = hash_(DEFAULT_PASS, u.salt);
    u.doiMK = '1'; u.khoa = '0'; u.sai = '0';
    updateRow_('USERS', u);
    return { ok: true, message: 'Đã đặt lại mật khẩu ' + u.username + ' về ' + DEFAULT_PASS };
  });
}

/* ============================ BOOTSTRAP ============================ */

function catalogPublic_() {
  return readAll_('CATALOG').filter(function (c) { return c.an !== '1'; }).map(function (c) {
    return { id: c.id, ma: c.ma, ten: c.ten, loai: c.loai, mon: c.mon, khoi: c.khoi, kho: c.kho,
      soLuong: Number(c.soLuong) || 0, donVi: c.donVi, diemTruong: c.diemTruong, ghiChu: c.ghiChu };
  });
}

function bootstrap_(me) {
  autoClose_();
  var out = { ok: true, version: VERSION, me: publicUser_(me), settings: settings_(), sites: SITES, toList: TO_LIST,
    catalog: catalogPublic_(), today: today_() };
  if (has_(me, 'GV')) { out.hayDung = hayDung_(me); out.rules = myRules_(me); }
  if (has_(me, 'NVTB') || has_(me, 'ADMIN')) {
    out.teachers = listUsers_().filter(function (u) { return u.vaiTro.indexOf('GV') >= 0; })
      .map(function (u) { return { username: u.username, hoTen: u.hoTen, to: u.to, diemTruong: u.diemTruong }; });
  }
  return out;
}

/* ============================ CATALOG ============================ */

function importCatalogRows_(site, rows, replace) {
  var all = readAll_('CATALOG');
  var mine = all.filter(function (c) { return c.diemTruong === site && c.an !== '1'; });
  var byMa = {}, byTen = {};
  mine.forEach(function (c) { if (c.ma) byMa[c.ma.toUpperCase()] = c; byTen[c.ten.toLowerCase()] = c; });
  var seen = {}, toAdd = [], toUpd = [], stamp = now_();
  rows.forEach(function (r) {
    var ten = String(r.ten || '').trim();
    if (!ten) return;
    var ma = String(r.ma || '').trim();
    var ex = (ma && byMa[ma.toUpperCase()]) || (!ma && byTen[ten.toLowerCase()]) || null;
    var data = { ma: ma, ten: ten, loai: String(r.loai || '').trim(), mon: String(r.mon || '').trim(),
      khoi: String(r.khoi || '').trim(), kho: String(r.kho || '').trim(), soLuong: String(Number(r.soLuong) || 0),
      donVi: String(r.donVi || '').trim(), ghiChu: String(r.ghiChu || '').trim(), diemTruong: site, an: '0', capNhat: stamp };
    if (ex) {
      Object.keys(data).forEach(function (k) { if (k !== 'ma' || data.ma) ex[k] = data[k]; });
      seen[ex.id] = 1;
      toUpd.push(ex);
    } else {
      data.id = uid_();
      if (!data.ma) data.ma = site + '-' + data.id.slice(0, 6).toUpperCase();
      toAdd.push(data);
      byTen[ten.toLowerCase()] = data;
      if (data.ma) byMa[data.ma.toUpperCase()] = data;
    }
  });
  var hidden = 0;
  if (replace) {
    mine.forEach(function (c) {
      if (!seen[c.id] && toUpd.indexOf(c) < 0) { c.an = '1'; c.capNhat = stamp; toUpd.push(c); hidden++; }
    });
  }
  if (toUpd.length) updateRows_('CATALOG', toUpd);
  appendRows_('CATALOG', toAdd);
  return { added: toAdd.length, updated: toUpd.length - hidden, hidden: hidden };
}

function importCatalog_(me, site, rows, mode) {
  if (!SITES[site]) throw new Error('Chưa chọn điểm trường.');
  if (!canSite_(me, site)) throw new Error('Bạn chỉ được cập nhật danh mục của điểm trường mình phụ trách.');
  return withLock_(function () {
    var r = importCatalogRows_(site, rows || [], mode === 'replace');
    r.ok = true;
    r.catalog = catalogPublic_();
    return r;
  });
}

function saveCatalogItem_(me, item) {
  if (!item || !SITES[item.diemTruong]) throw new Error('Thiếu điểm trường.');
  if (!canSite_(me, item.diemTruong)) throw new Error('Bạn không phụ trách điểm trường này.');
  return withLock_(function () {
    if (item.id) {
      var c = readAll_('CATALOG').filter(function (x) { return x.id === item.id; })[0];
      if (!c) return { ok: false, error: 'Không tìm thấy thiết bị.' };
      ['ma', 'ten', 'loai', 'mon', 'khoi', 'kho', 'donVi', 'ghiChu'].forEach(function (k) { if (item[k] !== undefined) c[k] = String(item[k]); });
      if (item.soLuong !== undefined) c.soLuong = String(Number(item.soLuong) || 0);
      if (item.xoa) c.an = '1';
      c.capNhat = now_();
      updateRow_('CATALOG', c);
    } else {
      importCatalogRows_(item.diemTruong, [item], false);
    }
    return { ok: true, catalog: catalogPublic_() };
  });
}

/* ============================ GIÁO VIÊN ============================ */

function activeItem_(it) { return it.trangThai !== 'HUY' && it.trangThai !== 'TU_CHOI'; }
function slotKey_(it) { return (it.catalogId || ('TUDO:' + String(it.tenTB || '').toLowerCase())) + '|' + it.ngay + '|' + it.buoi + '|' + it.tiet; }

function weekDates_(tuanBD) {
  var p = tuanBD.split('-');
  var d0 = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12);
  var out = [];
  for (var i = 0; i < 7; i++) out.push(Utilities.formatDate(new Date(d0.getTime() + i * 864e5), TZ, 'yyyy-MM-dd'));
  return out;
}

/** Số lượng đã được người khác đặt theo từng thiết bị/tiết trong tuần. */
function usageForWeek_(tuanBD, exceptUser, items) {
  var days = weekDates_(tuanBD);
  var use = {};
  (items || readAll_('ITEMS')).forEach(function (it) {
    if (!activeItem_(it) || days.indexOf(it.ngay) < 0 || it.username === exceptUser) return;
    var k = slotKey_(it);
    use[k] = (use[k] || 0) + Number(it.soLuong || 0);
  });
  return use;
}

function myWeek_(me, tuanBD) {
  autoClose_();
  var w = readAll_('WEEKS').filter(function (x) { return x.username === me.username && x.tuanBD === tuanBD; })[0] || null;
  var allItems = readAll_('ITEMS');
  var items = w ? allItems.filter(function (it) { return it.weekId === w.id && it.trangThai !== 'HUY'; }) : [];
  var lbg = null;
  if (w && w.lbg) { try { lbg = JSON.parse(w.lbg); } catch (e) { lbg = null; } }
  if (w) delete w.lbg;
  return { ok: true, week: w, items: items.map(cleanItem_), lbg: lbg, usage: usageForWeek_(tuanBD, me.username, allItems), hayDung: hayDung_(me),
    phieu: w ? readAll_('PHIEU').filter(function (x) { return x.weekId === w.id && x.trangThai !== 'HUY'; }).map(function (x) { delete x._row; return x; }) : [] };
}

function myWeeks_(me) {
  var items = readAll_('ITEMS');
  var weeks = readAll_('WEEKS').filter(function (x) { return x.username === me.username; }).map(function (w) {
    var its = items.filter(function (it) { return it.weekId === w.id && it.trangThai !== 'HUY'; });
    return { id: w.id, tuanBD: w.tuanBD, soTuan: w.soTuan, trangThai: w.trangThai, guiLuc: w.guiLuc,
      soThietBi: its.length, daGiao: its.filter(function (i) { return i.trangThai === 'DA_GIAO' || i.trangThai === 'DA_TRA'; }).length,
      tuChoi: its.filter(function (i) { return i.trangThai === 'TU_CHOI'; }).length };
  });
  weeks.sort(function (a, b) { return a.tuanBD < b.tuanBD ? 1 : -1; });
  return { ok: true, weeks: weeks };
}

function cleanItem_(it) { var o = {}; COLS.ITEMS.forEach(function (c) { o[c] = it[c]; }); return o; }

function weekNumber_(st, tuanBD) {
  var a = weekDates_(st.NGAY_TUAN1)[0].split('-'), b = tuanBD.split('-');
  var da = new Date(Number(a[0]), Number(a[1]) - 1, Number(a[2]), 12), db = new Date(Number(b[0]), Number(b[1]) - 1, Number(b[2]), 12);
  return Math.round((db - da) / (7 * 864e5)) + 1;
}

function checkMonday_(tuanBD) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(tuanBD || ''))) throw new Error('Chưa chọn tuần.');
  var p = tuanBD.split('-');
  var d = new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]), 12);
  if (d.getDay() !== 1) throw new Error('Ngày bắt đầu tuần phải là thứ Hai.');
}

function getOrCreateWeek_(me, tuanBD, soTuan) {
  var w = readAll_('WEEKS').filter(function (x) { return x.username === me.username && x.tuanBD === tuanBD; })[0];
  if (!w) {
    w = { id: uid_(), username: me.username, hoTen: me.hoTen, to: me.to, tuanBD: tuanBD, soTuan: String(soTuan || ''),
      trangThai: '', soThietBi: '0', guiLuc: '', lbg: '' };
    appendRow_('WEEKS', w);
    w = readAll_('WEEKS').filter(function (x) { return x.id === w.id; })[0];
  }
  return w;
}

function submitWeek_(me, req) {
  checkMonday_(req.tuanBD);
  var st = settings_();
  var days = weekDates_(req.tuanBD);
  var list = req.items || [];
  if (!list.length) return { ok: false, error: 'Chưa chọn thiết bị nào. Nếu tuần này không mượn, hãy bấm "Tuần này không mượn".' };
  if (!findUser_(me.username).chuKy) return { ok: false, error: 'Bạn chưa có chữ ký. Vào ☰ → Hồ sơ của tôi để tải ảnh chữ ký lên.', needSig: true };
  return withLock_(function () {
    var catalog = {};
    readAll_('CATALOG').forEach(function (c) { catalog[c.id] = c; });
    var allItems = readAll_('ITEMS');
    var w = getOrCreateWeek_(me, req.tuanBD, req.soTuan || weekNumber_(st, req.tuanBD));
    var old = allItems.filter(function (it) { return it.weekId === w.id && it.trangThai !== 'HUY'; });
    var locked = old.filter(function (it) { return it.trangThai === 'DA_GIAO' || it.trangThai === 'DA_TRA'; });
    var lockedKeys = {};
    locked.forEach(function (it) { lockedKeys[slotKey_(it)] = it; });
    var use = usageForWeek_(req.tuanBD, me.username, allItems);
    // cộng các mục đã khóa của chính mình (vẫn chiếm chỗ)
    locked.forEach(function (it) { if (activeItem_(it)) { var k = slotKey_(it); use[k] = (use[k] || 0) + Number(it.soLuong || 0); } });

    var errors = [], toAdd = [], need = {};
    list.forEach(function (x) {
      var custom = !x.catalogId;
      var c = custom ? null : catalog[x.catalogId];
      var tenTu = String(x.tenTB || '').trim().slice(0, 150);
      if (custom && !tenTu) { errors.push('Thiếu tên thiết bị ghi tay.'); return; }
      if (!custom && (!c || c.an === '1')) { errors.push('Thiết bị không còn trong danh mục: ' + (x.tenTB || x.catalogId)); return; }
      if (custom) c = { id: '', ma: '', ten: tenTu, loai: '', soLuong: '0', donVi: '', diemTruong: SITES[x.site] ? x.site : 'DC' };
      if (days.indexOf(x.ngay) < 0) { errors.push('Ngày ' + x.ngay + ' không thuộc tuần đã chọn.'); return; }
      var it = { catalogId: c.id, tenTB: c.ten, ngay: x.ngay, buoi: x.buoi === 'Chiều' ? 'Chiều' : 'Sáng', tiet: String(Number(x.tiet) || 1) };
      var k = slotKey_(it);
      if (lockedKeys[k]) return; // đã giao/đã trả: giữ nguyên
      var sl = Math.max(1, Number(x.soLuong) || 1);
      need[k] = (need[k] || 0) + sl;
      var avail = Number(c.soLuong || 0) - (use[k] || 0);
      if (!custom && need[k] > avail) {
        errors.push(c.ten + ' – ' + fmtDay_(x.ngay) + ' ' + it.buoi.toLowerCase() + ' tiết ' + it.tiet + ': chỉ còn ' +
          Math.max(0, avail) + ' ' + (c.donVi || '') + ' (đã có người đăng ký).');
        return;
      }
      toAdd.push({
        id: uid_(), weekId: w.id, username: me.username, hoTen: me.hoTen, to: me.to, diemTruong: c.diemTruong,
        ngay: it.ngay, buoi: it.buoi, tiet: it.tiet, lop: String(x.lop || ''), mon: String(x.mon || ''),
        tietPPCT: String(x.tietPPCT || ''), tenBai: String(x.tenBai || ''), catalogId: c.id, maTB: c.ma, tenTB: c.ten,
        soLuong: String(sl), hinhThuc: x.hinhThuc || (c.loai === 'Phòng học bộ môn' ? 'Phòng thực hành' : 'Tại lớp'),
        gioDuKien: periodTime_(st, it.buoi, Number(it.tiet)), trangThai: 'CHO_GIAO', gioGiao: '', nguoiGiao: '',
        gioTra: '', nguoiNhan: '', tinhTrang: '', lyDo: '', taoLuc: now_()
      });
    });
    if (errors.length) return { ok: false, error: 'Chưa gửi được:\n• ' + errors.join('\n• '), conflicts: errors };

    var cancel = old.filter(function (it) { return it.trangThai === 'CHO_GIAO' || it.trangThai === 'TU_CHOI'; });
    cancel.forEach(function (it) { it.trangThai = 'HUY'; });
    if (cancel.length) updateRows_('ITEMS', cancel);
    appendRows_('ITEMS', toAdd);

    w.trangThai = 'DA_GUI';
    w.soTuan = String(req.soTuan || w.soTuan || weekNumber_(st, req.tuanBD));
    w.soThietBi = String(toAdd.length + locked.length);
    w.guiLuc = now_();
    w.hoTen = me.hoTen; w.to = me.to;
    if (req.lbg) w.lbg = JSON.stringify(req.lbg).slice(0, 49000);
    updateRow_('WEEKS', w);
    upsertPhieu_(me, w, toAdd.concat(locked));
    return { ok: true, message: 'Đã gửi ' + (toAdd.length + locked.length) + ' lượt mượn cho tuần ' + w.soTuan + '.', weekId: w.id };
  });
}

/** Tải lịch báo giảng sau khi đã gửi phiếu: cập nhật tên bài, tiết PPCT vào các lượt đã gửi (không phải gửi/duyệt lại) */
function updateLessons_(me, req) {
  checkMonday_(req.tuanBD);
  return withLock_(function () {
    var w = readAll_('WEEKS').filter(function (x) { return x.username === me.username && x.tuanBD === req.tuanBD; })[0];
    if (!w) return { ok: true, n: 0 };
    var map = {};
    (req.rows || []).forEach(function (r) { map[r.ngay + '|' + r.buoi + '|' + Number(r.tiet) + '|' + r.lop] = r; });
    var ch = readAll_('ITEMS').filter(function (it) {
      if (it.weekId !== w.id || it.trangThai === 'HUY') return false;
      var r = map[it.ngay + '|' + it.buoi + '|' + Number(it.tiet) + '|' + it.lop];
      if (!r) return false;
      var bai = String(r.bai || '').slice(0, 200), pp = String(r.ppct || '').slice(0, 20);
      if (it.tenBai === bai && String(it.tietPPCT) === pp) return false;
      it.tenBai = bai; it.tietPPCT = pp; if (r.mon) it.mon = String(r.mon).slice(0, 60);
      return true;
    });
    if (ch.length) updateRows_('ITEMS', ch);
    if (req.lbg) { w.lbg = JSON.stringify(req.lbg).slice(0, 49000); updateRow_('WEEKS', w); }
    return { ok: true, n: ch.length };
  });
}

function noBorrow_(me, req) {
  checkMonday_(req.tuanBD);
  var st = settings_();
  return withLock_(function () {
    var w = getOrCreateWeek_(me, req.tuanBD, req.soTuan || weekNumber_(st, req.tuanBD));
    var items = readAll_('ITEMS').filter(function (it) { return it.weekId === w.id && it.trangThai !== 'HUY'; });
    if (items.some(function (it) { return it.trangThai === 'DA_GIAO' || it.trangThai === 'DA_TRA'; })) {
      return { ok: false, error: 'Tuần này đã có thiết bị được giao, không thể chuyển sang "Không mượn".' };
    }
    items.forEach(function (it) { it.trangThai = 'HUY'; });
    if (items.length) updateRows_('ITEMS', items);
    upsertPhieu_(me, w, []);
    w.trangThai = 'KHONG_MUON';
    w.soThietBi = '0';
    if (req.lbg) w.lbg = JSON.stringify(req.lbg).slice(0, 49000);
    w.guiLuc = now_();
    w.soTuan = String(req.soTuan || w.soTuan || weekNumber_(st, req.tuanBD));
    w.hoTen = me.hoTen; w.to = me.to;
    updateRow_('WEEKS', w);
    return { ok: true, message: 'Đã báo: tuần ' + w.soTuan + ' không mượn thiết bị.' };
  });
}

function fmtDay_(ymd) { var p = String(ymd).split('-'); return p[2] + '/' + p[1]; }

/* ============================ NV THIẾT BỊ ============================ */

function siteDay_(me, site, ngay) {
  autoClose_();
  if (!canSite_(me, site)) throw new Error('Bạn không phụ trách điểm trường này.');
  var items = readAll_('ITEMS').filter(function (it) { return it.diemTruong === site && it.trangThai !== 'HUY'; });
  var day = items.filter(function (it) { return it.ngay === ngay; }).map(cleanItem_);
  var chuaTra = items.filter(function (it) { return it.trangThai === 'DA_GIAO' && it.ngay < ngay; }).map(cleanItem_);
  return { ok: true, items: day, chuaTra: chuaTra };
}

function setItems_(me, ids, op, tinhTrang, lyDo) {
  if (!ids || !ids.length) return { ok: false, error: 'Chưa chọn mục nào.' };
  return withLock_(function () {
    var all = readAll_('ITEMS');
    var map = {};
    all.forEach(function (it) { map[it.id] = it; });
    var who = me.hoTen, t = now_(), changed = [];
    ids.forEach(function (id) {
      var it = map[id];
      if (!it || it.trangThai === 'HUY') return;
      if (!canSite_(me, it.diemTruong)) throw new Error('Bạn không phụ trách điểm trường của thiết bị này.');
      if (op === 'GIAO' && it.trangThai === 'CHO_GIAO') {
        it.trangThai = 'DA_GIAO'; it.gioGiao = t; it.nguoiGiao = who;
      } else if (op === 'TRA' && (it.trangThai === 'DA_GIAO' || it.trangThai === 'CHO_GIAO')) {
        if (!it.gioGiao) { it.gioGiao = t; it.nguoiGiao = who; }
        it.trangThai = 'DA_TRA'; it.gioTra = t; it.nguoiNhan = who; it.tinhTrang = tinhTrang || 'Bình thường';
        if (lyDo) it.lyDo = lyDo;
      } else if (op === 'TU_CHOI' && it.trangThai === 'CHO_GIAO') {
        it.trangThai = 'TU_CHOI'; it.lyDo = lyDo || ''; it.nguoiGiao = who; it.gioGiao = '';
      } else if (op === 'HOAN_TAC') {
        if (it.trangThai === 'DA_TRA') { it.trangThai = 'DA_GIAO'; it.gioTra = ''; it.nguoiNhan = ''; it.tinhTrang = ''; }
        else if (it.trangThai === 'DA_GIAO' || it.trangThai === 'TU_CHOI') {
          it.trangThai = 'CHO_GIAO'; it.gioGiao = ''; it.nguoiGiao = ''; it.lyDo = '';
        } else return;
      } else return;
      changed.push(it);
    });
    if (changed.length) updateRows_('ITEMS', changed);
    return { ok: true, changed: changed.length };
  });
}

function weekStatus_(me, tuanBD, site) {
  autoClose_();
  if (!(has_(me, 'NVTB') || has_(me, 'ADMIN'))) throw new Error('Bạn không có quyền xem.');
  if (site && !canSite_(me, site)) throw new Error('Bạn không phụ trách điểm trường này.');
  var weeks = readAll_('WEEKS').filter(function (w) { return w.tuanBD === tuanBD; });
  var items = readAll_('ITEMS').filter(function (it) {
    return it.trangThai !== 'HUY' && (!site || it.diemTruong === site) && weeks.some(function (w) { return w.id === it.weekId; });
  });
  var phieus = readAll_('PHIEU').filter(function (x) { return x.tuanBD === tuanBD && x.trangThai !== 'HUY'; });
  var out = listUsers_().filter(function (u) {
    return u.vaiTro.indexOf('GV') >= 0 && (!site || u.diemTruong.split(',').indexOf(site) >= 0);
  }).map(function (u) {
    var w = weeks.filter(function (x) { return x.username === u.username; })[0];
    var n = items.filter(function (it) { return it.username === u.username; }).length;
    var tt = !w || !w.trangThai ? 'CHUA_BAO' : (w.trangThai === 'DA_GUI' && n === 0 ? 'KHONG_MUON_O_DAY' : w.trangThai);
    var ph = w ? phieus.filter(function (x) { return x.weekId === w.id && (!site || x.diemTruong === site); }) : [];
    return { username: u.username, hoTen: u.hoTen, to: u.to, trangThai: tt, soThietBi: n, guiLuc: w ? w.guiLuc : '',
      phieu: ph.map(function (x) { return { id: x.id, diemTruong: x.diemTruong, trangThai: x.trangThai, nvHoTen: x.nvHoTen, nvKyLuc: x.nvKyLuc, lyDo: x.lyDo, gvKyLuc: x.gvKyLuc }; }) };
  });
  return { ok: true, list: out };
}

/* ============================ BÁO CÁO / SỔ ============================ */

function report_(me, req) {
  autoClose_();
  var isMgr = has_(me, 'NVTB') || has_(me, 'ADMIN');
  var from = req.from || '0000-00-00', to = req.to || '9999-99-99';
  var items = readAll_('ITEMS').filter(function (it) {
    if (it.trangThai === 'HUY') return false;
    if (it.ngay < from || it.ngay > to) return false;
    if (!isMgr && it.username !== me.username) return false;
    if (isMgr && !has_(me, 'ADMIN') && me.sites.indexOf(it.diemTruong) < 0) return false;
    if (req.site && it.diemTruong !== req.site) return false;
    if (req.to2 && it.to !== req.to2) return false;
    if (req.username && it.username !== req.username) return false;
    return true;
  }).map(cleanItem_);
  items.sort(function (a, b) {
    var x = a.ngay + (a.buoi === 'Sáng' ? '1' : '2') + ('0' + a.tiet).slice(-2) + a.hoTen;
    var y = b.ngay + (b.buoi === 'Sáng' ? '1' : '2') + ('0' + b.tiet).slice(-2) + b.hoTen;
    return x < y ? -1 : x > y ? 1 : 0;
  });
  var siteOf = {};
  readAll_('USERS').forEach(function (u) { siteOf[u.username] = (u.diemTruong || '').split(','); });
  var siteFilter = req.site || (isMgr && !has_(me, 'ADMIN') ? null : '');
  var weeks = readAll_('WEEKS').filter(function (w) {
    if (!w.trangThai) return false;
    if (isMgr && !has_(me, 'ADMIN') && !(siteOf[w.username] || []).some(function (s) { return me.sites.indexOf(s) >= 0; })) return false;
    if (siteFilter && (siteOf[w.username] || []).indexOf(siteFilter) < 0) return false;
    var end = weekDates_(w.tuanBD)[6];
    if (end < from || w.tuanBD > to) return false;
    if (!isMgr && w.username !== me.username) return false;
    if (req.to2 && w.to !== req.to2) return false;
    if (req.username && w.username !== req.username) return false;
    return true;
  }).map(function (w) { return { tuanBD: w.tuanBD, soTuan: w.soTuan, username: w.username, hoTen: w.hoTen, to: w.to, trangThai: w.trangThai, guiLuc: w.guiLuc }; });
  return { ok: true, items: items, weeks: weeks };
}

/* ============================ TRẢ THIẾT BỊ ============================ */

/** Giáo viên tự báo đã trả thiết bị của mình. */
function gvReturn_(me, ids, tinhTrang, lyDo) {
  if (!ids || !ids.length) return { ok: false, error: 'Chưa chọn thiết bị.' };
  var today = today_(), st = settings_();
  return withLock_(function () {
    var all = readAll_('ITEMS'), t = now_(), changed = [];
    all.forEach(function (it) {
      if (ids.indexOf(it.id) < 0 || it.username !== me.username) return;
      if ((it.trangThai !== 'DA_GIAO' && it.trangThai !== 'CHO_GIAO') || it.ngay > today) return;
      if (!it.gioGiao) { it.gioGiao = fmtDmy_(it.ngay) + ' ' + (it.buoi === 'Chiều' ? (st.GIO_MUON_CHIEU || '13:50') : (st.GIO_MUON_SANG || '06:50')); it.nguoiGiao = 'Tự động (đầu buổi)'; }
      it.trangThai = 'DA_TRA'; it.gioTra = t; it.nguoiNhan = 'GV tự báo trả';
      it.tinhTrang = tinhTrang || 'Bình thường';
      if (lyDo) it.lyDo = lyDo;
      changed.push(it);
    });
    if (changed.length) updateRows_('ITEMS', changed);
    return { ok: true, changed: changed.length };
  });
}

function fmtDmy_(ymd) { var p = String(ymd).split('-'); return p[2] + '/' + p[1] + '/' + p[0]; }

/**
 * Tự động theo buổi (mượn đầu buổi, trả cuối buổi):
 *  - Từ giờ vào tiết 1 của buổi: thiết bị "Chờ giao" -> "Đã giao", giờ mượn = GIO_MUON_SANG/CHIEU.
 *  - Từ giờ tan buổi: thiết bị chưa trả -> "Đã trả", giờ trả = GIO_TRA_SANG/CHIEU.
 *  - Các ngày trước: mọi thiết bị chưa trả đều được ghi trả.
 * Chạy khi có người mở app (tối đa 2 phút/lần).
 */
function autoClose_() {
  var cache = CacheService.getScriptCache();
  if (cache.get('autoClose')) return;
  cache.put('autoClose', '1', 120);
  var st = settings_(), today = today_();
  var nowHM = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  var cfg = {
    'Sáng': { vao: st.GIO_SANG || '07:00', muon: st.GIO_MUON_SANG || '06:50', tra: st.GIO_TRA_SANG || '11:35' },
    'Chiều': { vao: st.GIO_CHIEU || '14:00', muon: st.GIO_MUON_CHIEU || '13:50', tra: st.GIO_TRA_CHIEU || '17:35' }
  };
  function action(it) {
    if (it.trangThai !== 'DA_GIAO' && it.trangThai !== 'CHO_GIAO') return '';
    var c = cfg[it.buoi] || cfg['Sáng'];
    if (it.ngay < today || (it.ngay === today && nowHM >= c.tra)) return 'TRA';
    if (it.ngay === today && it.trangThai === 'CHO_GIAO' && nowHM >= c.vao) return 'GIAO';
    return '';
  }
  var todo = readAll_('ITEMS').filter(function (it) { return !!action(it); });
  if (!todo.length) return;
  withLock_(function () {
    var fresh = {}, list = [];
    readAll_('ITEMS').forEach(function (it) { fresh[it.id] = it; });
    todo.forEach(function (x) {
      var it = fresh[x.id];
      if (!it) return;
      var act = action(it);
      if (!act) return;
      var c = cfg[it.buoi] || cfg['Sáng'], d = fmtDmy_(it.ngay);
      if (!it.gioGiao) { it.gioGiao = d + ' ' + c.muon; it.nguoiGiao = 'Tự động (đầu buổi)'; }
      if (act === 'GIAO') it.trangThai = 'DA_GIAO';
      else {
        it.trangThai = 'DA_TRA'; it.gioTra = d + ' ' + c.tra;
        it.nguoiNhan = 'Tự động (cuối buổi)'; it.tinhTrang = 'Bình thường (tự động)';
      }
      list.push(it);
    });
    if (list.length) updateRows_('ITEMS', list);
  });
}

/* ============================ GỢI Ý & MƯỢN CỐ ĐỊNH ============================ */

function keyN_(s) { return noAccent_(s).toLowerCase().replace(/\s+/g, ' ').trim(); }
function normBai_(b) {
  return keyN_(b).replace(/\(\s*t\s*\d+\s*\)/g, ' ').replace(/tiet\s*\d+/g, ' ').replace(/[^a-z0-9]+/g, ' ').trim();
}
function gradeOf_(lop) { var m = String(lop || '').match(/^(\d+)/); return m ? m[1] : ''; }

/** Thiết bị giáo viên hay mượn nhất (tối đa 8). */
function hayDung_(me) {
  var cnt = {};
  readAll_('ITEMS').forEach(function (it) {
    if (it.username !== me.username || it.trangThai === 'HUY' || it.trangThai === 'TU_CHOI' || !it.catalogId) return;
    var k = keyN_(it.tenTB);
    var o = cnt[k] || (cnt[k] = { ten: it.tenTB, ma: it.maTB, n: 0 });
    o.n++;
  });
  return Object.keys(cnt).map(function (k) { return cnt[k]; }).sort(function (a, b) { return b.n - a.n; }).slice(0, 8);
}

/** Gợi ý theo bài: các thiết bị đã được mượn cho cùng môn + khối + tiết PPCT (hoặc cùng tên bài). */
function hints_(me, keys) {
  var want = {};
  (keys || []).forEach(function (k) { want[k] = 1; });
  var agg = {};
  readAll_('ITEMS').forEach(function (it) {
    if (it.trangThai === 'HUY' || it.trangThai === 'TU_CHOI' || !it.catalogId) return;
    var base = keyN_(it.mon) + '|' + gradeOf_(it.lop) + '|';
    var ks = [];
    if (it.tietPPCT) ks.push(base + 'p' + String(it.tietPPCT).trim());
    var nb = normBai_(it.tenBai);
    if (nb) ks.push(base + 'b' + nb);
    ks.forEach(function (k) {
      if (!want[k]) return;
      var a = agg[k] || (agg[k] = {});
      var tk = keyN_(it.tenTB);
      var o = a[tk] || (a[tk] = { ten: it.tenTB, ma: it.maTB, n: 0, sl: {}, lessons: {} });
      o.lessons[it.username + it.ngay + it.buoi + it.tiet] = 1;
      o.sl[it.soLuong] = (o.sl[it.soLuong] || 0) + 1;
    });
  });
  var out = {};
  Object.keys(agg).forEach(function (k) {
    out[k] = Object.keys(agg[k]).map(function (t) {
      var o = agg[k][t], best = Object.keys(o.sl).sort(function (a, b) { return o.sl[b] - o.sl[a]; })[0];
      return { ten: o.ten, ma: o.ma, n: Object.keys(o.lessons).length, soLuong: Number(best) || 1 };
    }).sort(function (a, b) { return b.n - a.n; }).slice(0, 6);
  });
  return { ok: true, hints: out };
}

function myRules_(me) {
  return readAll_('RULES').filter(function (r) { return r.username === me.username; }).map(function (r) {
    return { id: r.id, mon: r.mon, khoi: r.khoi, tenTB: r.tenTB, maTB: r.maTB, soLuong: Number(r.soLuong) || 1, hinhThuc: r.hinhThuc };
  });
}

function saveRule_(me, r) {
  if (!r || !r.mon || !r.tenTB) return { ok: false, error: 'Thiếu môn hoặc thiết bị.' };
  return withLock_(function () {
    var all = readAll_('RULES').filter(function (x) { return x.username === me.username; });
    var dup = all.filter(function (x) { return keyN_(x.mon) === keyN_(r.mon) && String(x.khoi) === String(r.khoi || '') && keyN_(x.tenTB) === keyN_(r.tenTB); })[0];
    if (dup) {
      dup.soLuong = String(Math.max(1, Number(r.soLuong) || 1)); dup.hinhThuc = r.hinhThuc || dup.hinhThuc; dup.maTB = r.maTB || dup.maTB;
      updateRow_('RULES', dup);
    } else {
      appendRow_('RULES', { id: uid_(), username: me.username, mon: String(r.mon).trim(), khoi: String(r.khoi || ''), tenTB: String(r.tenTB).trim(),
        maTB: String(r.maTB || ''), soLuong: String(Math.max(1, Number(r.soLuong) || 1)), hinhThuc: r.hinhThuc || 'Tại lớp', taoLuc: now_() });
    }
    return { ok: true, rules: myRules_(me) };
  });
}

function deleteRule_(me, id) {
  return withLock_(function () {
    var sh = sh_('RULES');
    var r = readAll_('RULES').filter(function (x) { return x.id === id && x.username === me.username; })[0];
    if (r) sh.deleteRow(r._row);
    return { ok: true, rules: myRules_(me) };
  });
}

/* ============================ HỒ SƠ GIÁO VIÊN ============================ */

/** Giáo viên tự cài: tổ, điểm trường dạy, môn dạy, thiết bị của tôi. */
function saveProfile_(me, p) {
  p = p || {};
  return withLock_(function () {
    sh_('USERS').getRange(1, 1, 1, COLS.USERS.length).setValues([COLS.USERS]);
    var u = findUser_(me.username);
    var lanDau = !publicUser_(u).hsXong;
    // Thiết lập lần đầu: người dùng tự chọn vai trò Giáo viên hoặc NV thiết bị (chỉ chọn được một lần)
    if (lanDau && (p.vaiTro === 'GV' || p.vaiTro === 'NVTB')) {
      if (p.vaiTro === 'NVTB') {
        var one = normSites_(p.diemTruong);
        if (!one || one.indexOf(',') >= 0) return { ok: false, error: 'NV thiết bị chọn đúng một điểm trường phụ trách.' };
        u.vaiTro = 'NVTB'; u.diemTruong = one; u.to = ''; u.hsXong = '1';
        updateRow_('USERS', u);
        return { ok: true, me: publicUser_(u) };
      }
      if (!normTo_(p.to)) return { ok: false, error: 'Chọn tổ chuyên môn.' };
    }
    if (p.to !== undefined) u.to = normTo_(p.to);
    if (p.diemTruong !== undefined) {
      var ds = normSites_(p.diemTruong);
      if (!ds) return { ok: false, error: 'Chọn ít nhất một điểm trường.' };
      // NV thiết bị không được tự đổi điểm trường phụ trách
      if (u.vaiTro.indexOf('NVTB') < 0 || u.vaiTro.indexOf('ADMIN') >= 0) u.diemTruong = ds;
    }
    u.hsXong = '1';
    if (p.monDay !== undefined) u.monDay = String(p.monDay).split(',').map(function (x) { return x.trim(); }).filter(Boolean).slice(0, 10).join(', ');
    if (p.yeuThich !== undefined) {
      var list = (p.yeuThich || []).slice(0, 40).map(function (x) { return { ten: String(x.ten || '').slice(0, 150), ma: String(x.ma || '').slice(0, 40) }; })
        .filter(function (x) { return x.ten; });
      u.yeuThich = JSON.stringify(list);
    }
    updateRow_('USERS', u);
    return { ok: true, me: publicUser_(u) };
  });
}

/* ============================ KÝ PHIẾU ============================ */

function saveSignature_(me, dataUrl) {
  dataUrl = String(dataUrl || '');
  if (dataUrl && !/^data:image\/png;base64,/.test(dataUrl)) return { ok: false, error: 'Ảnh chữ ký không hợp lệ.' };
  if (dataUrl.length > 45000) return { ok: false, error: 'Ảnh chữ ký quá lớn, hãy chụp gần hơn hoặc cắt sát chữ ký.' };
  return withLock_(function () {
    sh_('USERS').getRange(1, 1, 1, COLS.USERS.length).setValues([COLS.USERS]);
    var u = findUser_(me.username);
    u.chuKy = dataUrl;
    updateRow_('USERS', u);
    return { ok: true, me: publicUser_(u) };
  });
}

/** Tạo/cập nhật phiếu (mỗi giáo viên × tuần × điểm trường), GV đã ký, chờ NV duyệt. */
function upsertPhieu_(me, w, items) {
  var sites = {};
  items.forEach(function (it) { if (it.trangThai !== 'HUY') sites[it.diemTruong] = (sites[it.diemTruong] || 0) + 1; });
  var all = readAll_('PHIEU').filter(function (x) { return x.weekId === w.id; });
  var t = now_(), upd = [], add = [];
  all.forEach(function (x) {
    if (sites[x.diemTruong]) {
      x.gvKyLuc = t; x.trangThai = 'CHO_DUYET'; x.nvUsername = ''; x.nvHoTen = ''; x.nvKyLuc = ''; x.lyDo = '';
      x.soThietBi = String(sites[x.diemTruong]); x.hoTen = me.hoTen; x.to = me.to; x.soTuan = w.soTuan;
      delete sites[x.diemTruong];
    } else if (x.trangThai !== 'HUY') x.trangThai = 'HUY';
    upd.push(x);
  });
  Object.keys(sites).forEach(function (site) {
    add.push({ id: uid_(), weekId: w.id, username: me.username, hoTen: me.hoTen, to: me.to, diemTruong: site, tuanBD: w.tuanBD,
      soTuan: w.soTuan, gvKyLuc: t, trangThai: 'CHO_DUYET', nvUsername: '', nvHoTen: '', nvKyLuc: '', lyDo: '', soThietBi: String(sites[site]) });
  });
  if (upd.length) updateRows_('PHIEU', upd);
  appendRows_('PHIEU', add);
}

function sigOf_(username) {
  if (!username) return '';
  var u = readAll_('USERS').filter(function (x) { return x.username === username; })[0];
  return u ? (u.chuKy || '') : '';
}

/** Chi tiết phiếu kèm chữ ký hai bên (GV chủ phiếu, NV phụ trách điểm, quản trị). */
function phieuDetail_(me, id) {
  var p = readAll_('PHIEU').filter(function (x) { return x.id === id; })[0];
  if (!p) return { ok: false, error: 'Không tìm thấy phiếu.' };
  if (p.username !== me.username && !canSite_(me, p.diemTruong)) throw new Error('Bạn không có quyền xem phiếu này.');
  var items = readAll_('ITEMS').filter(function (it) { return it.weekId === p.weekId && it.diemTruong === p.diemTruong && it.trangThai !== 'HUY'; }).map(cleanItem_);
  delete p._row;
  return { ok: true, phieu: p, items: items, sigGV: sigOf_(p.username), sigNV: p.trangThai === 'CHO_DUYET' ? '' : sigOf_(p.nvUsername) };
}

/** NV thiết bị đồng ý & ký, hoặc từ chối (kèm lý do) một hay nhiều phiếu. */
function phieuApprove_(me, ids, op, lyDo) {
  if (!ids || !ids.length) return { ok: false, error: 'Chưa chọn phiếu.' };
  if (op === 'TU_CHOI' && !String(lyDo || '').trim()) return { ok: false, error: 'Nhập lý do từ chối.' };
  return withLock_(function () {
    var me2 = findUser_(me.username);
    if (op === 'DUYET' && !me2.chuKy) return { ok: false, error: 'Bạn chưa có chữ ký. Vào ☰ → Hồ sơ / chữ ký để tải ảnh chữ ký lên trước.' };
    var t = now_(), phs = readAll_('PHIEU'), its = readAll_('ITEMS'), updP = [], updI = [];
    phs.forEach(function (p) {
      if (ids.indexOf(p.id) < 0 || p.trangThai === 'HUY') return;
      if (!canSite_(me, p.diemTruong)) throw new Error('Bạn không phụ trách điểm trường của phiếu này.');
      p.trangThai = op === 'DUYET' ? 'DA_DUYET' : (op === 'TU_CHOI' ? 'TU_CHOI' : 'CHO_DUYET');
      p.nvUsername = op === 'HOAN_TAC' ? '' : me.username; p.nvHoTen = op === 'HOAN_TAC' ? '' : me.hoTen;
      p.nvKyLuc = op === 'HOAN_TAC' ? '' : t; p.lyDo = op === 'TU_CHOI' ? String(lyDo).trim() : '';
      updP.push(p);
      its.forEach(function (it) {
        if (it.weekId !== p.weekId || it.diemTruong !== p.diemTruong) return;
        if (op === 'TU_CHOI' && it.trangThai === 'CHO_GIAO') { it.trangThai = 'TU_CHOI'; it.lyDo = p.lyDo; it.nguoiGiao = me.hoTen; updI.push(it); }
        if (op === 'HOAN_TAC' && it.trangThai === 'TU_CHOI') { it.trangThai = 'CHO_GIAO'; it.lyDo = ''; it.nguoiGiao = ''; updI.push(it); }
      });
    });
    if (updP.length) updateRows_('PHIEU', updP);
    if (updI.length) updateRows_('ITEMS', updI);
    return { ok: true, changed: updP.length };
  });
}

/** Tất cả phiếu của một tuần ở một điểm trường, kèm chữ ký (để NV xuất Excel một lần). */
function phieuBundle_(me, tuanBD, site) {
  if (!canSite_(me, site)) throw new Error('Bạn không phụ trách điểm trường này.');
  var phs = readAll_('PHIEU').filter(function (x) { return x.tuanBD === tuanBD && x.diemTruong === site && x.trangThai !== 'HUY'; });
  var need = {};
  phs.forEach(function (p) { need[p.username] = 1; if (p.nvUsername && p.trangThai !== 'CHO_DUYET') need[p.nvUsername] = 1; });
  var sigs = {};
  readAll_('USERS').forEach(function (u) { if (need[u.username] && u.chuKy) sigs[u.username] = u.chuKy; });
  return { ok: true, phieu: phs.map(function (p) { delete p._row; return p; }), sigs: sigs };
}

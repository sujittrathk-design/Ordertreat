/**
 * Order & Loading Dashboard — Google Apps Script backend
 * ฐานข้อมูล = Google Sheets (dataWeb)
 */
const SS_ID = '1aXvKQMUUmIVttUnb_06qLhNVMikr_Fbfy-mnuazAO6Y';
const TZ = 'Asia/Bangkok';

// ชื่อแท็บในชีต (แก้ให้ตรงกับไฟล์จริงได้ที่นี่)
const SHEETS = {
  order: 'Ordertreat',
  load: 'Export',            // แท็บข้อมูลการโหลด (ถ้าชื่อไม่ตรง แก้ตรงนี้)
  notes: 'Ordertreat_Notes'  // สร้างอัตโนมัติ เก็บ inspec / Remark
};

// หัวคอลัมน์ในแท็บ Export (ต้องตรงกับชีตจริงทุกตัวอักษร)
const COLS = {
  date: 'Date',
  ton: 'จำนวนตัน',
  box: 'จำนวนตู้',
  wait: 'จำนวนที่รอโหลด',
  pct: '%',
  plan: 'จำนวนที่รอโหลดจากวางแผน'
};
const WH_COLS = ['ระยอง', 'WH32', 'ศรีราชา']; // คอลัมน์คลัง

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Order & Loading Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ---------- helpers ---------- */
const book_ = () => SpreadsheetApp.openById(SS_ID);
const num_ = v => parseFloat(String(v).replace(/,/g, '')) || 0;
const iso_ = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');

// รองรับ Date จริง, yyyy-mm-dd, และข้อความ dd/mm/yyyy หรือ dd/mm
function toDate_(v) {
  if (v instanceof Date) return v;
  const s = String(v || '').trim();
  const iso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  const m = s.match(/^(\d{1,2})\/(\d{1,2})(?:\/(\d{4}))?$/);
  if (!m) return null;
  return new Date(+(m[3] || new Date().getFullYear()), +m[2] - 1, +m[1]);
}

function table_(name) {
  const sh = book_().getSheetByName(name);
  if (!sh) throw new Error('ไม่พบแท็บ "' + name + '"');
  const v = sh.getDataRange().getValues();
  return { head: v[0].map(s => String(s).trim()), rows: v.slice(1) };
}

/* ---------- หน้า 1: Ordertreat ---------- */
function getOrderData() {
  const { head, rows } = table_(SHEETS.order);
  const c = n => head.indexOf(n);
  const iNet = c('Net order wgh'), iRem = c('ค้างผลิต'), iMon = c('เดือน'),
        iLoad = c('Loading Date'), iPack = c('วันที่แพ็ค');
  if ([iNet, iRem, iMon, iLoad].includes(-1)) throw new Error('หัวคอลัมน์ในแท็บ Ordertreat ไม่ครบ');

  const MON = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const now = new Date();
  const months = [0, 1].map(k => {
    const d = new Date(now.getFullYear(), now.getMonth() + k, 1);
    return { key: MON[d.getMonth()], year: d.getFullYear(), month: d.getMonth(), order: 0, remain: 0 };
  });

  const groups = {};
  const packMax = {}; // วันที่แพ็คช้าสุด แยกตามวันโหลด (นับทุกแถว)

  rows.forEach(r => {
    if (!r[iNet] && !r[iRem]) return;
    const net = num_(r[iNet]) / 1000, rem = num_(r[iRem]);
    const m = months.find(x => x.key === String(r[iMon]).trim());
    if (m) { m.order += net; m.remain += rem; }

    const ld = toDate_(r[iLoad]);
    const key = ld ? iso_(ld) : String(r[iLoad] || '-');

    const pk = iPack > -1 ? toDate_(r[iPack]) : null;
    if (pk) {
      const p = iso_(pk);
      if (p > (packMax[key] || '')) packMax[key] = p;
    }

    if (rem > 0) {
      const g = groups[key] || (groups[key] = { load: key, remain: 0, pack: '' });
      g.remain += rem;
    }
  });

  Object.values(groups).forEach(g => { g.pack = packMax[g.load] || ''; });

  months.forEach(m => {
    m.order = +m.order.toFixed(2);
    m.remain = +m.remain.toFixed(2);
    m.done = +(m.order - m.remain).toFixed(2);
    m.pct = m.order ? Math.round(m.done / m.order * 100) : 0;
  });

  const list = Object.values(groups).sort((a, b) => a.load.localeCompare(b.load));
  list.forEach(g => g.remain = +g.remain.toFixed(2));
  return { months: months, groups: list, notes: getNotes_() };
}
/* ---------- inspec / Remark ---------- */

function clearNotes() {
  const sh = notesSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  return true;
}
/* ---------- inspec / Remark ---------- */
const key_ = v => v instanceof Date ? iso_(v) : String(v);

function notesSheet_() {
  const bk = book_();
  let sh = bk.getSheetByName(SHEETS.notes);
  if (!sh) {
    sh = bk.insertSheet(SHEETS.notes);
    sh.appendRow(['loading_date', 'inspec', 'remark', 'updated_at']);
    sh.setFrozenRows(1);
  }
  sh.getRange('A:A').setNumberFormat('@'); // เก็บวันที่เป็นข้อความ ไม่ให้ Sheets แปลงเอง
  return sh;
}

function getNotes_() {
  const v = notesSheet_().getDataRange().getValues().slice(1), out = {};
  v.forEach(r => { if (r[0] !== '') out[key_(r[0])] = { inspec: r[1], remark: r[2] }; });
  return out;
}

// รับเฉพาะแถวที่ผู้ใช้แก้ แล้ว "อัปเดตเฉพาะแถวนั้น" ไม่ไปแตะของคนอื่น
function saveNotes(items) { // items = [{date, inspec, remark}]
  const sh = notesSheet_(), lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const last = sh.getLastRow();
    const data = last > 1 ? sh.getRange(2, 1, last - 1, 4).getValues() : [];
    const pos = {};
    data.forEach((r, i) => pos[key_(r[0])] = i);
    items.forEach(it => {
      const row = [it.date, it.inspec, it.remark, new Date()];
      if (it.date in pos) data[pos[it.date]] = row;
      else { pos[it.date] = data.length; data.push(row); }
    });
    const out = data.filter(r => r[1] || r[2]); // แถวที่ว่างทั้งคู่จะถูกลบ
    if (last > 1) sh.getRange(2, 1, last - 1, 4).clearContent();
    if (out.length) sh.getRange(2, 1, out.length, 4).setValues(out);
    return items.length;
  } finally { lock.releaseLock(); }
}
/* ---------- หน้า 2: ข้อมูลการโหลดสินค้า ---------- */
function pct_(v) {
  if (v === '' || v === null || v === undefined) return null;
  if (typeof v === 'number') return v <= 1 ? v * 100 : v;
  const n = parseFloat(String(v).replace('%', ''));
  return isNaN(n) ? null : n;
}

function getLoadData() {
  const { head, rows } = table_(SHEETS.load);
  const ix = n => head.indexOf(n);
  const iD = ix(COLS.date), iT = ix(COLS.ton), iB = ix(COLS.box),
        iQ = ix(COLS.wait), iP = ix(COLS.plan), iPct = ix(COLS.pct);
  if (iD < 0) throw new Error('ไม่พบคอลัมน์ "' + COLS.date + '" ในแท็บ ' + SHEETS.load + ' (หัวคอลัมน์ที่เจอ: ' + head.join(', ') + ')');

  const whIdx = WH_COLS.map(n => [n, ix(n)]).filter(x => x[1] > -1);
  const days = [], wh = {};

  rows.forEach(r => {
    const d = toDate_(r[iD]);
    if (!d) return;
    const k = iso_(d);
    days.push({
      date: k,
      ton: +num_(iT > -1 ? r[iT] : 0).toFixed(1),
      box: num_(iB > -1 ? r[iB] : 0),
      wait: num_(iQ > -1 ? r[iQ] : 0),
      plan: num_(iP > -1 ? r[iP] : 0),
      pct: iPct > -1 ? pct_(r[iPct]) : null
    });
    wh[k] = {};
    whIdx.forEach(([n, i]) => { wh[k]['คลัง' + n] = num_(r[i]); });
  });
  return { days: days, wh: wh };
}
//add
/* ---------- อาหารรอโหลด ---------- */
const WAIT_SHEET = 'รอโหลด'; // ชื่อแท็บ (ถ้าไม่ตรง แก้ตรงนี้)

// ชื่อหัวคอลัมน์ที่เป็นไปได้ของแต่ละช่อง (ใส่เพิ่มได้ ไม่สนตัวพิมพ์เล็ก/ใหญ่)
const WAIT_COLS = {
  date:     ['วันที่โหลด', 'ETL', 'Loading Date'],
  country:  ['ประเทศ', 'Country'],
  customer: ['ชื่อลูกค้า', 'Customer'],
  pcg:      ['PCG', 'LOT'],
  sku:      ['SKU'],
  order:    ['ปริมาณ Order', 'ปริมาณ order', 'Order', 'ปริมาณ Order'],
  remain:   ['ปริมาณรอผลิต', 'ค้างผลิต', 'Remaining'],
  dept:     ['หน่วยงาน'],
  remark:   ['Remark']
};

function pick_(head, names) {
  for (const n of names) {
    const i = head.findIndex(h => h.toLowerCase() === n.toLowerCase());
    if (i > -1) return i;
  }
  return -1;
}

function getWaitData() {
  const { head, rows } = table_(WAIT_SHEET);
  const ix = {};
  Object.keys(WAIT_COLS).forEach(k => ix[k] = pick_(head, WAIT_COLS[k]));
  if (ix.date < 0 && ix.sku < 0) {
    throw new Error('ไม่พบคอลัมน์วันที่โหลด/SKU ในแท็บ "' + WAIT_SHEET + '" (หัวคอลัมน์ที่เจอ: ' + head.join(', ') + ')');
  }
  const get = (r, k) => ix[k] > -1 ? r[ix[k]] : '';

  const out = rows
    .filter(r => get(r, 'sku') !== '' || get(r, 'pcg') !== '')
    .map(r => {
      const d = toDate_(get(r, 'date'));
      return {
        date: d ? iso_(d) : String(get(r, 'date') || ''),
        country: String(get(r, 'country')),
        customer: String(get(r, 'customer')),
        pcg: String(get(r, 'pcg')),
        sku: String(get(r, 'sku')),
        order: num_(get(r, 'order')),
        remain: num_(get(r, 'remain')),
        dept: String(get(r, 'dept')),
        remark: String(get(r, 'remark'))
      };
    });
  out.sort((a, b) => b.date.localeCompare(a.date));
  return out;
}

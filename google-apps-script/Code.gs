/**
 * Wes – Influencer Takip
 * Google E-Tablo arka ucu. Bu kodu tablonun Uzantılar > Apps Script bölümüne yapıştırın.
 *
 * - Kayıtlar "Influencerlar" sayfasında, her influencer bir satır.
 * - Hiçbir satır silinmez: sayfa sadece ekler, günceller veya "Arşivde" işaretler.
 * - Her değişiklikten önceki hâl "Gecmis" sayfasına yazılır.
 * - Ekip şifresi "Ayarlar" sayfasının B1 hücresindedir; oradan değiştirilebilir.
 */

const SHEET_NAME = 'Influencerlar';
const HIST_NAME = 'Gecmis';
const SET_NAME = 'Ayarlar';
const DEFAULT_PASSWORD = 'atewes4465';

const FIELDS = ['id', 'ad', 'soyad', 'telefon', 'instagram', 'adres', 'boy', 'kilo', 'beden', 'urunler',
  'kargo', 'firma', 'takip', 'video', 'duzenli', 'notlar', 'arsiv', 'createdAt', 'updatedAt'];
const HEADERS = ['ID', 'Ad', 'Soyad', 'Telefon', 'Instagram', 'Adres', 'Boy (cm)', 'Kilo (kg)', 'Beden', 'Gönderilen ürünler',
  'Kargo durumu', 'Kargo firması', 'Takip no', 'Video durumu', 'Düzenli çalışma', 'Notlar', 'Arşivde', 'Eklenme', 'Son güncelleme'];
const EDITABLE = FIELDS.filter(f => !['id', 'createdAt', 'updatedAt'].includes(f));

/** Tabloyu ilk kez hazırlar. Apps Script'te bir kez "Çalıştır" ile çalıştırın. */
function kurulum() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME, 0);
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#dfeae5');
  sh.setFrozenRows(1);
  // Telefonun başındaki 0 kaybolmasın diye metin olarak sakla; tarihler tarih biçiminde.
  sh.getRange(2, 1, sh.getMaxRows() - 1, 17).setNumberFormat('@');
  sh.getRange(2, 18, sh.getMaxRows() - 1, 2).setNumberFormat('dd.mm.yyyy hh:mm');
  sh.setColumnWidth(1, 90);
  [6, 10, 16].forEach(c => sh.setColumnWidth(c, 260));

  let hist = ss.getSheetByName(HIST_NAME);
  if (!hist) hist = ss.insertSheet(HIST_NAME);
  hist.getRange(1, 1, 1, 4).setValues([['Zaman', 'Kayıt ID', 'İşlem', 'Değişiklikten önceki hâli']]).setFontWeight('bold');
  hist.setFrozenRows(1);

  let set = ss.getSheetByName(SET_NAME);
  if (!set) set = ss.insertSheet(SET_NAME);
  if (!set.getRange('B1').getValue()) {
    set.getRange('A1:B1').setValues([['Ekip şifresi', DEFAULT_PASSWORD]]);
    set.getRange('A3').setValue('Şifreyi değiştirmek için B1 hücresini düzenleyin. Ekip yeni şifreyle tekrar giriş yapar.');
  }
  const def = ss.getSheetByName('Sayfa1') || ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(def);
}

function password_() {
  const set = SpreadsheetApp.getActive().getSheetByName(SET_NAME);
  return String((set && set.getRange('B1').getValue()) || DEFAULT_PASSWORD).trim();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function rowToObj_(row) {
  const o = {};
  FIELDS.forEach((f, i) => {
    let v = row[i];
    if (v instanceof Date) v = v.getTime();
    if (f === 'arsiv') v = v === true || v === 'TRUE' || v === 'Evet';
    o[f] = v === undefined || v === null ? '' : v;
  });
  return o;
}

function readAll_(sh) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, FIELDS.length).getValues().map(rowToObj_).filter(o => o.id);
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (String(p.key || '').trim() !== password_()) return json_({ ok: false, error: 'sifre' });
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) return json_({ ok: false, error: 'kurulum' });
  return json_({ ok: true, items: readAll_(sh) });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'istek' }); }
  if (String(body.key || '').trim() !== password_()) return json_({ ok: false, error: 'sifre' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const sh = ss.getSheetByName(SHEET_NAME);
    const hist = ss.getSheetByName(HIST_NAME);
    if (!sh) return json_({ ok: false, error: 'kurulum' });
    const item = body.item || {};
    const now = new Date();

    if (body.action !== 'save') return json_({ ok: false, error: 'istek' });

    if (item.id) {
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
      const idx = ids.indexOf(String(item.id));
      if (idx === -1) return json_({ ok: false, error: 'bulunamadi' });
      const rowNum = idx + 2;
      const range = sh.getRange(rowNum, 1, 1, FIELDS.length);
      const old = range.getValues()[0];
      const oldObj = rowToObj_(old);
      hist.appendRow([now, item.id, body.islem || 'duzenlendi', JSON.stringify(oldObj)]);
      const next = old.slice();
      EDITABLE.forEach(f => {
        if (Object.prototype.hasOwnProperty.call(item, f)) {
          const i = FIELDS.indexOf(f);
          next[i] = f === 'arsiv' ? (item[f] ? 'Evet' : '') : String(item[f] == null ? '' : item[f]);
        }
      });
      next[FIELDS.indexOf('updatedAt')] = now;
      range.setValues([next]);
      return json_({ ok: true, item: rowToObj_(next) });
    }

    const row = FIELDS.map(f => {
      if (f === 'id') return Utilities.getUuid().slice(0, 8);
      if (f === 'createdAt' || f === 'updatedAt') return now;
      if (f === 'arsiv') return '';
      return String(item[f] == null ? '' : item[f]);
    });
    const r = sh.getLastRow() + 1;
    if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
    sh.getRange(r, 1, 1, 17).setNumberFormat('@');
    sh.getRange(r, 18, 1, 2).setNumberFormat('dd.mm.yyyy hh:mm');
    sh.getRange(r, 1, 1, FIELDS.length).setValues([row]);
    hist.appendRow([now, row[0], 'eklendi', '']);
    return json_({ ok: true, item: rowToObj_(row) });
  } finally {
    lock.releaseLock();
  }
}

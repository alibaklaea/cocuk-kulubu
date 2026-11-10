/**
 * Çocuk Kulübü Uygulaması · Google Drive arka ucu (Google Apps Script)
 * Veriler bu betiğin bağlı olduğu Google E-Tablo'da durur. Her okul yalnızca kendi satırlarını görür.
 * Kayıt kodu: "Ayarlar" sayfasındaki B1 hücresi (okullara bu kodu siz verirsiniz).
 * Dekontlar Google Drive'ın yazı okuma (OCR) özelliğiyle okunur; dosya okunduktan hemen sonra silinir, yalnızca bilgiler kalır.
 */

var SAYFALAR = {
  Ayarlar: ['Ayar', 'Değer'],
  Okullar: ['okulId', 'okulAdi', 'kullanici', 'sifreOzet', 'tuz', 'kayitZamani'],
  Oturumlar: ['token', 'okulId', 'sonKullanim'],
  Ogrenciler: ['okulId', 'ogrId', 'adSoyad', 'sinif', 'sube', 'veliTel', 'dersler', 'ucret', 'durum', 'dekontNo', 'odemeTarihi', 'tutar', 'not'],
  Dekontlar: ['okulId', 'dekId', 'dekontNo', 'tarih', 'tutar', 'gonderen', 'aciklama', 'ogrIdler', 'durum', 'adaylar', 'dosyaAdi', 'yuklenme'],
  Veri: ['okulId', 'anahtar', 'json', 'guncelleme']
};
var OTURUM_GUN = 30;

// ---------- kurulum (bir kez Çalıştır) ----------
function kurulum() {
  var ss = tablo_();
  Object.keys(SAYFALAR).forEach(function (ad) {
    var sh = ss.getSheetByName(ad) || ss.insertSheet(ad);
    if (sh.getLastRow() === 0) { sh.appendRow(SAYFALAR[ad]); sh.setFrozenRows(1); sh.getRange(1, 1, 1, SAYFALAR[ad].length).setFontWeight('bold'); }
    if (ad !== 'Ayarlar') sh.getRange(1, 1, sh.getMaxRows(), SAYFALAR[ad].length).setNumberFormat('@'); // telefonun baştaki 0'ı ve numaralar metin kalsın
  });
  var ay = ss.getSheetByName('Ayarlar');
  if (!ay.getRange('B2').getValue()) {
    ay.getRange('A2:B2').setValues([['Kayıt kodu (okullara siz verin)', 'KULUP-' + Math.floor(1000 + Math.random() * 9000)]]);
  }
  var s1 = ss.getSheetByName('Sayfa1') || ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  Logger.log('Kurulum tamam. Kayıt kodu: ' + ay.getRange('B2').getValue());
}

function tablo_() {
  var id = PropertiesService.getScriptProperties().getProperty('TABLO_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) { ss = SpreadsheetApp.create('Çocuk Kulübü Uygulaması · Veriler'); }
  PropertiesService.getScriptProperties().setProperty('TABLO_ID', ss.getId());
  return ss;
}
function sayfa_(ad) { return tablo_().getSheetByName(ad); }

// ---------- web uç noktası ----------
function doGet() { return json_({ ok: true, ad: 'Çocuk Kulübü API' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var g = JSON.parse(e.postData.contents || '{}');
    var islem = ISLEMLER[g.islem];
    if (!islem) throw hata_('Bilinmeyen işlem.');
    var okul = null;
    if (islem.oturum !== false) okul = oturumOkul_(g.token);
    return json_({ ok: true, veri: islem.f(g, okul) });
  } catch (err) {
    return json_({ ok: false, hata: err && err.kullanici ? err.message : 'Sunucu hatası: ' + err });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hata_(m) { var e = new Error(m); e.kullanici = true; return e; }

// ---------- yardımcılar ----------
function satirlar_(ad) {
  var sh = sayfa_(ad), v = sh.getDataRange().getValues(), b = v[0];
  return v.slice(1).map(function (r, i) { var o = { _satir: i + 2 }; b.forEach(function (k, j) { o[k] = r[j]; }); return o; });
}
function ekle_(ad, o) { sayfa_(ad).appendRow(SAYFALAR[ad].map(function (k) { return o[k] === undefined ? '' : o[k]; })); }
function yaz_(ad, o) { sayfa_(ad).getRange(o._satir, 1, 1, SAYFALAR[ad].length).setValues([SAYFALAR[ad].map(function (k) { return o[k] === undefined ? '' : o[k]; })]); }
function sil_(ad, satirNolar) { var sh = sayfa_(ad); satirNolar.sort(function (a, b) { return b - a; }).forEach(function (n) { sh.deleteRow(n); }); }
function id_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function ozet_(sifre, tuz) {
  var h = sifre + '|' + tuz;
  for (var i = 0; i < 2000; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + tuz, Utilities.Charset.UTF_8));
  return h;
}
function kayitKodu_() { return String(sayfa_('Ayarlar').getRange('B2').getValue() || '').trim(); }
function kucuk_(s) { return String(s || '').trim().toLocaleLowerCase('tr'); }

function oturumOkul_(token) {
  if (!token) throw hata_('Oturum açın.');
  var o = satirlar_('Oturumlar').filter(function (x) { return x.token === token; })[0];
  if (!o || (new Date() - new Date(o.sonKullanim)) > OTURUM_GUN * 864e5) throw hata_('Oturum süresi doldu, yeniden giriş yapın.');
  o.sonKullanim = new Date().toISOString(); yaz_('Oturumlar', o);
  var okul = satirlar_('Okullar').filter(function (x) { return x.okulId === o.okulId; })[0];
  if (!okul) throw hata_('Okul bulunamadı.');
  return okul;
}
function oturumAc_(okul) {
  var t = Utilities.getUuid() + Utilities.getUuid();
  ekle_('Oturumlar', { token: t, okulId: okul.okulId, sonKullanim: new Date().toISOString() });
  return { token: t, okul: { okulId: okul.okulId, okulAdi: okul.okulAdi, kullanici: okul.kullanici } };
}

// ---------- işlemler ----------
var ISLEMLER = {
  kayit: { oturum: false, f: function (g) {
    if (!kayitKodu_() || String(g.kayitKodu || '').trim() !== kayitKodu_()) throw hata_('Kayıt kodu hatalı. Kodu kulüp koordinatöründen alın.');
    var ad = String(g.okulAdi || '').trim(), k = kucuk_(g.kullanici), s = String(g.sifre || '');
    if (ad.length < 3) throw hata_('Okul adını yazın.');
    if (!/^[a-z0-9._-]{3,30}$/.test(k)) throw hata_('Kullanıcı adı 3-30 karakter olmalı; yalnız küçük harf (Türkçe karakter olmadan), rakam, nokta, tire.');
    if (s.length < 6) throw hata_('Şifre en az 6 karakter olmalı.');
    if (satirlar_('Okullar').some(function (x) { return x.kullanici === k; })) throw hata_('Bu kullanıcı adı alınmış.');
    var tuz = id_() + id_(), okul = { okulId: id_(), okulAdi: ad, kullanici: k, sifreOzet: ozet_(s, tuz), tuz: tuz, kayitZamani: new Date().toISOString() };
    ekle_('Okullar', okul);
    return oturumAc_(okul);
  } },
  giris: { oturum: false, f: function (g) {
    var k = kucuk_(g.kullanici), okul = satirlar_('Okullar').filter(function (x) { return x.kullanici === k; })[0];
    if (!okul || ozet_(String(g.sifre || ''), okul.tuz) !== okul.sifreOzet) { Utilities.sleep(800); throw hata_('Kullanıcı adı ya da şifre hatalı.'); }
    return oturumAc_(okul);
  } },
  cikis: { f: function (g) {
    var o = satirlar_('Oturumlar').filter(function (x) { return x.token === g.token; });
    sil_('Oturumlar', o.map(function (x) { return x._satir; })); return true;
  } },
  sifreDegistir: { f: function (g, okul) {
    if (ozet_(String(g.eski || ''), okul.tuz) !== okul.sifreOzet) throw hata_('Mevcut şifre hatalı.');
    if (String(g.yeni || '').length < 6) throw hata_('Yeni şifre en az 6 karakter olmalı.');
    okul.tuz = id_() + id_(); okul.sifreOzet = ozet_(String(g.yeni), okul.tuz); yaz_('Okullar', okul); return true;
  } },

  // hesaplama uygulamasının verisi (tarayıcıdaki kayıtlar)
  veriGetir: { f: function (g, okul) { // uzun kayıtlar "anahtar#0, #1…" parçalarında durur
    var p = {}; satirlar_('Veri').forEach(function (x) { if (x.okulId !== okul.okulId) return; var m = String(x.anahtar).match(/^(.*)#(\d+)$/); if (!m) return; (p[m[1]] = p[m[1]] || [])[+m[2]] = String(x.json); });
    var r = {}; Object.keys(p).forEach(function (a) { r[a] = p[a].join(''); }); return r;
  } },
  veriKaydet: { f: function (g, okul) {
    Object.keys(g.kayitlar || {}).forEach(function (a) {
      var json = String(g.kayitlar[a] || ''), parca = [];
      for (var i = 0; i < json.length || i === 0; i += 40000) parca.push(json.slice(i, i + 40000));
      sil_('Veri', satirlar_('Veri').filter(function (y) { return y.okulId === okul.okulId && String(y.anahtar).indexOf(a + '#') === 0; }).map(function (y) { return y._satir; }));
      var sh = sayfa_('Veri'), z = new Date().toISOString();
      sh.getRange(sh.getLastRow() + 1, 1, parca.length, 4).setValues(parca.map(function (x, i) { return [okul.okulId, a + '#' + i, x, z]; }));
    });
    return true;
  } },

  // öğrenci listesi
  ogrenciler: { f: function (g, okul) { return liste_(okul); } },
  ogrenciKaydet: { f: function (g, okul) {
    var o = g.ogrenci || {}, ad = String(o.adSoyad || '').trim(); if (!ad) throw hata_('Öğrenci adını yazın.');
    var tum = satirlar_('Ogrenciler'), x = o.ogrId ? tum.filter(function (y) { return y.okulId === okul.okulId && y.ogrId === o.ogrId; })[0] : null;
    var alanlar = ['adSoyad', 'sinif', 'sube', 'veliTel', 'dersler', 'ucret', 'durum', 'dekontNo', 'odemeTarihi', 'tutar', 'not'];
    if (!x) x = { okulId: okul.okulId, ogrId: id_(), durum: 'Ödenmedi' };
    alanlar.forEach(function (k) { if (o[k] !== undefined) x[k] = String(o[k]); });
    if (x._satir) yaz_('Ogrenciler', x); else ekle_('Ogrenciler', x);
    return liste_(okul);
  } },
  ogrenciSil: { f: function (g, okul) {
    sil_('Ogrenciler', satirlar_('Ogrenciler').filter(function (y) { return y.okulId === okul.okulId && (g.ogrIdler || []).indexOf(y.ogrId) >= 0; }).map(function (y) { return y._satir; }));
    return liste_(okul);
  } },
  ogrenciAktar: { f: function (g, okul) { // Excel'den toplu ekleme; aynı ad + sınıf varsa güncellenir
    var tum = satirlar_('Ogrenciler').filter(function (y) { return y.okulId === okul.okulId; }), yeni = 0, gunc = 0, sh = sayfa_('Ogrenciler'), ekRows = [];
    (g.ogrenciler || []).forEach(function (o) {
      var ad = String(o.adSoyad || '').trim(); if (!ad) return;
      var x = tum.filter(function (y) { return norm_(y.adSoyad) === norm_(ad) && String(y.sinif) === String(o.sinif || ''); })[0];
      if (x) { ['sube', 'veliTel', 'dersler', 'ucret'].forEach(function (k) { if (o[k]) x[k] = String(o[k]); }); yaz_('Ogrenciler', x); gunc++; }
      else { var y = { okulId: okul.okulId, ogrId: id_(), adSoyad: ad, sinif: String(o.sinif || ''), sube: String(o.sube || ''), veliTel: String(o.veliTel || ''), dersler: String(o.dersler || ''), ucret: String(o.ucret || ''), durum: 'Ödenmedi' };
        ekRows.push(SAYFALAR.Ogrenciler.map(function (k) { return y[k] === undefined ? '' : y[k]; })); tum.push(y); yeni++; }
    });
    if (ekRows.length) sh.getRange(sh.getLastRow() + 1, 1, ekRows.length, ekRows[0].length).setValues(ekRows);
    return { yeni: yeni, guncellenen: gunc, liste: liste_(okul) };
  } },

  // dekontlar
  dekontlar: { f: function (g, okul) { return dekListe_(okul); } },
  dekontYukle: { f: function (g, okul) {
    var bayt = Utilities.base64Decode(g.dosya), blob = Utilities.newBlob(bayt, g.tur || 'application/pdf', g.ad || 'dekont');
    var metin = ocr_(blob), b = dekontAyristir_(metin), ogrs = liste_(okul);
    if (b.dekontNo && satirlar_('Dekontlar').some(function (d) { return d.okulId === okul.okulId && String(d.dekontNo) === String(b.dekontNo); }))
      return { yinelenen: true, bilgi: b, metin: metin.slice(0, 1500) };
    var es = eslestir_(b, metin, ogrs), d = { okulId: okul.okulId, dekId: id_(), dekontNo: b.dekontNo, tarih: b.tarih, tutar: b.tutar, gonderen: b.gonderen,
      aciklama: b.aciklama, ogrIdler: es.kesin.join(','), durum: es.kesin.length ? 'Eşleşti' : 'Kontrol bekliyor', adaylar: es.adaylar.join(','), dosyaAdi: g.ad || '', yuklenme: new Date().toISOString() };
    ekle_('Dekontlar', d);
    if (es.kesin.length) odendiIsaretle_(okul, es.kesin, d);
    return { dekont: d, liste: liste_(okul), dekontlar: dekListe_(okul), metin: es.kesin.length ? '' : metin.slice(0, 1500) };
  } },
  dekontAta: { f: function (g, okul) { // elle eşleştirme ya da düzeltme
    var d = satirlar_('Dekontlar').filter(function (x) { return x.okulId === okul.okulId && x.dekId === g.dekId; })[0]; if (!d) throw hata_('Dekont bulunamadı.');
    var eski = String(d.ogrIdler || '').split(',').filter(String);
    if (eski.length) odemeKaldir_(okul, eski, d.dekontNo);
    var yeni = (g.ogrIdler || []).filter(String);
    ['dekontNo', 'tarih', 'tutar'].forEach(function (k) { if (g[k] !== undefined) d[k] = String(g[k]); });
    d.ogrIdler = yeni.join(','); d.durum = yeni.length ? 'Eşleşti (elle)' : 'Kontrol bekliyor'; yaz_('Dekontlar', d);
    if (yeni.length) odendiIsaretle_(okul, yeni, d);
    return { liste: liste_(okul), dekontlar: dekListe_(okul) };
  } },
  dekontSil: { f: function (g, okul) {
    var d = satirlar_('Dekontlar').filter(function (x) { return x.okulId === okul.okulId && x.dekId === g.dekId; })[0]; if (!d) throw hata_('Dekont bulunamadı.');
    var ids = String(d.ogrIdler || '').split(',').filter(String); if (ids.length) odemeKaldir_(okul, ids, d.dekontNo);
    sil_('Dekontlar', [d._satir]); return { liste: liste_(okul), dekontlar: dekListe_(okul) };
  } }
};

function temiz_(x) { var o = {}; Object.keys(x).forEach(function (k) { if (k !== '_satir' && k !== 'okulId') o[k] = x[k] instanceof Date ? Utilities.formatDate(x[k], 'Europe/Istanbul', 'dd.MM.yyyy') : x[k]; }); return o; }
function liste_(okul) { return satirlar_('Ogrenciler').filter(function (y) { return y.okulId === okul.okulId; }).map(temiz_); }
function dekListe_(okul) { return satirlar_('Dekontlar').filter(function (y) { return y.okulId === okul.okulId; }).map(temiz_).reverse(); }
function odendiIsaretle_(okul, ids, d) {
  satirlar_('Ogrenciler').forEach(function (y) {
    if (y.okulId !== okul.okulId || ids.indexOf(y.ogrId) < 0) return;
    y.durum = 'Ödendi'; y.dekontNo = d.dekontNo || '(numara okunamadı)'; y.odemeTarihi = d.tarih; y.tutar = d.tutar;
    if (ids.length > 1) y.not = 'Kardeşlerle tek dekont'; yaz_('Ogrenciler', y);
  });
}
function odemeKaldir_(okul, ids, no) {
  satirlar_('Ogrenciler').forEach(function (y) {
    if (y.okulId !== okul.okulId || ids.indexOf(y.ogrId) < 0 || String(y.dekontNo) !== String(no || '(numara okunamadı)')) return;
    y.durum = 'Ödenmedi'; y.dekontNo = ''; y.odemeTarihi = ''; y.tutar = ''; if (y.not === 'Kardeşlerle tek dekont') y.not = ''; yaz_('Ogrenciler', y);
  });
}

// ---------- dekont okuma ----------
function ocr_(blob) {
  // Drive API (Gelişmiş Hizmet, v3) gerekir: dosyayı Google Doküman'a çevirirken yazıyı okur, sonra siler.
  var f = Drive.Files.create({ name: 'gecici-dekont-' + id_(), mimeType: 'application/vnd.google-apps.document' }, blob, { ocrLanguage: 'tr', fields: 'id' });
  try { return DocumentApp.openById(f.id).getBody().getText(); }
  finally { try { Drive.Files.remove(f.id); } catch (x) { DriveApp.getFileById(f.id).setTrashed(true); } }
}

function deger_(metin, etiketler) {
  var satir = metin.split(/\r?\n/);
  for (var e = 0; e < etiketler.length; e++) {
    var re = new RegExp('(?:^|\\s)' + etiketler[e] + '\\s*[:：]?\\s*(.*)$', 'i');
    for (var i = 0; i < satir.length; i++) {
      var m = satir[i].match(re);
      if (m) {
        var v = m[1].split(/\s{3,}|\s+(?:İŞLEM|VALÖR|REFERANS|TARİH|TUTAR|IBAN)\b/i)[0].trim();
        if (!v && satir[i + 1]) v = satir[i + 1].trim();
        if (v && v !== '-' && v !== ':') return v;
      }
    }
  }
  return '';
}
function dekontAyristir_(metin) {
  var T = metin.replace(/ /g, ' ');
  var no = deger_(T, ['DEKONT NO', 'SORGU NO', 'İŞLEM NO', 'İŞLEM REFERANS(?:I|ı)?(?: NO)?', 'REFERANS NO', 'FAST SORGU NO', 'BİMREF[- ]SERİSIRANO', 'BELGE NO', 'REFERANS', 'İşlem No', 'Dekont No']);
  no = (no.match(/[A-Z0-9][A-Z0-9.\-\/]{4,}/i) || [''])[0];
  var tarihHam = deger_(T, ['İŞLEM TARİHİ', 'İşlem Tarihi', 'TARİH', 'Tarih']) || T;
  var tm = tarihHam.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/) || T.match(/(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{4})/);
  var tarih = tm ? ('0' + tm[1]).slice(-2) + '.' + ('0' + tm[2]).slice(-2) + '.' + tm[3] : '';
  var tu = ''; // etiketin geçtiği satırın tamamında ilk para değeri
  ['İŞLEM TUTARI', 'İşlem Tutarı', 'TOPLAM', 'TUTAR', 'Tutar', 'Gönderilen Tutar'].some(function (et) {
    return T.split(/\r?\n/).some(function (s) {
      var i = s.indexOf(et); if (i < 0) return false;
      var m = s.slice(i).match(/\d{1,3}(?:[.,]\d{3})*[.,]\d{2}|\d+(?=\s*(?:TL|₺))/); if (m) tu = m[0]; return !!m;
    });
  });
  return { dekontNo: no, tarih: tarih, tutar: tutarNorm_(tu), gonderen: deger_(T, ['GÖNDEREN(?: AD[I]? ?SOYAD[I]?)?', 'Gönderen', 'GÖNDEREN ADI', 'AD SOYAD']).replace(/\s+IBAN.*$/i, ''),
    aciklama: deger_(T, ['AÇIKLAMA', 'Açıklama', 'ÖDEME AÇIKLAMASI', 'Mesaj']) };
}
function tutarNorm_(s) {
  if (!s) return '';
  var m = s.match(/^(.*?)[.,](\d{2})$/), tam = m ? m[1] : s, kurus = m ? m[2] : '00';
  tam = tam.replace(/[.,]/g, '');
  return Number(tam).toLocaleString('tr-TR') + ',' + kurus;
}

// ---------- öğrenci eşleştirme ----------
function norm_(s) {
  return String(s || '').toLocaleUpperCase('tr').replace(/İ/g, 'I').replace(/Ş/g, 'S').replace(/Ğ/g, 'G').replace(/Ü/g, 'U').replace(/Ö/g, 'O').replace(/Ç/g, 'C').replace(/Â/g, 'A')
    .replace(/[^A-Z0-9 ]/g, ' ').replace(/\s+/g, ' ').trim();
}
function kelimeler_(s) { return norm_(s).split(' ').filter(function (w) { return w.length >= 2; }); }
function icerir_(metinK, adK) { return adK.every(function (w) { return metinK.indexOf(w) >= 0; }); }
function eslestir_(b, metin, ogrs) {
  var ack = kelimeler_(b.aciklama), tum = kelimeler_(metin), gon = kelimeler_(b.gonderen);
  var acikTam = ogrs.filter(function (o) { var k = kelimeler_(o.adSoyad); return k.length >= 2 && icerir_(ack, k); });
  var adlar = acikTam.map(function (o) { return norm_(o.adSoyad); });
  if (acikTam.length && adlar.every(function (a, i) { return adlar.indexOf(a) === i; })) return { kesin: acikTam.map(function (o) { return o.ogrId; }), adaylar: [] };
  if (acikTam.length) return { kesin: [], adaylar: acikTam.map(function (o) { return o.ogrId; }) }; // aynı adlı iki öğrenci: elle seçilsin
  // açıklamada yoksa: tüm metinde tam ad (gönderen kendisi değilse) ve tek öğrenci
  var metinTam = ogrs.filter(function (o) { var k = kelimeler_(o.adSoyad); return k.length >= 2 && icerir_(tum, k) && !icerir_(gon, k); });
  if (metinTam.length === 1) return { kesin: [metinTam[0].ogrId], adaylar: [] };
  // aday puanı: açıklamada geçen ad kelimeleri ×2, gönderenle aynı soyad +1
  var puan = ogrs.map(function (o) {
    var k = kelimeler_(o.adSoyad), p = 0; k.forEach(function (w) { if (ack.indexOf(w) >= 0) p += 2; else if (tum.indexOf(w) >= 0) p += 1; });
    if (gon.length && k.length && gon.indexOf(k[k.length - 1]) >= 0) p += 1;
    return { id: o.ogrId, p: p };
  }).filter(function (x) { return x.p >= 2; }).sort(function (a, b) { return b.p - a.p; });
  return { kesin: [], adaylar: puan.slice(0, 5).map(function (x) { return x.id; }) };
}

// Elle deneme: Apps Script düzenleyicide çalıştırın, Günlük'te sonucu görün
function dekontTesti() {
  var ornek = 'GÖNDEREN : AYŞE YILMAZ\nİŞLEM TARİHİ : 07/11/2026 - 10:12\nSORGU NO : 1661057775\nİŞLEM TUTARI  (TL) :   1,500.00\nAÇIKLAMA : Ali Can Yılmaz çocuk kulübü';
  var b = dekontAyristir_(ornek);
  Logger.log(JSON.stringify(b));
  Logger.log(JSON.stringify(eslestir_(b, ornek, [{ ogrId: '1', adSoyad: 'ALİ CAN YILMAZ' }, { ogrId: '2', adSoyad: 'ELİF YILMAZ' }])));
}

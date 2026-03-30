function api_kasaAcilisKaydet(p) {
  // Devreden çıkarıldı — FİNANS_KASA artık kullanılmıyor
  // Açılış bakiyesi gerekirse FİNANS_CARİ'ye GELİR kaydı olarak eklenir
  try {
    var kasaAdi = String(p.kasaAdi || "").trim();
    var bakiye = parseFloat(p.bakiye || 0);
    if (!kasaAdi) return { ok: false, msg: "Kasa adı zorunludur." };
    if (bakiye <= 0) return { ok: false, msg: "Bakiye sıfırdan büyük olmalı." };
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok: false, msg: "FİNANS_CARİ bulunamadı" };
    var simdi = new Date();
    sh.appendRow([
      "FIN_ACILIS_" + simdi.getTime(),
      simdi, "GELİR", "Kasa Açılış",
      kasaAdi, "Açılış bakiyesi",
      bakiye, 0, bakiye, bakiye,
      "SISTEM", String(p.kullanici || ""),
      "TAMAMLANDI", ""
    ]);
    return { ok: true, msg: kasaAdi + " açılış bakiyesi kaydedildi." };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_kasaGuncelle_(kasaAdi, tip, tutar) {
  // Devreden çıkarıldı — bakiyeler artık FİNANS_CARİ'den canlı hesaplanıyor
  return;
}

function api_kasaListesiGetir() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, kasalar: [] };
    var data = sh.getDataRange().getValues();
    var kasaMap = {};
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var tip = String(data[i][2] || "").toUpperCase();
      var kasa = String(data[i][4] || "NAKİT").trim();
      var tutar = parseFloat(data[i][6]) || 0;
      var gelirMi = tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1;
      if (!kasaMap[kasa]) kasaMap[kasa] = 0;
      kasaMap[kasa] += gelirMi ? tutar : -tutar;
    }
    var kasalar = [];
    for (var k in kasaMap) {
      kasalar.push({ kasa: k, tip: k, bakiye: Math.round(kasaMap[k] * 100) / 100 });
    }
    return { ok: true, kasalar: kasalar };
  } catch(e) { return { ok: false, kasalar: [], msg: e.message }; }
}

function api_nakitPozisyonu() {
  try {
    var res = api_kasaListesiGetir();
    var toplam = 0;
    (res.kasalar || []).forEach(function(k) { toplam += k.bakiye || 0; });
    return { ok: true, toplamBanka: 0, toplamKasa: toplam, toplamNakit: toplam, bankaListesi: [], kasaListesi: res.kasalar || [] };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_kasaHareketleri(kasaAdi, donemAy, donemYil) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, hareketler: [], gelirToplam: 0, giderToplam: 0 };
    var data = sh.getDataRange().getValues();
    var hareketler = [], gelirToplam = 0, giderToplam = 0;
    var buAy  = donemAy  ? parseInt(donemAy) - 1  : new Date().getMonth();
    var buYil = donemYil ? parseInt(donemYil)      : new Date().getFullYear();
    for (var i = data.length - 1; i >= 1; i--) {
      if (!data[i][0]) continue;
      var dt = data[i][1] instanceof Date ? data[i][1] : new Date(String(data[i][1]).split(".").reverse().join("-"));
      if (isNaN(dt) || dt.getMonth() !== buAy || dt.getFullYear() !== buYil) continue;
      var tip = String(data[i][2] || "").toUpperCase();
      var tutar = parseFloat(data[i][6]) || 0;
      if (tip.indexOf("GEL") > -1 || tip.indexOf("TAH") > -1) gelirToplam += tutar;
      else giderToplam += tutar;
      var dt2 = data[i][1] instanceof Date ? data[i][1] : new Date(String(data[i][1]).split(".").reverse().join("-"));
var tarihStr = isNaN(dt2) ? String(data[i][1]||"") : ('0'+dt2.getDate()).slice(-2) + '.' + ('0'+(dt2.getMonth()+1)).slice(-2) + '.' + dt2.getFullYear();
hareketler.push({ 
  id       : String(data[i][0]||""), 
  tarih    : tarihStr, 
  tip      : tip, 
  kategori : String(data[i][3]||""),
  kasa     : String(data[i][4]||""),
  aciklama : String(data[i][5]||""),
  tutar    : tutar 
});
    }
    return { ok: true, hareketler: hareketler, gelirToplam: gelirToplam, giderToplam: giderToplam };
  } catch(e) { return { ok: false, hareketler: [], gelirToplam: 0, giderToplam: 0, msg: e.message }; }
}

function api_muhasebeKpiKartlari() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok:true, ayGelir:0, ayGider:0, ayNet:0, toplamNakit:0, toplamKasa:0, toplamBanka:0, acikAlacak:0, bekleyenCek:0 };

    var data = sh.getDataRange().getValues();
    var simdi = new Date(), buAy = simdi.getMonth(), buYil = simdi.getFullYear();
    var kasaMap = {}, ayGelir = 0, ayGider = 0;

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var tip = String(data[i][2] || "").toUpperCase();
      var kasa = String(data[i][4] || "NAKİT").trim();
      var tutar = parseFloat(data[i][6]) || 0;
      var gelirMi = tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1;

      if (!kasaMap[kasa]) kasaMap[kasa] = 0;
      kasaMap[kasa] += gelirMi ? tutar : -tutar;

      var dt = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
      if (!isNaN(dt) && dt.getMonth() === buAy && dt.getFullYear() === buYil) {
        if (gelirMi) ayGelir += tutar; else ayGider += tutar;
      }
    }

    var toplamKasa = 0;
    for (var k in kasaMap) toplamKasa += kasaMap[k];

    // Açık alacak — FİRMA_CARİ'den
    var acikAlacak = 0;
    try {
      var shFC = ss.getSheetByName("FİRMA_CARİ");
      if (shFC && shFC.getLastRow() > 1) {
        var fcData = shFC.getDataRange().getValues();
        var firmaBakiye = {};
        for (var j = 1; j < fcData.length; j++) {
          var fId = String(fcData[j][2] || "").trim();
          if (fId) firmaBakiye[fId] = parseFloat(fcData[j][8]) || 0;
        }
        for (var f in firmaBakiye) {
          if (firmaBakiye[f] > 0) acikAlacak += firmaBakiye[f];
        }
      }
    } catch(e2) {}

    // Bekleyen çek — CEKLER_SENETLER'den
    var bekleyenCek = 0;
    try {
      var shCek = ss.getSheetByName("CEKLER_SENETLER");
      if (shCek && shCek.getLastRow() > 1) {
        var cekData = shCek.getDataRange().getValues();
        for (var c = 1; c < cekData.length; c++) {
          if (String(cekData[c][6] || "").toUpperCase() !== "ODENDİ") {
            bekleyenCek += parseFloat(cekData[c][4]) || 0;
          }
        }
      }
    } catch(e3) {}

    return {
      ok: true,
      ayGelir: Math.round(ayGelir * 100) / 100,
      ayGider: Math.round(ayGider * 100) / 100,
      ayNet: Math.round((ayGelir - ayGider) * 100) / 100,
      toplamNakit: Math.round(toplamKasa * 100) / 100,
      toplamKasa: Math.round(toplamKasa * 100) / 100,
      toplamBanka: 0,
      acikAlacak: Math.round(acikAlacak * 100) / 100,
      bekleyenCek: Math.round(bekleyenCek * 100) / 100
    };
  } catch(e) { return { ok:false, msg:e.message }; }
}

function toplamBankaBakiyesi() {
  try {
    var res = api_kasaListesiGetir();
    var toplam = 0;
    (res.kasalar || []).forEach(function(k) { toplam += k.bakiye || 0; });
    return toplam;
  } catch(e) { return 0; }
}

function posTahsilatiEkle(params) {
  try {
    var sheet = getSheet(SHEETS.POS);
    var id = yeniID('P', sheet);
    var tutar = params.tutar || 0;
    var komisyon = params.komisyon || 0;
    var netTutar = tutar - komisyon;
    sheet.appendRow([
      id,
      params.tarih || bugun(),
      params.aciklama || '',
      tutar,
      komisyon,
      netTutar,
      params.kaynakId || ''
    ]);
    Logger.log('💳 POS tahsilat: ' + paraBirim(tutar) + ' | ID: ' + id);
    return id;
  } catch(e) {
    Logger.log('POS tahsilat hatası: ' + e.message);
    return null;
  }
}

function eftGiris(bankaKod, tutar, aciklama, karsiHesap, kaynakId) {
  try {
    var sheet = getSheet(SHEETS.BANKA_HAREKET);
    var mevcutBakiye = bankaBakiyesiAl(bankaKod);
    var yeniBakiye = mevcutBakiye + tutar;
    var id = yeniID('BH', sheet);
    sheet.appendRow([
      id, bugun(), bankaKod, aciklama,
      'Giriş', tutar, yeniBakiye,
      karsiHesap || '', kaynakId || ''
    ]);
    Logger.log('🏦 EFT Giriş: ' + bankaKod + ' | ' + paraBirim(tutar));
    return id;
  } catch(e) {
    Logger.log('EFT giriş hatası: ' + e.message);
    return null;
  }
}

function eftCikis(bankaKod, tutar, aciklama, karsiHesap, kaynakId) {
  try {
    var sheet = getSheet(SHEETS.BANKA_HAREKET);
    var mevcutBakiye = bankaBakiyesiAl(bankaKod);
    var yeniBakiye = mevcutBakiye - tutar;
    var id = yeniID('BH', sheet);
    sheet.appendRow([
      id, bugun(), bankaKod, aciklama,
      'Çıkış', tutar, yeniBakiye,
      karsiHesap || '', kaynakId || ''
    ]);
    Logger.log('🏦 EFT Çıkış: ' + bankaKod + ' | ' + paraBirim(tutar));
    return id;
  } catch(e) {
    Logger.log('EFT çıkış hatası: ' + e.message);
    return null;
  }
}

function bankaBakiyesiAl(bankaKod) {
  try {
    var sheet = getSheet(SHEETS.BANKA_HAREKET);
    var sonSatir = sheet.getLastRow();
    if (sonSatir <= 1) return 0;
    var data = sheet.getDataRange().getValues();
    var bakiye = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][COL.BANKA_HAREKET.BANKA_KOD - 1]) === String(bankaKod)) {
        bakiye = data[i][COL.BANKA_HAREKET.BAKIYE - 1] || 0;
        break;
      }
    }
    return bakiye;
  } catch(e) {
    return 0;
  }
}

function api_bankaHesaplariListele() {
  try {
    var res = api_kasaListesiGetir();
    var liste = [];
    (res.kasalar || []).forEach(function(k) {
      if (k.kasa === "NAKİT") return;
      liste.push({ hesapAdi: k.kasa, banka: k.tip, bakiye: k.bakiye, durum: "AKTİF" });
    });
    return { ok: true, liste: liste };
  } catch(e) { return { ok: false, liste: [], msg: e.message }; }
}

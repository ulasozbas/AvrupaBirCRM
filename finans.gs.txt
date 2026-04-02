// ==========================================
// FİNANS İŞLEM KAYDET
// ==========================================
function api_finansIslemKaydet(p) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok: false, msg: "FİNANS_CARİ sekmesi bulunamadı" };
    var simdi = new Date();
    var tutar = parseFloat(p.tutar) || 0;
    var kdv = parseFloat(p.kdv) || 0;
    var kdvsiz = kdv > 0 ? tutar / (1 + kdv/100) : tutar;
    var netTutar = tutar - (tutar - kdvsiz);
    var tip = String(p.tip || "GİDER").toUpperCase();
    sh.appendRow([
      "FIN_" + simdi.getTime(),
      simdi,
      tip,
      String(p.kategori || p.odemeTipi || ""),
      String(p.kasa || "NAKİT"),
      String(p.aciklama || p.firma || p.firmaAdi || ""),
      tutar, kdv, kdvsiz, netTutar,
      String(p.fisTip || "BANKO"),
      String(p.kullanici || p.kaydedenEmail || ""),
      "TAMAMLANDI", ""
    ]);
    var tipKasa = String(p.tip || p.odemeTipi || "").toUpperCase();
    if (!tipKasa || (tipKasa !== "GELİR" && tipKasa !== "GİDER")) tipKasa = "GELİR";
    api_kasaGuncelle_(String(p.kasa || "NAKİT"), tipKasa, tutar);
    var emoji = tip.indexOf("GELİR") > -1 ? "💰" : "💸";
    chatBildir_(emoji + " *" + tip + "*\nFirma: " + String(p.aciklama || p.firmaAdi || "-") + "\nTutar: " + tutar.toLocaleString("tr-TR") + " TL");
    logYaz("FINANS_ISLEM", tip + " | " + tutar + " TL | " + String(p.aciklama || p.firmaAdi || "-"), "FINANS", "FIN_" + simdi.getTime(), String(p.aciklama || p.firmaAdi || ""), "");
    return { ok: true };
  } catch(e) { return { ok: false, msg: e.message }; }
}


// ==========================================
// TOPLAM KASA
// ==========================================
function api_getToplamKasa() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok:true, toplam:0, bakiye:0, kasalar:[], ayGelir:0, ayGider:0 };

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

    var toplam = 0, kasalar = [];
    for (var k in kasaMap) {
      var b = Math.round(kasaMap[k] * 100) / 100;
      toplam += b;
      kasalar.push({ kasa: k, bakiye: b });
    }

    return { ok:true, toplam:Math.round(toplam*100)/100, bakiye:Math.round(toplam*100)/100, kasalar:kasalar, ayGelir:Math.round(ayGelir*100)/100, ayGider:Math.round(ayGider*100)/100 };
  } catch(e) { return { ok:false, toplam:0, msg:e.message }; }
}

// ==========================================
// FİNANS LİSTESİ
// ==========================================
function api_getFinansList(p) {
  try {
    p = p || {};
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return [];
    var data = sh.getDataRange().getValues();
    if (data.length < 2) return [];
    var from = p.from ? new Date(p.from) : null;
    var to   = p.to   ? new Date(p.to + "T23:59:59") : null;
    var tip  = String(p.tip || "");
    var list = [];
    for (var i=data.length-1; i>=1; i--) {
      if (!data[i][0]) continue;
      if (from || to) { var dt=new Date(data[i][1]); if(!isNaN(dt.getTime())){if(from&&dt<from)continue;if(to&&dt>to)continue;} }
      if (tip && String(data[i][2]||"").indexOf(tip)<0) continue;
      list.push({ id:String(data[i][0]||""), tarih:String(data[i][1]||""), tip:String(data[i][2]||""), kategori:String(data[i][3]||""), kasa:String(data[i][4]||""), firma:String(data[i][5]||""), kisi:parseInt(data[i][6])||0, aciklama:String(data[i][7]||""), tutar:parseFloat(data[i][8])||0, kdv:parseFloat(data[i][9])||0, kaydeden:String(data[i][12]||""), durum:String(data[i][13]||"") });
      if (list.length >= 50) break;
    }
    return list;
  } catch(e) { return []; }
}

function api_finansIslemSil(id) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok:false, msg:"FİNANS_CARİ bulunamadı" };
    var data = sh.getDataRange().getValues();
    for (var i=1; i<data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        var tip=String(data[i][2]||""), tutar=parseFloat(data[i][8])||0, kasa=String(data[i][4]||"ANA_KASA");
        var tersTip = (tip.indexOf("GELİR")>-1||tip.indexOf("TAHSİLAT")>-1) ? "GİDER" : "GELİR";
        api_kasaGuncelle_(kasa, tersTip, tutar);
        sh.deleteRow(i+1);
        logYaz("FINANS_SIL", "Silindi | " + tip + " | " + tutar + " TL | " + kasa, "FINANS", String(id), "", "");
        return { ok:true };
      }
    }
    return { ok:false, msg:"Kayıt bulunamadı" };
  } catch(e) { return { ok:false, msg:e.message }; }
}

// ==========================================
// FİNANS DEFTERİ
// ==========================================
function api_getFinansDefteri(arananTarih) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var shCari = ss.getSheetByName("FİNANS_CARI") || ss.getSheetByName("FİNANS CARİ") || ss.getSheetByName("FINANS_CARI");
    if (!shCari) return { hata:true, msg:"FİNANS_CARI sayfası bulunamadı!" };
    var data = shCari.getDataRange().getDisplayValues();
    if (data.length < 2) return [];
    var head = data[0].map(function(h){ return String(h).toLowerCase().replace(/[\s\.\-\_₺\(\)\+]/g,""); });
    var cxTarih=head.findIndex(function(h){return h.includes("tarih");}); if(cxTarih===-1)cxTarih=1;
    var cxTip=head.findIndex(function(h){return h.includes("tip");}); if(cxTip===-1)cxTip=2;
    var cxKategori=head.findIndex(function(h){return h.includes("kategori");}); if(cxKategori===-1)cxKategori=3;
    var cxAciklama=head.findIndex(function(h){return h.includes("firma")||h.includes("açıklama")||h.includes("aciklama");}); if(cxAciklama===-1)cxAciklama=5;
    var cxTutar=head.findIndex(function(h){return h.includes("tutar")||h.includes("kdv");}); if(cxTutar===-1)cxTutar=6;
    var cxPersonel=head.findIndex(function(h){return h.includes("yapan")||h.includes("personel");}); if(cxPersonel===-1)cxPersonel=11;
    var list = [], maxSayac = 0;
    var temizAranan = String(arananTarih||"").trim().replace(/\//g,".");
    for (var i=data.length-1; i>=1; i--) {
      var islemTarihi = String(data[i][cxTarih]).trim();
      if (!islemTarihi) continue;
      var excelTarihi = islemTarihi.replace(/\//g,".");
      if (temizAranan && !excelTarihi.includes(temizAranan)) continue;
      var hamTutar = String(data[i][cxTutar]).replace(/[^\d\,\.-]/g,"");
      if (hamTutar.includes(",")) hamTutar = hamTutar.replace(/\./g,"").replace(",",".");
      list.push({ tarih:islemTarihi, tip:String(data[i][cxTip]).toUpperCase()||"BİLİNMİYOR", kategori:String(data[i][cxKategori])||"Diğer", aciklama:String(data[i][cxAciklama])||"-", tutar:parseFloat(hamTutar)||0, personel:String(data[i][cxPersonel]||"Bilinmiyor").trim() });
      maxSayac++;
      if (maxSayac >= 500) break;
    }
    return list;
  } catch(e) { return { hata:true, msg:"Sunucu Hatası: "+e.message }; }
}

// ==========================================
// KDV DURUMU
// ==========================================
function api_kdvDurumuHesapla() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok:false, msg:"FİNANS_CARİ bulunamadı" };
    var simdi=new Date(), buAy=simdi.getMonth(), buYil=simdi.getFullYear();
    var data = sh.getDataRange().getValues();
    var toplananKdv=0, odenenKdv=0;
    for (var i=1; i<data.length; i++) {
      var tarih=data[i][1], dt=(tarih instanceof Date)?tarih:new Date(String(tarih).split(".").reverse().join("-"));
      if (isNaN(dt)||dt.getMonth()!==buAy||dt.getFullYear()!==buYil) continue;
      var tip=String(data[i][2]||""), tutar=parseFloat(data[i][6])||0, kdvOran=parseFloat(data[i][7])||0;
      if (kdvOran<=0) continue;
      var kdvTutar = tutar - (tutar/(1+kdvOran/100));
      if (tip==="GELİR") toplananKdv+=kdvTutar; else odenenKdv+=kdvTutar;
    }
    return { ok:true, toplananKdv:Math.round(toplananKdv*100)/100, odenenKdv:Math.round(odenenKdv*100)/100, netKdv:Math.round((toplananKdv-odenenKdv)*100)/100 };
  } catch(e) { return { ok:false, msg:e.message }; }
}

function api_fisiYapayZekayaOkut(base64Data, mimeType) {
  try {
    var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + GEMINI_API_KEY;
    var payload = { contents:[{parts:[{text:"Bu bir fiş veya fatura fotoğrafı. Sadece JSON döndür: {\"firma\": \"firma adı\", \"tutar\": toplam tutar sayı, \"kdv\": kdv oranı sayı, \"tarih\": \"gg.aa.yyyy\", \"aciklama\": \"kısa açıklama\"}"},{inline_data:{mime_type:mimeType,data:base64Data}}]}] };
    var response = UrlFetchApp.fetch(url, { method:"POST", contentType:"application/json", payload:JSON.stringify(payload), muteHttpExceptions:true });
    var json = JSON.parse(response.getContentText());
    var text = json.candidates[0].content.parts[0].text.replace(/```json/g,"").replace(/```/g,"").trim();
    var parsed = JSON.parse(text);
    return { ok:true, firma:String(parsed.firma||""), tutar:parseFloat(parsed.tutar||0), kdv:parseInt(parsed.kdv||0), tarih:String(parsed.tarih||""), aciklama:String(parsed.aciklama||"") };
  } catch(e) { return { ok:false, msg:"Fiş okunamadı: "+e.message }; }
}

function api_gelirKaydet(p) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok: false, msg: "FİNANS_CARİ bulunamadı" };
    var simdi = new Date();
    var tutar = parseFloat(p.tutar) || 0;
    var odemeTipi = String(p.odemeTipi || "NAKİT").toUpperCase();
    var kasaAdi = odemeTipi === "KREDİ KARTI" ? "KREDİ KARTI" :
                  (odemeTipi === "HAVALE" || odemeTipi === "EFT") ? "HAVALE/EFT" : "NAKİT";
    sh.appendRow([
      "FIN_" + simdi.getTime(),
      simdi,
      "GELİR",
      odemeTipi,
      kasaAdi,
      String(p.firmaAdi || p.aciklama || ""),
      tutar, 0, tutar, tutar,
      "BANKO",
      String(p.kaydedenEmail || ""),
      "TAMAMLANDI", ""
    ]);
    api_kasaGuncelle_(kasaAdi, "GELİR", tutar);
    chatBildir_("💰 *Gelir Kaydedildi*\nFirma: " + String(p.firmaAdi||"-") + "\nTutar: " + tutar.toLocaleString("tr-TR") + " ₺\nÖdeme: " + odemeTipi + "\nKaydeden: " + String(p.kaydedenEmail||""));
    logYaz("GELIR_KAYIT", "GELİR | " + tutar + " TL | " + String(p.firmaAdi || "-"), "FINANS", "FIN_" + simdi.getTime(), String(p.firmaAdi || ""), "");
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}
function api_firmaCariEkle(p) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    
    // FİRMA_CARİ'ye kaydet (mevcut)
    var sh = ss.getSheetByName("FİRMA_CARİ");
    if (!sh) return { ok: false, msg: "FİRMA_CARİ tablosu bulunamadı" };
    var simdi = new Date();
    var tarih = Utilities.formatDate(simdi, "Europe/Istanbul", "dd.MM.yyyy HH:mm");
    var id = "CARİ_" + simdi.getTime();
    var tip = String(p.tip || "ALACAK").toUpperCase();
    var borc = tip === "BORC" ? parseFloat(p.tutar || 0) : 0;
    var alacak = tip === "ALACAK" ? parseFloat(p.tutar || 0) : 0;
    var data = sh.getDataRange().getValues();
    var mevcutBakiye = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]||"").trim() === String(p.firmaId||"").trim()) {
        mevcutBakiye = parseFloat(data[i][8]||0);
        break;
      }
    }
    var yeniBakiye = mevcutBakiye + alacak - borc;
    sh.appendRow([id, tarih, String(p.firmaId||""), String(p.firmaAdi||""), tip, String(p.aciklama||""), borc, alacak, Math.round(yeniBakiye*100)/100, String(p.faturaNo||""), String(p.kaydeden||""), "AÇIK"]);

    // FİNANS_CARİ'ye de kaydet (banko görünümü için)
    var shFinans = ss.getSheetByName("FİNANS_CARİ");
    if (shFinans) {
      shFinans.appendRow([
        id,                                      // A - ID
        simdi,                                   // B - Tarih
        "GELİR",                                 // C - Tip
        "Cari Tahsilat",                         // D - Kategori
        "CARİ",                                  // E - Kasa
        String(p.firmaAdi || p.aciklama || ""),  // F - Açıklama
        Number(p.tutar || 0),                    // G - Tutar
        0,                                       // H - KDV%
        Number(p.tutar || 0),                    // I - KDVsiz
        Number(p.tutar || 0),                    // J - NetTutar
        "BANKO",                                 // K - FisTip
        String(p.kaydeden || ""),                // L - Kaydeden
        "TAMAMLANDI"                             // M - Durum
      ]);
    }

    chatBildir_("📋 *Cari Tahsilat*\nFirma: " + String(p.firmaAdi||"") + "\nTutar: " + Number(p.tutar||0).toLocaleString("tr-TR") + " ₺\nKaydeden: " + String(p.kaydeden||""));
    return { ok: true, id: id, yeniBakiye: yeniBakiye };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}
function api_firmaCariGetir(firmaId) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİRMA_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, liste: [], bakiye: 0 };
    var data = sh.getDataRange().getValues();
    var liste = [];
    var bakiye = 0;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2]||"").trim() !== String(firmaId||"").trim()) continue;
      liste.push({
        id: String(data[i][0]||""),
        tarih: String(data[i][1]||""),
        tip: String(data[i][4]||""),
        aciklama: String(data[i][5]||""),
        borc: parseFloat(data[i][6]||0),
        alacak: parseFloat(data[i][7]||0),
        bakiye: parseFloat(data[i][8]||0),
        faturaNo: String(data[i][9]||""),
        durum: String(data[i][11]||"")
      });
      bakiye = parseFloat(data[i][8]||0);
    }
    return { ok: true, liste: liste, bakiye: bakiye };
  } catch(e) {
    return { ok: false, liste: [], bakiye: 0, msg: e.message };
  }
}

function api_firmaCariTahsilat(p) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİRMA_CARİ");
    if (!sh) return { ok: false, msg: "FİRMA_CARİ bulunamadı" };
    var data = sh.getDataRange().getValues();
    var mevcutBakiye = 0;
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]||"").trim() === String(p.firmaId||"").trim()) {
        mevcutBakiye = parseFloat(data[i][8]||0);
        break;
      }
    }
    var tahsilat = parseFloat(p.tutar || 0);
    var yeniBakiye = mevcutBakiye - tahsilat;
    var simdi = new Date();
    sh.appendRow([
      "TAH_" + simdi.getTime(),
      Utilities.formatDate(simdi, "Europe/Istanbul", "dd.MM.yyyy HH:mm"),
      String(p.firmaId || ""),
      String(p.firmaAdi || ""),
      "TAHSİLAT",
      String(p.aciklama || "Tahsilat"),
      tahsilat, 0,
      Math.round(yeniBakiye * 100) / 100,
      String(p.faturaNo || ""),
      String(p.kaydeden || ""),
      "KAPANDI"
    ]);
    // Kasaya da ekle
    var odemeTipi = String(p.odemeTipi || "NAKİT").toUpperCase();
    var kasaAdi = odemeTipi === "KREDİ KARTI" ? "KREDİ KARTI" :
                  (odemeTipi === "HAVALE" || odemeTipi === "EFT") ? "HAVALE/EFT" : "NAKİT";
    api_kasaGuncelle_(kasaAdi, "GELİR", tahsilat);
    return { ok: true, yeniBakiye: yeniBakiye };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function api_giderKaydet(p) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok: false, msg: "FİNANS_CARİ bulunamadı" };
    var simdi = new Date();
    var tutar = parseFloat(p.tutar) || 0;
    var odemeTipi = String(p.odemeTipi || "NAKİT").toUpperCase();

    // Ödeme tipine göre kasa belirle
    var kasaAdi;
    if (odemeTipi === "KREDİ KARTI") kasaAdi = "KREDİ KARTI";
    else if (odemeTipi === "HAVALE" || odemeTipi === "EFT") kasaAdi = "HAVALE/EFT";
    else if (odemeTipi === "ÇEK") kasaAdi = "ÇEK";
    else kasaAdi = "NAKİT";

    sh.appendRow([
      "FIN_" + simdi.getTime(),
      simdi,
      "GİDER",
      String(p.kategori || "Diğer"),
      kasaAdi,
      String(p.aciklama || ""),
      tutar, 0, tutar, tutar,
      "BANKO",
      String(p.kaydedenEmail || ""),
      "TAMAMLANDI", ""
    ]);

    // Çek ise ayrı tabloya ekle
    if (odemeTipi === "ÇEK") {
      var shCek = ss.getSheetByName("CEKLER_SENETLER");
      if (!shCek) {
        shCek = ss.insertSheet("CEKLER_SENETLER");
        shCek.appendRow(["ID","Tarih","VadeTarihi","Aciklama","Tutar","Tip","Durum","Kaydeden"]);
        shCek.getRange(1,1,1,8).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      }
      shCek.appendRow([
        "CEK_" + simdi.getTime(),
        Utilities.formatDate(simdi, "Europe/Istanbul", "dd.MM.yyyy"),
        String(p.vadeTarihi || ""),
        String(p.aciklama || ""),
        tutar, "GİDER", "BEKLEMEDE",
        String(p.kaydedenEmail || "")
      ]);
    } else {
      // Kasadan düş
      api_kasaGuncelle_(kasaAdi, "GİDER", tutar);
    }

    chatBildir_("💸 *GİDER*\nAçıklama: " + String(p.aciklama||"-") + "\nTutar: " + tutar.toLocaleString("tr-TR") + " ₺\nÖdeme: " + odemeTipi + "\nKaydeden: " + String(p.kaydedenEmail||""));
    logYaz("GIDER_KAYIT", "GİDER | " + tutar + " TL | " + String(p.aciklama || "-"), "FINANS", "FIN_" + simdi.getTime(), String(p.aciklama || ""), "");
    return { ok: true };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function api_cekSenetListesiGetir() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("CEKLER_SENETLER");
    if (!sh || sh.getLastRow() < 2) return { ok: true, liste: [], toplam: 0 };
    var data = sh.getDataRange().getValues();
    var liste = [];
    var toplam = 0;
    for (var i = 1; i < data.length; i++) {
      var durum = String(data[i][6]||"").toUpperCase();
      if (durum === "ODENDI") continue;
      var tutar = parseFloat(data[i][4]||0);
      toplam += tutar;
      liste.push({
        id: String(data[i][0]||""),
        tarih: String(data[i][1]||""),
        vadeTarihi: String(data[i][2]||""),
        aciklama: String(data[i][3]||""),
        tutar: tutar,
        tip: String(data[i][5]||""),
        durum: durum,
        kaydeden: String(data[i][7]||"")
      });
    }
    liste.sort(function(a,b){ return a.vadeTarihi > b.vadeTarihi ? 1 : -1; });
    return { ok: true, liste: liste, toplam: Math.round(toplam*100)/100 };
  } catch(e) {
    return { ok: false, liste: [], toplam: 0, msg: e.message };
  }
}

function api_cekOdendi(id) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("CEKLER_SENETLER");
    if (!sh) return { ok: false, msg: "CEKLER_SENETLER bulunamadı" };
    var data = sh.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]||"") === String(id)) {
        sh.getRange(i+1, 7).setValue("ODENDİ");
        var tutar = parseFloat(data[i][4]||0);
        api_kasaGuncelle_("NAKİT", "GİDER", tutar);
        return { ok: true };
      }
    }
    return { ok: false, msg: "Çek bulunamadı" };
  } catch(e) {
    return { ok: false, msg: e.message };
  }
}

function api_karZararHesapla(donem) {
  try {
    donem = donem || "aylik";
    if (donem === "gunluk" || donem === "haftalik") donem = "aylik"; // güvenlik
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, gelir: 0, gider: 0, net: 0 };

    var simdi = new Date();
    var data = sh.getDataRange().getValues();
    var gelir = 0, gider = 0, gelirOnceki = 0, giderOnceki = 0;

    // Sabit giderleri al
    var sabitGiderAylik = 0;
    try {
      var sgSh = ss.getSheetByName("SABIT_GIDER");
      if (sgSh && sgSh.getLastRow() > 1) {
        var sgData = sgSh.getDataRange().getValues();
        for (var s = 1; s < sgData.length; s++) {
          var aktif = sgData[s][4];
          if (aktif === true || String(aktif).toUpperCase() === "TRUE") {
            sabitGiderAylik += parseFloat(sgData[s][3]) || 0;
          }
        }
      }
    } catch(e) {}

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dt = data[i][1] instanceof Date ? data[i][1]
        : new Date(String(data[i][1]).split(".").reverse().join("-"));
      if (isNaN(dt)) continue;
      var tutar = parseFloat(data[i][6]) || 0;
      var tip = String(data[i][2] || "").toUpperCase();
      var gelirMi = tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1;

      if (_donemIcindeMi(dt, simdi, donem)) {
        if (gelirMi) gelir += tutar; else gider += tutar;
      }
      if (_oncekiDonemMi(dt, simdi, donem)) {
        if (gelirMi) gelirOnceki += tutar; else giderOnceki += tutar;
      }
    }

    // Sabit gideri ekle
    var sabitGider = donem === "yillik" ? sabitGiderAylik * 12 : sabitGiderAylik;
    gider += sabitGider;
    giderOnceki += sabitGider;

    var net = gelir - gider;
    var netOnceki = gelirOnceki - giderOnceki;
    var degisim = netOnceki !== 0
      ? Math.round(((net - netOnceki) / Math.abs(netOnceki)) * 100) : 0;

    return {
      ok           : true,
      donem        : donem,
      gelir        : Math.round(gelir * 100) / 100,
      gider        : Math.round(gider * 100) / 100,
      net          : Math.round(net * 100) / 100,
      sabitGider   : Math.round(sabitGider * 100) / 100,
      onceki       : { gelir: gelirOnceki, gider: giderOnceki, net: netOnceki },
      degisim      : degisim
    };
  } catch(e) { return { ok: false, msg: e.message }; }
}
function _donemIcindeMi(dt, simdi, donem) {
  if (donem === "aylik") {
    return dt.getMonth() === simdi.getMonth() && dt.getFullYear() === simdi.getFullYear();
  } else if (donem === "yillik") {
    return dt.getFullYear() === simdi.getFullYear();
  }
  return false;
}

function _oncekiDonemMi(dt, simdi, donem) {
  if (donem === "aylik") {
    var oncekiAy = new Date(simdi.getFullYear(), simdi.getMonth() - 1, 1);
    return dt.getMonth() === oncekiAy.getMonth() && dt.getFullYear() === oncekiAy.getFullYear();
  } else if (donem === "yillik") {
    return dt.getFullYear() === simdi.getFullYear() - 1;
  }
  return false;
}
function api_kacakRadariKontrol() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh) return { ok: false, msg: "FİNANS_CARİ bulunamadı" };

    var LIMITLER = {
      "Araç Yakıt & Otoyol"            : 5000,
      "Ofis ve Mutfak Gideri"           : 3000,
      "Personel Yemek"                  : 4000,
      "Kırtasiye ve Toner"              : 2000,
      "Dışarıdan Hizmet / Laboratuvar"  : 10000,
      "Maaş / Avans Ödemesi"            : 100000,
      "Vergi ve SGK Ödemeleri"          : 50000,
      "Kira"                            : 20000,
      "Elektrik / Su / Doğalgaz"        : 5000,
      "Diğer"                           : 5000
    };

    var simdi = new Date();
    var buAy = simdi.getMonth();
    var buYil = simdi.getFullYear();
    var oncekiAy = buAy === 0 ? 11 : buAy - 1;
    var oncekiYil = buAy === 0 ? buYil - 1 : buYil;

    var data = sh.getDataRange().getValues();
    var buAyToplam = {}, oncekiAyToplam = {};

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][2] || "").toUpperCase() !== "GİDER") continue;
      var dt = data[i][1] instanceof Date ? data[i][1]
        : new Date(String(data[i][1]).split(".").reverse().join("-"));
      if (isNaN(dt)) continue;
      var kat = String(data[i][3] || "Diğer");
      var tutar = parseFloat(data[i][6]) || 0;

      if (dt.getMonth() === buAy && dt.getFullYear() === buYil) {
        buAyToplam[kat] = (buAyToplam[kat] || 0) + tutar;
      }
      if (dt.getMonth() === oncekiAy && dt.getFullYear() === oncekiYil) {
        oncekiAyToplam[kat] = (oncekiAyToplam[kat] || 0) + tutar;
      }
    }

    var alarmlar = [];

    // Limit aşımı kontrolü
    for (var kat in buAyToplam) {
      var limit = LIMITLER[kat] || 5000;
      if (buAyToplam[kat] > limit) {
        alarmlar.push({
          tip      : "LIMIT_ASIMI",
          kategori : kat,
          harcanan : Math.round(buAyToplam[kat]),
          limit    : limit,
          asim     : Math.round(buAyToplam[kat] - limit),
          mesaj    : kat + " kategorisi limiti " + Math.round(buAyToplam[kat] - limit).toLocaleString("tr-TR") + " TL aştı!"
        });
      }
    }

    // Anormal artış kontrolü (geçen aya göre %50+ artış)
    for (var kat in buAyToplam) {
      var onceki = oncekiAyToplam[kat] || 0;
      if (onceki === 0) continue;
      var artis = ((buAyToplam[kat] - onceki) / onceki) * 100;
      if (artis >= 50) {
        alarmlar.push({
          tip      : "ANORMAL_ARTIS",
          kategori : kat,
          harcanan : Math.round(buAyToplam[kat]),
          onceki   : Math.round(onceki),
          artisYuzde: Math.round(artis),
          mesaj    : kat + " geçen aya göre %" + Math.round(artis) + " arttı!"
        });
      }
    }

    alarmlar.sort(function(a, b) { return b.harcanan - a.harcanan; });

    return {
      ok         : true,
      alarmlar   : alarmlar,
      alarmSayisi: alarmlar.length,
      buAyGider  : buAyToplam,
      oncekiAyGider: oncekiAyToplam
    };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_yapayzeka_harcamaAnalizi() {
  try {
    var GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!GEMINI_API_KEY) return { ok: false, msg: "API_KEY tanımlı değil" };

    // Kaçak radar verisini al
    var radar = api_kacakRadariKontrol();
    if (!radar.ok) return { ok: false, msg: radar.msg };

    // Kar/zarar verisini al
    var karZarar = api_karZararHesapla("aylik");
    if (!karZarar.ok) return { ok: false, msg: karZarar.msg };

    // Gemini'ye gönderilecek özet
    var ozet = "Bu ay gelir: " + karZarar.gelir.toLocaleString("tr-TR") + " TL\n"
      + "Bu ay gider: " + karZarar.gider.toLocaleString("tr-TR") + " TL\n"
      + "Net kar/zarar: " + karZarar.net.toLocaleString("tr-TR") + " TL\n"
      + "Geçen aya göre değişim: %" + karZarar.degisim + "\n\n"
      + "Kategori bazlı giderler:\n";

    for (var kat in radar.buAyGider) {
      ozet += "- " + kat + ": " + Math.round(radar.buAyGider[kat]).toLocaleString("tr-TR") + " TL";
      if (radar.oncekiAyGider[kat]) {
        ozet += " (geçen ay: " + Math.round(radar.oncekiAyGider[kat]).toLocaleString("tr-TR") + " TL)";
      }
      ozet += "\n";
    }

    if (radar.alarmlar.length > 0) {
      ozet += "\nTespit edilen alarmlar:\n";
      radar.alarmlar.forEach(function(a) {
        ozet += "- " + a.mesaj + "\n";
      });
    }

    var prompt = "Sen bir şirket mali müşavirisin. Aşağıdaki OSGB şirketinin bu ayki mali verilerini analiz et.\n\n"
      + ozet + "\n\n"
      + "Lütfen şunları yap:\n"
      + "1. Kritik uyarıları listele\n"
      + "2. Tasarruf önerilerini listele\n"
      + "3. Genel mali sağlık değerlendirmesi yap\n"
      + "Türkçe, kısa ve net cevap ver. Madde madde yaz.";

    var url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + GEMINI_API_KEY;
    var payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.3, maxOutputTokens: 1024 }
    };

    var response = UrlFetchApp.fetch(url, {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var json = JSON.parse(response.getContentText());
    if (!json || !json.candidates || !json.candidates[0]) {
  Logger.log("AvrupaBirAI yanıtı: " + response.getContentText());
  return { ok: false, msg: "AvrupaAI yanıt vermedi: " + response.getContentText().substring(0, 200) };
}
var analiz = json.candidates[0].content.parts[0].text;

    // Sonucu logla
    logYaz("YZ_ANALIZ", "Harcama analizi tamamlandı. Alarm sayısı: " + radar.alarmSayisi, "FINANS", "", "", "");

    return {
      ok        : true,
      analiz    : analiz,
      alarmlar  : radar.alarmlar,
      karZarar  : karZarar,
      olusturmaTarihi: Utilities.formatDate(new Date(), "Europe/Istanbul", "dd.MM.yyyy HH:mm")
    };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_aylikOdemeSkor() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    
    // 1. Sabit gider toplamı (hedef)
    var hedef = 0;
    var kalemler = [];
    try {
      var sgSh = ss.getSheetByName("SABIT_GIDER");
      if (sgSh && sgSh.getLastRow() > 1) {
        var sgData = sgSh.getDataRange().getValues();
        for (var s = 1; s < sgData.length; s++) {
          var aktif = sgData[s][4];
          if (aktif === true || String(aktif).toUpperCase() === "TRUE") {
            var t = parseFloat(sgData[s][3]) || 0;
            hedef += t;
            kalemler.push({ kategori: String(sgData[s][1] || ""), aciklama: String(sgData[s][2] || ""), tutar: t });
          }
        }
      }
    } catch(e) {}

    // 2. Bu ay toplam gelir (FİNANS_CARİ'den)
    var shCari = ss.getSheetByName("FİNANS_CARİ");
    var ayGelir = 0, ayGider = 0;
    if (shCari && shCari.getLastRow() > 1) {
      var data = shCari.getDataRange().getValues();
      var simdi = new Date(), buAy = simdi.getMonth(), buYil = simdi.getFullYear();
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        var dt = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
        if (isNaN(dt) || dt.getMonth() !== buAy || dt.getFullYear() !== buYil) continue;
        var tip = String(data[i][2] || "").toUpperCase();
        var tutar = parseFloat(data[i][6]) || 0;
        if (tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1) ayGelir += tutar;
        else ayGider += tutar;
      }
    }

    // 3. Skor hesapla
    var skor = hedef > 0 ? Math.min(Math.round((ayGelir / hedef) * 100), 200) : 0;
    var kalan = hedef - ayGelir;

    return {
      ok: true,
      hedef: Math.round(hedef * 100) / 100,
      ayGelir: Math.round(ayGelir * 100) / 100,
      ayGider: Math.round(ayGider * 100) / 100,
      skor: skor,
      kalan: Math.round(kalan * 100) / 100,
      kalemler: kalemler
    };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_ceoFinansOzet() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var res = api_kasaListesiGetir();
    var kasalar = res.kasalar || [];
    var nakitBakiye = 0, kkartiBakiye = 0, havaleBakiye = 0;
    var bankaListesi = [];

    kasalar.forEach(function(k) {
      var ad = String(k.kasa || "").toUpperCase();
      if (ad === "NAKİT") nakitBakiye = k.bakiye || 0;
      else if (ad === "KREDİ KARTI") kkartiBakiye = k.bakiye || 0;
      else {
        havaleBakiye += k.bakiye || 0;
        bankaListesi.push({ ad: k.kasa, bakiye: k.bakiye || 0, durum: "AKTİF" });
      }
    });

    // Açık alacak — FİRMA_CARİ
    var toplamFatura = 0, toplamTahsilat = 0, kalanAlacak = 0;
    try {
      var shFC = ss.getSheetByName("FİRMA_CARİ");
      if (shFC && shFC.getLastRow() > 1) {
        var fcData = shFC.getDataRange().getValues();
        var firmaBakiye = {};
        for (var j = 1; j < fcData.length; j++) {
          var fId = String(fcData[j][2] || "").trim();
          var tip = String(fcData[j][4] || "").toUpperCase();
        var borc = parseFloat(fcData[j][5]) || 0;
        var alacak = parseFloat(fcData[j][6]) || 0;
        if (tip === "ALACAK" || tip === "FATURA") toplamFatura += alacak;
        if (tip === "TAHSİLAT") toplamTahsilat += borc;
        if (fId) firmaBakiye[fId] = parseFloat(fcData[j][7]) || 0;
        }
        for (var f in firmaBakiye) {
          if (firmaBakiye[f] > 0) kalanAlacak += firmaBakiye[f];
        }
      }
    } catch(e2) {}

    var toplamKullanilabilir = nakitBakiye + havaleBakiye;

    return {
      nakitBakiye: Math.round(nakitBakiye * 100) / 100,
      bankaAktif: Math.round(havaleBakiye * 100) / 100,
      bankaBloke: 0,
      toplamKullanilabilir: Math.round(toplamKullanilabilir * 100) / 100,
      kkartiBakiye: Math.round(kkartiBakiye * 100) / 100,
      toplamFatura: Math.round(toplamFatura * 100) / 100,
      toplamTahsilat: Math.round(toplamTahsilat * 100) / 100,
      kalanAlacak: Math.round(kalanAlacak * 100) / 100,
      bankaListesi: bankaListesi
    };
  } catch(e) { return {}; }
}

function api_sonIslemLoglari(limit) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("LOGS");
    if (!sh || sh.getLastRow() < 2) return { ok: true, liste: [] };
    var data = sh.getDataRange().getValues();
    var liste = [];
    for (var i = data.length - 1; i >= 1; i--) {
      var islem = String(data[i][10] || "");
      if (islem.indexOf("FINANS") < 0 && islem.indexOf("GELIR") < 0 && islem.indexOf("GIDER") < 0 && islem.indexOf("KASA") < 0) continue;
      liste.push({
        tarih: String(data[i][9] || ""),
        kullanici: String(data[i][1] || ""),
        islem: islem,
        detay: String(data[i][5] || ""),
        entity: String(data[i][11] || "")
      });
      if (liste.length >= (limit || 20)) break;
    }
    return { ok: true, liste: liste };
  } catch(e) { return { ok: false, liste: [], msg: e.message }; }
}

function api_gelirKacagiTespit() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("ISG_HIZMETLER");
    if (!sh || sh.getLastRow() < 2) return { ok: true, liste: [], toplam: 0 };

    var data = sh.getDataRange().getValues();
    var head = data[0];
    var bugun = new Date(); bugun.setHours(0,0,0,0);
    var liste = [];

    for (var i = 1; i < data.length; i++) {
      var durum = String(data[i][11] || "").toUpperCase(); // L = Durum
      if (durum !== "AKTİF" && durum !== "AKTIF") continue;

      var firma = String(data[i][1] || ""); // B = FirmaAdi
      var sozBit = data[i][10]; // K = SozlesmeBit
      var sonZiyaret = data[i][25]; // Z = SonZiyaret
      var calisan = parseInt(data[i][3]) || 0; // D = CalisanSayisi

      // Tarihleri parse et
      var bitTarih = sozBit instanceof Date ? sozBit : new Date(String(sozBit));
      var ziyaretTarih = sonZiyaret instanceof Date ? sonZiyaret : new Date(String(sonZiyaret));

      // 1970 = hiç girilmemiş
      var sozBitGecerli = !isNaN(bitTarih) && bitTarih.getFullYear() > 2000;
      var ziyaretGecerli = !isNaN(ziyaretTarih) && ziyaretTarih.getFullYear() > 2000;

      var sorunlar = [];

      // Sözleşme bitmiş ama hâlâ AKTİF
      if (sozBitGecerli && bitTarih < bugun) {
        var gecikmeGun = Math.round((bugun - bitTarih) / (1000*60*60*24));
        sorunlar.push("Sözleşme " + gecikmeGun + " gün önce bitti");
      }

      // Sözleşme hiç tanımlanmamış (1970)
      if (!sozBitGecerli) {
        sorunlar.push("Sözleşme tarihi tanımsız");
      }

      // 60 günden fazla ziyaret yok
      if (ziyaretGecerli) {
        var ziyaretFark = Math.round((bugun - ziyaretTarih) / (1000*60*60*24));
        if (ziyaretFark > 60) {
          sorunlar.push(ziyaretFark + " gündür ziyaret yok");
        }
      }

      if (sorunlar.length > 0) {
        liste.push({
          firma: firma,
          calisan: calisan,
          sozBit: sozBitGecerli ? Utilities.formatDate(bitTarih, "Europe/Istanbul", "dd.MM.yyyy") : "Tanımsız",
          sonZiyaret: ziyaretGecerli ? Utilities.formatDate(ziyaretTarih, "Europe/Istanbul", "dd.MM.yyyy") : "Yok",
          sorunlar: sorunlar
        });
      }
    }

    liste.sort(function(a, b) { return b.sorunlar.length - a.sorunlar.length; });

    return { ok: true, liste: liste.slice(0, 50), toplam: liste.length };
  } catch(e) { return { ok: false, liste: [], toplam: 0, msg: e.message }; }
}

function api_firmaKarlilikAnalizi() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, liste: [] };

    var data = sh.getDataRange().getValues();
    var firmaMap = {};

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var firma = String(data[i][5] || "").trim();
      if (!firma || firma === "-") continue;
      var tip = String(data[i][2] || "").toUpperCase();
      var tutar = parseFloat(data[i][6]) || 0;
      var gelirMi = tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1;

      if (!firmaMap[firma]) firmaMap[firma] = { gelir: 0, gider: 0, islem: 0 };
      if (gelirMi) firmaMap[firma].gelir += tutar;
      else firmaMap[firma].gider += tutar;
      firmaMap[firma].islem++;
    }

    // ISG ziyaret maliyeti tahmini (personel başına günlük maliyet)
    var GUNLUK_MALIYET = 1500; // TL — yaklaşık personel+yol+zaman
    try {
      var shISG = ss.getSheetByName("ISG_HIZMETLER");
      if (shISG && shISG.getLastRow() > 1) {
        var isgData = shISG.getDataRange().getValues();
        for (var j = 1; j < isgData.length; j++) {
          var fAdi = String(isgData[j][1] || "").trim();
          var aylikDk = (parseFloat(isgData[j][6]) || 0) + (parseFloat(isgData[j][7]) || 0) + (parseFloat(isgData[j][8]) || 0);
          if (!fAdi || aylikDk <= 0) continue;
          var aylikZiyaret = Math.ceil(aylikDk / 480); // 480dk = 1 tam gün
          var aylikMaliyet = aylikZiyaret * GUNLUK_MALIYET;
          if (!firmaMap[fAdi]) firmaMap[fAdi] = { gelir: 0, gider: 0, islem: 0 };
          firmaMap[fAdi].gider += aylikMaliyet;
        }
      }
    } catch(e2) {}

    var liste = [];
    for (var f in firmaMap) {
      var net = firmaMap[f].gelir - firmaMap[f].gider;
      var marj = firmaMap[f].gelir > 0 ? Math.round((net / firmaMap[f].gelir) * 100) : -100;
      liste.push({
        firma: f,
        gelir: Math.round(firmaMap[f].gelir),
        gider: Math.round(firmaMap[f].gider),
        net: Math.round(net),
        marj: marj,
        islem: firmaMap[f].islem
      });
    }

    liste.sort(function(a, b) { return a.net - b.net; }); // En zararlı üste

    return { ok: true, liste: liste.slice(0, 30) };
  } catch(e) { return { ok: false, liste: [], msg: e.message }; }
}

function api_nakitAkisTahmini() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var simdi = new Date();

    // 1. Mevcut toplam kasa
    var kasaRes = api_kasaListesiGetir();
    var mevcutKasa = 0;
    (kasaRes.kasalar || []).forEach(function(k) { mevcutKasa += k.bakiye || 0; });

    // 2. Aylık sabit gider
    var sabitGider = 0;
    try {
      var sgSh = ss.getSheetByName("SABIT_GIDER");
      if (sgSh && sgSh.getLastRow() > 1) {
        var sgData = sgSh.getDataRange().getValues();
        for (var s = 1; s < sgData.length; s++) {
          var aktif = sgData[s][4];
          if (aktif === true || String(aktif).toUpperCase() === "TRUE") {
            sabitGider += parseFloat(sgData[s][3]) || 0;
          }
        }
      }
    } catch(e) {}

    // 3. Son 3 ay ortalama gelir/gider (FİNANS_CARİ)
    var shCari = ss.getSheetByName("FİNANS_CARİ");
    var aylikGelirler = {}, aylikGiderler = {};
    if (shCari && shCari.getLastRow() > 1) {
      var data = shCari.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0]) continue;
        var dt = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
        if (isNaN(dt)) continue;
        var anahtar = dt.getFullYear() + "-" + (dt.getMonth() + 1);
        var tip = String(data[i][2] || "").toUpperCase();
        var tutar = parseFloat(data[i][6]) || 0;
        var gelirMi = tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1;
        if (gelirMi) aylikGelirler[anahtar] = (aylikGelirler[anahtar] || 0) + tutar;
        else aylikGiderler[anahtar] = (aylikGiderler[anahtar] || 0) + tutar;
      }
    }

    // Son 3 ay ortalaması
    var son3Ay = [];
    for (var m = 1; m <= 3; m++) {
      var d = new Date(simdi.getFullYear(), simdi.getMonth() - m, 1);
      son3Ay.push(d.getFullYear() + "-" + (d.getMonth() + 1));
    }
    var ortGelir = 0, ortGider = 0, aySayisi = 0;
    son3Ay.forEach(function(a) {
      if (aylikGelirler[a] || aylikGiderler[a]) {
        ortGelir += aylikGelirler[a] || 0;
        ortGider += aylikGiderler[a] || 0;
        aySayisi++;
      }
    });
    if (aySayisi > 0) { ortGelir = ortGelir / aySayisi; ortGider = ortGider / aySayisi; }

    // 4. Bekleyen çekler (30/60/90 gün içinde vadesi gelen)
    var cek30 = 0, cek60 = 0, cek90 = 0;
    try {
      var shCek = ss.getSheetByName("CEKLER_SENETLER");
      if (shCek && shCek.getLastRow() > 1) {
        var cekData = shCek.getDataRange().getValues();
        for (var c = 1; c < cekData.length; c++) {
          if (String(cekData[c][6] || "").toUpperCase() === "ODENDİ") continue;
          var vade = cekData[c][2] instanceof Date ? cekData[c][2] : new Date(String(cekData[c][2]).split(".").reverse().join("-"));
          if (isNaN(vade)) continue;
          var farkGun = Math.round((vade - simdi) / (1000*60*60*24));
          var ct = parseFloat(cekData[c][4]) || 0;
          if (farkGun >= 0 && farkGun <= 30) cek30 += ct;
          else if (farkGun > 30 && farkGun <= 60) cek60 += ct;
          else if (farkGun > 60 && farkGun <= 90) cek90 += ct;
        }
      }
    } catch(e3) {}

    // 5. Tahmin hesapla
    var aylikNet = ortGelir - ortGider;
    var gun30 = mevcutKasa + aylikNet - cek30 - sabitGider;
    var gun60 = gun30 + aylikNet - cek60 - sabitGider;
    var gun90 = gun60 + aylikNet - cek90 - sabitGider;

    return {
      ok: true,
      mevcutKasa: Math.round(mevcutKasa),
      sabitGider: Math.round(sabitGider),
      ortGelir: Math.round(ortGelir),
      ortGider: Math.round(ortGider),
      aylikNet: Math.round(aylikNet),
      cekler: { gun30: Math.round(cek30), gun60: Math.round(cek60), gun90: Math.round(cek90) },
      tahmin: { gun30: Math.round(gun30), gun60: Math.round(gun60), gun90: Math.round(gun90) }
    };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_sozlesmeYenilemeUyari() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("ISG_HIZMETLER");
    if (!sh || sh.getLastRow() < 2) return { ok: true, yaklasan: [], biten: [] };

    var data = sh.getDataRange().getValues();
    var bugun = new Date(); bugun.setHours(0,0,0,0);
    var yaklasan = [], biten = [];

    for (var i = 1; i < data.length; i++) {
      var durum = String(data[i][11] || "").toUpperCase();
      if (durum !== "AKTİF" && durum !== "AKTIF") continue;

      var firma = String(data[i][1] || "");
      var calisan = parseInt(data[i][3]) || 0;
      var sozBit = data[i][10];
      var bitTarih = sozBit instanceof Date ? sozBit : new Date(String(sozBit));
      if (isNaN(bitTarih) || bitTarih.getFullYear() < 2001) continue;

      var kalanGun = Math.round((bitTarih - bugun) / (1000*60*60*24));
      var bitStr = Utilities.formatDate(bitTarih, "Europe/Istanbul", "dd.MM.yyyy");

      if (kalanGun < 0) {
        biten.push({ firma: firma, calisan: calisan, bitTarih: bitStr, gecikme: Math.abs(kalanGun) });
      } else if (kalanGun <= 60) {
        yaklasan.push({ firma: firma, calisan: calisan, bitTarih: bitStr, kalanGun: kalanGun });
      }
    }

    yaklasan.sort(function(a, b) { return a.kalanGun - b.kalanGun; });
    biten.sort(function(a, b) { return b.gecikme - a.gecikme; });

    return { ok: true, yaklasan: yaklasan.slice(0, 20), biten: biten.slice(0, 20), toplamYaklasan: yaklasan.length, toplamBiten: biten.length };
  } catch(e) { return { ok: false, yaklasan: [], biten: [], msg: e.message }; }
}
function api_parasutCariGuncelle() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var shFatura = ss.getSheetByName("PARASUT_FATURALAR");
    if (!shFatura || shFatura.getLastRow() < 2) return { ok: true, msg: "Fatura yok", eklenen: 0 };

    var shCari = ss.getSheetByName("FİRMA_CARİ");
    if (!shCari) return { ok: false, msg: "FİRMA_CARİ bulunamadı" };

    // Mevcut cari ID'leri topla
    var cariData = shCari.getDataRange().getValues();
    var mevcutIdler = {};
    for (var c = 1; c < cariData.length; c++) {
      mevcutIdler[String(cariData[c][0] || "")] = true;
    }

    // Firma adı → ID eşleştirme
    var firmaMap = {};
    try {
      var shFirma = ss.getSheetByName("FIRMALAR");
      if (shFirma && shFirma.getLastRow() > 1) {
        var fData = shFirma.getDataRange().getValues();
        for (var fi = 1; fi < fData.length; fi++) {
          var fAdi = String(fData[fi][1] || "").trim().toUpperCase();
          if (fAdi) firmaMap[fAdi] = String(fData[fi][0] || "");
        }
      }
    } catch(ef) {}

    // Firma bakiyeleri hesapla (mevcut cariden)
    var firmaBakiye = {};
    for (var cb = 1; cb < cariData.length; cb++) {
      var fId = String(cariData[cb][2] || "").trim();
      if (fId) firmaBakiye[fId] = parseFloat(cariData[cb][8]) || 0;
    }

    var faturaData = shFatura.getDataRange().getValues();
    var yeniSatirlar = [];
    var eklenen = 0;
    var simdi = Utilities.formatDate(new Date(), "Europe/Istanbul", "dd.MM.yyyy HH:mm");

    for (var i = 1; i < faturaData.length; i++) {
      var parasutId = String(faturaData[i][0] || "");
      var firmaAdi = String(faturaData[i][3] || "").trim();
      var faturaNo = String(faturaData[i][4] || "");
      var faturaTarihi = String(faturaData[i][5] || "");
      var brutTutar = parseFloat(faturaData[i][9]) || 0;

      if (!firmaAdi || brutTutar <= 0) continue;

      var cariId = "PARASUT_" + parasutId;

      // Zaten eklenmişse atla
      if (mevcutIdler[cariId]) continue;

      var firmaId = firmaMap[firmaAdi.toUpperCase()] || "";

      // Mevcut bakiye + yeni fatura tutarı
      var mevcutBak = firmaBakiye[firmaId] || 0;
      var yeniBakiye = mevcutBak + brutTutar;
      firmaBakiye[firmaId] = yeniBakiye;

      yeniSatirlar.push([
        cariId,
        faturaTarihi || simdi,
        firmaId,
        firmaAdi,
        "ALACAK",
        "Fatura: " + faturaNo,
        0,
        brutTutar,
        yeniBakiye,
        faturaNo,
        "PARASUT_SYNC",
        "AÇIK"
      ]);
      mevcutIdler[cariId] = true;
      eklenen++;
    }

    // Toplu yazma
    if (yeniSatirlar.length > 0) {
      shCari.getRange(shCari.getLastRow() + 1, 1, yeniSatirlar.length, 12).setValues(yeniSatirlar);
    }

    Logger.log("Parasut Cari: " + eklenen + " fatura eklendi (sadece alacak, ödeme yok)");
    return { ok: true, msg: eklenen + " fatura cariye eklendi", eklenen: eklenen };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_kasaMutabakat() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, kasalar: [], bugunIslem: 0 };

    var data = sh.getDataRange().getValues();
    var bugun = new Date();
    var bugunStr = ('0' + bugun.getDate()).slice(-2) + '.' + ('0' + (bugun.getMonth() + 1)).slice(-2) + '.' + bugun.getFullYear();
    var kasaMap = {}, bugunGelir = {}, bugunGider = {}, bugunIslem = 0;

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var tip = String(data[i][2] || "").toUpperCase();
      var kasa = String(data[i][4] || "NAKİT").trim();
      var tutar = parseFloat(data[i][6]) || 0;
      var gelirMi = tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1;

      if (!kasaMap[kasa]) kasaMap[kasa] = 0;
      kasaMap[kasa] += gelirMi ? tutar : -tutar;

      // Bugünkü işlemler
      var dt = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
      if (!isNaN(dt) && dt.getDate() === bugun.getDate() && dt.getMonth() === bugun.getMonth() && dt.getFullYear() === bugun.getFullYear()) {
        if (!bugunGelir[kasa]) bugunGelir[kasa] = 0;
        if (!bugunGider[kasa]) bugunGider[kasa] = 0;
        if (gelirMi) bugunGelir[kasa] += tutar;
        else bugunGider[kasa] += tutar;
        bugunIslem++;
      }
    }

    var kasalar = [];
    for (var k in kasaMap) {
      kasalar.push({
        kasa: k,
        sistemBakiye: Math.round(kasaMap[k] * 100) / 100,
        bugunGelir: Math.round((bugunGelir[k] || 0) * 100) / 100,
        bugunGider: Math.round((bugunGider[k] || 0) * 100) / 100,
        bugunNet: Math.round(((bugunGelir[k] || 0) - (bugunGider[k] || 0)) * 100) / 100
      });
    }

    return { ok: true, tarih: bugunStr, kasalar: kasalar, bugunIslem: bugunIslem };
  } catch(e) { return { ok: false, msg: e.message }; }
}

function api_aylikTrend() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName("FİNANS_CARİ");
    if (!sh || sh.getLastRow() < 2) return { ok: true, aylar: [] };

    var data = sh.getDataRange().getValues();
    var simdi = new Date();
    var ayMap = {};

    // Son 6 ay için boş hazırla
    for (var m = 5; m >= 0; m--) {
      var d = new Date(simdi.getFullYear(), simdi.getMonth() - m, 1);
      var anahtar = d.getFullYear() + "-" + ('0' + (d.getMonth() + 1)).slice(-2);
      var ayAdi = ["Oca","Şub","Mar","Nis","May","Haz","Tem","Ağu","Eyl","Eki","Kas","Ara"][d.getMonth()] + " " + d.getFullYear();
      ayMap[anahtar] = { ay: ayAdi, gelir: 0, gider: 0 };
    }

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var dt = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
      if (isNaN(dt)) continue;
      var anahtar2 = dt.getFullYear() + "-" + ('0' + (dt.getMonth() + 1)).slice(-2);
      if (!ayMap[anahtar2]) continue;

      var tip = String(data[i][2] || "").toUpperCase();
      var tutar = parseFloat(data[i][6]) || 0;
      if (tip.indexOf("GELİR") > -1 || tip.indexOf("TAHSİLAT") > -1) ayMap[anahtar2].gelir += tutar;
      else ayMap[anahtar2].gider += tutar;
    }

    var aylar = [];
    var anahtarlar = Object.keys(ayMap).sort();
    anahtarlar.forEach(function(k) {
      var a = ayMap[k];
      a.net = Math.round((a.gelir - a.gider) * 100) / 100;
      a.gelir = Math.round(a.gelir * 100) / 100;
      a.gider = Math.round(a.gider * 100) / 100;
      aylar.push(a);
    });

    return { ok: true, aylar: aylar };
  } catch(e) { return { ok: false, aylar: [], msg: e.message }; }
}

function api_isgUzmanPerformans() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var shPlan = ss.getSheetByName("ISG_PLANLAR");
    var simdi = new Date(), buAy = simdi.getMonth(), buYil = simdi.getFullYear();
    var personelMap = {};

    // ISG planlarından ziyaret verileri
    if (shPlan && shPlan.getLastRow() > 1) {
      var data = shPlan.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var dt = data[i][1] instanceof Date ? data[i][1] : new Date(data[i][1]);
        if (isNaN(dt) || dt.getMonth() !== buAy || dt.getFullYear() !== buYil) continue;
        var email = String(data[i][4] || "").toLowerCase().trim();
        var durum = String(data[i][8] || "").toUpperCase();
        var dk = parseFloat(data[i][7]) || 0;
        if (!email) continue;
        if (!personelMap[email]) personelMap[email] = { ad: String(data[i][5] || email.split("@")[0]), ziyaret: 0, tamamlanan: 0, toplam: 0, dakika: 0, firmalar: {} };
        personelMap[email].toplam++;
        personelMap[email].dakika += dk;
        var firma = String(data[i][3] || "");
        if (firma) personelMap[email].firmalar[firma] = true;
        if (durum === "TAMAMLANDI") {
          personelMap[email].tamamlanan++;
          personelMap[email].ziyaret++;
        }
      }
    }

    // ISG ziyaret kayıtlarından
    var shZiyaret = ss.getSheetByName("ISG_ZIYARETLER");
    if (shZiyaret && shZiyaret.getLastRow() > 1) {
      var zData = shZiyaret.getDataRange().getValues();
      for (var j = 1; j < zData.length; j++) {
        var dt2 = zData[j][1] instanceof Date ? zData[j][1] : new Date(zData[j][1]);
        if (isNaN(dt2) || dt2.getMonth() !== buAy || dt2.getFullYear() !== buYil) continue;
        var email2 = String(zData[j][3] || "").toLowerCase().trim();
        var dk2 = parseFloat(zData[j][6]) || 0;
        var firma2 = String(zData[j][2] || "");
        if (!email2) continue;
        if (!personelMap[email2]) personelMap[email2] = { ad: email2.split("@")[0], ziyaret: 0, tamamlanan: 0, toplam: 0, dakika: 0, firmalar: {} };
        personelMap[email2].ziyaret++;
        personelMap[email2].dakika += dk2;
        if (firma2) personelMap[email2].firmalar[firma2] = true;
      }
    }

    var liste = [];
    for (var e in personelMap) {
      var p = personelMap[e];
      var basariOrani = p.toplam > 0 ? Math.round((p.tamamlanan / p.toplam) * 100) : 0;
      liste.push({
        email: e,
        ad: p.ad,
        ziyaret: p.ziyaret,
        tamamlanan: p.tamamlanan,
        toplam: p.toplam,
        basariOrani: basariOrani,
        dakika: Math.round(p.dakika),
        firmaSayisi: Object.keys(p.firmalar).length
      });
    }
    liste.sort(function(a, b) { return b.ziyaret - a.ziyaret; });

    return { ok: true, liste: liste };
  } catch(e) { return { ok: false, liste: [], msg: e.message }; }
}


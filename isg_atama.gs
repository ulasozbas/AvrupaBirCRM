/* =============================================
   ISG AKILLI ATAMA — BACKEND
   Sheet: ISG_PERSONEL_KAPASITE
   Sütunlar: A:PERSONEL | B:SERTIFIKA | C:SINIF | D:DAKIKA | E:FIRMA | F:SGK_NO | G:CALISAN | H:TEHLIKE | I:TARIH
   ============================================= */

var ISG_KAP_SHEET = 'ISG_PERSONEL_KAPASITE';
var ISG_KAP_LIMIT = 11700;

function api_isgKapasiteOzet() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName(ISG_KAP_SHEET);
    if (!sh) return { ok: false, msg: 'ISG_PERSONEL_KAPASITE sheet bulunamadı. Önce Excel import edin.' };
    var son = sh.getLastRow();
    if (son < 2) return { ok: false, msg: 'Sheet boş. Önce Excel import edin.' };

    var rows = sh.getRange(2, 1, son - 1, 9).getValues();
    var pMap = {};
    var verimsiz = [];
    var tehlike = { 'Az Tehlikeli': 0, 'Tehlikeli': 0, 'Çok Tehlikeli': 0 };
    var firmaSet = {};

    for (var i = 0; i < rows.length; i++) {
      var ad = String(rows[i][0]).trim();
      var sertifika = String(rows[i][1]).trim();
      var sinif = String(rows[i][2]).trim();
      var dk = Number(rows[i][3]) || 0;
      var firma = String(rows[i][4]).trim();
      var tehlikeSinifi = String(rows[i][7]).trim();

      if (!ad) continue;
      firmaSet[firma] = 1;
      if (tehlike.hasOwnProperty(tehlikeSinifi)) tehlike[tehlikeSinifi]++;

      if (!pMap[ad]) {
        pMap[ad] = { ad: ad, sertifika: sertifika, sinif: sinif, dk: 0, firma: 0 };
      }
      pMap[ad].dk += dk;
      pMap[ad].firma++;

      if (sinif === 'A' && tehlikeSinifi === 'Az Tehlikeli') {
        verimsiz.push({ personel: ad, firma: firma, dk: dk, sinif: 'A' });
      }
    }

    var liste = [];
    for (var k in pMap) liste.push(pMap[k]);
    liste.sort(function(a, b) { return b.dk - a.dk; });

    var toplamDk = 0;
    var kritik = 0;
    for (var j = 0; j < liste.length; j++) {
      toplamDk += liste[j].dk;
      if ((liste[j].dk / ISG_KAP_LIMIT * 100) >= 95) kritik++;
    }

    return {
      ok: true,
      liste: liste,
      verimsiz: verimsiz,
      tehlike: tehlike,
      ozet: {
        toplamDk: toplamDk,
        toplamKapasite: ISG_KAP_LIMIT * liste.length,
        doluluk: Math.round(toplamDk / (ISG_KAP_LIMIT * liste.length) * 1000) / 10,
        kritik: kritik,
        limit: ISG_KAP_LIMIT,
        personelSayisi: liste.length,
        firmaSayisi: Object.keys(firmaSet).length
      }
    };
  } catch (e) {
    return { ok: false, msg: 'Hata: ' + e.message };
  }
}

function api_isgKapasiteImport(satirlar) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName(ISG_KAP_SHEET);
    if (!sh) {
      sh = ss.insertSheet(ISG_KAP_SHEET);
      var headers = ['PERSONEL', 'SERTIFIKA', 'SINIF', 'DAKIKA', 'FIRMA', 'SGK_NO', 'CALISAN', 'TEHLIKE', 'TARIH'];
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    if (sh.getLastRow() > 1) {
      sh.getRange(2, 1, sh.getLastRow() - 1, 9).clear();
    }
    if (satirlar && satirlar.length > 0) {
      sh.getRange(2, 1, satirlar.length, satirlar[0].length).setValues(satirlar);
    }
    return { ok: true, msg: satirlar.length + ' kayıt import edildi.' };
  } catch (e) {
    return { ok: false, msg: 'Import hatası: ' + e.message };
  }
}

function api_isgSinifMap() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName('ISG_PERSONEL_KAPASITE');
    if (!sh || sh.getLastRow() < 2) return { ok: true, map: {} };
    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
    var map = {};
    for (var i = 0; i < rows.length; i++) {
      var ad = String(rows[i][0]).trim();
      if (ad && !map[ad]) map[ad] = { sinif: String(rows[i][2]).trim(), dk: Number(rows[i][3]) || 0 };
    }
    var sonuc = {};
    for (var k in map) {
      sonuc[k] = map[k].sinif;
    }
    return { ok: true, map: sonuc };
  } catch (e) {
    return { ok: true, map: {} };
  }
}

/* =============================================
   FAZ 3 — MALİYET / KÂRLILIK ANALİZİ
   ============================================= */

function api_isgKarlilikOzet() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);

    // 1) PERSONEL_IK → maaş bilgisi
    var ikSh = ss.getSheetByName('PERSONEL_IK');
    if (!ikSh || ikSh.getLastRow() < 2) return { ok: false, msg: 'PERSONEL_IK sheet bulunamadı.' };
    var ikData = ikSh.getRange(2, 1, ikSh.getLastRow() - 1, 6).getValues();
    var maasMap = {};
    for (var i = 0; i < ikData.length; i++) {
      var ad = String(ikData[i][1]).trim();
      var maas = Number(String(ikData[i][5]).replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
      if (ad && maas > 0) maasMap[ad] = maas;
    }

    // 2) ISG_PERSONEL_KAPASITE → atama dakikaları
    var kapSh = ss.getSheetByName('ISG_PERSONEL_KAPASITE');
    if (!kapSh || kapSh.getLastRow() < 2) return { ok: false, msg: 'ISG_PERSONEL_KAPASITE sheet boş.' };
    var kapData = kapSh.getRange(2, 1, kapSh.getLastRow() - 1, 9).getValues();

    var perMap = {};
    var firmaPerMap = {};
    for (var j = 0; j < kapData.length; j++) {
      var per = String(kapData[j][0]).trim();
      var sinif = String(kapData[j][2]).trim();
      var dk = Number(kapData[j][3]) || 0;
      var firma = String(kapData[j][4]).trim();
      if (!per) continue;

      if (!perMap[per]) perMap[per] = { ad: per, sinif: sinif, toplamDk: 0, firmaSayisi: 0, firmalar: {} };
      perMap[per].toplamDk += dk;
      if (!perMap[per].firmalar[firma]) {
        perMap[per].firmalar[firma] = 0;
        perMap[per].firmaSayisi++;
      }
      perMap[per].firmalar[firma] += dk;

      if (!firmaPerMap[firma]) firmaPerMap[firma] = [];
      firmaPerMap[firma].push({ personel: per, dk: dk });
    }

    // 3) PARASUT_FATURALAR → firma gelirleri
    var fatSh = ss.getSheetByName('PARASUT_FATURALAR');
    var gelirMap = {};
    if (fatSh && fatSh.getLastRow() > 1) {
      var fatHeaders = fatSh.getRange(1, 1, 1, fatSh.getLastColumn()).getValues()[0];
      var colFirma = -1, colBrut = -1, colDurum = -1;
      for (var h = 0; h < fatHeaders.length; h++) {
        var hdr = String(fatHeaders[h]).trim();
        if (hdr === 'FirmaAdi') colFirma = h;
        if (hdr === 'BrutTutar') colBrut = h;
        if (hdr === 'Durum') colDurum = h;
      }
      if (colFirma > -1 && colBrut > -1) {
        var fatData = fatSh.getRange(2, 1, fatSh.getLastRow() - 1, fatSh.getLastColumn()).getValues();
        for (var f = 0; f < fatData.length; f++) {
          var fAdi = String(fatData[f][colFirma]).trim();
          var brut = Number(fatData[f][colBrut]) || 0;
          if (fAdi && brut > 0) {
            if (!gelirMap[fAdi]) gelirMap[fAdi] = 0;
            gelirMap[fAdi] += brut;
          }
        }
      }
    }

    // 4) Personel P&L hesapla
    var sonuc = [];
    for (var k in perMap) {
      var p = perMap[k];
      var maas = 0;
      // İsim eşleştirme (büyük/küçük harf)
      for (var m in maasMap) {
        if (m.toUpperCase().trim() === k.toUpperCase().trim()) { maas = maasMap[m]; break; }
      }

      // Personele ait firma gelirlerini hesapla
      var toplamGelir = 0;
      var firmaDetay = [];
      for (var firma in p.firmalar) {
        var fDk = p.firmalar[firma];
        // Firma gelirini bu personele oranla
        var firmaGelir = gelirMap[firma] || 0;
        var firmaTopDk = 0;
        if (firmaPerMap[firma]) {
          for (var x = 0; x < firmaPerMap[firma].length; x++) firmaTopDk += firmaPerMap[firma][x].dk;
        }
        var payGelir = firmaTopDk > 0 ? Math.round(firmaGelir * fDk / firmaTopDk) : 0;
        toplamGelir += payGelir;
        if (payGelir > 0 || fDk > 0) {
          firmaDetay.push({ firma: firma, dk: fDk, gelir: payGelir });
        }
      }

      // Dakika başına maliyet
      var dkMaliyet = p.toplamDk > 0 ? Math.round(maas / p.toplamDk * 100) / 100 : 0;
      var dkGelir = p.toplamDk > 0 ? Math.round(toplamGelir / p.toplamDk * 100) / 100 : 0;
      var kar = toplamGelir - maas;
      var durum = kar > 0 ? 'KARLI' : (kar === 0 ? 'DENGEDE' : 'ZARAR');
      if (maas === 0 && toplamGelir === 0) durum = 'VERİYOK';

      sonuc.push({
        ad: k,
        sinif: p.sinif,
        toplamDk: p.toplamDk,
        firmaSayisi: p.firmaSayisi,
        maas: maas,
        gelir: toplamGelir,
        kar: kar,
        dkMaliyet: dkMaliyet,
        dkGelir: dkGelir,
        durum: durum,
        firmaDetay: firmaDetay.sort(function(a, b) { return b.gelir - a.gelir; }).slice(0, 10)
      });
    }

    sonuc.sort(function(a, b) { return b.kar - a.kar; });

    var topGelir = 0, topMaas = 0;
    for (var s = 0; s < sonuc.length; s++) { topGelir += sonuc[s].gelir; topMaas += sonuc[s].maas; }

    return {
      ok: true,
      liste: sonuc,
      ozet: {
        toplamGelir: topGelir,
        toplamMaas: topMaas,
        toplamKar: topGelir - topMaas,
        personelSayisi: sonuc.length
      }
    };
  } catch (e) {
    return { ok: false, msg: 'Hata: ' + e.message };
  }
}

/* =============================================
   FAZ 4 — GEMİNİ AI ATAMA ÖNERİSİ
   ============================================= */

function api_isgAtamaOneri(firmaAdi, tehlikeSinifi, calisanSayisi) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) return { ok: false, msg: "Gemini API anahtarı bulunamadı" };

    // Kapasite verisini çek
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var kapSh = ss.getSheetByName('ISG_PERSONEL_KAPASITE');
    var perMap = {};
    if (kapSh && kapSh.getLastRow() > 1) {
      var rows = kapSh.getRange(2, 1, kapSh.getLastRow() - 1, 4).getValues();
      for (var i = 0; i < rows.length; i++) {
        var ad = String(rows[i][0]).trim();
        var sinif = String(rows[i][2]).trim();
        var dk = Number(rows[i][3]) || 0;
        if (!ad) continue;
        if (!perMap[ad]) perMap[ad] = { sinif: sinif, dk: 0 };
        perMap[ad].dk += dk;
      }
    }

    // Maaş verisini çek
    var ikSh = ss.getSheetByName('PERSONEL_IK');
    var maasMap = {};
    if (ikSh && ikSh.getLastRow() > 1) {
      var ikRows = ikSh.getRange(2, 1, ikSh.getLastRow() - 1, 6).getValues();
      for (var m = 0; m < ikRows.length; m++) {
        var ikAd = String(ikRows[m][1]).trim();
        var maas = Number(String(ikRows[m][5]).replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
        if (ikAd) maasMap[ikAd] = maas;
      }
    }

    // Personel özet listesi oluştur
    var personelOzet = [];
    for (var k in perMap) {
      var p = perMap[k];
      var kalan = 11700 - p.dk;
      var pct = Math.round(p.dk / 11700 * 100);
      var m2 = 0;
      for (var mk in maasMap) {
        if (mk.toUpperCase().trim() === k.toUpperCase().trim()) { m2 = maasMap[mk]; break; }
      }
      personelOzet.push(k + " | " + p.sinif + " | %" + pct + " dolu | " + kalan + " dk boş | Maaş: ₺" + m2);
    }

    var prompt = "Sen bir ISG (İş Sağlığı ve Güvenliği) uzman atama danışmanısın.\n\n" +
      "FİRMA BİLGİSİ:\n" +
      "- Firma: " + firmaAdi + "\n" +
      "- Tehlike Sınıfı: " + tehlikeSinifi + "\n" +
      "- Çalışan Sayısı: " + (calisanSayisi || "Bilinmiyor") + "\n\n" +
      "YASAL KURALLAR:\n" +
      "- C Sınıfı İGU → Sadece Az Tehlikeli\n" +
      "- B Sınıfı İGU → Az Tehlikeli + Tehlikeli\n" +
      "- A Sınıfı İGU → Hepsi\n" +
      "- Hekim → Hepsi\n\n" +
      "MEVCUT PERSONEL (Ad | Sınıf | Doluluk | Boş Kapasite | Maaş):\n" +
      personelOzet.join("\n") + "\n\n" +
      "Bu firmaya İGU ve Hekim ataması öner. Şu JSON formatında döndür (başka hiçbir şey yazma):\n" +
      "{\n" +
      "  \"iguOneri\": {\"ad\": \"Önerilen kişi\", \"sebep\": \"Kısa neden\"},\n" +
      "  \"hekimOneri\": {\"ad\": \"Önerilen kişi\", \"sebep\": \"Kısa neden\"},\n" +
      "  \"tahminiDk\": {\"igu\": 120, \"hekim\": 60},\n" +
      "  \"uyarilar\": [\"Varsa uyarı mesajları\"],\n" +
      "  \"optimizasyon\": \"Maliyet veya kapasite ile ilgili kısa öneri\"\n" +
      "}";

    var payload = {
      contents: [{ parts: [{ text: prompt }] }]
    };

    var response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey,
      {
        method: "POST",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      }
    );

    var json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0]) return { ok: false, msg: "AI yanıt vermedi" };
    var text = json.candidates[0].content.parts[0].text;
    text = text.replace(/```json|```/g, "").trim();
    var sonuc = JSON.parse(text);

    return { ok: true, oneri: sonuc };
  } catch (e) {
    return { ok: false, msg: "AI Hata: " + e.message };
  }
}

/* =============================================
   FAZ 5 — COĞRAFİ OPTİMİZASYON
   ============================================= */

function api_isgCografiVeri() {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var sh = ss.getSheetByName('ISG_HIZMETLER');
    if (!sh || sh.getLastRow() < 2) return { ok: false, msg: 'ISG_HIZMETLER sheet bulunamadı.' };

    var rows = sh.getRange(2, 1, sh.getLastRow() - 1, 27).getValues();
    var firmalar = [];
    var perMap = {};
    var gpsYok = 0;
    var gpsVar = 0;

    for (var i = 0; i < rows.length; i++) {
      var firmaId = String(rows[i][0]).trim();
      var firmaAdi = String(rows[i][1]).trim();
      var tehlike = String(rows[i][2]).trim();
      var calisan = Number(rows[i][3]) || 0;
      var iguDk = Number(rows[i][6]) || 0;
      var hekimDk = Number(rows[i][7]) || 0;
      var adres = String(rows[i][17]).trim();
      var enlem = Number(rows[i][18]) || 0;
      var boylam = Number(rows[i][19]) || 0;
      var personel = String(rows[i][14]).trim();
      var hekim = String(rows[i][21]).trim();
      var durum = String(rows[i][11]).trim();

      if (!firmaAdi) continue;

      var hasGps = enlem > 0 && boylam > 0;
      if (hasGps) gpsVar++; else gpsYok++;

      var firma = {
        id: firmaId,
        ad: firmaAdi,
        tehlike: tehlike,
        calisan: calisan,
        iguDk: iguDk,
        hekimDk: hekimDk,
        adres: adres,
        enlem: enlem,
        boylam: boylam,
        personel: personel,
        hekim: hekim,
        durum: durum,
        hasGps: hasGps
      };
      firmalar.push(firma);

      if (personel) {
        if (!perMap[personel]) perMap[personel] = { firmalar: [], topDk: 0 };
        perMap[personel].firmalar.push({ ad: firmaAdi, enlem: enlem, boylam: boylam, dk: iguDk, tehlike: tehlike, hasGps: hasGps });
        perMap[personel].topDk += iguDk;
      }
    }

    var perListe = [];
    for (var k in perMap) {
      var gpsli = perMap[k].firmalar.filter(function(f) { return f.hasGps; });
      perListe.push({
        ad: k,
        firmaSayisi: perMap[k].firmalar.length,
        gpsli: gpsli.length,
        topDk: perMap[k].topDk,
        firmalar: perMap[k].firmalar
      });
    }
    perListe.sort(function(a, b) { return b.firmaSayisi - a.firmaSayisi; });

    return {
      ok: true,
      firmalar: firmalar.filter(function(f) { return f.hasGps; }),
      personeller: perListe,
      ozet: {
        toplamFirma: firmalar.length,
        gpsVar: gpsVar,
        gpsYok: gpsYok,
        personelSayisi: perListe.length
      }
    };
  } catch (e) {
    return { ok: false, msg: 'Hata: ' + e.message };
  }
}

/* =============================================
   FAZ 6 — SİMÜLASYON & TAHMİN
   ============================================= */

function api_isgSimulasyon(aylikFirma, dkFirma, aySayisi) {
  try {
    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var kapSh = ss.getSheetByName('ISG_PERSONEL_KAPASITE');
    if (!kapSh || kapSh.getLastRow() < 2) return { ok: false, msg: 'Kapasite verisi yok.' };
    var rows = kapSh.getRange(2, 1, kapSh.getLastRow() - 1, 4).getValues();
    var perMap = {};
    for (var i = 0; i < rows.length; i++) {
      var ad = String(rows[i][0]).trim();
      var sinif = String(rows[i][2]).trim();
      var dk = Number(rows[i][3]) || 0;
      if (!ad) continue;
      if (!perMap[ad]) perMap[ad] = { sinif: sinif, dk: 0 };
      perMap[ad].dk += dk;
    }

    var topYeniDk = aylikFirma * dkFirma;
    var sonuclar = [];
    var aylikTrend = [];

    for (var k in perMap) {
      var p = perMap[k];
      var kalan = 11700 - p.dk;
      var mevcutPct = Math.round(p.dk / 11700 * 1000) / 10;
      var aylar = [];
      var simDk = p.dk;
      for (var a = 0; a <= aySayisi; a++) {
        aylar.push({ ay: a, dk: simDk, pct: Math.round(simDk / 11700 * 1000) / 10 });
        simDk += Math.round(topYeniDk / Object.keys(perMap).length);
        if (simDk > 11700) simDk = 11700;
      }
      var sonPct = aylar[aylar.length - 1].pct;
      var dolacakAy = -1;
      for (var b = 0; b < aylar.length; b++) {
        if (aylar[b].pct >= 95) { dolacakAy = b; break; }
      }
      var doldurmaFirma = kalan > 0 ? Math.ceil(kalan / dkFirma) : 0;

      sonuclar.push({
        ad: k,
        sinif: p.sinif,
        mevcutDk: p.dk,
        mevcutPct: mevcutPct,
        sonPct: sonPct,
        kalan: kalan,
        dolacakAy: dolacakAy,
        doldurmaFirma: doldurmaFirma,
        aylar: aylar
      });
    }
    sonuclar.sort(function(a, b) { return b.sonPct - a.sonPct; });

    var kritikSayisi = sonuclar.filter(function(s) { return s.dolacakAy > -1; }).length;
    var topKalan = sonuclar.reduce(function(s, p) { return s + p.kalan; }, 0);
    var topAlabilecekFirma = Math.floor(topKalan / dkFirma);

    return {
      ok: true,
      liste: sonuclar,
      ozet: {
        kritikSayisi: kritikSayisi,
        topKalanDk: topKalan,
        topAlabilecekFirma: topAlabilecekFirma,
        personelSayisi: sonuclar.length
      }
    };
  } catch (e) {
    return { ok: false, msg: 'Hata: ' + e.message };
  }
}

function api_isgSimSoru(soru) {
  try {
    var apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!apiKey) return { ok: false, msg: "Gemini API key yok" };

    var ss = SpreadsheetApp.openById(CRM_SHEET_ID);
    var kapSh = ss.getSheetByName('ISG_PERSONEL_KAPASITE');
    var perOzet = [];
    if (kapSh && kapSh.getLastRow() > 1) {
      var rows = kapSh.getRange(2, 1, kapSh.getLastRow() - 1, 4).getValues();
      var pm = {};
      for (var i = 0; i < rows.length; i++) {
        var ad = String(rows[i][0]).trim();
        var sinif = String(rows[i][2]).trim();
        var dk = Number(rows[i][3]) || 0;
        if (!ad) continue;
        if (!pm[ad]) pm[ad] = { sinif: sinif, dk: 0 };
        pm[ad].dk += dk;
      }
      for (var k in pm) {
        var kalan = 11700 - pm[k].dk;
        perOzet.push(k + " | " + pm[k].sinif + " | %" + Math.round(pm[k].dk / 11700 * 100) + " dolu | " + kalan + " dk boş");
      }
    }

    var prompt = "Sen ISG kapasite planlama uzmanısın. Limit: 11.700 dk/kişi/ay.\n\n" +
      "MEVCUT PERSONEL:\n" + perOzet.join("\n") + "\n\n" +
      "YASAL KURALLAR: C→Sadece Az Tehlikeli, B→Az+Tehlikeli, A→Hepsi, Hekim→Hepsi\n\n" +
      "SORU: " + soru + "\n\n" +
      "Kısa ve net Türkçe cevap ver. Sayısal tahminlerle destekle.";

    var response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey,
      { method: "POST", contentType: "application/json", payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }), muteHttpExceptions: true }
    );
    var json = JSON.parse(response.getContentText());
    if (!json.candidates || !json.candidates[0]) return { ok: false, msg: "AI yanıt vermedi" };
    var text = json.candidates[0].content.parts[0].text;
    return { ok: true, cevap: text };
  } catch (e) {
    return { ok: false, msg: "AI Hata: " + e.message };
  }
}


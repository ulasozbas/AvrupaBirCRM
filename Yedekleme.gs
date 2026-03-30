function tumDosyalariYedekle() {
  var scriptId = ScriptApp.getScriptId();
  var url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";
  var token = ScriptApp.getOAuthToken();
  
  var response = UrlFetchApp.fetch(url, {
    method: "GET",
    headers: { "Authorization": "Bearer " + token },
    muteHttpExceptions: true
  });
  
  if (response.getResponseCode() !== 200) {
    Logger.log("HATA: " + response.getContentText());
    return;
  }
  
  var data = JSON.parse(response.getContentText());
  var files = data.files || [];
  
  var tarih = Utilities.formatDate(new Date(), "Europe/Istanbul", "yyyy-MM-dd_HHmm");
  var klasorAdi = "CRM_YEDEK_" + tarih;
  
  var anaKlasor;
  var aramaSonuc = DriveApp.getFoldersByName("AVRUPABIR CRM YEDEK");
  if (aramaSonuc.hasNext()) {
    anaKlasor = aramaSonuc.next();
  } else {
    anaKlasor = DriveApp.createFolder("AVRUPABIR CRM YEDEK");
  }
  
  var yedekKlasor = anaKlasor.createFolder(klasorAdi);
  
  var sayac = 0;
  files.forEach(function(f) {
    var isim = f.name || "isimsiz";
    var tip = f.type || "SERVER_JS";
    var icerik = f.source || "";
    var uzanti = tip === "HTML" ? ".html" : ".gs";
    var dosyaAdi = isim + uzanti;
    
    yedekKlasor.createFile(dosyaAdi, icerik, MimeType.PLAIN_TEXT);
    sayac++;
    Logger.log("✅ " + dosyaAdi + " (" + icerik.length + " karakter)");
  });
  
  Logger.log("═══════════════════════════════════");
  Logger.log("✅ YEDEK TAMAMLANDI!");
  Logger.log("📁 Klasör: " + klasorAdi);
  Logger.log("📄 Dosya sayısı: " + sayac);
  Logger.log("📍 Konum: Drive > AVRUPABIR CRM YEDEK > " + klasorAdi);
}
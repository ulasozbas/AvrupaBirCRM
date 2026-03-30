function yedekDriveaDosyaDosya() {
  var scriptId = ScriptApp.getScriptId();
  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(
    'https://script.googleapis.com/v1/projects/' + scriptId + '/content',
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  
  if (res.getResponseCode() !== 200) {
    Logger.log('HATA: ' + res.getResponseCode());
    Logger.log('Script ID: ' + scriptId);
    Logger.log('API açık mı kontrol edin: https://script.google.com/home/usersettings');
    return;
  }
  
  var data = JSON.parse(res.getContentText());
  if (!data.files) { Logger.log('Dosya yok'); return; }
  
  var folder;
  var folders = DriveApp.getFoldersByName('AvrupaBirCRM_GIT');
  if (folders.hasNext()) {
    folder = folders.next();
    var eski = folder.getFiles();
    while (eski.hasNext()) eski.next().setTrashed(true);
  } else {
    folder = DriveApp.createFolder('AvrupaBirCRM_GIT');
  }
  
  data.files.forEach(function(f) {
    var ext = f.type === 'HTML' ? '.html' : f.type === 'JSON' ? '.json' : '.gs';
    folder.createFile(f.name + ext, f.source, MimeType.PLAIN_TEXT);
  });
  
  Logger.log('✅ ' + data.files.length + ' dosya → AvrupaBirCRM_GIT klasörüne yazıldı');
  Logger.log('📂 ' + folder.getUrl());
}
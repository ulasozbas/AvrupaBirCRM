// ============================================================
//  FATURA_IMPORT.GS — Paraşüt Fatura Import Modülü
//
//  Paraşüt'ten çekilen OSGB/Mobil faturalar bu modülle
//  sisteme aktarılır ve firma carilerine otomatik işlenir.
// ============================================================

/**
 * Paraşütten kopyalanan fatura verilerini toplu işler.
 * Faturalar JSON dizi olarak verilir.
 *
 * @param {Array} faturaListesi - [{faturaNo, tarih, firmaAd, firmaId, hizmetTipi, tutar, kdv}, ...]
 * @returns {Object} İşlem özeti
 */
function parasutFaturalariImport(faturaListesi) {
  if (!Array.isArray(faturaListesi) || faturaListesi.length === 0) {
    throw new Error('Geçerli fatura listesi gönderilmedi.');
  }

  const sonuc = {
    islenen   : 0,
    atlanan   : 0,
    hatalar   : [],
    faturalar : [],
  };

  faturaListesi.forEach((fatura, index) => {
    try {
      const id = _tekFaturaIsle(fatura);
      sonuc.islenen++;
      sonuc.faturalar.push({ id, faturaNo: fatura.faturaNo });
    } catch(e) {
      sonuc.atlanan++;
      sonuc.hatalar.push(`[${index + 1}] ${fatura.faturaNo || '?'}: ${e.message}`);
      Logger.log(`⚠️ Import hatası: ${e.message}`);
    }
  });

  Logger.log(`📦 Paraşüt import tamamlandı: ${sonuc.islenen} işlendi, ${sonuc.atlanan} atlandı.`);
  return sonuc;
}

/**
 * Tek bir faturayı işler (yardımcı).
 */
function _tekFaturaIsle(fatura) {
  const {
    faturaNo,
    tarih,
    firmaAd,
    firmaId,
    hizmetTipi = 'OSGB',
    tutar,
    kdv = 0,
  } = fatura;

  if (!faturaNo)  throw new Error('Fatura numarası boş.');
  if (!firmaAd)   throw new Error('Firma adı boş.');
  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('Geçersiz tutar.');

  const toplam = tutar + kdv;

  // Daha önce import edilmiş mi kontrol et
  if (_faturaVarMi(faturaNo)) {
    throw new Error(`Fatura zaten mevcut: ${faturaNo}`);
  }

  // Firma ID'yi çöz (verilmemişse firma adına göre bul)
  const gercekFirmaId = firmaId || _firmaIdBul(firmaAd);

  // FATURA sheet'e ekle
  const fatSheet = getSheet(SHEETS.FATURA);
  const id = yeniID('F', fatSheet);

  fatSheet.appendRow([
    id,
    faturaNo,
    tarih ? tarihFormat(tarih) : bugun(),
    gercekFirmaId,
    firmaAd,
    hizmetTipi,
    tutar,
    kdv,
    toplam,
    'Açık',
    '',     // Tahsilat tarihi
    true,   // Paraşütten import edildi
  ]);

  // Firma carisine alacak olarak yaz
  if (gercekFirmaId) {
    cariAlacakEkle(gercekFirmaId, toplam, `${hizmetTipi} Fatura: ${faturaNo}`, faturaNo);
  } else {
    // Firma sistemde yoksa yeni kayıt oluştur
    Logger.log(`⚠️ Firma bulunamadı: "${firmaAd}" — Cari oluşturuluyor...`);
    const yeniFirmaId = _firmaOlustur(firmaAd);
    // FATURA sheet'teki firma ID'yi güncelle
    const sonSatir = fatSheet.getLastRow();
    fatSheet.getRange(sonSatir, 4).setValue(yeniFirmaId);
    cariAlacakEkle(yeniFirmaId, toplam, `${hizmetTipi} Fatura: ${faturaNo}`, faturaNo);
  }

  return id;
}

/**
 * Fatura numarasına göre daha önce import edilip edilmediğini kontrol eder.
 */
function _faturaVarMi(faturaNo) {
  const sheet = getSheet(SHEETS.FATURA);
  const data = sheet.getDataRange().getValues().slice(1);
  return data.some(r => r[1] === faturaNo); // 2. sütun: Fatura No
}

/**
 * Firma adına göre FIRMA sheet'ten ID bulur.
 */
function _firmaIdBul(firmaAd) {
  try {
    const sheet = getSheet(SHEETS.FIRMA);
    const data = sheet.getDataRange().getValues().slice(1);
    const satir = data.find(r =>
      r[1].toString().toLowerCase().trim() === firmaAd.toLowerCase().trim()
    );
    return satir ? satir[0] : null;
  } catch(e) {
    return null;
  }
}

/**
 * Sistemde olmayan firmayı otomatik oluşturur.
 */
function _firmaOlustur(firmaAd) {
  const sheet = getSheet(SHEETS.FIRMA);
  const id = yeniID('FR', sheet);
  sheet.appendRow([
    id,
    firmaAd,
    '',    // Vergi no (sonra eklenecek)
    '',    // Yetkili
    '',    // Telefon
    '',    // E-posta
    '',    // Adres
    'OSGB Anlaşma',
    true,  // Aktif
    'Paraşütten otomatik oluşturuldu',
  ]);
  Logger.log(`🏢 Yeni firma oluşturuldu: ${firmaAd} | ID: ${id}`);
  return id;
}

// ── MANUEL FATURA ────────────────────────────────────────────

/**
 * Tek fatura manuel ekler (Paraşüt dışı).
 *
 * @param {Object} params
 */
function faturaManuelEkle(params) {
  const {
    faturaNo,
    firmaId,
    firmaAd,
    hizmetTipi = 'OSGB',
    tutar,
    kdvOrani,
    tarih,
  } = params;

  // Hizmet tipine göre doğru KDV oranını al (config.gs'deki KDV_ORANLARI tablosundan)
  const hesaplanan = kdvHesapla(hizmetTipi, tutar);
  const kdvTutar   = kdvOrani !== undefined ? tutar * (kdvOrani / 100) : hesaplanan.kdvTutar;
  const toplam     = tutar + kdvTutar;

  return _tekFaturaIsle({
    faturaNo,
    tarih     : tarih || bugun(),
    firmaAd,
    firmaId,
    hizmetTipi,
    tutar,
    kdv       : kdvTutar,
  });
}

// ── FATURA TAHSİLAT ───────────────────────────────────────────

/**
 * Açık faturayı tahsil edildi olarak işaretler.
 * Aynı zamanda firmanın cari borcunu sıfırlar (tahsilat yapar).
 *
 * @param {string} faturaId
 * @param {string} odemeTipi  - ODEME_TIPI sabitlerinden biri
 * @param {string} [bankaKod] - EFT ise hangi banka
 */
function faturaTahsilatYap(faturaId, odemeTipi, bankaKod = '') {
  const sheet = getSheet(SHEETS.FATURA);
  const data = sheet.getDataRange().getValues();

  let satirNo = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === faturaId) { satirNo = i + 1; break; }
  }
  if (satirNo === -1) throw new Error(`Fatura bulunamadı: ${faturaId}`);

  const satirVeri = data[satirNo - 1];
  const firmaId   = satirVeri[COL.FATURA.FIRMA_ID - 1];
  const toplam    = satirVeri[COL.FATURA.TOPLAM - 1];
  const durum     = satirVeri[COL.FATURA.DURUM - 1];

  if (durum === 'Tahsil Edildi') throw new Error('Bu fatura zaten tahsil edilmiş.');

  // Fatura durumunu güncelle
  sheet.getRange(satirNo, COL.FATURA.DURUM).setValue('Tahsil Edildi');
  sheet.getRange(satirNo, COL.FATURA.TAHSILAT_TAR).setValue(bugun());

  // Firma carisinden tahsilat yap (bakiye sıfırlanır)
  cariTahsilatYap(firmaId, toplam, odemeTipi, bankaKod, `Fatura tahsilat: ${faturaId}`, faturaId);

  Logger.log(`✅ Fatura tahsil edildi: ${faturaId} | ${paraBirim(toplam)}`);
}

// ── FATURA LİSTELEME ─────────────────────────────────────────

/**
 * Açık faturaları döner.
 */
function acikFaturalari() {
  const sheet = getSheet(SHEETS.FATURA);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .map(r => ({
      id           : r[0],
      faturaNo     : r[1],
      tarih        : tarihFormat(r[2]),
      firmaId      : r[3],
      firmaAd      : r[4],
      hizmetTipi   : r[5],
      tutar        : r[6],
      kdv          : r[7],
      toplam       : r[8],
      durum        : r[9],
      tahsilatTar  : tarihFormat(r[10]),
      parasuten    : r[11],
    }))
    .filter(f => f.durum === 'Açık');
}
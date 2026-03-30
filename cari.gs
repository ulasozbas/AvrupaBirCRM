// ============================================================
//  CARI.GS — Firma Cari Hesap Modülü
//
//  - Anlaşmalı firmalar iş giriş ücretini kendileri karşılar
//  - Tahsilat sırasında cari'ye alacak yazılır
//  - OSGB/Mobil faturalar da cariye işlenir
//  - Aylık fatura kesilince tahsilat yapılır, bakiye sıfırlanır
// ============================================================

// ── CARİ HESAP OKU ───────────────────────────────────────────

/**
 * Bir firmanın mevcut cari bakiyesini döner.
 * Bakiye = Toplam alacak (faturalandırılan) - Toplam tahsilat
 * Pozitif = firma bize borçlu | Negatif = firmadan fazla aldık
 *
 * @param {string} firmaId
 * @returns {number} Cari bakiye
 */
function cariBakiyesiAl(firmaId) {
  const sheet = getSheet(SHEETS.CARI);
  const data = sheet.getDataRange().getValues().slice(1);
  const satirIndex = data.findIndex(r => r[COL.CARI.FIRMA_ID - 1] === firmaId);
  if (satirIndex === -1) return 0;
  return data[satirIndex][COL.CARI.BAKIYE - 1] || 0;
}

/**
 * Tüm firma carilerini döner.
 */
function tumCarileriAl() {
  const sheet = getSheet(SHEETS.CARI);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(r => ({
    id          : r[COL.CARI.ID - 1],
    firmaId     : r[COL.CARI.FIRMA_ID - 1],
    firmaAd     : r[COL.CARI.FIRMA_AD - 1],
    toplamBorc  : r[COL.CARI.TOPLAM_BORC - 1] || 0,
    toplamOdeme : r[COL.CARI.TOPLAM_ODEME - 1] || 0,
    bakiye      : r[COL.CARI.BAKIYE - 1] || 0,
    sonHareket  : tarihFormat(r[COL.CARI.SON_HAREKET - 1]),
    notlar      : r[COL.CARI.NOTLAR - 1] || '',
  }));
}

/**
 * Bakiyesi sıfırdan büyük (bize borçlu) olan carileri döner.
 */
function borcluCariler() {
  return tumCarileriAl().filter(c => c.bakiye > 0);
}

// ── CARİ HAREKET ─────────────────────────────────────────────

/**
 * Cari'ye alacak ekler (müşteri bize borçlandı).
 * Ayaktan tahsilat "Cari" seçilince veya fatura işlenince çağrılır.
 *
 * @param {string} firmaId
 * @param {number} tutar
 * @param {string} aciklama
 * @param {string} [belgeNo]   - Fatura no veya tahsilat ID'si
 */
function cariAlacakEkle(firmaId, tutar, aciklama, belgeNo = '') {
  if (!firmaId) throw new Error('Firma ID boş olamaz.');
  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('Tutar geçersiz.');

  const firmaAd = _firmaAdAl(firmaId);
  _cariHareketEkle(firmaId, firmaAd, 'Alacak', tutar, aciklama, belgeNo);

  Logger.log(`📊 Cari alacak: ${firmaAd} | ${paraBirim(tutar)} | ${belgeNo}`);
}

/**
 * Cariden tahsilat yapar (firma ödeme yaptı, bakiye düşer).
 *
 * @param {string} firmaId
 * @param {number} tutar
 * @param {string} odemeTipi    - 'Nakit' | 'Kredi Kartı (POS)' | 'Havale / EFT'
 * @param {string} [bankaKod]   - EFT ise hangi banka
 * @param {string} [aciklama]
 * @param {string} [belgeNo]
 */
function cariTahsilatYap(firmaId, tutar, odemeTipi, bankaKod = '', aciklama = '', belgeNo = '') {
  if (!firmaId) throw new Error('Firma ID boş olamaz.');
  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('Tutar geçersiz.');

  const mevcutBakiye = cariBakiyesiAl(firmaId);
  const firmaAd = _firmaAdAl(firmaId);

  // Cari'ye tahsilat hareketi yaz
  _cariHareketEkle(firmaId, firmaAd, 'Tahsilat', tutar, aciklama || `Cari tahsilat: ${firmaAd}`, belgeNo);

  // Ödeme tipine göre hesaba yaz
  const hareketAciklama = `Cari tahsilat: ${firmaAd} | ${belgeNo}`;

  switch (odemeTipi) {
    case ODEME_TIPI.NAKIT:
      kasaGiris(hareketAciklama, tutar, 'Cari Tahsilat', belgeNo);
      break;
    case ODEME_TIPI.KART:
      posTahsilatiEkle({ tutar, aciklama: hareketAciklama });
      break;
    case ODEME_TIPI.EFT:
      if (!bankaKod) throw new Error('EFT için banka kodu zorunlu.');
      eftGiris(bankaKod, tutar, hareketAciklama, firmaAd);
      break;
    default:
      throw new Error(`Geçersiz ödeme tipi: ${odemeTipi}`);
  }

  Logger.log(`💰 Cari tahsilat: ${firmaAd} | ${paraBirim(tutar)} | Yeni bakiye: ${paraBirim(mevcutBakiye - tutar)}`);
}

/**
 * Cari hareket ekler ve cari özet tablosunu günceller (ortak yardımcı).
 */
function _cariHareketEkle(firmaId, firmaAd, tip, tutar, aciklama, belgeNo = '') {
  // 1. CARI_HAREKET sheet'e satır ekle
  const hareketSheet = getSheet(SHEETS.CARI_HAREKET);
  const hId = yeniID('CH', hareketSheet);
  const mevcutBakiye = cariBakiyesiAl(firmaId);
  const yeniBakiye = tip === 'Alacak'
    ? mevcutBakiye + tutar
    : mevcutBakiye - tutar;

  hareketSheet.appendRow([
    hId,
    bugun(),
    firmaId,
    tip,
    aciklama,
    tutar,
    belgeNo,
    yeniBakiye,
  ]);

  // 2. CARI özet tablosunu güncelle
  _cariOzetGuncelle(firmaId, firmaAd, tip, tutar, yeniBakiye);
}

/**
 * CARI sheet'teki özet satırı günceller (yoksa ekler).
 */
function _cariOzetGuncelle(firmaId, firmaAd, tip, tutar, yeniBakiye) {
  const sheet = getSheet(SHEETS.CARI);
  const data = sheet.getDataRange().getValues();

  // Firma bu sheet'te var mı?
  let satirNo = -1;
  for (let i = 1; i < data.length; i++) {
    if (data[i][COL.CARI.FIRMA_ID - 1] === firmaId) {
      satirNo = i + 1; // 1-tabanlı
      break;
    }
  }

  if (satirNo === -1) {
    // Yeni firma → satır ekle
    const id = yeniID('CR', sheet);
    const baslangicBorc  = tip === 'Alacak'   ? tutar : 0;
    const baslangicOdeme = tip === 'Tahsilat' ? tutar : 0;
    sheet.appendRow([id, firmaId, firmaAd, baslangicBorc, baslangicOdeme, yeniBakiye, bugun(), '']);
  } else {
    // Mevcut satırı güncelle
    const mevcutBorc  = sheet.getRange(satirNo, COL.CARI.TOPLAM_BORC).getValue() || 0;
    const mevcutOdeme = sheet.getRange(satirNo, COL.CARI.TOPLAM_ODEME).getValue() || 0;

    if (tip === 'Alacak') {
      sheet.getRange(satirNo, COL.CARI.TOPLAM_BORC).setValue(mevcutBorc + tutar);
    } else {
      sheet.getRange(satirNo, COL.CARI.TOPLAM_ODEME).setValue(mevcutOdeme + tutar);
    }
    sheet.getRange(satirNo, COL.CARI.BAKIYE).setValue(yeniBakiye);
    sheet.getRange(satirNo, COL.CARI.SON_HAREKET).setValue(bugun());
  }
}

// ── FATURA KESİM ─────────────────────────────────────────────

/**
 * Firmaya aylık fatura keser.
 * Firmanın CARI_HAREKET'indeki açık bakiyesini faturalandırır.
 * (Paraşütten import fonksiyonu fatura_import.gs'de ayrıca var.)
 *
 * @param {string} firmaId
 * @param {string} faturaNo
 * @param {number} [tutarOverride] - Belirtilmezse mevcut cari bakiyesi kullanılır
 * @returns {string} Fatura ID'si
 */
function faturaKes(firmaId, faturaNo, tutarOverride) {
  const mevcutBakiye = cariBakiyesiAl(firmaId);
  const faturalanacakTutar = tutarOverride || mevcutBakiye;

  if (faturalanacakTutar <= 0) {
    throw new Error(`${firmaId} firmasının fatura kesilecek bakiyesi yok. Bakiye: ${paraBirim(mevcutBakiye)}`);
  }

  const firmaAd = _firmaAdAl(firmaId);
  const kdvOrani = parseFloat(ayarOku('KDV_ORANI') || '20') / 100;
  const kdvTutar = faturalanacakTutar * kdvOrani;
  const toplam   = faturalanacakTutar + kdvTutar;

  const fatSheet = getSheet(SHEETS.FATURA);
  const id = yeniID('F', fatSheet);

  fatSheet.appendRow([
    id,
    faturaNo || id,
    bugun(),
    firmaId,
    firmaAd,
    'Çeşitli Hizmetler',
    faturalanacakTutar,
    kdvTutar,
    toplam,
    'Açık',
    '',      // Tahsilat tarihi (henüz yok)
    false,   // Paraşütten mi (hayır, elle kesildi)
  ]);

  Logger.log(`🧾 Fatura kesildi: ${faturaNo} | ${firmaAd} | ${paraBirim(toplam)}`);
  return id;
}

// ── DASHBOARD VERİSİ ──────────────────────────────────────────

function cariDashboardVerisi() {
  const cariler = tumCarileriAl();
  return {
    toplamAlacak : cariler.reduce((s, c) => s + Math.max(0, c.bakiye), 0),
    borcluFirma  : cariler.filter(c => c.bakiye > 0).length,
    tumCariler   : cariler.slice(0, 10),
  };
}

// ── YARDIMCI ─────────────────────────────────────────────────

function _firmaAdAl(firmaId) {
  try {
    const sheet = getSheet(SHEETS.FIRMA);
    const data = sheet.getDataRange().getValues().slice(1);
    const satir = data.find(r => r[0] === firmaId);
    return satir ? satir[1] : firmaId;
  } catch(e) {
    return firmaId;
  }
}
// ============================================================
//  TAHSILAT.GS — Ayaktan Müşteri Tahsilat Modülü
//
//  4 ödeme tipi:
//  A) Nakit      → Kasa'ya giriş
//  B) Kredi Kartı → POS kaydı → İşbank POS hesabına
//  C) Havale/EFT → Seçilen banka hesabına giriş
//  D) Cari       → Firmanın cari hesabına alacak
// ============================================================

// ── ANA TAHSİLAT FONKSİYONU ──────────────────────────────────

/**
 * Ayaktan müşteri tahsilatını kaydeder.
 * Ödeme tipine göre ilgili modülü otomatik tetikler.
 *
 * @param {Object} params
 * @param {string} params.musteriAd     - Hasta/müşteri adı
 * @param {string} params.tcKimlik      - TC kimlik veya pasaport no
 * @param {string} params.hizmet        - Hizmet türü (İşe Giriş Muayenesi vb.)
 * @param {number} params.tutar         - Tutar
 * @param {string} params.odemeTipi     - 'Nakit' | 'Kredi Kartı (POS)' | 'Havale / EFT' | 'Cari Hesap'
 * @param {string} [params.bankaKod]    - EFT için hangi banka (Havale/EFT seçilince zorunlu)
 * @param {string} [params.firmaId]     - Cari için firma ID (Cari seçilince zorunlu)
 * @param {string} [params.aciklama]
 * @param {number} [params.posKomisyon] - POS komisyonu (opsiyonel)
 * @param {string} [params.belgeNo]     - Fiş/makbuz no
 * @param {Date|string} [params.tarih]  - Boş = bugün
 * @returns {string} Tahsilat ID'si
 */
function tahsilatEkle(params) {
  const {
    musteriAd,
    tcKimlik = '',
    hizmet,
    tutar,
    odemeTipi,
    bankaKod = '',
    firmaId = '',
    aciklama = '',
    posKomisyon = 0,
    belgeNo = '',
    tarih,
  } = params;

  // ── Validasyon ──
  if (!musteriAd) throw new Error('Müşteri adı boş olamaz.');
  if (!hizmet)    throw new Error('Hizmet türü boş olamaz.');
  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('Tutar geçersiz.');

  const gecerliTipler = Object.values(ODEME_TIPI);
  if (!gecerliTipler.includes(odemeTipi)) {
    throw new Error(`Geçersiz ödeme tipi: "${odemeTipi}". Geçerli tipler: ${gecerliTipler.join(', ')}`);
  }

  if (odemeTipi === ODEME_TIPI.EFT && !bankaKod) {
    throw new Error('Havale/EFT için banka kodu zorunludur.');
  }
  if (odemeTipi === ODEME_TIPI.CARI && !firmaId) {
    throw new Error('Cari hesap için firma ID zorunludur.');
  }

  // ── TAHSILAT sheet'e kayıt ──
  const sheet = getSheet(SHEETS.TAHSILAT);
  const id = yeniID('T', sheet);
  const kayitTarihi = tarih ? tarihFormat(tarih) : bugun();

  sheet.appendRow([
    id,
    kayitTarihi,
    musteriAd,
    tcKimlik,
    hizmet,
    tutar,
    odemeTipi,
    bankaKod,
    firmaId,
    aciklama,
    aktifKullanici(),
  ]);

  // ── Ödeme tipine göre ilgili modüle işle ──
  const hareketAciklama = `Tahsilat: ${musteriAd} — ${hizmet}`;

  switch (odemeTipi) {
    case ODEME_TIPI.NAKIT:
      // A) Nakit → Kasa
      kasaGiris(hareketAciklama, tutar, 'Tahsilat', belgeNo, id);
      break;

    case ODEME_TIPI.KART:
      // B) Kredi Kartı → POS → İşbank
      posTahsilatiEkle({
        tutar,
        aciklama : hareketAciklama,
        komisyon : posKomisyon,
        kaynakId : id,
        tarih    : kayitTarihi,
      });
      break;

    case ODEME_TIPI.EFT:
      // C) Havale/EFT → Seçilen banka hesabına
      eftGiris(bankaKod, tutar, hareketAciklama, musteriAd, id);
      break;

    case ODEME_TIPI.CARI:
      // D) Cari → Firmanın cari hesabına alacak yaz
      cariAlacakEkle(firmaId, tutar, hareketAciklama, id);
      break;
  }

  Logger.log(`✅ Tahsilat kaydedildi: ${id} | ${musteriAd} | ${paraBirim(tutar)} | ${odemeTipi}`);
  return id;
}

// ── TAHSİLAT LİSTELEME ───────────────────────────────────────

/**
 * Tüm tahsilatları döner.
 * @param {number} limit - Kaç kayıt (0 = hepsi)
 */
function tahsilatlariAl(limit = 0) {
  const sheet = getSheet(SHEETS.TAHSILAT);
  const sonSatir = sheet.getLastRow();
  if (sonSatir <= 1) return [];

  const baslangic = limit > 0 ? Math.max(2, sonSatir - limit + 1) : 2;
  const data = sheet.getRange(baslangic, 1, sonSatir - baslangic + 1, 11).getValues();

  return data.map(r => ({
    id          : r[0],
    tarih       : tarihFormat(r[1]),
    musteriAd   : r[2],
    tcKimlik    : r[3],
    hizmet      : r[4],
    tutar       : r[5],
    odemeTipi   : r[6],
    bankaKod    : r[7],
    firmaId     : r[8],
    aciklama    : r[9],
    kullanici   : r[10],
  })).reverse();
}

/**
 * Günlük tahsilat özeti.
 * @param {string} [tarih] - GG.AA.YYYY (boş = bugün)
 */
function tahsilatGunlukOzet(tarih) {
  const arananTarih = tarih || bugun();
  const hepsini = tahsilatlariAl();
  const gunTahsilat = hepsini.filter(t => t.tarih === arananTarih);

  const ozet = {
    tarih,
    toplamTutar : gunTahsilat.reduce((s, t) => s + t.tutar, 0),
    adet        : gunTahsilat.length,
    nakit       : 0,
    kart        : 0,
    eft         : 0,
    cari        : 0,
  };

  gunTahsilat.forEach(t => {
    switch (t.odemeTipi) {
      case ODEME_TIPI.NAKIT: ozet.nakit += t.tutar; break;
      case ODEME_TIPI.KART : ozet.kart  += t.tutar; break;
      case ODEME_TIPI.EFT  : ozet.eft   += t.tutar; break;
      case ODEME_TIPI.CARI : ozet.cari  += t.tutar; break;
    }
  });

  return ozet;
}

/**
 * Aylık tahsilat özeti.
 * @param {string} [yilAy] - YYYY-MM (boş = bu ay)
 */
function tahsilatAylikOzet(yilAy) {
  const buAy = yilAy || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const hepsini = tahsilatlariAl();

  const ayTahsilat = hepsini.filter(t => {
    if (!t.tarih) return false;
    const parts = t.tarih.split('.');
    return parts.length >= 3 && `${parts[2]}-${parts[1]}` === buAy;
  });

  return {
    donem       : buAy,
    toplamTutar : ayTahsilat.reduce((s, t) => s + t.tutar, 0),
    adet        : ayTahsilat.length,
    nakit       : ayTahsilat.filter(t => t.odemeTipi === ODEME_TIPI.NAKIT).reduce((s, t) => s + t.tutar, 0),
    kart        : ayTahsilat.filter(t => t.odemeTipi === ODEME_TIPI.KART ).reduce((s, t) => s + t.tutar, 0),
    eft         : ayTahsilat.filter(t => t.odemeTipi === ODEME_TIPI.EFT  ).reduce((s, t) => s + t.tutar, 0),
    cari        : ayTahsilat.filter(t => t.odemeTipi === ODEME_TIPI.CARI ).reduce((s, t) => s + t.tutar, 0),
  };
}

// ── DASHBOARD VERİSİ ──────────────────────────────────────────

function tahsilatDashboardVerisi() {
  return {
    gunlukOzet   : tahsilatGunlukOzet(),
    aylikOzet    : tahsilatAylikOzet(),
    sonTahsilatlar: tahsilatlariAl(10),
  };
}
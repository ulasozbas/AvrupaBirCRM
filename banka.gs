// ============================================================
//  BANKA.GS — Banka Hesapları & POS Modülü
//  - Birden fazla banka hesabı takibi
//  - İşbankası POS tahsilatları
//  - EFT / Havale giriş-çıkışları
//  - Anlık bakiye hesaplama
// ============================================================

// ── BANKA TANIM ──────────────────────────────────────────────

/**
 * Tanımlı banka hesaplarını döner.
 * BANKA sheet'ini okur.
 */
function bankaliHesaplariAl() {
  const sheet = getSheet(SHEETS.BANKA);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map(r => ({
    kod        : r[0],
    ad         : r[1],
    iban       : r[2],
    basBakiye  : r[3] || 0,
    guncelBak  : r[4] || 0,
    aktif      : r[5] === true || r[5] === 'TRUE',
    notlar     : r[6] || '',
  })).filter(b => b.aktif);
}

/**
 * Belirli bir banka hesabının anlık bakiyesini döner.
 * BANKA_HAREKET sheet'inden hesaplar.
 * @param {string} bankaKod - Banka kodu (örn: 'ISBANK_POS')
 */
function bankaBakiyesiAl(bankaKod) {
  const sheet = getSheet(SHEETS.BANKA_HAREKET);
  const sonSatir = sheet.getLastRow();
  if (sonSatir <= 1) {
    // Harekat yoksa BANKA sheet'teki başlangıç bakiyesini döndür
    return _bankaBaslangicBakiye(bankaKod);
  }

  const data = sheet.getDataRange().getValues().slice(1);
  // O bankaya ait son hareketteki bakiyeyi bul
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][COL.BANKA_HAREKET.BANKA_KOD - 1] === bankaKod) {
      return data[i][COL.BANKA_HAREKET.BAKIYE - 1] || 0;
    }
  }
  return _bankaBaslangicBakiye(bankaKod);
}

/**
 * Tüm aktif banka hesaplarının toplam bakiyesini döner.
 */
function toplamBankaBakiyesi() {
  const bankalar = bankaliHesaplariAl();
  return bankalar.reduce((toplam, banka) => toplam + bankaBakiyesiAl(banka.kod), 0);
}

/**
 * BANKA sheet'teki başlangıç bakiyesini döner (yardımcı).
 */
function _bankaBaslangicBakiye(bankaKod) {
  const sheet = getSheet(SHEETS.BANKA);
  const data = sheet.getDataRange().getValues().slice(1);
  const banka = data.find(r => r[0] === bankaKod);
  return banka ? (banka[3] || 0) : 0;
}

// ── BANKA HAREKET ────────────────────────────────────────────

/**
 * Banka hesabına hareket ekler. Bakiyeyi anlık günceller.
 *
 * @param {Object} params
 * @param {string} params.bankaKod    - Hangi banka hesabı
 * @param {string} params.tip         - 'Giriş' veya 'Çıkış'
 * @param {number} params.tutar       - Tutar
 * @param {string} params.aciklama    - Açıklama
 * @param {string} [params.karsiHesap] - Karşı IBAN veya isim
 * @param {string} [params.kaynakId]  - İlgili kayıt ID
 * @param {Date|string} [params.tarih]
 * @returns {string} Oluşturulan hareket ID'si
 */
function bankaHareketiEkle(params) {
  const { bankaKod, tip, tutar, aciklama, karsiHesap = '', kaynakId = '', tarih } = params;

  if (!['Giriş', 'Çıkış'].includes(tip)) throw new Error(`Geçersiz tip: "${tip}"`);
  if (!bankaKod) throw new Error('Banka kodu boş olamaz.');
  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('Tutar pozitif olmalıdır.');

  const sheet = getSheet(SHEETS.BANKA_HAREKET);
  const mevcutBakiye = bankaBakiyesiAl(bankaKod);
  const yeniBakiye = tip === 'Giriş' ? mevcutBakiye + tutar : mevcutBakiye - tutar;

  const id = yeniID('BH', sheet);
  const kayitTarihi = tarih ? tarihFormat(tarih) : bugun();

  sheet.appendRow([
    id,
    kayitTarihi,
    bankaKod,
    aciklama,
    tip,
    tutar,
    yeniBakiye,
    karsiHesap,
    kaynakId,
  ]);

  // BANKA sheet'teki güncel bakiyeyi de güncelle
  _bankaGuncelBakiyeYaz(bankaKod, yeniBakiye);

  Logger.log(`🏦 Banka ${tip} [${bankaKod}]: ${paraBirim(tutar)} | Yeni bakiye: ${paraBirim(yeniBakiye)}`);
  return id;
}

/**
 * BANKA sheet'teki "Güncel Bakiye" sütununu günceller.
 */
function _bankaGuncelBakiyeYaz(bankaKod, yeniBakiye) {
  const sheet = getSheet(SHEETS.BANKA);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === bankaKod) {
      sheet.getRange(i + 1, 5).setValue(yeniBakiye); // 5. sütun: Güncel Bakiye
      return;
    }
  }
}

/**
 * EFT / Havale ile para GELDİĞİNDE çağrılır.
 * Otomatik olarak ilgili IBAN'ın banka hesabına işler.
 *
 * @param {string} bankaKod   - Paranın geldiği banka hesabı
 * @param {number} tutar
 * @param {string} aciklama
 * @param {string} [karsiHesap] - Gönderenin IBAN'ı veya adı
 * @param {string} [kaynakId]
 */
function eftGiris(bankaKod, tutar, aciklama, karsiHesap = '', kaynakId = '') {
  return bankaHareketiEkle({
    bankaKod,
    tip: 'Giriş',
    tutar,
    aciklama: aciklama || 'EFT / Havale Geliri',
    karsiHesap,
    kaynakId,
  });
}

/**
 * EFT / Havale ile ödeme YAPILDIĞINDA çağrılır.
 *
 * @param {string} bankaKod   - Paranın çıktığı banka hesabı
 * @param {number} tutar
 * @param {string} aciklama
 * @param {string} [karsiIban] - Gönderilen IBAN
 * @param {string} [kaynakId]
 */
function eftCikis(bankaKod, tutar, aciklama, karsiIban = '', kaynakId = '') {
  return bankaHareketiEkle({
    bankaKod,
    tip: 'Çıkış',
    tutar,
    aciklama: aciklama || 'EFT / Havale Ödemesi',
    karsiHesap: karsiIban,
    kaynakId,
  });
}

// ── POS ──────────────────────────────────────────────────────

/**
 * POS (kredi kartı) tahsilatı kaydeder.
 * - POS sheet'e kayıt ekler
 * - İşbankası POS hesabına banka hareketi olarak yazar
 *
 * @param {Object} params
 * @param {number} params.tutar       - Tahsilat tutarı
 * @param {string} params.aciklama    - Açıklama
 * @param {number} [params.komisyon]  - POS komisyonu (TL, opsiyonel)
 * @param {string} [params.kaynakId]  - Tahsilat ID'si
 * @param {Date|string} [params.tarih]
 * @returns {string} POS kayıt ID'si
 */
function posTahsilatiEkle(params) {
  const { tutar, aciklama, komisyon = 0, kaynakId = '', tarih } = params;

  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('POS tutarı geçersiz.');

  const posSheet = getSheet(SHEETS.POS);
  const netTutar = tutar - komisyon;
  const id = yeniID('P', posSheet);
  const kayitTarihi = tarih ? tarihFormat(tarih) : bugun();

  // POS sheet'e kayıt
  posSheet.appendRow([
    id,
    kayitTarihi,
    aciklama,
    tutar,
    komisyon,
    netTutar,
    kaynakId,
  ]);

  // POS hesabını tutan bankaya da yaz
  const posBankaKod = ayarOku('POS_BANKA') || 'ISBANK_POS';
  bankaHareketiEkle({
    bankaKod   : posBankaKod,
    tip        : 'Giriş',
    tutar      : netTutar,  // Komisyon düşülmüş tutar
    aciklama   : `POS Tahsilat: ${aciklama}`,
    kaynakId   : id,
  });

  Logger.log(`💳 POS tahsilat: ${paraBirim(tutar)} (Komisyon: ${paraBirim(komisyon)}, Net: ${paraBirim(netTutar)}) | ID: ${id}`);
  return id;
}

/**
 * POS hareketlerini döner.
 * @param {number} limit
 */
function posHareketleriAl(limit = 0) {
  const sheet = getSheet(SHEETS.POS);
  const sonSatir = sheet.getLastRow();
  if (sonSatir <= 1) return [];

  const baslangic = limit > 0 ? Math.max(2, sonSatir - limit + 1) : 2;
  const data = sheet.getRange(baslangic, 1, sonSatir - baslangic + 1, 7).getValues();

  return data.map(r => ({
    id       : r[0],
    tarih    : tarihFormat(r[1]),
    aciklama : r[2],
    tutar    : r[3],
    komisyon : r[4],
    netTutar : r[5],
    kaynakId : r[6],
  })).reverse();
}

/**
 * Ayın toplam POS tahsilatını döner.
 */
function posAylikToplam(yilAy) {
  const buAy = yilAy || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const hareketler = posHareketleriAl();

  return hareketler
    .filter(h => {
      if (!h.tarih) return false;
      const parts = h.tarih.split('.');
      return parts.length >= 3 && `${parts[2]}-${parts[1]}` === buAy;
    })
    .reduce((s, h) => s + h.netTutar, 0);
}

// ── BANKA HAREKET LISTESI ─────────────────────────────────────

/**
 * Belirli bir banka hesabının hareketlerini döner.
 * @param {string} bankaKod
 * @param {number} limit
 */
function bankaHareketleriAl(bankaKod, limit = 0) {
  const sheet = getSheet(SHEETS.BANKA_HAREKET);
  const sonSatir = sheet.getLastRow();
  if (sonSatir <= 1) return [];

  const data = sheet.getRange(2, 1, sonSatir - 1, 9).getValues();
  let filtered = bankaKod ? data.filter(r => r[COL.BANKA_HAREKET.BANKA_KOD - 1] === bankaKod) : data;
  if (limit > 0) filtered = filtered.slice(-limit);

  return filtered.map(r => ({
    id         : r[0],
    tarih      : tarihFormat(r[1]),
    bankaKod   : r[2],
    aciklama   : r[3],
    tip        : r[4],
    tutar      : r[5],
    bakiye     : r[6],
    karsiHesap : r[7],
    kaynakId   : r[8],
  })).reverse();
}

// ── DASHBOARD VERİSİ ──────────────────────────────────────────

/**
 * Dashboard için banka verilerini JSON döner.
 */
function bankaDashboardVerisi() {
  const bankalar = bankaliHesaplariAl();
  const detaylar = bankalar.map(b => ({
    kod    : b.kod,
    ad     : b.ad,
    bakiye : bankaBakiyesiAl(b.kod),
  }));

  return {
    toplamBankaBakiye: toplamBankaBakiyesi(),
    bankalar         : detaylar,
    sonPosIslemleri  : posHareketleriAl(5),
  };
}
// ============================================================
//  GIDER.GS — Gider Kayıt Modülü
//
//  Her gider kaydında ödeme kaynağı seçilir:
//  - Kasa         → Kasa bakiyesi düşer
//  - Banka (EFT)  → İlgili banka hesabı düşer
//  - Kredi Kartı  → POS borçlanma (negatif hareket)
//  - Çek          → Çek/Senet modülüne eklenir
// ============================================================

// ── ÖDEME KAYNAKLARI ─────────────────────────────────────────
const GIDER_KAYNAGI = {
  KASA  : 'Kasa',
  BANKA : 'Banka (EFT)',
  KART  : 'Kredi Kartı',
  CEK   : 'Çek',
};

// ── GIDER KATEGORİLERİ ────────────────────────────────────────
const GIDER_KATEGORI = {
  KIRA        : 'Kira',
  ELEKTRIK    : 'Elektrik',
  SU          : 'Su',
  INTERNET    : 'İnternet',
  TELEFON     : 'Telefon',
  MAAS        : 'Maaş',
  SGK         : 'SGK',
  VERGI       : 'Vergi',
  KIRTASIYE   : 'Kırtasiye',
  TEMIZLIK    : 'Temizlik/Hijyen',
  TEKNIK_SERV : 'Teknik Servis',
  MALZEME     : 'Sarf Malzeme',
  ARAC        : 'Araç/Ulaşım',
  TEMSIL      : 'Temsil/Ağırlama',
  DIGER       : 'Diğer',
};

// ── ANA GİDER FONKSİYONU ─────────────────────────────────────

/**
 * Gider kaydeder ve ödeme kaynağını anlık günceller.
 *
 * @param {Object} params
 * @param {string} params.kategori      - Gider kategorisi
 * @param {string} params.aciklama      - Açıklama
 * @param {number} params.tutar         - Tutar
 * @param {string} params.odemeKaynagi  - 'Kasa' | 'Banka (EFT)' | 'Kredi Kartı' | 'Çek'
 * @param {string} [params.bankaKod]    - Banka EFT ise hangi hesap
 * @param {string} [params.cekAlacakli] - Çek ise kime (alacaklı)
 * @param {string} [params.cekVade]     - Çek vadesi (GG.AA.YYYY)
 * @param {string} [params.cekBanka]    - Çekin hangi bankadan kesileceği
 * @param {string} [params.belgeNo]     - Fiş/fatura no
 * @param {Date|string} [params.tarih]  - Boş = bugün
 * @returns {string} Gider ID'si
 */
function giderEkle(params) {
  const {
    kategori,
    aciklama,
    tutar,
    odemeKaynagi,
    bankaKod = '',
    cekAlacakli = '',
    cekVade = '',
    cekBanka = '',
    belgeNo = '',
    tarih,
  } = params;

  // ── Validasyon ──
  if (!kategori)    throw new Error('Kategori boş olamaz.');
  if (!aciklama)    throw new Error('Açıklama boş olamaz.');
  if (typeof tutar !== 'number' || tutar <= 0) throw new Error('Tutar geçersiz.');
  if (!odemeKaynagi) throw new Error('Ödeme kaynağı seçilmelidir.');

  if (odemeKaynagi === GIDER_KAYNAGI.BANKA && !bankaKod) {
    throw new Error('Banka EFT için banka kodu zorunludur.');
  }
  if (odemeKaynagi === GIDER_KAYNAGI.CEK && !cekAlacakli) {
    throw new Error('Çek için alacaklı adı zorunludur.');
  }
  if (odemeKaynagi === GIDER_KAYNAGI.CEK && !cekVade) {
    throw new Error('Çek için vade tarihi zorunludur.');
  }

  // ── GIDER sheet'e kayıt ──
  const sheet = getSheet(SHEETS.GIDER);
  const id = yeniID('G', sheet);
  const kayitTarihi = tarih ? tarihFormat(tarih) : bugun();

  sheet.appendRow([
    id,
    kayitTarihi,
    kategori,
    aciklama,
    tutar,
    odemeKaynagi,
    bankaKod,
    belgeNo,
    aktifKullanici(),
  ]);

  // ── Ödeme kaynağını güncelle ──
  const hareketAciklama = `Gider: ${kategori} — ${aciklama}`;

  switch (odemeKaynagi) {
    case GIDER_KAYNAGI.KASA:
      // Nakit çıkış → kasa düşer
      kasaCikis(hareketAciklama, tutar, kategori, belgeNo, id);
      break;

    case GIDER_KAYNAGI.BANKA:
      // EFT çıkışı → banka hesabı düşer
      eftCikis(bankaKod, tutar, hareketAciklama, '', id);
      break;

    case GIDER_KAYNAGI.KART:
      // Kredi kartı → POS'a borçlanma olarak kayıt
      // (Not: POS tahsilat değil, borç/harcama kaydı)
      _posGiderKaydi(tutar, hareketAciklama, id, kayitTarihi);
      break;

    case GIDER_KAYNAGI.CEK:
      // Çek → Çek/Senet modülüne ekle (vade bekler)
      cekEkle({
        alacakli          : cekAlacakli,
        tutar,
        vade              : cekVade,
        banka             : cekBanka,
        aciklama          : hareketAciklama,
        duzenlenmeTarihi  : kayitTarihi,
      });
      break;
  }

  Logger.log(`💸 Gider kaydedildi: ${id} | ${kategori} | ${paraBirim(tutar)} | ${odemeKaynagi}`);
  return id;
}

/**
 * POS ile yapılan gideri (borçlanma) kaydeder.
 * POS sheet'e negatif işlem olarak girer.
 */
function _posGiderKaydi(tutar, aciklama, kaynakId, tarih) {
  const posSheet = getSheet(SHEETS.POS);
  const id = yeniID('PG', posSheet);
  posSheet.appendRow([
    id,
    tarih || bugun(),
    `[GİDER] ${aciklama}`,
    -tutar,   // Negatif = gider
    0,
    -tutar,
    kaynakId,
  ]);
  // Not: POS banka bakiyesini düşürmüyoruz çünkü
  // kart harcamaları banka bakiyesinden değil, kart borcundan düşer.
  // İstenirse bankaBakiyesi - tutar yapılabilir.
}

// ── SABİT GİDER ──────────────────────────────────────────────

/**
 * Aylık sabit giderleri döner.
 * Dashboard'daki "Planlanan Aylık Sabit Gider" kutusunda görünür.
 */
function sabitGiderleriAl() {
  const sheet = getSheet(SHEETS.SABIT_GIDER);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1)
    .map(r => ({
      id         : r[0],
      kategori   : r[1],
      aciklama   : r[2],
      aylikTutar : r[3] || 0,
      aktif      : r[4] === true || r[4] === 'TRUE',
      sonOdeme   : tarihFormat(r[5]),
      notlar     : r[6] || '',
    }))
    .filter(g => g.aktif);
}

/**
 * Toplam aylık sabit gider tutarını döner.
 * (Kira + Fatura + Abonelikler vb.)
 */
function toplamAylikSabitGider() {
  return sabitGiderleriAl().reduce((s, g) => s + g.aylikTutar, 0);
}

/**
 * Sabit gider ekler veya günceller.
 */
function sabitGiderEkle(kategori, aciklama, aylikTutar) {
  const sheet = getSheet(SHEETS.SABIT_GIDER);
  const id = yeniID('SG', sheet);
  sheet.appendRow([id, kategori, aciklama, aylikTutar, true, '', '']);
  Logger.log(`📌 Sabit gider eklendi: ${kategori} | ${paraBirim(aylikTutar)}/ay`);
  return id;
}

// ── GİDER LİSTELEME ──────────────────────────────────────────

/**
 * Giderleri döner.
 * @param {number} limit
 */
function giderleriAl(limit = 0) {
  const sheet = getSheet(SHEETS.GIDER);
  const sonSatir = sheet.getLastRow();
  if (sonSatir <= 1) return [];

  const baslangic = limit > 0 ? Math.max(2, sonSatir - limit + 1) : 2;
  const data = sheet.getRange(baslangic, 1, sonSatir - baslangic + 1, 9).getValues();

  return data.map(r => ({
    id           : r[0],
    tarih        : tarihFormat(r[1]),
    kategori     : r[2],
    aciklama     : r[3],
    tutar        : r[4],
    odemeKaynagi : r[5],
    bankaKod     : r[6],
    belgeNo      : r[7],
    kullanici    : r[8],
  })).reverse();
}

/**
 * Aylık gider özeti, kategoriye göre gruplu.
 * @param {string} [yilAy] - YYYY-MM
 */
function giderAylikOzet(yilAy) {
  const buAy = yilAy || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
  const giderler = giderleriAl();

  const ayGiderler = giderler.filter(g => {
    if (!g.tarih) return false;
    const parts = g.tarih.split('.');
    return parts.length >= 3 && `${parts[2]}-${parts[1]}` === buAy;
  });

  // Kategoriye göre grupla
  const kategoriler = {};
  ayGiderler.forEach(g => {
    kategoriler[g.kategori] = (kategoriler[g.kategori] || 0) + g.tutar;
  });

  return {
    donem       : buAy,
    toplamTutar : ayGiderler.reduce((s, g) => s + g.tutar, 0),
    adet        : ayGiderler.length,
    kategoriler,
  };
}

// ── DASHBOARD VERİSİ ──────────────────────────────────────────

function giderDashboardVerisi() {
  return {
    aylikOzet        : giderAylikOzet(),
    sabitGiderler    : sabitGiderleriAl(),
    toplamSabitGider : toplamAylikSabitGider(),
    sonGiderler      : giderleriAl(10),
  };
}
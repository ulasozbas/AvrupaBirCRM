// ============================================================
//  DASHBOARD.GS — Ana Dashboard Modülü
//
//  4 büyük kart:
//  1. Kasa Nakit      — anlık nakit bakiye
//  2. Banka Nakit     — tüm banka hesapları toplamı
//  3. Borçlar         — kart + çek + SGK + vergi
//  4. Planlanan Aylık Sabit Gider
//
//  Skorboard:
//  - Aylık sabit gider hedefine karşı gerçekleşen gelir
//  - Para girdikçe ilerleme çubuğu dolar
//  - Yetersizse kırmızı, yeterliyse yeşil
// ============================================================

// ── ANA DASHBOARD VERİSİ ──────────────────────────────────────

/**
 * Tüm dashboard verilerini tek seferde toplar ve döner.
 * HTML sidebar veya web app bu fonksiyonu çağırır.
 *
 * @returns {Object} Dashboard için tam veri paketi
 */
function dashboardVerisiAl() {
  try {
    const simdi  = new Date();
    const yilAy  = Utilities.formatDate(simdi, Session.getScriptTimeZone(), 'yyyy-MM');

    // ── 4 Büyük Kart ──
    const kasaNakit       = kasaBakiyesiAl();
    const bankaNakit      = toplamBankaBakiyesi();
    const toplamBorclar   = _toplamBorcHesapla();
    const aylikSabitGider = toplamAylikSabitGider();

    // ── Skorboard ──
    const skorboard       = _skorboardHesapla(yilAy, aylikSabitGider);

    // ── Ek bilgiler ──
    const bekleyenCekler  = cekDashboardVerisi();
    const cariAlacaklar   = cariDashboardVerisi();
    const aylikTahsilat   = tahsilatAylikOzet(yilAy);
    const aylikGider      = giderAylikOzet(yilAy);

    // Son hareketler
    const sonKasaHareketler  = kasaHareketleriAl(8);
    const sonTahsilatlar     = tahsilatlariAl(8);

    return {
      guncelleme       : tarihFormat(simdi),
      donem            : yilAy,

      // 4 kart
      kartlar: {
        kasaNakit,
        bankaNakit,
        toplamNakit    : kasaNakit + bankaNakit,
        toplamBorclar,
        aylikSabitGider,
      },

      // Skorboard
      skorboard,

      // Özet bilgiler
      bekleyenCekler   : {
        adet           : bekleyenCekler.bekleyenAdet,
        tutar          : bekleyenCekler.toplamBekleyenTutar,
        yaklasanlar    : bekleyenCekler.yaklasanCekler,
      },
      cariAlacaklar    : {
        toplamAlacak   : cariAlacaklar.toplamAlacak,
        borcluFirma    : cariAlacaklar.borcluFirma,
      },
      aylikTahsilat,
      aylikGider,
      sonKasaHareketler,
      sonTahsilatlar,
    };
  } catch(e) {
    Logger.log(`Dashboard verisi hatası: ${e.message}`);
    throw e;
  }
}

// ── BORÇ HESAPLA ─────────────────────────────────────────────

/**
 * Tüm borç kalemlerini toplar.
 * - Bekleyen çekler
 * - Ödenmemiş SGK / Vergi
 * - (Kart borcu şimdilik ayrı takip edilmez, ileride eklenebilir)
 */
function _toplamBorcHesapla() {
  let toplam = 0;

  // Bekleyen çekler
  try {
    toplam += bekleyenCeklerToplam();
  } catch(e) { /* sheet yoksa geç */ }

  // Ödenmemiş SGK / Vergi
  try {
    toplam += _odenmemisSgkVergi();
  } catch(e) { /* sheet yoksa geç */ }

  return toplam;
}

/**
 * ÖDENMEMIŞ SGK ve vergi toplamı.
 */
function _odenmemisSgkVergi() {
  try {
    const sheet = getSheet(SHEETS.SGK_VERGI);
    const data = sheet.getDataRange().getValues().slice(1);
    return data
      .filter(r => r[6] === 'Bekliyor' || r[6] === 'Gecikmiş')
      .reduce((s, r) => s + (r[3] || 0), 0);
  } catch(e) {
    return 0;
  }
}

// ── SKORBOARD HESAPLA ─────────────────────────────────────────

/**
 * Aylık ödeme skorboard verisi.
 *
 * Hedef: Aylık sabit giderleri karşılamak için gereken tutar
 * Gerçekleşen: Bu ay gelen toplam gelir (tahsilat + cari tahsilat)
 *
 * Progress = Gerçekleşen / Hedef × 100
 * Yeşil = %100+, Sarı = %60-99, Kırmızı = <%60
 *
 * @param {string} yilAy         - YYYY-MM
 * @param {number} hedefTutar    - Aylık sabit gider toplamı
 */
function _skorboardHesapla(yilAy, hedefTutar) {
  // Bu ay gelen toplam tahsilat
  let gerceklesen = 0;

  try {
    const tahsilatOzet = tahsilatAylikOzet(yilAy);
    gerceklesen += tahsilatOzet.toplamTutar;
  } catch(e) { /* geç */ }

  const yuzde = hedefTutar > 0
    ? Math.min(Math.round((gerceklesen / hedefTutar) * 100), 999)
    : 0;

  const durum = yuzde >= 100 ? 'yeterli'
              : yuzde >= 60  ? 'dikkat'
              : 'yetersiz';

  const kalan = Math.max(0, hedefTutar - gerceklesen);

  return {
    donem         : yilAy,
    hedefTutar,
    gerceklesen,
    kalan,
    yuzde,
    durum,         // 'yeterli' | 'dikkat' | 'yetersiz'
    mesaj          : _skorboardMesaji(durum, kalan, hedefTutar),
  };
}

function _skorboardMesaji(durum, kalan, hedef) {
  if (durum === 'yeterli') return '✅ Bu ay sabit giderler karşılandı!';
  if (durum === 'dikkat')  return `⚠️ Dikkat: Sabit giderler için ${paraBirim(kalan)} daha girmeli.`;
  return `🔴 Yetersiz: Sabit giderler için ${paraBirim(kalan)} eksik. Hedef: ${paraBirim(hedef)}`;
}

// ── HTML SIDEBAR / WEB APP ────────────────────────────────────

/**
 * Google Sheets menüsünden Dashboard'u açar.
 */
function dashboardAc() {
  const html = HtmlService
    .createHtmlOutputFromFile('dashboard_ui')
    .setTitle('OSGB Finansal Dashboard')
    .setWidth(900);
  SpreadsheetApp.getUi().showSidebar(html);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💼 OSGB Finans')
    .addItem('📊 Dashboard Aç', 'dashboardAc')
    .addSeparator()
    .addSubMenu(ui.createMenu('💵 Tahsilat')
      .addItem('Yeni Tahsilat Ekle', 'tahsilatFormAc')
    )
    .addSubMenu(ui.createMenu('💸 Gider')
      .addItem('Yeni Gider Ekle', 'giderFormAc')
    )
    .addSubMenu(ui.createMenu('📄 Çek / Senet')
      .addItem('Yeni Çek Ekle', 'cekFormAc')
      .addItem('Bekleyen Çekleri Gör', 'bekleyenCekleriGoster')
    )
    .addSubMenu(ui.createMenu('🏢 Firma / Cari')
      .addItem('Cari Listesi', 'cariListesiGoster')
      .addItem('Paraşüt Fatura Import', 'parasutImportAc')
    )
    .addSeparator()
    .addItem('⚙️ Kurulum / Ayarlar', 'sistemKurulum')
    .addToUi();
}

// Placeholder fonksiyonlar (UI modülünde implement edilecek)
function tahsilatFormAc()     { SpreadsheetApp.getUi().alert('Tahsilat formu açılıyor...'); }
function giderFormAc()        { SpreadsheetApp.getUi().alert('Gider formu açılıyor...'); }
function cekFormAc()          { SpreadsheetApp.getUi().alert('Çek formu açılıyor...'); }
function bekleyenCekleriGoster() { SpreadsheetApp.getUi().alert(JSON.stringify(beklenyenCekler(), null, 2)); }
function cariListesiGoster()  { SpreadsheetApp.getUi().alert(JSON.stringify(tumCarileriAl(), null, 2)); }
function parasutImportAc()    { SpreadsheetApp.getUi().alert('Paraşüt import açılıyor...'); }
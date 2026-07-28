/**
 * ============================================================
 * Code.gs — Google Apps Script Backend
 * Portfolio Content Creator Website
 * ============================================================
 *
 * SETUP INSTRUCTIONS:
 * 1. Buka Google Apps Script di https://script.google.com
 * 2. Paste seluruh kode ini ke file Code.gs
 * 3. Ganti SPREADSHEET_ID di bawah dengan ID Google Spreadsheet Anda
 * 4. Klik Deploy > New Deployment > Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. Copy URL deployment dan paste ke index.html (variabel GAS_URL)
 * 6. Untuk GitHub Pages / cross-origin: tambahkan ?callback=NAMA_FUNGSI
 *    untuk menggunakan mode JSONP (contoh: ?action=all&callback=myFunc)
 *
 * STRUKTUR SPREADSHEET (buat sheet berikut):
 * Sheet: Settings, Home, About, Brands, Analytics,
 *        Portfolio, Testimonials, FAQ, Contact,
 *        RateCard Packages, RateCard Services
 *
 * Setiap sheet memiliki 2 kolom: Key | Value
 * Kecuali: Brands, Analytics, Portfolio, Testimonials, FAQ,
 *          RateCard Packages, RateCard Services
 * yang memiliki format row (header di baris 1)
 * ============================================================
 */

// ============================================================
// KONFIGURASI — GANTI HANYA BAGIAN INI
// ============================================================
// Gunakan var agar kompatibel dengan semua versi Google Apps Script engine
var SPREADSHEET_ID = 'YOUR_GOOGLE_SPREADSHEET_ID';

// ============================================================
// MAIN HANDLER
// ============================================================

/**
 * doGet — menangani semua permintaan GET
 *
 * Parameter URL:
 *   ?action=settings|home|about|brands|analytics|portfolio|
 *           testimonials|faq|contact|ratecard|all
 *   ?callback=namaFungsi   → aktifkan mode JSONP (wajib untuk GitHub Pages)
 *
 * Contoh pemanggilan dari GitHub Pages (fetch biasa akan CORS-blocked):
 *   <script>
 *     function handleData(resp) { console.log(resp); }
 *   </script>
 *   <script src="https://script.google.com/macros/s/XXXX/exec?action=all&callback=handleData"></script>
 *
 * Contoh pemanggilan tanpa JSONP (dari domain yang sama / sudah pakai no-cors):
 *   fetch('https://script.google.com/macros/s/XXXX/exec?action=all')
 */
function doGet(e) {
  var params   = (e && e.parameter) ? e.parameter : {};
  var action   = params.action   ? params.action.toLowerCase() : 'all';
  var callback = params.callback ? params.callback.trim()      : '';

  var result;
  try {
    switch (action) {
      case 'settings':     result = getSettings();     break;
      case 'home':         result = getHome();          break;
      case 'about':        result = getAbout();         break;
      case 'brands':       result = getBrands();        break;
      case 'analytics':    result = getAnalytics();     break;
      case 'portfolio':    result = getPortfolio();     break;
      case 'testimonials': result = getTestimonials();  break;
      case 'faq':          result = getFAQ();           break;
      case 'contact':      result = getContact();       break;
      case 'ratecard':     result = getRateCard();      break;
      case 'all':          result = getAllData();        break;
      default:
        result = {
          error: 'Action tidak valid: ' + action,
          available: ['settings','home','about','brands','analytics',
                      'portfolio','testimonials','faq','contact','ratecard','all']
        };
    }
  } catch (err) {
    result = { error: err.message };
  }

  // Tentukan status sukses secara eksplisit
  var isSuccess = !(result && result.error);
  var payload = JSON.stringify({
    success:   isSuccess,
    data:      result,
    timestamp: new Date().toISOString()
  });

  // Mode JSONP — wajib untuk panggilan dari GitHub Pages / cross-origin
  if (callback && /^[a-zA-Z_$][a-zA-Z0-9_$]*$/.test(callback)) {
    return ContentService
      .createTextOutput(callback + '(' + payload + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  // Mode JSON biasa
  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doPost — menangani permintaan POST (opsional, untuk form contact dll.)
 * Body JSON: { action: 'contact', name: '...', email: '...', message: '...' }
 */
function doPost(e) {
  var params = {};
  try {
    if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      params = e.parameter;
    }
  } catch (err) {
    params = (e && e.parameter) ? e.parameter : {};
  }

  var action = params.action ? params.action.toLowerCase() : '';
  var result;

  try {
    if (action === 'contact' || action === 'pesan') {
      result = saveContactMessage(params);
    } else {
      result = { error: 'Action POST tidak dikenali: ' + action };
    }
  } catch (err) {
    result = { error: err.message };
  }

  var payload = JSON.stringify({
    success:   !(result && result.error),
    data:      result,
    timestamp: new Date().toISOString()
  });

  return ContentService
    .createTextOutput(payload)
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Menyimpan pesan kontak ke sheet "Messages" (dibuat otomatis jika belum ada)
 * Kolom: Timestamp | Name | Email | Phone | Subject | Message
 */
function saveContactMessage(params) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName('Messages');

  // Buat sheet jika belum ada
  if (!sheet) {
    sheet = ss.insertSheet('Messages');
    sheet.appendRow(['Timestamp','Nama','Email','Telepon','Subjek','Pesan']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }

  var name    = params.name    || params.nama    || '';
  var email   = params.email   || '';
  var phone   = params.phone   || params.telepon || '';
  var subject = params.subject || params.subjek  || '';
  var message = params.message || params.pesan   || '';

  if (!name && !email) {
    throw new Error('Nama atau email wajib diisi.');
  }

  sheet.appendRow([
    new Date().toISOString(),
    name, email, phone, subject, message
  ]);

  return { message: 'Pesan berhasil dikirim. Terima kasih, ' + (name || 'Anda') + '!' };
}

// ============================================================
// HELPER: Open Spreadsheet
// ============================================================
function getSpreadsheet() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet(name) {
  var ss    = getSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan di Spreadsheet.');
  return sheet;
}

/**
 * Membaca sheet dengan format Key-Value (2 kolom)
 * Kolom A = Key, Kolom B = Value
 * Mengembalikan object { key: value }
 */
function readKeyValueSheet(sheetName) {
  var sheet   = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return {};

  var data   = sheet.getRange(1, 1, lastRow, 2).getValues();
  var result = {};

  for (var i = 0; i < data.length; i++) {
    var key   = String(data[i][0] || '').trim();
    var value = data[i][1];
    if (key && key !== '') {
      result[toCamelCase(key)] = (value !== null && value !== undefined) ? String(value) : '';
    }
  }
  return result;
}

/**
 * Membaca sheet dengan format tabel (baris pertama = header)
 * Mengembalikan array of objects
 */
function readTableSheet(sheetName) {
  var sheet   = getSheet(sheetName);
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  var allData = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = allData[0].map(function(h) { return toCamelCase(String(h || '').trim()); });
  var rows    = [];

  for (var i = 1; i < allData.length; i++) {
    var row     = {};
    var hasData = false;
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) {
        var val        = allData[i][j];
        row[headers[j]] = (val !== null && val !== undefined) ? String(val) : '';
        if (row[headers[j]] !== '') hasData = true;
      }
    }
    if (hasData) rows.push(row);
  }
  return rows;
}

/**
 * Konversi string ke camelCase
 * Contoh: "Site Name" -> "siteName", "WhatsApp" -> "whatsapp"
 * "Jam Operasional" -> "jamOperasional" (mendukung karakter non-ASCII)
 */
function toCamelCase(str) {
  if (!str) return '';
  // Normalisasi: ganti tanda baca umum dengan spasi, tapi JANGAN hapus huruf aksen
  return str
    .trim()
    .replace(/[^\w\s]/g, ' ')   // hapus tanda baca, sisakan huruf+angka+spasi
    .replace(/\s+/g, ' ')       // normalisasi spasi ganda
    .trim()
    .split(' ')
    .map(function(word, index) {
      if (!word) return '';
      var lower = word.toLowerCase();
      // Kata pertama seluruhnya lowercase, kata berikutnya kapital di depan
      return index === 0 ? lower : lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join('');
}

// ============================================================
// SETTINGS
// ============================================================
/**
 * Sheet: Settings
 * Format: Key | Value
 *
 * Keys yang digunakan:
 * Site Name | Sarah Amanda
 * Logo | https://drive.google.com/...
 * Favicon | https://...
 * WhatsApp | 6281234567890
 * WA Template | Halo, Saya ingin...
 * PDF Rate Card | https://drive.google.com/...
 * Instagram | https://instagram.com/...
 * TikTok | https://tiktok.com/...
 * YouTube | https://youtube.com/...
 * Threads | https://threads.net/...
 * Facebook | https://facebook.com/...
 * LinkedIn | https://linkedin.com/...
 * Website | https://...
 * Email | hello@example.com
 * Phone | +62 812 ...
 * Address | Jakarta Selatan, Indonesia
 * Google Maps | <iframe...> atau URL
 * QR WhatsApp | https://drive.google.com/...
 * Copyright | © 2025 Sarah Amanda
 * Primary Color | #059669
 * Secondary Color | #84a98c
 */
function getSettings() {
  try {
    var d = readKeyValueSheet('Settings');
    return {
      siteName:       d.siteName || d.siteNama || '',
      logo:           toDirectLink(d.logo || ''),
      favicon:        toDirectLink(d.favicon || ''),
      whatsapp:       (d.whatsapp || '').replace(/\D/g,''),
      waTemplate:     d.waTemplate || d.templateWhatsapp || d.pesanWhatsapp || 'Halo,\n\nSaya ingin melakukan kerja sama.\n\nNama :\nPerusahaan :\nProduk :\nJenis Campaign :\nBudget :\nDeadline :\n\nTerima kasih.',
      pdfRateCard:    toDirectLink(d.pdfRateCard || d.rateCard || ''),
      instagram:      d.instagram || '',
      tiktok:         d.tiktok || '',
      youtube:        d.youtube || '',
      threads:        d.threads || '',
      facebook:       d.facebook || '',
      linkedin:       d.linkedin || '',
      website:        d.website || '',
      email:          d.email || '',
      phone:          d.phone || d.nomorHp || '',
      address:        d.address || d.alamat || '',
      googleMaps:     d.googleMaps || d.maps || '',
      qrWhatsapp:     toDirectLink(d.qrWhatsapp || d.qr || ''),
      copyright:      d.copyright || ('© ' + new Date().getFullYear() + ' Portfolio. All rights reserved.'),
      primaryColor:   d.primaryColor || '#059669',
      secondaryColor: d.secondaryColor || '#84a98c',
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// HOME
// ============================================================
/**
 * Sheet: Home
 * Format: Key | Value
 *
 * Keys:
 * Name | Sarah Amanda
 * Profession | Lifestyle Creator
 * Tagline | Creating Content That Connects
 * Description | I help brands...
 * Photo | https://drive.google.com/...
 * Followers | 150K+
 * Engagement Rate | 5.8%
 * Brands Worked | 120+
 * Monthly Reach | 1.2M+
 */
function getHome() {
  try {
    var d = readKeyValueSheet('Home');
    return {
      name:           d.name || d.nama || '',
      profession:     d.profession || d.profesi || '',
      tagline:        d.tagline || '',
      description:    d.description || d.deskripsi || '',
      photo:          toDirectLink(d.photo || d.foto || ''),
      followers:      d.followers || d.totalFollowers || '',
      engagementRate: d.engagementRate || d.engagement || '',
      brandsWorked:   d.brandsWorked || d.brand || '',
      monthlyReach:   d.monthlyReach || d.reach || '',
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// ABOUT
// ============================================================
/**
 * Sheet: About
 * Format: Key | Value
 *
 * Keys:
 * Photo | https://drive.google.com/...
 * Title | Turning Ideas Into Impactful Content
 * Description | Saya adalah content creator...
 * Experience | 5+
 * Campaign | 120+
 * Repeat Client | 95%
 * Videos Created | 500+
 * Achievement | 10+
 */
function getAbout() {
  try {
    var d = readKeyValueSheet('About');
    return {
      photo:         toDirectLink(d.photo || d.foto || ''),
      title:         d.title || d.judul || '',
      description:   d.description || d.deskripsi || '',
      experience:    d.experience || d.pengalaman || '',
      campaign:      d.campaign || d.campaignCompleted || '',
      repeatClient:  d.repeatClient || d.repeat || '',
      videosCreated: d.videosCreated || d.videos || '',
      achievement:   d.achievement || '',
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// BRANDS
// ============================================================
/**
 * Sheet: Brands
 * Format: Tabel (Header di baris 1)
 *
 * Kolom: Name | Logo | Website | Order
 * Contoh:
 * Name | Logo | Website | Order
 * Tokopedia | https://logo.url | https://tokopedia.com | 1
 */
function getBrands() {
  try {
    var rows = readTableSheet('Brands');
    return rows
      .filter(function(r) { return r.name || r.logo; })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(r) { return {
        name:    r.name || r.nama || '',
        logo:    toDirectLink(r.logo || ''),
        website: r.website || '',
      }; });
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// ANALYTICS
// ============================================================
/**
 * Sheet: Analytics
 * Format: Tabel (Header di baris 1)
 *
 * Kolom: Platform | Handle | Followers | Engagement Rate | Reach | Age Range | Female Percent | Male Percent | Order
 * Contoh:
 * Platform | Handle | Followers | Engagement Rate | Reach | Age Range | Female Percent | Order
 * Instagram | @sarahamanda | 150K | 5.8% | 1.2M | 18-34 | 68 | 1
 */
function getAnalytics() {
  try {
    var rows = readTableSheet('Analytics');
    return rows
      .filter(function(r) { return r.platform; })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(r) { return {
        platform:       r.platform || '',
        handle:         r.handle || '',
        followers:      r.followers || '',
        engagementRate: r.engagementRate || r.engagement || '',
        reach:          r.reach || '',
        ageRange:       r.ageRange || r.age || '18-34',
        femalePercent:  r.femalePercent || r.female || r.wanita || '65',
        malePercent:    r.malePercent || r.male || r.pria || '35',
      }; });
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// PORTFOLIO
// ============================================================
/**
 * Sheet: Portfolio
 * Format: Tabel (Header di baris 1)
 *
 * Kolom: Title | Category | Description | Thumbnail | Link | Status | Order
 * Contoh:
 * Title | Category | Description | Thumbnail | Link | Status | Order
 * Skincare Review | Beauty | Honest review... | https://img.url | https://ig.url | active | 1
 *
 * Status: active / inactive (inactive = tidak ditampilkan)
 */
function getPortfolio() {
  try {
    var rows = readTableSheet('Portfolio');
    return rows
      .filter(function(r) { return r.title && (r.status || 'active').toLowerCase() !== 'inactive'; })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(r) { return {
        title:       r.title || r.judul || '',
        category:    r.category || r.kategori || '',
        description: r.description || r.deskripsi || '',
        thumbnail:   toDirectLink(r.thumbnail || ''),
        link:        r.link || '',
        status:      r.status || 'active',
        order:       parseInt(r.order || 0),
      }; });
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// TESTIMONIALS
// ============================================================
/**
 * Sheet: Testimonials
 * Format: Tabel (Header di baris 1)
 *
 * Kolom: Name | Brand | Rating | Text | Avatar | Order
 * Contoh:
 * Name | Brand | Rating | Text | Avatar | Order
 * Wardah | Beauty Brand | 5 | Sangat profesional... | https://img.url | 1
 *
 * Rating: 1-5
 */
function getTestimonials() {
  try {
    var rows = readTableSheet('Testimonials');
    return rows
      .filter(function(r) { return r.name && (r.text || r.testimonial); })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(r) { return {
        name:   r.name || r.nama || '',
        brand:  r.brand || r.company || r.perusahaan || '',
        rating: parseInt(r.rating || 5),
        text:   r.text || r.testimonial || r.review || '',
        avatar: toDirectLink(r.avatar || ''),
      }; });
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// FAQ
// ============================================================
/**
 * Sheet: FAQ
 * Format: Tabel (Header di baris 1)
 *
 * Kolom: Question | Answer | Order
 * Contoh:
 * Question | Answer | Order
 * Berapa lama proses... | Proses pembuatan... | 1
 */
function getFAQ() {
  try {
    var rows = readTableSheet('FAQ');
    return rows
      .filter(function(r) { return r.question && r.answer; })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(r) { return {
        question: r.question || r.pertanyaan || '',
        answer:   r.answer || r.jawaban || '',
      }; });
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// CONTACT
// ============================================================
/**
 * Sheet: Contact
 * Format: Key | Value
 *
 * Keys:
 * Phone | +62 812 3456 7890
 * Email | hello@example.com
 * Address | Jakarta Selatan, Indonesia
 * Google Maps | <iframe ...> atau embed URL
 * QR WhatsApp | https://drive.google.com/...
 * Hours Mon-Fri | 09:00 - 18:00
 * Hours Saturday | 10:00 - 15:00
 * Hours Sunday | Closed
 */
function getContact() {
  try {
    var d     = readKeyValueSheet('Contact');
    var hours = {};
    if (d.hoursMonFri || d.mondayFriday || d.seninJumat)
      hours['Mon – Fri']  = d.hoursMonFri || d.mondayFriday || d.seninJumat || '09:00 – 18:00';
    if (d.hoursSaturday || d.sabtu)
      hours['Saturday']   = d.hoursSaturday || d.sabtu || '10:00 – 15:00';
    if (d.hoursSunday || d.minggu)
      hours['Sunday']     = d.hoursSunday || d.minggu || 'Closed';
    return {
      phone:      d.phone || d.nomorHp || d.nomor || '',
      email:      d.email || '',
      address:    d.address || d.alamat || '',
      googleMaps: d.googleMaps || d.maps || d.peta || '',
      qrWhatsapp: toDirectLink(d.qrWhatsapp || d.qr || ''),
      hours:      Object.keys(hours).length ? hours : { 'Mon – Fri': '09:00 – 18:00', 'Saturday': '10:00 – 15:00', 'Sunday': 'Closed' },
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ============================================================
// RATE CARD
// ============================================================
/**
 * Sheet: RateCard (atau Rate Card)
 *
 * Sub-sheet 1: Sheet bernama "RateCard Packages" atau "Rate Card Packages"
 * Format: Tabel
 * Kolom: Tier | Name | Price | Features | Featured | Order
 * Contoh:
 * Tier | Name | Price | Features | Featured | Order
 * STARTER | Story Package | Rp250.000 | 1 Instagram Story\nMention\nMax 3 Slide | false | 1
 * PROFESSIONAL | Bundle Package | Rp1.250.000 | Instagram Reels\nEditing+Caption\n... | true | 2
 *
 * Sub-sheet 2: Sheet bernama "RateCard Services" atau "Additional Services"
 * Format: Tabel
 * Kolom: Name | Order
 * Contoh:
 * Name | Order
 * UGC Content | 1
 * Photography | 2
 *
 * ATAU: Jika menggunakan satu sheet "RateCard":
 * Format Key-Value untuk packages sederhana
 */
function getRateCard() {
  try {
    var packages           = [];
    var additionalServices = [];

    try { packages = readTableSheet('RateCard Packages'); } catch (e) {}
    if (!packages.length) { try { packages = readTableSheet('Rate Card Packages'); } catch (e) {} }
    if (!packages.length) { try { packages = readTableSheet('Packages');           } catch (e) {} }
    if (!packages.length) { try { var r1 = readTableSheet('RateCard');   if (r1.length) packages = r1; } catch (e) {} }
    if (!packages.length) { try { var r2 = readTableSheet('Rate Card');  if (r2.length) packages = r2; } catch (e) {} }

    packages = packages
      .filter(function(p) { return p.name || p.nama; })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(p) { return {
        tier:     p.tier || p.tipe || '',
        name:     p.name || p.nama || '',
        price:    p.price || p.harga || '',
        features: (p.features || p.fitur || p.items || '').split(/\n|\\n/).map(function(f) { return f.trim(); }).filter(Boolean),
        featured: String(p.featured || p.popular || '').toLowerCase() === 'true',
      }; });

    try { additionalServices = readTableSheet('RateCard Services');   } catch (e) {}
    if (!additionalServices.length) { try { additionalServices = readTableSheet('Additional Services'); } catch (e) {} }
    if (!additionalServices.length) { try { additionalServices = readTableSheet('Layanan Tambahan');    } catch (e) {} }

    additionalServices = additionalServices
      .filter(function(s) { return s.name || s.nama; })
      .sort(function(a, b) { return parseInt(a.order||0) - parseInt(b.order||0); })
      .map(function(s) { return { name: s.name || s.nama || '' }; });

    return { packages: packages, additionalServices: additionalServices };
  } catch (e) {
    return { error: e.message, packages: [], additionalServices: [] };
  }
}

// ============================================================
// ALL DATA (single request untuk pre-loading)
// ============================================================
function getAllData() {
  return {
    settings:     getSettings(),
    home:         getHome(),
    about:        getAbout(),
    brands:       getBrands(),
    analytics:    getAnalytics(),
    portfolio:    getPortfolio(),
    testimonials: getTestimonials(),
    faq:          getFAQ(),
    contact:      getContact(),
    ratecard:     getRateCard(),
  };
}

// ============================================================
// HELPER: Convert Google Drive share link to direct image URL
// ============================================================
/**
 * Mengkonversi berbagai format Google Drive link menjadi
 * URL gambar langsung yang bisa digunakan di <img src="">
 *
 * Mendukung:
 * - https://drive.google.com/file/d/FILE_ID/view
 * - https://drive.google.com/open?id=FILE_ID
 * - https://drive.google.com/uc?id=FILE_ID
 * - URL biasa (dikembalikan apa adanya)
 */
function toDirectLink(url) {
  if (!url || url.trim() === '') return '';
  url = url.trim();
  if (!url.includes('drive.google.com')) return url;
  var match1 = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (match1) return 'https://drive.google.com/uc?export=view&id=' + match1[1];
  var match2 = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (match2) return 'https://drive.google.com/uc?export=view&id=' + match2[1];
  if (url.includes('/uc?')) return url;
  return url;
}

// ============================================================
// HELPER: Test functions (jalankan dari editor Apps Script)
// ============================================================

/** Jalankan fungsi ini untuk mengetes semua data sekaligus */
function testAll() {
  var result = getAllData();
  Logger.log(JSON.stringify(result, null, 2));
}

function testSettings()     { Logger.log(JSON.stringify(getSettings(),     null, 2)); }
function testHome()         { Logger.log(JSON.stringify(getHome(),         null, 2)); }
function testAbout()        { Logger.log(JSON.stringify(getAbout(),        null, 2)); }
function testBrands()       { Logger.log(JSON.stringify(getBrands(),       null, 2)); }
function testAnalytics()    { Logger.log(JSON.stringify(getAnalytics(),    null, 2)); }
function testPortfolio()    { Logger.log(JSON.stringify(getPortfolio(),    null, 2)); }
function testTestimonials() { Logger.log(JSON.stringify(getTestimonials(), null, 2)); }
function testFAQ()          { Logger.log(JSON.stringify(getFAQ(),          null, 2)); }
function testContact()      { Logger.log(JSON.stringify(getContact(),      null, 2)); }
function testRateCard()     { Logger.log(JSON.stringify(getRateCard(),     null, 2)); }

/**
 * Simulasi doGet dengan parameter tertentu
 * Contoh: jalankan testDoGet('portfolio') atau testDoGet('all')
 */
function testDoGet(actionParam) {
  var fakeEvent = { parameter: { action: actionParam || 'all' } };
  var output    = doGet(fakeEvent);
  Logger.log(output.getContent());
}

/**
 * Simulasi JSONP — cek apakah output terbungkus callback dengan benar
 */
function testJSONP() {
  var fakeEvent = { parameter: { action: 'settings', callback: 'myCallback' } };
  var output    = doGet(fakeEvent);
  Logger.log(output.getContent()); // harus dimulai dengan: myCallback({...});
}

/**
 * Tes toCamelCase dengan berbagai input
 */
function testCamelCase() {
  var cases = [
    'Site Name',          // → siteName
    'WhatsApp',           // → whatsapp
    'Jam Operasional',    // → jamOperasional
    'Senin – Jumat',      // → seninJumat
    'Hours Mon-Fri',      // → hoursMonFri
    'Female Percent',     // → femalePercent
    'PDF Rate Card',      // → pdfRateCard
    'Google Maps',        // → googleMaps
  ];
  cases.forEach(function(c) {
    Logger.log('"' + c + '" → "' + toCamelCase(c) + '"');
  });
}

// ============================================================
//  CLOVERTEES — Proyek 5: Public Order Form Backend
//  GAS Project BARU (terpisah dari admin dashboard)
//
//  Sheets ID : 133ENleDfF5-8Tt2c5bFHSIPKmjtYtSR89vCaFua7v-c
//  Drive Root: 1N91DmawIYiSRjfCJlHCfnW-clxezASd_
//
//  Struktur folder Drive:
//  📁 Clovertees (root)
//    └── 📁 2025-04
//          └── 📁 CLV-0426-0001
//                └── 📁 Item-1_Kaos_BLACK
//                      ├── depan_namafile.png
//                      ├── belakang_namafile.png
//                      └── lengan_namafile.png
// ============================================================

// ==================== CONFIGURATION ====================
var CFG = {
  SPREADSHEET_ID : '133ENleDfF5-8Tt2c5bFHSIPKmjtYtSR89vCaFua7v-c',
  SHEET_ORDER    : 'Orderan',
  SHEET_DESIGN   : 'Desain',
  DRIVE_ROOT_ID  : '1N91DmawIYiSRjfCJlHCfnW-clxezASd_',
};

// ==================== ENTRY POINTS ====================
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ result: 'ok', service: 'Clovertees P5 API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var raw  = e.postData ? e.postData.contents : '';
    var data = JSON.parse(raw);
    var result = processOrder(data);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==================== MAIN PROCESSOR ====================
function processOrder(data) {
  var ss          = SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
  var orderSheet  = ss.getSheetByName(CFG.SHEET_ORDER);
  var designSheet = ss.getSheetByName(CFG.SHEET_DESIGN);

  // 1. Generate No. Pesanan
  var noPesanan = generateNoPesanan(orderSheet);
  var timestamp = new Date();
  var tz        = Session.getScriptTimeZone();

  // 2. Buat folder Drive: Root → YYYY-MM → No.Pesanan
  var monthFolder  = getOrCreateFolder(CFG.DRIVE_ROOT_ID, formatMonthFolder(timestamp));
  var orderFolder  = getOrCreateFolder(monthFolder.getId(), noPesanan);

  // 3. Upload semua file & kumpulkan link per item
  var items      = data.items || [];
  var totalPcs   = 0;
  var itemLinks  = []; // array of { fileDepan, fileBelakang, fileLengan }
  var detailPesananArr = [];

  items.forEach(function(item, idx) {
    var pcs = calcPcs(item.sizes);
    totalPcs += pcs;

    // Rekap untuk Detail Pesanan
    var jenisLabel = (item.jenis || 'ITEM').toUpperCase();
    var warnaLabel = (item.warna || '-').toUpperCase();
    var ukD = formatSizes(item.sizes ? item.sizes.dewasa : {});
    var ukA = formatSizes(item.sizes ? item.sizes.anak   : {});
    var textDetail = (idx + 1) + '. ' + jenisLabel + '\n'
                   + warnaLabel + '\n'
                   + 'Dewasa: ' + ukD + '\n'
                   + 'Kids: '   + ukA;
    detailPesananArr.push(textDetail);

    // Nama folder item: "Item-1_Kaos_BLACK"
    var safeJenis = sanitizeName(item.jenis || 'Item');
    var safeWarna = sanitizeName(item.warna || 'NoColor');
    var itemFolderName = 'Item-' + (idx + 1) + '_' + safeJenis + '_' + safeWarna;
    var itemFolder = getOrCreateFolder(orderFolder.getId(), itemFolderName);

    var links = { fileDepan: '', fileBelakang: '', fileLengan: '' };

    ['depan', 'belakang', 'lengan'].forEach(function(sisi) {
      var sisiData = item.desain ? item.desain[sisi] : null;
      if (!sisiData || !sisiData.files || !sisiData.files.length) return;

      var uploaded = [];
      sisiData.files.forEach(function(f, fi) {
        if (!f.data || !f.mimeType) return; // file kosong / terlalu besar
        try {
          var prefix   = sisi + (sisiData.files.length > 1 ? '_' + (fi + 1) : '') + '_';
          var fileName = prefix + (f.name || ('file_' + (fi + 1)));
          var blob     = Utilities.newBlob(Utilities.base64Decode(f.data), f.mimeType, fileName);
          var driveFile = itemFolder.createFile(blob);
          driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          uploaded.push(driveFile.getUrl());
        } catch (uploadErr) {
          Logger.log('Upload error [' + sisi + ' fi=' + fi + ']: ' + uploadErr);
        }
      });

      var linkKey = 'file' + sisi.charAt(0).toUpperCase() + sisi.slice(1); // fileDepan dll
      links[linkKey] = uploaded.join('\n');
    });

    itemLinks.push(links);
  });

  // 4. Simpan ke tab Orderan
  // Kolom: No.Pesanan | Timestamp | Nama | WhatsApp | Instagram | Alamat | Detail Pesanan | Total PCS | Ongkir | Total Harga
  var tsFormatted = Utilities.formatDate(timestamp, tz, 'dd/MM/yyyy HH:mm:ss');
  var detailPesananFull = detailPesananArr.join('\n\n') + '\n\nTotal PCS : ' + totalPcs + ' Pcs';

  orderSheet.appendRow([
    noPesanan,
    tsFormatted,
    data.nama     || '',
    data.whatsapp || '',
    data.instagram|| '',
    data.alamat   || '',
    detailPesananFull, // Detail Pesanan — otomatis diisi
    totalPcs,
    '',            // Ongkir — diisi admin
    '',            // Total Harga — diisi admin
    'Pending',     // Status (kolom K)
  ]);

  // 5. Simpan ke tab Desain (1 baris per item)
  // Kolom: No.Pesanan | Nama | Jenis Baju | Warna | Ukuran Dewasa | Ukuran Anak |
  //        PCS Item | Desain Depan | Desain Belakang | Desain Lengan |
  //        File Depan | File Belakang | File Lengan
  items.forEach(function(item, idx) {
    var pcs     = calcPcs(item.sizes);
    var links   = itemLinks[idx] || {};

    var ukDewasa = formatSizes(item.sizes ? item.sizes.dewasa : {});
    var ukAnak   = formatSizes(item.sizes ? item.sizes.anak   : {});

    var txtDepan    = item.desain && item.desain.depan    ? (item.desain.depan.text    || '') : '';
    var txtBelakang = item.desain && item.desain.belakang ? (item.desain.belakang.text || '') : '';
    var txtLengan   = item.desain && item.desain.lengan   ? (item.desain.lengan.text   || '') : '';

    designSheet.appendRow([
      noPesanan,
      data.nama || '',
      item.jenis || '',
      item.warna || '',
      ukDewasa,
      ukAnak,
      pcs,
      txtDepan,
      txtBelakang,
      txtLengan,
      links.fileDepan    || '',
      links.fileBelakang || '',
      links.fileLengan   || '',
    ]);
  });

  return {
    result     : 'success',
    no_pesanan : noPesanan,
    total_pcs  : totalPcs,
    folder_url : orderFolder.getUrl(),
  };
}

// ==================== GENERATE NO. PESANAN ====================
function generateNoPesanan(orderSheet) {
  var now    = new Date();
  var mm     = String(now.getMonth() + 1).padStart(2, '0');
  var yy     = String(now.getFullYear()).slice(-2);
  var prefix = 'CLV-' + mm + yy + '-';

  var lastRow = orderSheet.getLastRow();
  var maxSeq  = 0;

  if (lastRow >= 2) {
    var col = orderSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    col.forEach(function(r) {
      var val = String(r[0] || '').trim();
      if (val.startsWith(prefix)) {
        var parts = val.split('-');
        var seq   = parseInt(parts[2], 10);
        if (!isNaN(seq) && seq > maxSeq) maxSeq = seq;
      }
    });
  }

  return prefix + String(maxSeq + 1).padStart(4, '0');
}

// ==================== DRIVE HELPERS ====================
function getOrCreateFolder(parentId, folderName) {
  var parent  = DriveApp.getFolderById(parentId);
  var iter    = parent.getFoldersByName(folderName);
  if (iter.hasNext()) return iter.next();
  return parent.createFolder(folderName);
}

function formatMonthFolder(date) {
  // Format: "2025-04"
  var yyyy = date.getFullYear();
  var mm   = String(date.getMonth() + 1).padStart(2, '0');
  return yyyy + '-' + mm;
}

// ==================== CALC & FORMAT HELPERS ====================
function calcPcs(sizes) {
  if (!sizes) return 0;
  var total = 0;
  var dewasa = sizes.dewasa || {};
  var anak   = sizes.anak   || {};
  Object.keys(dewasa).forEach(function(k) { total += parseInt(dewasa[k]) || 0; });
  Object.keys(anak).forEach(function(k)   { total += parseInt(anak[k])   || 0; });
  return total;
}

function formatSizes(sizeObj) {
  if (!sizeObj) return '-';
  var parts = [];
  Object.keys(sizeObj).forEach(function(k) {
    var v = parseInt(sizeObj[k]) || 0;
    if (v > 0) parts.push(k + ':' + v);
  });
  return parts.length ? parts.join(', ') : '-';
}

function sanitizeName(str) {
  // Buang karakter yang tidak aman untuk nama folder
  return String(str).replace(/[^a-zA-Z0-9_\- ]/g, '').replace(/\s+/g, '_').substring(0, 40);
}

/**
 * Sistem Kasir + Dashboard untuk Spreadsheet Paket Sembako.
 *
 * Fitur utama:
 * 1) Menu Kasir + Dashboard.
 * 2) Form kasir (sidebar/modal/web app) untuk input transaksi.
 * 3) Validasi stok (stok tidak boleh minus) sebelum transaksi disimpan.
 * 4) Sinkronisasi otomatis Stok Akhir di sheet PRODUK.
 * 5) Rebuild REKAP_PRODUK dan REKAP_PELANGGAN.
 * 6) Dashboard chart (line/bar/pie) + KPI.
 */
const APP_CONFIG = {
  menuKasir: 'Kasir',
  menuDashboard: 'Dashboard',
  titleIndex: 'Paket Sembako',
  titleKasir: 'Form Kasir',
  titleProduk: 'Manajemen Produk',
  titleDashboard: 'Dashboard Penjualan',
  cachePrefix: 'dashboard:data:v',
  cacheVersionKey: 'dashboard_version',
  salesSnapshotBackfillRowKey: 'penjualan_snapshot_backfill_last_row',
  cacheTtlSeconds: 45,
  chartItemLimit: 20,
  pieItemLimit: 10,
  minMarginPercent: 12,
  maxAlertItems: 20,
  sheets: {
    PRODUK: 'PRODUK',
    PENJUALAN: 'PENJUALAN',
    MUTASI_STOK: 'MUTASI_STOK',
    REKAP_PRODUK: 'REKAP_PRODUK',
    REKAP_PELANGGAN: 'REKAP_PELANGGAN',
  },
};

const SHEET_HEADERS = {
  PRODUK: [
    'ID / SKU',
    'Kategori',
    'Nama Produk',
    'Harga Modal (Rp)',
    'Satuan',
    'Perkiraan Harga (Rp)',
    'Stok Awal',
    'Stok Akhir',
    'Modal',
    'Stok Minimum',
  ],
  PENJUALAN: [
    'Tanggal',
    'No. Invoice',
    'Nama Pelanggan',
    'SKU',
    'Nama Produk',
    'Satuan',
    'Harga Satuan (Rp)',
    'Qty',
    'Total (Rp)',
    'Catatan',
    'Harga Modal Satuan (Rp)',
    'HPP (Rp)',
    'Laba Kotor (Rp)',
  ],
  MUTASI_STOK: [
    'Tanggal',
    'Jenis Mutasi',
    'SKU',
    'Nama Produk',
    'Qty (+/-)',
    'Catatan',
  ],
  REKAP_PRODUK: [
    'SKU',
    'Nama Produk',
    'Satuan',
    'Harga Modal (Rp)',
    'Harga Jual (Rp)',
    'Qty Terjual',
    'Omzet (Rp)',
    'HPP (Rp)',
    'Laba Kotor (Rp)',
  ],
  REKAP_PELANGGAN: [
    'Nama Pelanggan',
    'Total Transaksi',
    'Total Qty',
    'Total Belanja (Rp)',
    'Total HPP (Rp)',
    'Total Profit (Rp)',
  ],
};

const FIELD_ALIASES = {
  sku: ['id / sku', 'id sku', 'sku', 'kode', 'kode produk'],
  kategori: ['kategori', 'category'],
  namaProduk: ['nama produk', 'produk', 'nama barang', 'barang', 'item'],
  hargaModal: ['harga modal (rp)', 'harga modal', 'modal', 'hpp unit', 'biaya pokok'],
  satuan: ['satuan', 'unit', 'uom'],
  hargaJual: ['perkiraan harga (rp)', 'harga jual (rp)', 'harga jual', 'harga', 'price'],
  stokAwal: ['stok awal', 'stock awal', 'stok masuk'],
  stokAkhir: ['stok akhir', 'stock akhir', 'sisa stok'],
  stokMinimum: ['stok minimum', 'minimal stok', 'minimum stok', 'min stok', 'reorder point'],
  modal: ['modal', 'nilai modal'],

  tanggal: ['tanggal', 'tgl', 'date', 'waktu', 'timestamp'],
  invoice: ['no. invoice', 'no invoice', 'invoice', 'nomor invoice', 'id transaksi'],
  jenisMutasi: ['jenis mutasi', 'tipe mutasi', 'type mutasi', 'jenis', 'type'],
  pelanggan: ['nama pelanggan', 'pelanggan', 'customer', 'pembeli', 'client'],
  hargaSatuan: ['harga satuan (rp)', 'harga satuan', 'harga', 'price'],
  hargaModalTransaksi: [
    'harga modal satuan (rp)',
    'harga modal satuan',
    'harga modal transaksi (rp)',
    'harga modal transaksi',
    'modal transaksi',
    'hpp unit transaksi',
    'hpp unit',
  ],
  qty: ['qty', 'jumlah', 'kuantitas', 'quantity'],
  total: ['total (rp)', 'total', 'omzet', 'grand total', 'nilai transaksi'],
  catatan: ['catatan', 'note', 'keterangan'],

  qtyTerjual: ['qty terjual', 'jumlah terjual', 'qty'],
  omzet: ['omzet (rp)', 'omzet', 'total belanja'],
  hpp: ['hpp (rp)', 'hpp', 'biaya pokok'],
  labaKotor: ['laba kotor (rp)', 'laba kotor', 'profit'],

  totalTransaksi: ['total transaksi', 'jumlah transaksi', 'trx'],
  totalQty: ['total qty', 'total quantity', 'qty total'],
  totalBelanja: ['total belanja (rp)', 'total belanja', 'belanja'],
  totalHpp: ['total hpp (rp)', 'total hpp', 'hpp total'],
  totalProfit: ['total profit (rp)', 'total profit', 'profit total'],
};

/**
 * Menu kustom saat spreadsheet dibuka.
 */
function onOpen() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ensureCoreSheets_(ss);
  } catch (err) {
    // Jangan hentikan pembuatan menu jika auto-setup gagal.
    Logger.log('onOpen ensureCoreSheets_ error: ' + err);
  }

  const ui = SpreadsheetApp.getUi();

  ui
    .createMenu(APP_CONFIG.menuKasir)
    .addItem('Buka Kasir (Sidebar)', 'showKasirSidebar')
    .addItem('Buka Kasir (Modal)', 'showKasirModal')
    .addItem('Buka Manajemen Produk', 'showProductManagerModal')
    .addSeparator()
    .addItem('Input Stok Masuk', 'inputStockMasuk')
    .addItem('Input Retur Barang', 'inputStockRetur')
    .addItem('Input Adjustment Stok', 'inputStockAdjustment')
    .addItem('Cek Alert Stok Minimum', 'showLowStockAlerts')
    .addSeparator()
    .addItem('Buat Data Dummy', 'generateDummyData')
    .addItem('Sinkronkan Stok + Rekap', 'syncRecapAndStock')
    .addItem('Setup Struktur Sheet', 'setupDatabaseSheets')
    .addToUi();

  ui
    .createMenu(APP_CONFIG.menuDashboard)
    .addItem('Buka Halaman Utama', 'showIndex')
    .addItem('Buka Dashboard', 'showDashboard')
    .addItem('Buka Sidebar Dashboard', 'showDashboardSidebar')
    .addSeparator()
    .addItem('Aktifkan Trigger Auto-Refresh', 'setupDashboardTriggers')
    .addToUi();
}

/**
 * onInstall untuk memastikan menu muncul setelah install.
 */
function onInstall(e) {
  onOpen(e);
}

/**
 * Menampilkan kasir dalam sidebar.
 */
function showKasirSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Kasir').setTitle(APP_CONFIG.titleKasir);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Menampilkan kasir dalam modal dialog.
 */
function showKasirModal() {
  const html = HtmlService.createHtmlOutputFromFile('Kasir')
    .setTitle(APP_CONFIG.titleKasir)
    .setWidth(520)
    .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, APP_CONFIG.titleKasir);
}

/**
 * Menampilkan manajemen produk dalam modal dialog.
 */
function showProductManagerModal() {
  const html = HtmlService.createHtmlOutputFromFile('Produk')
    .setTitle(APP_CONFIG.titleProduk)
    .setWidth(1180)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, APP_CONFIG.titleProduk);
}

/**
 * Input stok masuk via prompt sederhana.
 */
function inputStockMasuk() {
  promptStockMutationFlow_('MASUK');
}

/**
 * Input retur barang (stok bertambah) via prompt sederhana.
 */
function inputStockRetur() {
  promptStockMutationFlow_('RETUR');
}

/**
 * Input adjustment stok via prompt sederhana.
 * Qty boleh positif/negatif.
 */
function inputStockAdjustment() {
  promptStockMutationFlow_('ADJUSTMENT');
}

/**
 * Menampilkan daftar produk dengan stok di bawah batas minimum.
 */
function showLowStockAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets_(ss);

  const info = getLowStockAlerts();
  const ui = SpreadsheetApp.getUi();
  if (!info.items.length) {
    ui.alert('Tidak ada produk yang melewati batas stok minimum.');
    return;
  }

  const lines = info.items.slice(0, APP_CONFIG.maxAlertItems).map((item, idx) => {
    return (
      String(idx + 1) + '. ' + item.sku + ' - ' + item.namaProduk +
      ' | Stok: ' + formatPlainNumber_(item.stokTersedia) +
      ' | Minimum: ' + formatPlainNumber_(item.stokMinimum)
    );
  });

  ui.alert(
    'Alert Stok Minimum\n\n' +
      'Total produk hampir habis: ' + info.items.length + '\n\n' +
      lines.join('\n')
  );
}

/**
 * Flow input mutasi stok berbasis prompt.
 */
function promptStockMutationFlow_(mutationType) {
  const ui = SpreadsheetApp.getUi();

  const skuPrompt = ui.prompt(
    'Input ' + normalizeMutationLabel_(mutationType),
    'Masukkan SKU produk:',
    ui.ButtonSet.OK_CANCEL
  );
  if (skuPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const qtyHelp = mutationType === 'ADJUSTMENT'
    ? 'Masukkan Qty adjustment. Gunakan negatif untuk mengurangi stok, positif untuk menambah.'
    : 'Masukkan Qty (angka positif).';
  const qtyPrompt = ui.prompt('Qty', qtyHelp, ui.ButtonSet.OK_CANCEL);
  if (qtyPrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const notePrompt = ui.prompt('Catatan', 'Catatan transaksi (opsional):', ui.ButtonSet.OK_CANCEL);
  if (notePrompt.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  const result = saveStockMutation({
    type: mutationType,
    sku: skuPrompt.getResponseText(),
    qty: qtyPrompt.getResponseText(),
    note: notePrompt.getResponseText(),
  });

  ui.alert(
    'Mutasi stok tersimpan.\n\n' +
      'Jenis: ' + normalizeMutationLabel_(result.mutationType) + '\n' +
      'SKU: ' + result.sku + '\n' +
      'Qty: ' + formatPlainNumber_(result.qtySigned) + '\n' +
      'Stok Baru: ' + formatPlainNumber_(result.stokSetelah)
  );
}

/**
 * Menampilkan index utama dalam modal.
 */
function showIndex() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_CONFIG.titleIndex)
    .setWidth(980)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, APP_CONFIG.titleIndex);
}

/**
 * Menampilkan dashboard dalam modal.
 */
function showDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle(APP_CONFIG.titleDashboard)
    .setWidth(1240)
    .setHeight(760);
  SpreadsheetApp.getUi().showModalDialog(html, APP_CONFIG.titleDashboard);
}

/**
 * Menampilkan dashboard dalam sidebar.
 */
function showDashboardSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard').setTitle(APP_CONFIG.titleDashboard);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Endpoint Web App + API:
 * - default: index
 * - ?view=kasir: kasir
 * - ?view=produk: manajemen produk
 * - ?view=dashboard: dashboard
 * - ?action=status|kasir-options|product-crud-data|dashboard-data|low-stock-alerts (JSON GET API)
 */
function doGet(e) {
  if (isApiRequest_(e)) {
    return handleApiRequest_(e, 'GET');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  try {
    ensureCoreSheets_(ss);
  } catch (err) {
    // Tetap render halaman agar user bisa lihat status.
    Logger.log('doGet ensureCoreSheets_ error: ' + err);
  }

  const view = safeText_(e && e.parameter && e.parameter.view).toLowerCase();
  let template = 'Index';
  if (view === 'kasir') {
    template = 'Kasir';
  } else if (view === 'produk') {
    template = 'Produk';
  } else if (view === 'dashboard') {
    template = 'Dashboard';
  }

  let title = APP_CONFIG.titleIndex;
  if (template === 'Kasir') {
    title = APP_CONFIG.titleKasir;
  } else if (template === 'Produk') {
    title = APP_CONFIG.titleProduk;
  } else if (template === 'Dashboard') {
    title = APP_CONFIG.titleDashboard;
  }

  return HtmlService.createHtmlOutputFromFile(template)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Endpoint JSON API untuk operasi write dari Netlify/front-end eksternal.
 * Gunakan query `action`, body JSON text/plain, contoh:
 * POST .../exec?action=save-transaction
 * body: {"customerName":"Budi","sku":"BRS-001","qty":2,"note":""}
 * POST .../exec?action=save-stock-mutation
 * body: {"type":"MASUK","sku":"BRS-001","qty":10,"note":"Restok"}
 * POST .../exec?action=save-product
 * body: {"sku":"BRS-003","kategori":"Beras","namaProduk":"Beras 2kg","hargaModal":25000,"satuan":"sak","hargaJual":30000,"stokAwal":50,"stokMinimum":10}
 */
function doPost(e) {
  return handleApiRequest_(e, 'POST');
}

/**
 * Cek apakah request ditujukan ke endpoint API JSON.
 */
function isApiRequest_(e) {
  return !!safeText_(e && e.parameter && e.parameter.action);
}

/**
 * Router API sederhana (GET/POST) untuk kebutuhan Netlify fetch.
 */
function handleApiRequest_(e, method) {
  const action = safeText_(e && e.parameter && e.parameter.action)
    .toLowerCase()
    .replace(/[_\s]+/g, '-');

  if (!action) {
    return jsonOutput_({
      ok: false,
      error: 'Parameter action wajib diisi.',
      method: method,
      generatedAt: new Date().toISOString(),
    });
  }

  try {
    const data = executeApiAction_(action, method, e);
    return jsonOutput_({
      ok: true,
      action: action,
      method: method,
      generatedAt: new Date().toISOString(),
      data: data,
    });
  } catch (err) {
    return jsonOutput_({
      ok: false,
      action: action,
      method: method,
      generatedAt: new Date().toISOString(),
      error: err && err.message ? err.message : String(err),
    });
  }
}

/**
 * Eksekusi action API.
 */
function executeApiAction_(action, method, e) {
  if (method === 'GET') {
    if (action === 'status' || action === 'app-status' || action === 'connection-status') {
      return getAppConnectionStatus();
    }
    if (action === 'kasir-options') {
      return getKasirOptions();
    }
    if (action === 'product-crud-data' || action === 'products-data' || action === 'product-data') {
      return getProductCrudData();
    }
    if (action === 'dashboard-data') {
      return getDashboardData(toBoolean_(e && e.parameter && e.parameter.forceRefresh));
    }
    if (action === 'low-stock-alerts') {
      return getLowStockAlerts();
    }
    throw new Error('Action GET tidak dikenal: ' + action);
  }

  if (method === 'POST') {
    const body = parseApiPayload_(e);
    const payload = body && body.payload ? body.payload : body;

    if (action === 'save-transaction' || action === 'save-kasir-transaction') {
      return saveKasirTransaction(payload || {});
    }
    if (action === 'save-stock-mutation' || action === 'stock-mutation') {
      return saveStockMutation(payload || {});
    }
    if (action === 'save-product' || action === 'upsert-product' || action === 'product-upsert') {
      return saveProductItem(payload || {});
    }
    if (action === 'delete-product' || action === 'product-delete') {
      return deleteProductItem(payload || {});
    }
    if (
      action === 'save-products-batch' ||
      action === 'save-product-batch' ||
      action === 'upsert-products-batch' ||
      action === 'products-batch-upsert'
    ) {
      return saveProductBatch(payload || {});
    }
    if (action === 'sync-recap') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ensureCoreSheets_(ss);
      const result = rebuildRecapAndStock_(ss);
      bumpDashboardVersion_();
      return result;
    }
    if (action === 'setup-sheets') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      return { messages: ensureCoreSheets_(ss, { forceRecapHeaders: true }) };
    }
    if (action === 'generate-dummy-data') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      ensureCoreSheets_(ss, { forceRecapHeaders: true });

      const produkSheet = findSheetByName_(ss, APP_CONFIG.sheets.PRODUK);
      const penjualanSheet = findSheetByName_(ss, APP_CONFIG.sheets.PENJUALAN);
      const mutasiSheet = findSheetByName_(ss, APP_CONFIG.sheets.MUTASI_STOK);

      const dummyProducts = buildDummyProductCatalog_();
      const productRows = dummyProducts.map((item) => ([
        item.sku,
        item.kategori,
        item.namaProduk,
        item.hargaModal,
        item.satuan,
        item.hargaJual,
        item.stokAwal,
        item.stokAwal,
        item.hargaModal * item.stokAwal,
        item.stokMinimum,
      ]));
      const salesRows = buildDummySalesRows_(dummyProducts);

      overwriteSheetWithRows_(produkSheet, SHEET_HEADERS.PRODUK, productRows);
      overwriteSheetWithRows_(penjualanSheet, SHEET_HEADERS.PENJUALAN, salesRows);
      overwriteSheetWithRows_(mutasiSheet, SHEET_HEADERS.MUTASI_STOK, []);

      const recap = rebuildRecapAndStock_(ss);
      bumpDashboardVersion_();

      return {
        totalProduk: productRows.length,
        totalTransaksi: salesRows.length,
        totalPelanggan: recap.totalPelanggan,
      };
    }
    throw new Error('Action POST tidak dikenal: ' + action);
  }

  throw new Error('Method tidak didukung: ' + method);
}

/**
 * Parse payload dari body request API.
 */
function parseApiPayload_(e) {
  const raw = safeText_(e && e.postData && e.postData.contents);
  if (!raw) {
    return {};
  }

  try {
    return JSON.parse(raw);
  } catch (err) {
    // lanjut parse sebagai query string.
  }

  const form = parseQueryString_(raw);
  if (form.payload) {
    try {
      form.payload = JSON.parse(form.payload);
    } catch (err) {
      // payload tetap string.
    }
  }
  return form;
}

/**
 * Parse query-string sederhana (a=1&b=2).
 */
function parseQueryString_(text) {
  const result = {};
  String(text || '')
    .split('&')
    .forEach((part) => {
      if (!part) {
        return;
      }
      const idx = part.indexOf('=');
      const key = idx > -1 ? part.slice(0, idx) : part;
      const val = idx > -1 ? part.slice(idx + 1) : '';
      const decodedKey = decodeURIComponent(String(key).replace(/\+/g, ' '));
      const decodedVal = decodeURIComponent(String(val).replace(/\+/g, ' '));
      if (decodedKey) {
        result[decodedKey] = decodedVal;
      }
    });
  return result;
}

/**
 * Konversi string bool dari query API.
 */
function toBoolean_(value) {
  const text = safeText_(value).toLowerCase();
  return text === '1' || text === 'true' || text === 'yes' || text === 'y';
}

/**
 * Output JSON untuk API.
 */
function jsonOutput_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload || {}))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Trigger ringan untuk invalidasi cache dashboard saat edit data.
 */
function onEdit() {
  bumpDashboardVersion_();
}

/**
 * Trigger ringan untuk invalidasi cache saat ada perubahan struktur.
 */
function onChange() {
  bumpDashboardVersion_();
}

/**
 * Membuat trigger installable onChange untuk refresh dashboard lebih stabil.
 */
function setupDashboardTriggers() {
  const ss = SpreadsheetApp.getActive();
  const ssId = ss.getId();

  const exists = ScriptApp.getProjectTriggers().some((trigger) => {
    return (
      trigger.getHandlerFunction() === 'onChange' &&
      trigger.getTriggerSourceId &&
      trigger.getTriggerSourceId() === ssId
    );
  });

  if (!exists) {
    ScriptApp.newTrigger('onChange').forSpreadsheet(ss).onChange().create();
  }

  SpreadsheetApp.getUi().alert(
    exists
      ? 'Trigger auto-refresh sudah aktif.'
      : 'Trigger auto-refresh berhasil dibuat.'
  );
}

/**
 * Setup sheet + header standar jika sheet kosong / belum ada.
 */
function setupDatabaseSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const messages = ensureCoreSheets_(ss, { forceRecapHeaders: true });

  SpreadsheetApp.getUi().alert('Setup selesai.\n\n' + (messages.join('\n') || 'Tidak ada perubahan.'));
}

/**
 * Mengisi data dummy ke PRODUK dan PENJUALAN lalu rebuild seluruh rekap.
 * Dipakai untuk testing awal dashboard/kasir.
 */
function generateDummyData() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert(
    'Buat Data Dummy',
    'Data pada sheet PRODUK, PENJUALAN, MUTASI_STOK, REKAP_PRODUK, dan REKAP_PELANGGAN akan diganti data contoh.\n\nLanjutkan?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) {
    ui.alert('Pembuatan data dummy dibatalkan.');
    return;
  }

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss, { forceRecapHeaders: true });

    const produkSheet = findSheetByName_(ss, APP_CONFIG.sheets.PRODUK);
    const penjualanSheet = findSheetByName_(ss, APP_CONFIG.sheets.PENJUALAN);
    const mutasiSheet = findSheetByName_(ss, APP_CONFIG.sheets.MUTASI_STOK);

    const dummyProducts = buildDummyProductCatalog_();
    const productRows = dummyProducts.map((item) => ([
      item.sku,
      item.kategori,
      item.namaProduk,
      item.hargaModal,
      item.satuan,
      item.hargaJual,
      item.stokAwal,
      item.stokAwal,
      item.hargaModal * item.stokAwal,
      item.stokMinimum,
    ]));

    const salesRows = buildDummySalesRows_(dummyProducts);

    overwriteSheetWithRows_(produkSheet, SHEET_HEADERS.PRODUK, productRows);
    overwriteSheetWithRows_(penjualanSheet, SHEET_HEADERS.PENJUALAN, salesRows);
    overwriteSheetWithRows_(mutasiSheet, SHEET_HEADERS.MUTASI_STOK, []);

    const recap = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    ui.alert(
      'Data dummy berhasil dibuat.\n\n' +
        'Produk: ' + productRows.length + '\n' +
        'Transaksi: ' + salesRows.length + '\n' +
        'Pelanggan rekap: ' + recap.totalPelanggan
    );
  } finally {
    lock.releaseLock();
  }
}

/**
 * Endpoint data untuk UI kasir.
 */
function getKasirOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoSetupMessages = ensureCoreSheets_(ss);
  const warnings = [];
  const state = collectBusinessState_(ss, warnings);
  const alerts = buildProductAlerts_(state.products);

  const products = state.products
    .map((item) => ({
      sku: item.sku,
      namaProduk: item.namaProduk,
      kategori: item.kategori,
      satuan: item.satuan,
      hargaJual: round2_(item.hargaJual),
      hargaModal: round2_(item.hargaModal),
      stokTersedia: round2_(item.stockCalculated),
      stokMinimum: round2_(item.stokMinimum),
    }))
    .sort((a, b) => a.namaProduk.localeCompare(b.namaProduk));

  const todayKey = Utilities.formatDate(new Date(), state.timezone, 'yyyy-MM-dd');
  const todaySales = state.sales.filter((row) => row.dateKey === todayKey);

  return {
    products: products,
    summary: {
      transaksiHariIni: countDistinctTransactions_(todaySales),
      omzetHariIni: round2_(todaySales.reduce((sum, row) => sum + row.total, 0)),
      labaHariIni: round2_(todaySales.reduce((sum, row) => sum + row.profit, 0)),
    },
    meta: {
      generatedAt: new Date().toISOString(),
      warnings: autoSetupMessages.concat(warnings, alerts.lowStockWarnings, alerts.pricingWarnings),
      lowStockItems: alerts.lowStockItems,
      pricingAlerts: alerts.pricingItems,
      minMarginPercent: APP_CONFIG.minMarginPercent,
    },
  };
}

/**
 * Endpoint data produk untuk UI CRUD manajemen barang.
 */
function getProductCrudData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets_(ss);

  const warnings = [];
  const state = collectBusinessState_(ss, warnings);
  const alerts = buildProductAlerts_(state.products);
  const totalModalStokAkhir = round2_(
    state.products.reduce((sum, item) => {
      return sum + round2_(item.hargaModal * item.stockCalculated);
    }, 0)
  );

  const items = state.products
    .map((item) => ({
      rowNumber: item.rowNumber,
      sku: item.sku,
      kategori: item.kategori || 'Umum',
      namaProduk: item.namaProduk,
      satuan: item.satuan || 'pcs',
      hargaModal: round2_(item.hargaModal),
      hargaJual: round2_(item.hargaJual),
      stokAwal: round2_(item.stokAwal),
      stokAkhir: round2_(item.stockCalculated),
      stokMinimum: round2_(item.stokMinimum),
      modal: round2_(item.stokAwal * item.hargaModal),
      marginPercent: calcMarginPercent_(item.hargaJual, item.hargaModal),
    }))
    .sort((a, b) => a.namaProduk.localeCompare(b.namaProduk));

  return {
    generatedAt: new Date().toISOString(),
    total: items.length,
    items: items,
    summary: {
      totalProduk: items.length,
      totalLowStock: alerts.lowStockItems.length,
      totalPricingAlerts: alerts.pricingItems.length,
      totalModalStokAkhir: totalModalStokAkhir,
    },
    lowStockItems: alerts.lowStockItems,
    pricingAlerts: alerts.pricingItems,
    warnings: warnings.concat(alerts.lowStockWarnings, alerts.pricingWarnings),
  };
}

/**
 * Create / update 1 produk berdasarkan SKU.
 * Jika SKU sudah ada maka update, jika belum ada maka create.
 */
function saveProductItem(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);

    const result = upsertProductRows_(ss, [payload || {}], { allowEmpty: false });
    const recap = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    return {
      success: true,
      action: result.items[0] ? result.items[0].action : 'updated',
      product: result.items[0] || null,
      createdCount: result.createdCount,
      updatedCount: result.updatedCount,
      recap: recap,
      warnings: (result.warnings || []).concat(recap.warnings || []),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Batch create / update produk dari array items.
 * @param {{items:Array<Object>}} payload
 */
function saveProductBatch(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);

    const items = Array.isArray(payload && payload.items) ? payload.items : [];
    if (!items.length) {
      throw new Error('Data batch kosong. Tambahkan minimal 1 baris produk.');
    }

    const result = upsertProductRows_(ss, items, { allowEmpty: false });
    const recap = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    return {
      success: true,
      totalInput: items.length,
      createdCount: result.createdCount,
      updatedCount: result.updatedCount,
      items: result.items,
      recap: recap,
      warnings: (result.warnings || []).concat(recap.warnings || []),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Hapus produk berdasarkan SKU.
 * Jika SKU sudah dipakai pada transaksi/mutasi, wajib forceDelete=true.
 */
function deleteProductItem(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);

    const skuInput = safeText_(payload && payload.sku);
    const forceDelete = toBoolean_(payload && payload.forceDelete);
    if (!skuInput) {
      throw new Error('SKU wajib diisi untuk hapus produk.');
    }

    const warnings = [];
    const state = collectBusinessState_(ss, warnings);
    const skuKey = normalizeSku_(skuInput);
    const product = state.productsBySku[skuKey];
    if (!product) {
      throw new Error('SKU tidak ditemukan pada sheet PRODUK: ' + skuInput);
    }

    const salesRefCount = state.sales.filter((row) => row.skuKey === skuKey).length;
    const mutationRefCount = state.mutations.filter((row) => row.skuKey === skuKey).length;
    if ((salesRefCount > 0 || mutationRefCount > 0) && !forceDelete) {
      throw new Error(
        'SKU sudah dipakai di ' + salesRefCount + ' baris penjualan dan ' + mutationRefCount +
        ' baris mutasi. Ulangi dengan forceDelete=true jika tetap ingin menghapus.'
      );
    }

    if (!state.productsContext.sheet) {
      throw new Error('Sheet PRODUK tidak ditemukan.');
    }

    state.productsContext.sheet.deleteRow(product.rowNumber);

    const recap = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    return {
      success: true,
      deletedSku: product.sku,
      namaProduk: product.namaProduk,
      references: {
        penjualan: salesRefCount,
        mutasi: mutationRefCount,
      },
      forceDelete: forceDelete,
      recap: recap,
      warnings: warnings.concat(recap.warnings || []),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Upsert baris produk ke sheet PRODUK.
 */
function upsertProductRows_(ss, inputItems, options) {
  const opts = options || {};
  const warnings = [];
  const ctx = getSheetContext_(ss, APP_CONFIG.sheets.PRODUK, warnings);
  if (!ctx.sheet) {
    throw new Error('Sheet PRODUK tidak ditemukan.');
  }

  const cols = resolveColumns_(ctx.headers, {
    sku: FIELD_ALIASES.sku,
    kategori: FIELD_ALIASES.kategori,
    namaProduk: FIELD_ALIASES.namaProduk,
    hargaModal: FIELD_ALIASES.hargaModal,
    satuan: FIELD_ALIASES.satuan,
    hargaJual: FIELD_ALIASES.hargaJual,
    stokAwal: FIELD_ALIASES.stokAwal,
    stokAkhir: FIELD_ALIASES.stokAkhir,
    modal: FIELD_ALIASES.modal,
    stokMinimum: FIELD_ALIASES.stokMinimum,
  });

  ensureRequiredColumns_(
    cols,
    ['sku', 'namaProduk', 'hargaModal', 'satuan', 'hargaJual', 'stokAwal', 'stokAkhir', 'modal', 'stokMinimum'],
    APP_CONFIG.sheets.PRODUK
  );

  const normalizedItems = (inputItems || []).map((item, idx) => normalizeProductPayloadItem_(item, idx));
  if (!normalizedItems.length && !opts.allowEmpty) {
    throw new Error('Tidak ada data produk untuk disimpan.');
  }

  const existingBySku = {};
  ctx.rows.forEach((row) => {
    const sku = safeText_(row.values[cols.sku]);
    if (!sku) {
      return;
    }
    existingBySku[normalizeSku_(sku)] = {
      rowNumber: row.rowNumber,
      values: row.values,
    };
  });

  const rowLength = Math.max(ctx.headers.length, SHEET_HEADERS.PRODUK.length);
  const updates = [];
  const newRows = [];
  const seenInputSku = {};
  const savedItems = [];

  normalizedItems.forEach((item) => {
    if (seenInputSku[item.skuKey]) {
      throw new Error('SKU duplikat pada input batch: ' + item.sku);
    }
    seenInputSku[item.skuKey] = true;

    const existing = existingBySku[item.skuKey];
    const rowValues = existing
      ? padRowToLength_(existing.values, rowLength)
      : new Array(rowLength).fill('');

    assignRowValue_(rowValues, cols.sku, item.sku);
    assignRowValue_(rowValues, cols.kategori, item.kategori);
    assignRowValue_(rowValues, cols.namaProduk, item.namaProduk);
    assignRowValue_(rowValues, cols.hargaModal, item.hargaModal);
    assignRowValue_(rowValues, cols.satuan, item.satuan);
    assignRowValue_(rowValues, cols.hargaJual, item.hargaJual);
    assignRowValue_(rowValues, cols.stokAwal, item.stokAwal);
    assignRowValue_(rowValues, cols.stokAkhir, item.hasStokAkhir ? item.stokAkhir : item.stokAwal);
    assignRowValue_(rowValues, cols.modal, round2_(item.stokAwal * item.hargaModal));
    assignRowValue_(rowValues, cols.stokMinimum, item.stokMinimum);

    if (existing) {
      updates.push({ rowNumber: existing.rowNumber, values: rowValues });
    } else {
      newRows.push(rowValues);
    }

    savedItems.push({
      action: existing ? 'updated' : 'created',
      sku: item.sku,
      kategori: item.kategori,
      namaProduk: item.namaProduk,
      satuan: item.satuan,
      hargaModal: item.hargaModal,
      hargaJual: item.hargaJual,
      stokAwal: item.stokAwal,
      stokMinimum: item.stokMinimum,
      stokAkhir: item.hasStokAkhir ? item.stokAkhir : item.stokAwal,
    });
  });

  updates.forEach((entry) => {
    ctx.sheet.getRange(entry.rowNumber, 1, 1, rowLength).setValues([entry.values]);
  });

  if (newRows.length) {
    const startRow = ctx.sheet.getLastRow() + 1;
    ctx.sheet.getRange(startRow, 1, newRows.length, rowLength).setValues(newRows);
  }

  return {
    createdCount: newRows.length,
    updatedCount: updates.length,
    items: savedItems,
    warnings: warnings,
  };
}

/**
 * Validasi dan normalisasi 1 payload produk.
 */
function normalizeProductPayloadItem_(payload, index) {
  const idx = Number(index || 0) + 1;
  const sku = normalizeSku_(payload && payload.sku);
  const namaProduk = safeText_(payload && payload.namaProduk);
  const kategori = safeText_(payload && payload.kategori) || 'Umum';
  const satuan = safeText_(payload && payload.satuan) || 'pcs';
  const hargaModal = round2_(toNumber_(payload && payload.hargaModal));
  const hargaJual = round2_(toNumber_(payload && payload.hargaJual));
  const stokAwal = round2_(toNumber_(payload && payload.stokAwal));
  const stokMinimum = round2_(toNumber_(payload && payload.stokMinimum));

  const hasStokAkhir = payload && payload.stokAkhir !== undefined && payload.stokAkhir !== null &&
    safeText_(payload.stokAkhir) !== '';
  const stokAkhir = hasStokAkhir ? round2_(toNumber_(payload.stokAkhir)) : 0;

  if (!sku) {
    throw new Error('SKU item ke-' + idx + ' wajib diisi.');
  }
  if (!namaProduk) {
    throw new Error('Nama produk item ke-' + idx + ' wajib diisi.');
  }
  if (hargaModal < 0) {
    throw new Error('Harga modal item ke-' + idx + ' tidak boleh negatif.');
  }
  if (hargaJual < 0) {
    throw new Error('Harga jual item ke-' + idx + ' tidak boleh negatif.');
  }
  if (stokAwal < 0) {
    throw new Error('Stok awal item ke-' + idx + ' tidak boleh negatif.');
  }
  if (stokMinimum < 0) {
    throw new Error('Stok minimum item ke-' + idx + ' tidak boleh negatif.');
  }
  if (hasStokAkhir && stokAkhir < 0) {
    throw new Error('Stok akhir item ke-' + idx + ' tidak boleh negatif.');
  }

  return {
    sku: sku,
    skuKey: sku,
    kategori: kategori,
    namaProduk: namaProduk,
    satuan: satuan,
    hargaModal: hargaModal,
    hargaJual: hargaJual,
    stokAwal: stokAwal,
    stokMinimum: stokMinimum,
    hasStokAkhir: hasStokAkhir,
    stokAkhir: stokAkhir,
  };
}

/**
 * Pad array row ke panjang tertentu.
 */
function padRowToLength_(row, targetLength) {
  const out = new Array(Math.max(0, targetLength)).fill('');
  const source = Array.isArray(row) ? row : [];
  for (let i = 0; i < out.length && i < source.length; i += 1) {
    out[i] = source[i];
  }
  return out;
}

/**
 * Simpan transaksi dari form kasir.
 * Mendukung:
 * - mode lama: {customerName, sku, qty, note}
 * - mode cart: {customerName, note, items:[{sku, qty, note?}]}
 * @param {{customerName:string, sku?:string, qty?:number|string, note?:string, items?:Array<{sku:string, qty:number|string, note?:string}>}} payload
 */
function saveKasirTransaction(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);

    const customerName = safeText_(payload && payload.customerName);
    const invoiceNote = safeText_(payload && payload.note);
    const items = normalizeTransactionItems_(payload);

    if (!customerName) {
      throw new Error('Nama pelanggan wajib diisi.');
    }
    if (!items.length) {
      throw new Error('Keranjang kosong. Tambahkan minimal 1 item.');
    }

    const warnings = [];
    const stateBefore = collectBusinessState_(ss, warnings);
    const salesCtx = stateBefore.salesContext;
    if (!salesCtx.sheet) {
      throw new Error('Sheet PENJUALAN tidak ditemukan.');
    }

    const salesCols = resolveColumns_(salesCtx.headers, {
      tanggal: FIELD_ALIASES.tanggal,
      invoice: FIELD_ALIASES.invoice,
      pelanggan: FIELD_ALIASES.pelanggan,
      sku: FIELD_ALIASES.sku,
      namaProduk: FIELD_ALIASES.namaProduk,
      satuan: FIELD_ALIASES.satuan,
      hargaSatuan: FIELD_ALIASES.hargaSatuan,
      hargaModalTransaksi: FIELD_ALIASES.hargaModalTransaksi,
      qty: FIELD_ALIASES.qty,
      total: FIELD_ALIASES.total,
      hpp: FIELD_ALIASES.hpp,
      labaKotor: FIELD_ALIASES.labaKotor,
      catatan: FIELD_ALIASES.catatan,
    });

    ensureRequiredColumns_(salesCols, ['tanggal', 'pelanggan', 'sku', 'qty', 'total'], APP_CONFIG.sheets.PENJUALAN);

    const requestedBySku = {};
    items.forEach((item) => {
      const skuKey = normalizeSku_(item.sku);
      requestedBySku[skuKey] = round2_((requestedBySku[skuKey] || 0) + round2_(item.qty));
    });

    const validatedItems = items.map((item) => {
      const skuKey = normalizeSku_(item.sku);
      const product = stateBefore.productsBySku[skuKey];
      if (!product) {
        throw new Error('SKU tidak ditemukan pada sheet PRODUK: ' + item.sku);
      }

      const availableStock = round2_(product.stockCalculated);
      const requestedQty = round2_(requestedBySku[skuKey] || 0);
      if (requestedQty > availableStock) {
        throw new Error(
          'Stok tidak cukup. Stok tersedia untuk ' + product.namaProduk + ': ' + formatPlainNumber_(availableStock)
        );
      }

      const hargaSatuan = round2_(product.hargaJual);
      const hargaModalSatuan = round2_(product.hargaModal);
      const total = round2_(hargaSatuan * round2_(item.qty));
      const hpp = round2_(round2_(item.qty) * hargaModalSatuan);
      const labaKotor = round2_(total - hpp);

      return {
        sku: product.sku,
        skuKey: skuKey,
        namaProduk: product.namaProduk,
        satuan: product.satuan,
        qty: round2_(item.qty),
        hargaSatuan: hargaSatuan,
        hargaModalSatuan: hargaModalSatuan,
        total: total,
        hpp: hpp,
        labaKotor: labaKotor,
        note: safeText_(item.note) || invoiceNote,
        stokSisa: round2_(availableStock - requestedQty),
      };
    });

    const invoiceNumber = generateInvoiceNumber_(stateBefore.timezone);
    const rowLength = Math.max(salesCtx.headers.length, SHEET_HEADERS.PENJUALAN.length);
    const now = new Date();
    const newRows = validatedItems.map((item) => {
      const newRow = new Array(rowLength).fill('');
      assignRowValue_(newRow, salesCols.tanggal, now);
      assignRowValue_(newRow, salesCols.invoice, invoiceNumber);
      assignRowValue_(newRow, salesCols.pelanggan, customerName);
      assignRowValue_(newRow, salesCols.sku, item.sku);
      assignRowValue_(newRow, salesCols.namaProduk, item.namaProduk);
      assignRowValue_(newRow, salesCols.satuan, item.satuan);
      assignRowValue_(newRow, salesCols.hargaSatuan, item.hargaSatuan);
      assignRowValue_(newRow, salesCols.hargaModalTransaksi, item.hargaModalSatuan);
      assignRowValue_(newRow, salesCols.qty, item.qty);
      assignRowValue_(newRow, salesCols.total, item.total);
      assignRowValue_(newRow, salesCols.hpp, item.hpp);
      assignRowValue_(newRow, salesCols.labaKotor, item.labaKotor);
      assignRowValue_(newRow, salesCols.catatan, item.note);
      return newRow;
    });

    if (newRows.length === 1) {
      salesCtx.sheet.appendRow(newRows[0]);
    } else {
      const startRow = salesCtx.sheet.getLastRow() + 1;
      salesCtx.sheet.getRange(startRow, 1, newRows.length, rowLength).setValues(newRows);
    }

    const syncResult = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();
    const postWarnings = [];
    const stateAfter = collectBusinessState_(ss, postWarnings);
    const alertsAfter = buildProductAlerts_(stateAfter.products);

    const totalQty = round2_(validatedItems.reduce((sum, item) => sum + item.qty, 0));
    const totalBelanja = round2_(validatedItems.reduce((sum, item) => sum + item.total, 0));

    return {
      success: true,
      message: 'Transaksi berhasil disimpan.',
      transaksi: {
        invoice: invoiceNumber,
        tanggal: now.toISOString(),
        pelanggan: customerName,
        totalItem: validatedItems.length,
        totalQty: totalQty,
        total: totalBelanja,
        items: validatedItems.map((item) => ({
          sku: item.sku,
          namaProduk: item.namaProduk,
          qty: item.qty,
          hargaSatuan: item.hargaSatuan,
          hargaModalSatuan: item.hargaModalSatuan,
          total: item.total,
          hpp: item.hpp,
          labaKotor: item.labaKotor,
          stokSisa: item.stokSisa,
        })),
      },
      recap: syncResult,
      warnings: warnings.concat(postWarnings, alertsAfter.lowStockWarnings, alertsAfter.pricingWarnings),
    };
  } finally {
    lock.releaseLock();
  }
}
/**
 * Jalankan sinkronisasi stok + rekap secara manual dari menu.
 */
function syncRecapAndStock() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);
    const result = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    SpreadsheetApp.getUi().alert(
      'Sinkronisasi selesai.\n\n' +
        'Produk: ' + result.totalProduk + '\n' +
        'Pelanggan: ' + result.totalPelanggan + '\n' +
        'Transaksi: ' + result.totalTransaksi + '\n' +
        'Mutasi Stok: ' + (result.totalMutasi || 0) + '\n' +
        'Alert Stok Minimum: ' + (result.totalAlertStokMinimum || 0)
    );

    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Simpan mutasi stok:
 * - MASUK: tambah stok
 * - RETUR: tambah stok
 * - ADJUSTMENT: tambah/kurang stok (qty bisa +/-)
 * @param {{type:string,sku:string,qty:number|string,note?:string}} payload
 */
function saveStockMutation(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);

    const mutationType = normalizeMutationType_(payload && payload.type);
    const skuInput = safeText_(payload && payload.sku);
    const qtyRaw = toNumber_(payload && payload.qty);
    const note = safeText_(payload && payload.note);

    if (['MASUK', 'RETUR', 'ADJUSTMENT'].indexOf(mutationType) === -1) {
      throw new Error('Tipe mutasi tidak valid. Gunakan MASUK, RETUR, atau ADJUSTMENT.');
    }

    if (!skuInput) {
      throw new Error('SKU wajib diisi.');
    }

    let qtySigned = 0;
    if (mutationType === 'ADJUSTMENT') {
      qtySigned = round2_(qtyRaw);
      if (!qtySigned) {
        throw new Error('Qty adjustment tidak boleh 0.');
      }
    } else {
      qtySigned = round2_(Math.abs(qtyRaw));
      if (!qtySigned) {
        throw new Error('Qty harus lebih dari 0.');
      }
    }

    const warnings = [];
    const stateBefore = collectBusinessState_(ss, warnings);
    const skuKey = normalizeSku_(skuInput);
    const product = stateBefore.productsBySku[skuKey];
    if (!product) {
      throw new Error('SKU tidak ditemukan pada sheet PRODUK.');
    }

    const stokSebelum = round2_(product.stockCalculated);
    const stokSetelah = round2_(stokSebelum + qtySigned);
    if (stokSetelah < 0) {
      throw new Error(
        'Mutasi menyebabkan stok minus. Stok saat ini ' + product.namaProduk + ': ' + formatPlainNumber_(stokSebelum)
      );
    }

    const mutCtx = getSheetContext_(ss, APP_CONFIG.sheets.MUTASI_STOK, warnings);
    if (!mutCtx.sheet) {
      throw new Error('Sheet MUTASI_STOK tidak ditemukan.');
    }

    const mutCols = resolveColumns_(mutCtx.headers, {
      tanggal: FIELD_ALIASES.tanggal,
      jenisMutasi: FIELD_ALIASES.jenisMutasi,
      sku: FIELD_ALIASES.sku,
      namaProduk: FIELD_ALIASES.namaProduk,
      qty: FIELD_ALIASES.qty,
      catatan: FIELD_ALIASES.catatan,
    });

    ensureRequiredColumns_(mutCols, ['tanggal', 'jenisMutasi', 'sku', 'qty'], APP_CONFIG.sheets.MUTASI_STOK);

    const rowLength = Math.max(mutCtx.headers.length, SHEET_HEADERS.MUTASI_STOK.length);
    const newRow = new Array(rowLength).fill('');
    assignRowValue_(newRow, mutCols.tanggal, new Date());
    assignRowValue_(newRow, mutCols.jenisMutasi, mutationType);
    assignRowValue_(newRow, mutCols.sku, product.sku);
    assignRowValue_(newRow, mutCols.namaProduk, product.namaProduk);
    assignRowValue_(newRow, mutCols.qty, qtySigned);
    assignRowValue_(newRow, mutCols.catatan, note);
    mutCtx.sheet.appendRow(newRow);

    const recap = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    return {
      success: true,
      mutationType: mutationType,
      sku: product.sku,
      namaProduk: product.namaProduk,
      qtySigned: qtySigned,
      stokSebelum: stokSebelum,
      stokSetelah: stokSetelah,
      recap: recap,
      warnings: warnings.concat(recap.warnings || []),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ambil daftar produk yang stoknya di bawah / sama dengan stok minimum.
 */
function getLowStockAlerts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets_(ss);

  const warnings = [];
  const state = collectBusinessState_(ss, warnings);
  const alerts = buildProductAlerts_(state.products);

  return {
    generatedAt: new Date().toISOString(),
    total: alerts.lowStockItems.length,
    items: alerts.lowStockItems,
    warnings: warnings.concat(alerts.lowStockWarnings, alerts.pricingWarnings),
  };
}

/**
 * Rebuild stok + rekap berdasarkan data PRODUK & PENJUALAN.
 */
function rebuildRecapAndStock_(ss) {
  const warnings = [];
  const state = collectBusinessState_(ss, warnings);
  const alerts = buildProductAlerts_(state.products);

  syncProductStockColumns_(state);

  const salesAggBySku = {};
  state.sales.forEach((row) => {
    const skuKey = row.skuKey;
    if (!salesAggBySku[skuKey]) {
      salesAggBySku[skuKey] = {
        qty: 0,
        omzet: 0,
        hpp: 0,
      };
    }
    salesAggBySku[skuKey].qty = round2_(salesAggBySku[skuKey].qty + row.qty);
    salesAggBySku[skuKey].omzet = round2_(salesAggBySku[skuKey].omzet + row.total);
    salesAggBySku[skuKey].hpp = round2_(salesAggBySku[skuKey].hpp + row.hpp);
  });

  const rekapProdukRows = state.products
    .map((item) => {
      const aggr = salesAggBySku[item.skuKey] || { qty: 0, omzet: 0, hpp: 0 };
      const qtyTerjual = round2_(aggr.qty);
      const omzet = round2_(aggr.omzet);
      const hpp = round2_(aggr.hpp);
      const labaKotor = round2_(omzet - hpp);
      const hargaModalAvg = qtyTerjual ? round2_(hpp / qtyTerjual) : round2_(item.hargaModal);
      const hargaJualAvg = qtyTerjual ? round2_(omzet / qtyTerjual) : round2_(item.hargaJual);
      return [
        item.sku,
        item.namaProduk,
        item.satuan,
        hargaModalAvg,
        hargaJualAvg,
        qtyTerjual,
        omzet,
        hpp,
        labaKotor,
      ];
    })
    .sort((a, b) => b[5] - a[5]);

  const pelangganRows = Object.keys(state.customerAgg)
    .map((name) => {
      const aggr = state.customerAgg[name];
      return [
        name,
        aggr.transaksi,
        round2_(aggr.qty),
        round2_(aggr.belanja),
        round2_(aggr.hpp),
        round2_(aggr.profit),
      ];
    })
    .sort((a, b) => b[3] - a[3]);

  writeTableSheet_(ss, APP_CONFIG.sheets.REKAP_PRODUK, SHEET_HEADERS.REKAP_PRODUK, rekapProdukRows);
  writeTableSheet_(ss, APP_CONFIG.sheets.REKAP_PELANGGAN, SHEET_HEADERS.REKAP_PELANGGAN, pelangganRows);

  return {
    totalProduk: rekapProdukRows.length,
    totalPelanggan: pelangganRows.length,
    totalTransaksi: countDistinctTransactions_(state.sales),
    totalMutasi: state.mutations.length,
    totalAlertStokMinimum: alerts.lowStockItems.length,
    warnings: warnings.concat(alerts.lowStockWarnings, alerts.pricingWarnings),
  };
}

/**
 * Endpoint data dashboard (pakai cache dokumen agar UI tetap cepat).
 */
function getDashboardData(forceRefresh) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets_(ss);

  const version = getDashboardVersion_();
  const cacheKey = APP_CONFIG.cachePrefix + version;
  const cache = CacheService.getDocumentCache();

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (err) {
        // cache invalid, hitung ulang.
      }
    }
  }

  const payload = buildDashboardPayload_();
  payload.meta.version = version;

  try {
    cache.put(cacheKey, JSON.stringify(payload), APP_CONFIG.cacheTtlSeconds);
  } catch (err) {
    // Jika cache penuh, abaikan tanpa menggagalkan UI.
  }

  return payload;
}

/**
 * Membentuk payload KPI + chart untuk dashboard.
 */
function buildDashboardPayload_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCoreSheets_(ss);

  const warnings = [];
  const state = collectBusinessState_(ss, warnings);
  const alerts = buildProductAlerts_(state.products);
  const todayKey = Utilities.formatDate(new Date(), state.timezone, 'yyyy-MM-dd');
  const totalModalStokAkhir = round2_(
    state.products.reduce((sum, item) => {
      return sum + round2_(item.hargaModal * item.stockCalculated);
    }, 0)
  );

  const omzetPerHari = {};
  const qtyPerProduk = {};
  const profitBySku = {};
  const profitByKategori = {};
  const pelangganMap = {};
  const pelangganTransaksiSet = {};
  const transactionMap = {};

  let omzetHariIni = 0;
  let totalLabaKotor = 0;

  state.sales.forEach((row) => {
    if (row.dateKey) {
      omzetPerHari[row.dateKey] = (omzetPerHari[row.dateKey] || 0) + row.total;
      if (row.dateKey === todayKey) {
        omzetHariIni += row.total;
      }
    }

    qtyPerProduk[row.namaProduk] = (qtyPerProduk[row.namaProduk] || 0) + row.qty;
    profitBySku[row.skuKey] = round2_((profitBySku[row.skuKey] || 0) + row.profit);

    const kategori = safeText_(
      (state.productsBySku[row.skuKey] && state.productsBySku[row.skuKey].kategori) || 'Tanpa Kategori'
    ) || 'Tanpa Kategori';
    profitByKategori[kategori] = round2_((profitByKategori[kategori] || 0) + row.profit);

    if (!pelangganMap[row.pelanggan]) {
      pelangganMap[row.pelanggan] = { pelanggan: row.pelanggan, belanja: 0, profit: 0, qty: 0, transaksi: 0 };
    }
    if (!pelangganTransaksiSet[row.pelanggan]) {
      pelangganTransaksiSet[row.pelanggan] = {};
    }

    pelangganMap[row.pelanggan].belanja += row.total;
    pelangganMap[row.pelanggan].profit += row.profit;
    pelangganMap[row.pelanggan].qty += row.qty;
    const transaksiKey = getSalesTransactionKey_(row);
    if (!pelangganTransaksiSet[row.pelanggan][transaksiKey]) {
      pelangganTransaksiSet[row.pelanggan][transaksiKey] = true;
      pelangganMap[row.pelanggan].transaksi += 1;
    }

    if (!transactionMap[transaksiKey]) {
      transactionMap[transaksiKey] = {
        key: transaksiKey,
        dateKey: row.dateKey,
        pelanggan: row.pelanggan,
        total: 0,
        profit: 0,
        qty: 0,
      };
    }
    transactionMap[transaksiKey].total = round2_(transactionMap[transaksiKey].total + row.total);
    transactionMap[transaksiKey].profit = round2_(transactionMap[transaksiKey].profit + row.profit);
    transactionMap[transaksiKey].qty = round2_(transactionMap[transaksiKey].qty + row.qty);

    totalLabaKotor += row.profit;
  });

  const transactions = Object.keys(transactionMap).map((key) => transactionMap[key]);
  const totalTransaksi = transactions.length;
  const totalOmzet = round2_(transactions.reduce((sum, trx) => sum + trx.total, 0));
  const aov = totalTransaksi ? round2_(totalOmzet / totalTransaksi) : 0;

  const customerTransactionCount = {};
  transactions.forEach((trx) => {
    customerTransactionCount[trx.pelanggan] = (customerTransactionCount[trx.pelanggan] || 0) + 1;
  });
  const totalCustomerForRepeat = Object.keys(customerTransactionCount).length;
  const repeatCustomerCount = Object.keys(customerTransactionCount).filter(
    (name) => customerTransactionCount[name] >= 2
  ).length;
  const repeatCustomerRate = totalCustomerForRepeat
    ? round2_((repeatCustomerCount / totalCustomerForRepeat) * 100)
    : 0;

  const repeatTrendMap = {};
  transactions.forEach((trx) => {
    if (!trx.dateKey) {
      return;
    }
    if ((customerTransactionCount[trx.pelanggan] || 0) >= 2) {
      repeatTrendMap[trx.dateKey] = (repeatTrendMap[trx.dateKey] || 0) + 1;
    }
  });
  const repeatTrendRows = Object.keys(repeatTrendMap)
    .sort()
    .map((dateKey) => ({
      dateKey: dateKey,
      label: formatDateKey_(dateKey, state.timezone),
      repeatTransaksi: round2_(repeatTrendMap[dateKey]),
    }));

  const omzetHarianRows = Object.keys(omzetPerHari)
    .sort()
    .map((dateKey) => ({
      dateKey: dateKey,
      label: formatDateKey_(dateKey, state.timezone),
      omzet: round2_(omzetPerHari[dateKey]),
    }));

  const qtyProdukRows = Object.keys(qtyPerProduk)
    .map((name) => ({ name: name, qty: round2_(qtyPerProduk[name]) }))
    .filter((row) => row.qty !== 0)
    .sort((a, b) => b.qty - a.qty)
    .slice(0, APP_CONFIG.chartItemLimit);

  const produkProfitRows = Object.keys(profitBySku)
    .map((skuKey) => {
      const product = state.productsBySku[skuKey];
      return {
        sku: product ? product.sku : skuKey,
        namaProduk: product ? product.namaProduk : skuKey,
        kategori: product ? product.kategori : '-',
        profit: round2_(profitBySku[skuKey]),
      };
    })
    .sort((a, b) => b.profit - a.profit);

  const topProfitProduct = produkProfitRows.length
    ? produkProfitRows[0]
    : { sku: '', namaProduk: '-', kategori: '-', profit: 0 };

  const profitKategoriRows = Object.keys(profitByKategori)
    .map((kategori) => ({
      kategori: kategori,
      profit: round2_(profitByKategori[kategori]),
    }))
    .sort((a, b) => b.profit - a.profit)
    .slice(0, APP_CONFIG.chartItemLimit);

  const pelangganRows = Object.keys(pelangganMap)
    .map((name) => ({
      pelanggan: name,
      belanja: round2_(pelangganMap[name].belanja),
      profit: round2_(pelangganMap[name].profit),
      qty: round2_(pelangganMap[name].qty),
      transaksi: pelangganMap[name].transaksi,
    }))
    .sort((a, b) => (b.belanja - a.belanja) || (b.profit - a.profit));

  const pieRows = compressTopRows_(
    pelangganRows.filter((row) => row.profit > 0).sort((a, b) => b.profit - a.profit),
    APP_CONFIG.pieItemLimit,
    'Lainnya'
  );

  return {
    meta: {
      generatedAt: new Date().toISOString(),
      timezone: state.timezone,
      sourceSheetStatus: {
        PRODUK: !!state.productsContext.sheet,
        PENJUALAN: !!state.salesContext.sheet,
        MUTASI_STOK: !!state.mutationContext.sheet,
        REKAP_PRODUK: !!findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PRODUK),
        REKAP_PELANGGAN: !!findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PELANGGAN),
      },
      warnings: warnings.concat(alerts.lowStockWarnings, alerts.pricingWarnings),
      lowStockItems: alerts.lowStockItems,
      pricingAlerts: alerts.pricingItems,
      minMarginPercent: APP_CONFIG.minMarginPercent,
    },
    kpi: {
      omzetHariIni: round2_(omzetHariIni),
      totalLabaKotor: round2_(totalLabaKotor),
      totalModalStokAkhir: totalModalStokAkhir,
      totalTransaksi: totalTransaksi,
      totalPelanggan: pelangganRows.length,
      aov: aov,
      repeatCustomerCount: repeatCustomerCount,
      repeatCustomerRate: repeatCustomerRate,
      produkPalingUntung: {
        sku: topProfitProduct.sku,
        namaProduk: topProfitProduct.namaProduk,
        kategori: topProfitProduct.kategori,
        profit: round2_(topProfitProduct.profit),
      },
    },
    charts: {
      omzetHarian: {
        labels: omzetHarianRows.map((row) => row.label),
        values: omzetHarianRows.map((row) => row.omzet),
      },
      qtyProduk: {
        labels: qtyProdukRows.map((row) => row.name),
        values: qtyProdukRows.map((row) => row.qty),
      },
      profitPelanggan: {
        labels: pieRows.map((row) => row.pelanggan),
        values: pieRows.map((row) => row.profit),
      },
      repeatCustomerTrend: {
        labels: repeatTrendRows.map((row) => row.label),
        values: repeatTrendRows.map((row) => row.repeatTransaksi),
      },
      profitKategori: {
        labels: profitKategoriRows.map((row) => row.kategori),
        values: profitKategoriRows.map((row) => row.profit),
      },
    },
    tables: {
      pelanggan: pelangganRows.slice(0, 100),
      produk: qtyProdukRows,
      produkProfit: produkProfitRows.slice(0, 100),
      profitKategori: profitKategoriRows,
    },
  };
}

/**
 * Auto-create sheet inti bila belum ada.
 * Secara default hanya membuat sheet/headers yang kosong.
 * @param {SpreadsheetApp.Spreadsheet} ss
 * @param {{forceRecapHeaders?: boolean}=} options
 * @return {string[]} daftar aksi yang dilakukan.
 */
function ensureCoreSheets_(ss, options) {
  const opts = options || {};
  const messages = [];

  ensureSheetExistsWithHeader_(ss, APP_CONFIG.sheets.PRODUK, SHEET_HEADERS.PRODUK, messages);
  ensureProdukMinimumStockColumn_(ss, messages);
  ensureSheetExistsWithHeader_(ss, APP_CONFIG.sheets.PENJUALAN, SHEET_HEADERS.PENJUALAN, messages);
  ensurePenjualanInvoiceColumn_(ss, messages);
  ensurePenjualanHistoricalColumns_(ss, messages);
  backfillPenjualanSnapshotRows_(ss, messages);
  ensureSheetExistsWithHeader_(ss, APP_CONFIG.sheets.MUTASI_STOK, SHEET_HEADERS.MUTASI_STOK, messages);
  ensureSheetExistsWithHeader_(
    ss,
    APP_CONFIG.sheets.REKAP_PRODUK,
    SHEET_HEADERS.REKAP_PRODUK,
    messages,
    !!opts.forceRecapHeaders
  );
  ensureSheetExistsWithHeader_(
    ss,
    APP_CONFIG.sheets.REKAP_PELANGGAN,
    SHEET_HEADERS.REKAP_PELANGGAN,
    messages,
    !!opts.forceRecapHeaders
  );

  return messages;
}

/**
 * Pastikan sheet PRODUK memiliki kolom stok minimum.
 */
function ensureProdukMinimumStockColumn_(ss, messages) {
  const sheet = findSheetByName_(ss, APP_CONFIG.sheets.PRODUK);
  if (!sheet) {
    return;
  }

  const lastCol = sheet.getLastColumn();
  const readCols = Math.max(lastCol, 1);
  const headers = sheet.getRange(1, 1, 1, readCols).getValues()[0].map((cell) => safeText_(cell));
  const cols = resolveColumns_(headers, { stokMinimum: FIELD_ALIASES.stokMinimum });
  if (cols.stokMinimum > -1) {
    return;
  }

  const targetCol = readCols + 1;
  if (sheet.getMaxColumns() < targetCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCol - sheet.getMaxColumns());
  }
  sheet.getRange(1, targetCol).setValue(SHEET_HEADERS.PRODUK[SHEET_HEADERS.PRODUK.length - 1]);
  messages.push('Menambahkan kolom stok minimum pada sheet ' + APP_CONFIG.sheets.PRODUK);
}

/**
 * Pastikan sheet PENJUALAN memiliki kolom invoice (untuk transaksi multi-item).
 * Kolom ditambahkan di ujung header agar aman terhadap data lama.
 */
function ensurePenjualanInvoiceColumn_(ss, messages) {
  const sheet = findSheetByName_(ss, APP_CONFIG.sheets.PENJUALAN);
  if (!sheet) {
    return;
  }

  const lastCol = sheet.getLastColumn();
  const readCols = Math.max(lastCol, 1);
  const headers = sheet.getRange(1, 1, 1, readCols).getValues()[0].map((cell) => safeText_(cell));
  const cols = resolveColumns_(headers, { invoice: FIELD_ALIASES.invoice });
  if (cols.invoice > -1) {
    return;
  }

  const targetCol = readCols + 1;
  if (sheet.getMaxColumns() < targetCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetCol - sheet.getMaxColumns());
  }
  sheet.getRange(1, targetCol).setValue(SHEET_HEADERS.PENJUALAN[1]);
  messages.push('Menambahkan kolom invoice pada sheet ' + APP_CONFIG.sheets.PENJUALAN);
}

/**
 * Pastikan sheet PENJUALAN memiliki kolom snapshot harga/modal per transaksi.
 * Kolom ditambahkan di ujung agar kompatibel dengan data lama.
 */
function ensurePenjualanHistoricalColumns_(ss, messages) {
  const sheet = findSheetByName_(ss, APP_CONFIG.sheets.PENJUALAN);
  if (!sheet) {
    return;
  }

  const lastCol = sheet.getLastColumn();
  const readCols = Math.max(lastCol, 1);
  const headers = sheet.getRange(1, 1, 1, readCols).getValues()[0].map((cell) => safeText_(cell));
  const cols = resolveColumns_(headers, {
    hargaModalTransaksi: FIELD_ALIASES.hargaModalTransaksi,
    hpp: FIELD_ALIASES.hpp,
    labaKotor: FIELD_ALIASES.labaKotor,
  });

  const missing = [];
  if (cols.hargaModalTransaksi < 0) {
    missing.push('Harga Modal Satuan (Rp)');
  }
  if (cols.hpp < 0) {
    missing.push('HPP (Rp)');
  }
  if (cols.labaKotor < 0) {
    missing.push('Laba Kotor (Rp)');
  }

  if (!missing.length) {
    return;
  }

  const targetLastCol = readCols + missing.length;
  if (sheet.getMaxColumns() < targetLastCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), targetLastCol - sheet.getMaxColumns());
  }

  missing.forEach((header, idx) => {
    sheet.getRange(1, readCols + idx + 1).setValue(header);
  });

  messages.push('Menambahkan kolom snapshot harga/modal pada sheet ' + APP_CONFIG.sheets.PENJUALAN);
}

/**
 * Lengkapi snapshot harga jual + modal pada baris PENJUALAN yang belum memiliki nilai.
 * Proses incremental memakai property row terakhir agar tidak membebani setiap request.
 */
function backfillPenjualanSnapshotRows_(ss, messages) {
  const sheet = findSheetByName_(ss, APP_CONFIG.sheets.PENJUALAN);
  if (!sheet) {
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return;
  }

  const props = PropertiesService.getDocumentProperties();
  let lastProcessedRow = Number(props.getProperty(APP_CONFIG.salesSnapshotBackfillRowKey) || '1');
  if (!Number.isFinite(lastProcessedRow) || lastProcessedRow < 1) {
    lastProcessedRow = 1;
  }
  if (lastProcessedRow > lastRow) {
    // Jika sheet sempat dibersihkan, proses ulang dari awal data.
    lastProcessedRow = 1;
  }

  const startRow = Math.max(2, lastProcessedRow + 1);
  if (startRow > lastRow) {
    props.setProperty(APP_CONFIG.salesSnapshotBackfillRowKey, String(lastRow));
    return;
  }

  const lastCol = sheet.getLastColumn();
  if (!lastCol) {
    return;
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map((cell) => safeText_(cell));
  const cols = resolveColumns_(headers, {
    sku: FIELD_ALIASES.sku,
    qty: FIELD_ALIASES.qty,
    hargaSatuan: FIELD_ALIASES.hargaSatuan,
    total: FIELD_ALIASES.total,
    hargaModalTransaksi: FIELD_ALIASES.hargaModalTransaksi,
    hpp: FIELD_ALIASES.hpp,
    labaKotor: FIELD_ALIASES.labaKotor,
  });

  const required = ['sku', 'qty', 'hargaSatuan', 'total', 'hargaModalTransaksi', 'hpp', 'labaKotor'];
  if (required.some((field) => !(cols[field] > -1))) {
    props.setProperty(APP_CONFIG.salesSnapshotBackfillRowKey, String(lastRow));
    return;
  }

  const productsContext = getSheetContext_(ss, APP_CONFIG.sheets.PRODUK, []);
  const productPricesBySku = {};

  if (productsContext.sheet) {
    const pCols = resolveColumns_(productsContext.headers, {
      sku: FIELD_ALIASES.sku,
      hargaModal: FIELD_ALIASES.hargaModal,
      hargaJual: FIELD_ALIASES.hargaJual,
    });

    productsContext.rows.forEach((item) => {
      const raw = item.values;
      const skuKey = normalizeSku_(raw[pCols.sku]);
      if (!skuKey) {
        return;
      }
      productPricesBySku[skuKey] = {
        hargaModal: round2_(toNumber_(raw[pCols.hargaModal])),
        hargaJual: round2_(toNumber_(raw[pCols.hargaJual])),
      };
    });
  }

  const hasValue = function (value) {
    return !(value === '' || value === null || value === undefined);
  };

  const rowCount = lastRow - startRow + 1;
  const values = sheet.getRange(startRow, 1, rowCount, lastCol).getValues();
  const updates = [];

  values.forEach((row, idx) => {
    if (!row.some((cell) => hasValue(cell))) {
      return;
    }

    const qty = round2_(toNumber_(row[cols.qty]));
    if (!qty || qty <= 0) {
      return;
    }

    const skuKey = normalizeSku_(row[cols.sku]);
    const product = productPricesBySku[skuKey] || { hargaModal: 0, hargaJual: 0 };
    let changed = false;

    let hargaSatuan = round2_(toNumber_(row[cols.hargaSatuan]));
    let total = round2_(toNumber_(row[cols.total]));
    const hasHargaSatuan = hasValue(row[cols.hargaSatuan]);
    const hasTotal = hasValue(row[cols.total]);

    if (!hasHargaSatuan && hasTotal && qty) {
      hargaSatuan = round2_(total / qty);
      row[cols.hargaSatuan] = hargaSatuan;
      changed = true;
    } else if (!hasHargaSatuan && !hasTotal) {
      hargaSatuan = round2_(product.hargaJual);
      row[cols.hargaSatuan] = hargaSatuan;
      changed = true;
    }

    if (!hasTotal) {
      total = round2_(hargaSatuan * qty);
      row[cols.total] = total;
      changed = true;
    }

    let hargaModalSatuan = round2_(toNumber_(row[cols.hargaModalTransaksi]));
    if (!hasValue(row[cols.hargaModalTransaksi])) {
      hargaModalSatuan = round2_(product.hargaModal);
      row[cols.hargaModalTransaksi] = hargaModalSatuan;
      changed = true;
    }

    let hpp = round2_(toNumber_(row[cols.hpp]));
    if (!hasValue(row[cols.hpp])) {
      hpp = round2_(qty * hargaModalSatuan);
      row[cols.hpp] = hpp;
      changed = true;
    }

    if (!hasValue(row[cols.labaKotor])) {
      row[cols.labaKotor] = round2_(total - hpp);
      changed = true;
    }

    if (changed) {
      updates.push({
        rowNumber: startRow + idx,
        values: row,
      });
    }
  });

  updates.forEach((entry) => {
    sheet.getRange(entry.rowNumber, 1, 1, lastCol).setValues([entry.values]);
  });

  props.setProperty(APP_CONFIG.salesSnapshotBackfillRowKey, String(lastRow));

  if (updates.length) {
    messages.push('Melengkapi snapshot transaksi PENJUALAN: ' + updates.length + ' baris.');
  }
}

/**
 * Status koneksi aplikasi terhadap sheet sumber.
 * Dipakai oleh Index (halaman utama Web App).
 * @return {Object}
 */
function getAppConnectionStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoSetupMessages = ensureCoreSheets_(ss);
  const warnings = [];
  const state = collectBusinessState_(ss, warnings);
  const alerts = buildProductAlerts_(state.products);
  const totalModalStokAkhir = round2_(
    state.products.reduce((sum, item) => {
      return sum + round2_(item.hargaModal * item.stockCalculated);
    }, 0)
  );

  const rekapProdukSheet = findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PRODUK);
  const rekapPelangganSheet = findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PELANGGAN);
  const mutasiSheet = findSheetByName_(ss, APP_CONFIG.sheets.MUTASI_STOK);

  return {
    meta: {
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      generatedAt: new Date().toISOString(),
      timezone: ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone(),
      warnings: autoSetupMessages.concat(warnings, alerts.lowStockWarnings, alerts.pricingWarnings),
    },
    sheets: {
      PRODUK: buildSheetStatusInfo_(state.productsContext.sheet),
      PENJUALAN: buildSheetStatusInfo_(state.salesContext.sheet),
      MUTASI_STOK: buildSheetStatusInfo_(mutasiSheet),
      REKAP_PRODUK: buildSheetStatusInfo_(rekapProdukSheet),
      REKAP_PELANGGAN: buildSheetStatusInfo_(rekapPelangganSheet),
    },
    stats: {
      totalProduk: state.products.length,
      totalTransaksi: countDistinctTransactions_(state.sales),
      totalPelanggan: Object.keys(state.customerAgg).length,
      totalAlertStokMinimum: alerts.lowStockItems.length,
      totalModalStokAkhir: totalModalStokAkhir,
    },
  };
}

/**
 * Bentuk status sederhana per sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet|null} sheet
 * @return {{exists:boolean,name:string,rowCount:number,lastRow:number,lastColumn:number}}
 */
function buildSheetStatusInfo_(sheet) {
  if (!sheet) {
    return {
      exists: false,
      name: '',
      rowCount: 0,
      lastRow: 0,
      lastColumn: 0,
    };
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  return {
    exists: true,
    name: sheet.getName(),
    rowCount: Math.max(0, lastRow - 1),
    lastRow: lastRow,
    lastColumn: lastColumn,
  };
}

/**
 * Dataset master produk dummy.
 */
function buildDummyProductCatalog_() {
  return [
    { sku: 'BRS-001', kategori: 'Beras', namaProduk: 'Beras Premium 5kg', hargaModal: 62000, satuan: 'sak', hargaJual: 75000, stokAwal: 120, stokMinimum: 30 },
    { sku: 'BRS-002', kategori: 'Beras', namaProduk: 'Beras Medium 5kg', hargaModal: 56000, satuan: 'sak', hargaJual: 69000, stokAwal: 110, stokMinimum: 28 },
    { sku: 'MYK-001', kategori: 'Minyak', namaProduk: 'Minyak Goreng 1L', hargaModal: 14500, satuan: 'pcs', hargaJual: 17500, stokAwal: 240, stokMinimum: 60 },
    { sku: 'GLA-001', kategori: 'Gula', namaProduk: 'Gula Pasir 1kg', hargaModal: 14500, satuan: 'pcs', hargaJual: 18000, stokAwal: 180, stokMinimum: 45 },
    { sku: 'TLP-001', kategori: 'Telur', namaProduk: 'Telur Ayam 1kg', hargaModal: 24000, satuan: 'kg', hargaJual: 29000, stokAwal: 130, stokMinimum: 35 },
    { sku: 'MIE-001', kategori: 'Mie', namaProduk: 'Mie Instan Goreng', hargaModal: 2600, satuan: 'pcs', hargaJual: 3500, stokAwal: 520, stokMinimum: 120 },
    { sku: 'TEH-001', kategori: 'Minuman', namaProduk: 'Teh Celup 25s', hargaModal: 7800, satuan: 'box', hargaJual: 10500, stokAwal: 95, stokMinimum: 20 },
    { sku: 'SUS-001', kategori: 'Susu', namaProduk: 'Susu Kental Manis', hargaModal: 9800, satuan: 'kaleng', hargaJual: 12500, stokAwal: 140, stokMinimum: 30 },
    { sku: 'KCP-001', kategori: 'Bumbu', namaProduk: 'Kecap Manis 600ml', hargaModal: 13500, satuan: 'botol', hargaJual: 17000, stokAwal: 85, stokMinimum: 22 },
    { sku: 'GRM-001', kategori: 'Bumbu', namaProduk: 'Garam 500gr', hargaModal: 3200, satuan: 'pcs', hargaJual: 5000, stokAwal: 210, stokMinimum: 55 },
    { sku: 'KOP-001', kategori: 'Minuman', namaProduk: 'Kopi Bubuk 200gr', hargaModal: 14200, satuan: 'pack', hargaJual: 18500, stokAwal: 120, stokMinimum: 30 },
    { sku: 'SBN-001', kategori: 'Sabun', namaProduk: 'Sabun Mandi Batang', hargaModal: 3400, satuan: 'pcs', hargaJual: 5000, stokAwal: 260, stokMinimum: 65 },
  ];
}

/**
 * Membangun transaksi dummy acak dan memastikan stok tidak minus.
 * @param {Array<Object>} products
 * @return {Array<Array<*>>}
 */
function buildDummySalesRows_(products) {
  const customers = [
    'Andi', 'Budi', 'Citra', 'Dewi', 'Eko', 'Fajar', 'Gita', 'Hendra',
    'Indah', 'Joko', 'Kiki', 'Lina', 'Maya', 'Nanda', 'Putri', 'Rudi',
    'Sari', 'Tono', 'Wawan', 'Yuni',
  ];
  const notes = ['', '', '', 'Langganan', 'COD', 'Antar sore', 'Repeat order'];

  const stockLeft = {};
  products.forEach((p) => {
    stockLeft[p.sku] = p.stokAwal;
  });

  const rows = [];
  const now = new Date();
  const startDate = new Date(now);
  startDate.setDate(startDate.getDate() - 40);
  let invoiceSeq = 0;

  const targetTransactions = 220;
  let guard = 0;

  while (rows.length < targetTransactions && guard < 4000) {
    guard += 1;
    const availableProducts = products.filter((p) => stockLeft[p.sku] > 0);
    if (!availableProducts.length) {
      break;
    }

    const product = availableProducts[Math.floor(Math.random() * availableProducts.length)];
    const maxQty = Math.min(4, stockLeft[product.sku]);
    const qty = Math.max(1, Math.floor(Math.random() * maxQty) + 1);
    stockLeft[product.sku] -= qty;

    const d = new Date(startDate);
    d.setDate(startDate.getDate() + Math.floor(Math.random() * 41));
    d.setHours(7 + Math.floor(Math.random() * 12), Math.floor(Math.random() * 60), 0, 0);

    const customer = customers[Math.floor(Math.random() * customers.length)];
    const total = round2_(qty * product.hargaJual);
    const hargaModalSatuan = round2_(product.hargaModal);
    const hpp = round2_(qty * hargaModalSatuan);
    const labaKotor = round2_(total - hpp);
    const note = notes[Math.floor(Math.random() * notes.length)];
    invoiceSeq += 1;
    const invoice = 'INV-DMY-' + String(invoiceSeq).padStart(4, '0');

    rows.push([
      d,
      invoice,
      customer,
      product.sku,
      product.namaProduk,
      product.satuan,
      product.hargaJual,
      qty,
      total,
      note,
      hargaModalSatuan,
      hpp,
      labaKotor,
    ]);
  }

  rows.sort((a, b) => a[0] - b[0]);
  return rows;
}

/**
 * Menimpa isi sheet dengan header + data baru.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {Array<string>} headers
 * @param {Array<Array<*>>} rows
 */
function overwriteSheetWithRows_(sheet, headers, rows) {
  if (!sheet) {
    return;
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  const minRows = Math.max(2, rows.length + 1);
  if (sheet.getMaxRows() < minRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
  }

  const clearCols = Math.max(sheet.getLastColumn(), headers.length);
  if (clearCols > 0 && sheet.getMaxRows() > 0) {
    sheet.getRange(1, 1, sheet.getMaxRows(), clearCols).clearContent();
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  sheet.setFrozenRows(1);
}

/**
 * Mengumpulkan state bisnis dari PRODUK dan PENJUALAN.
 */
function collectBusinessState_(ss, warnings) {
  const timezone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();

  const productsContext = getSheetContext_(ss, APP_CONFIG.sheets.PRODUK, warnings);
  const salesContext = getSheetContext_(ss, APP_CONFIG.sheets.PENJUALAN, warnings);
  const mutationContext = getSheetContext_(ss, APP_CONFIG.sheets.MUTASI_STOK, warnings);

  const productsInfo = parseProducts_(productsContext, warnings);
  const sales = parseSales_(salesContext, productsInfo.bySku, timezone, warnings);
  const mutations = parseStockMutations_(mutationContext, productsInfo.bySku, warnings);

  const soldBySku = {};
  const mutationBySku = {};
  const customerAgg = {};
  const customerTransactionSet = {};

  mutations.forEach((row) => {
    mutationBySku[row.skuKey] = round2_((mutationBySku[row.skuKey] || 0) + row.qtySigned);
  });

  sales.forEach((row) => {
    soldBySku[row.skuKey] = round2_((soldBySku[row.skuKey] || 0) + row.qty);

    if (!customerAgg[row.pelanggan]) {
      customerAgg[row.pelanggan] = {
        transaksi: 0,
        qty: 0,
        belanja: 0,
        hpp: 0,
        profit: 0,
      };
    }
    if (!customerTransactionSet[row.pelanggan]) {
      customerTransactionSet[row.pelanggan] = {};
    }

    const transaksiKey = getSalesTransactionKey_(row);
    if (!customerTransactionSet[row.pelanggan][transaksiKey]) {
      customerTransactionSet[row.pelanggan][transaksiKey] = true;
      customerAgg[row.pelanggan].transaksi += 1;
    }
    customerAgg[row.pelanggan].qty += row.qty;
    customerAgg[row.pelanggan].belanja += row.total;
    customerAgg[row.pelanggan].hpp += row.hpp;
    customerAgg[row.pelanggan].profit += row.profit;
  });

  productsInfo.rows.forEach((item) => {
    item.qtySold = round2_(soldBySku[item.skuKey] || 0);
    item.qtyMutation = round2_(mutationBySku[item.skuKey] || 0);
    item.stockCalculated = round2_(item.stokAwal + item.qtyMutation - item.qtySold);
  });

  return {
    timezone: timezone,
    productsContext: productsContext,
    salesContext: salesContext,
    mutationContext: mutationContext,
    productColumns: productsInfo.columns,
    products: productsInfo.rows,
    productsBySku: productsInfo.bySku,
    sales: sales,
    mutations: mutations,
    soldBySku: soldBySku,
    mutationBySku: mutationBySku,
    customerAgg: customerAgg,
  };
}
/**
 * Parsing data master produk.
 */
function parseProducts_(context, warnings) {
  const rows = [];
  const bySku = {};

  if (!context.sheet) {
    return {
      columns: {},
      rows: rows,
      bySku: bySku,
    };
  }

  const columns = resolveColumns_(context.headers, {
    sku: FIELD_ALIASES.sku,
    kategori: FIELD_ALIASES.kategori,
    namaProduk: FIELD_ALIASES.namaProduk,
    hargaModal: FIELD_ALIASES.hargaModal,
    satuan: FIELD_ALIASES.satuan,
    hargaJual: FIELD_ALIASES.hargaJual,
    stokAwal: FIELD_ALIASES.stokAwal,
    stokAkhir: FIELD_ALIASES.stokAkhir,
    modal: FIELD_ALIASES.modal,
    stokMinimum: FIELD_ALIASES.stokMinimum,
  });

  if (columns.sku < 0) {
    warnings.push('Kolom SKU pada sheet PRODUK tidak ditemukan.');
  }

  context.rows.forEach((item) => {
    const raw = item.values;
    const sku = safeText_(raw[columns.sku]);
    if (!sku) {
      return;
    }

    const skuKey = normalizeSku_(sku);
    const product = {
      rowNumber: item.rowNumber,
      sku: sku,
      skuKey: skuKey,
      kategori: safeText_(raw[columns.kategori]),
      namaProduk: safeText_(raw[columns.namaProduk]) || sku,
      hargaModal: round2_(toNumber_(raw[columns.hargaModal])),
      satuan: safeText_(raw[columns.satuan]) || '-',
      hargaJual: round2_(toNumber_(raw[columns.hargaJual])),
      stokAwal: round2_(toNumber_(raw[columns.stokAwal])),
      stokAkhirInput: round2_(toNumber_(raw[columns.stokAkhir])),
      stokMinimum: round2_(toNumber_(raw[columns.stokMinimum])),
      modal: round2_(toNumber_(raw[columns.modal])),
      qtySold: 0,
      qtyMutation: 0,
      stockCalculated: 0,
    };

    rows.push(product);
    bySku[skuKey] = product;
  });

  return {
    columns: columns,
    rows: rows,
    bySku: bySku,
  };
}

/**
 * Parsing transaksi penjualan.
 */
function parseSales_(context, productsBySku, timezone, warnings) {
  const result = [];

  if (!context.sheet) {
    return result;
  }

  const columns = resolveColumns_(context.headers, {
    tanggal: FIELD_ALIASES.tanggal,
    invoice: FIELD_ALIASES.invoice,
    pelanggan: FIELD_ALIASES.pelanggan,
    sku: FIELD_ALIASES.sku,
    namaProduk: FIELD_ALIASES.namaProduk,
    satuan: FIELD_ALIASES.satuan,
    hargaSatuan: FIELD_ALIASES.hargaSatuan,
    hargaModalTransaksi: FIELD_ALIASES.hargaModalTransaksi,
    qty: FIELD_ALIASES.qty,
    total: FIELD_ALIASES.total,
    hpp: FIELD_ALIASES.hpp,
    labaKotor: FIELD_ALIASES.labaKotor,
    catatan: FIELD_ALIASES.catatan,
  });

  if (columns.qty < 0) {
    warnings.push('Kolom Qty pada sheet PENJUALAN tidak ditemukan.');
  }
  if (columns.sku < 0) {
    warnings.push('Kolom SKU pada sheet PENJUALAN tidak ditemukan.');
  }

  const fallbackDateCol = columns.tanggal > -1
    ? columns.tanggal
    : guessDateColumn_(context.rows.map((item) => item.values));

  context.rows.forEach((item) => {
    const raw = item.values;
    const qty = round2_(toNumber_(raw[columns.qty]));
    if (!qty || qty <= 0) {
      return;
    }

    const skuRaw = safeText_(raw[columns.sku]);
    const skuKey = normalizeSku_(skuRaw);
    const product = productsBySku[skuKey] || null;

    const hasHargaSatuan = hasCellValue_(raw[columns.hargaSatuan]);
    let hargaSatuan = round2_(toNumber_(raw[columns.hargaSatuan]));
    if (!hasHargaSatuan && product) {
      hargaSatuan = round2_(product.hargaJual);
    }

    const hasTotal = hasCellValue_(raw[columns.total]);
    let total = round2_(toNumber_(raw[columns.total]));
    if (!hasTotal && hargaSatuan) {
      total = round2_(hargaSatuan * qty);
    }

    const hasHargaModalTransaksi = hasCellValue_(raw[columns.hargaModalTransaksi]);
    let hargaModalSatuan = round2_(toNumber_(raw[columns.hargaModalTransaksi]));
    if (!hasHargaModalTransaksi && product) {
      hargaModalSatuan = round2_(product.hargaModal);
    }

    const hasHpp = hasCellValue_(raw[columns.hpp]);
    let hpp = round2_(toNumber_(raw[columns.hpp]));
    if (!hasHpp) {
      hpp = round2_(qty * hargaModalSatuan);
    }

    const hasLabaKotor = hasCellValue_(raw[columns.labaKotor]);
    let profit = round2_(toNumber_(raw[columns.labaKotor]));
    if (!hasLabaKotor) {
      profit = round2_(total - hpp);
    }

    const tanggal = toDateOnly_(raw[fallbackDateCol]);
    const dateKey = tanggal ? Utilities.formatDate(tanggal, timezone, 'yyyy-MM-dd') : '';

    const invoice = safeText_(raw[columns.invoice]);
    const invoiceKey = normalizeInvoice_(invoice);
    const namaProduk = safeText_(raw[columns.namaProduk]) || (product ? product.namaProduk : skuRaw || 'Tanpa SKU');
    const pelanggan = safeText_(raw[columns.pelanggan]) || 'Tanpa Nama';

    result.push({
      rowNumber: item.rowNumber,
      date: tanggal,
      dateKey: dateKey,
      invoice: invoice,
      invoiceKey: invoiceKey,
      pelanggan: pelanggan,
      sku: skuRaw,
      skuKey: skuKey,
      namaProduk: namaProduk,
      qty: qty,
      hargaSatuan: hargaSatuan,
      hargaModalSatuan: hargaModalSatuan,
      total: total,
      hpp: hpp,
      profit: profit,
      catatan: safeText_(raw[columns.catatan]),
    });
  });

  return result;
}

/**
 * Parsing mutasi stok (stok masuk/retur/adjustment).
 */
function parseStockMutations_(context, productsBySku, warnings) {
  const result = [];

  if (!context.sheet) {
    return result;
  }

  const columns = resolveColumns_(context.headers, {
    tanggal: FIELD_ALIASES.tanggal,
    jenisMutasi: FIELD_ALIASES.jenisMutasi,
    sku: FIELD_ALIASES.sku,
    namaProduk: FIELD_ALIASES.namaProduk,
    qty: FIELD_ALIASES.qty,
    catatan: FIELD_ALIASES.catatan,
  });

  if (columns.sku < 0) {
    warnings.push('Kolom SKU pada sheet MUTASI_STOK tidak ditemukan.');
  }
  if (columns.qty < 0) {
    warnings.push('Kolom Qty pada sheet MUTASI_STOK tidak ditemukan.');
  }
  if (columns.jenisMutasi < 0) {
    warnings.push('Kolom Jenis Mutasi pada sheet MUTASI_STOK tidak ditemukan.');
  }

  context.rows.forEach((item) => {
    const raw = item.values;
    const skuRaw = safeText_(raw[columns.sku]);
    if (!skuRaw) {
      return;
    }

    const qtyRaw = round2_(toNumber_(raw[columns.qty]));
    if (!qtyRaw) {
      return;
    }

    const type = normalizeMutationType_(raw[columns.jenisMutasi]);
    let qtySigned = qtyRaw;
    if (type === 'MASUK' || type === 'RETUR') {
      qtySigned = round2_(Math.abs(qtyRaw));
    } else if (type === 'ADJUSTMENT') {
      qtySigned = round2_(qtyRaw);
    }

    const skuKey = normalizeSku_(skuRaw);
    const product = productsBySku[skuKey] || null;
    const namaProduk = safeText_(raw[columns.namaProduk]) || (product ? product.namaProduk : skuRaw);

    result.push({
      rowNumber: item.rowNumber,
      type: type,
      sku: skuRaw,
      skuKey: skuKey,
      namaProduk: namaProduk,
      qtySigned: qtySigned,
      catatan: safeText_(raw[columns.catatan]),
    });
  });

  return result;
}

/**
 * Update kolom Stok Akhir + Modal pada sheet PRODUK.
 */
function syncProductStockColumns_(state) {
  const ctx = state.productsContext;
  if (!ctx.sheet || !ctx.allRows.length) {
    return;
  }

  const cols = state.productColumns;
  if (cols.sku < 0 || cols.stokAwal < 0) {
    return;
  }

  const stockValues = [];
  const modalValues = [];

  ctx.allRows.forEach((rowItem) => {
    const row = rowItem.values;
    const sku = safeText_(row[cols.sku]);

    if (!sku) {
      stockValues.push([cols.stokAkhir > -1 ? row[cols.stokAkhir] : '']);
      modalValues.push([cols.modal > -1 ? row[cols.modal] : '']);
      return;
    }

    const skuKey = normalizeSku_(sku);
    const stokAwal = round2_(toNumber_(row[cols.stokAwal]));
    const hargaModal = round2_(toNumber_(row[cols.hargaModal]));
    const mutasiQty = round2_(state.mutationBySku[skuKey] || 0);
    const soldQty = round2_(state.soldBySku[skuKey] || 0);

    const stokAkhir = round2_(stokAwal + mutasiQty - soldQty);
    const modal = round2_(stokAwal * hargaModal);

    stockValues.push([stokAkhir]);
    modalValues.push([modal]);
  });

  if (cols.stokAkhir > -1 && stockValues.length) {
    ctx.sheet.getRange(2, cols.stokAkhir + 1, stockValues.length, 1).setValues(stockValues);
  }
  if (cols.modal > -1 && modalValues.length) {
    ctx.sheet.getRange(2, cols.modal + 1, modalValues.length, 1).setValues(modalValues);
  }
}

/**
 * Tulis tabel ringkasan ke sheet tujuan secara efisien.
 */
function writeTableSheet_(ss, sheetName, headers, rows) {
  const sheet = findSheetByName_(ss, sheetName) || ss.insertSheet(sheetName);

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  const requiredRows = rows.length + 1;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const clearCols = Math.max(headers.length, sheet.getLastColumn());
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, clearCols).clearContent();
  }

  if (rows.length) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.setFrozenRows(1);
}

/**
 * Ambil context sheet: header + rows + allRows.
 */
function getSheetContext_(ss, sheetName, warnings) {
  const sheet = findSheetByName_(ss, sheetName);
  if (!sheet) {
    warnings.push('Sheet ' + sheetName + ' tidak ditemukan.');
    return { sheet: null, headers: [], rows: [], allRows: [] };
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (!lastRow || !lastCol) {
    warnings.push('Sheet ' + sheetName + ' belum memiliki data.');
    return { sheet: sheet, headers: [], rows: [], allRows: [] };
  }

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = values[0].map((cell) => safeText_(cell));

  const allRows = values.slice(1).map((row, idx) => ({
    rowNumber: idx + 2,
    values: row,
  }));

  const rows = allRows.filter((item) => item.values.some((cell) => cell !== '' && cell !== null));

  return {
    sheet: sheet,
    headers: headers,
    rows: rows,
    allRows: allRows,
  };
}

/**
 * Membuat sheet jika belum ada, dan isi header bila baris header kosong.
 */
function ensureSheetExistsWithHeader_(ss, sheetName, headers, messages, forceHeader) {
  let sheet = findSheetByName_(ss, sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    messages.push('Membuat sheet: ' + sheetName);
  }

  if (sheet.getMaxColumns() < headers.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
  }

  if (sheet.getMaxRows() < 1) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 1);
  }

  const currentHeader = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const isEmptyHeader = currentHeader.every((cell) => safeText_(cell) === '');

  if (forceHeader || isEmptyHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    messages.push((forceHeader ? 'Set header ulang: ' : 'Set header: ') + sheetName);
  }
}

/**
 * Cari sheet by name (case-insensitive).
 */
function findSheetByName_(ss, name) {
  const target = String(name || '').trim().toLowerCase();
  return ss
    .getSheets()
    .find((sheet) => String(sheet.getName() || '').trim().toLowerCase() === target) || null;
}

/**
 * Mapping alias header -> index kolom.
 */
function resolveColumns_(headers, fieldMap) {
  const normalizedHeaders = headers.map((h) => normalizeHeader_(h));
  const resolved = {};

  Object.keys(fieldMap).forEach((key) => {
    resolved[key] = findColumnIndex_(normalizedHeaders, fieldMap[key]);
  });

  return resolved;
}
/**
 * Cari index kolom terbaik dari list alias.
 */
function findColumnIndex_(normalizedHeaders, aliases) {
  let bestIdx = -1;
  let bestScore = -1;

  (aliases || []).forEach((alias) => {
    const normalizedAlias = normalizeHeader_(alias);
    if (!normalizedAlias) {
      return;
    }

    normalizedHeaders.forEach((header, idx) => {
      if (!header) {
        return;
      }

      let score = -1;
      if (header === normalizedAlias) {
        score = 100;
      } else if (header.indexOf(normalizedAlias) === 0 || normalizedAlias.indexOf(header) === 0) {
        score = 75;
      } else if (header.indexOf(normalizedAlias) > -1 || normalizedAlias.indexOf(header) > -1) {
        score = 55;
      }

      if (score > bestScore) {
        bestScore = score;
        bestIdx = idx;
      }
    });
  });

  return bestIdx;
}

/**
 * Validasi index kolom wajib.
 */
function ensureRequiredColumns_(columnMap, requiredFields, sheetName) {
  const missing = requiredFields.filter((field) => !(columnMap[field] > -1));
  if (missing.length) {
    throw new Error(
      'Kolom wajib tidak ditemukan di sheet ' + sheetName + ': ' + missing.join(', ') +
      '. Jalankan menu "Kasir > Setup Struktur Sheet" lalu cek header.'
    );
  }
}

/**
 * Fallback index tanggal bila header tanggal tidak ditemukan.
 */
function guessDateColumn_(rows) {
  if (!rows.length) {
    return -1;
  }

  const sample = rows.slice(0, Math.min(rows.length, 30));
  const colCount = sample[0].length;

  let bestIdx = -1;
  let bestScore = 0;

  for (let col = 0; col < colCount; col += 1) {
    let validDates = 0;
    let nonEmpty = 0;

    for (let i = 0; i < sample.length; i += 1) {
      const value = sample[i][col];
      if (value === '' || value === null) {
        continue;
      }

      nonEmpty += 1;
      if (toDateOnly_(value)) {
        validDates += 1;
      }
    }

    const ratio = nonEmpty ? validDates / nonEmpty : 0;
    if (ratio > bestScore && ratio >= 0.6) {
      bestScore = ratio;
      bestIdx = col;
    }
  }

  return bestIdx;
}

/**
 * Normalize nama header agar mudah dicocokkan.
 */
function normalizeHeader_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[\-_\/()]+/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .trim();
}

/**
 * Normalisasi key SKU.
 */
function normalizeSku_(value) {
  return safeText_(value).toUpperCase();
}

/**
 * Normalisasi tipe mutasi stok.
 */
function normalizeMutationType_(value) {
  const text = safeText_(value).toUpperCase();
  if (!text) {
    return 'ADJUSTMENT';
  }
  if (text.indexOf('RETUR') > -1) {
    return 'RETUR';
  }
  if (text.indexOf('MASUK') > -1 || text === 'IN') {
    return 'MASUK';
  }
  if (text.indexOf('ADJ') > -1 || text.indexOf('SESUAI') > -1 || text.indexOf('PENYESUAIAN') > -1) {
    return 'ADJUSTMENT';
  }
  return text;
}

/**
 * Label ramah untuk tipe mutasi.
 */
function normalizeMutationLabel_(value) {
  const type = normalizeMutationType_(value);
  if (type === 'MASUK') {
    return 'Stok Masuk';
  }
  if (type === 'RETUR') {
    return 'Retur';
  }
  if (type === 'ADJUSTMENT') {
    return 'Adjustment';
  }
  return type;
}

/**
 * Hitung persentase margin kotor.
 */
function calcMarginPercent_(hargaJual, hargaModal) {
  const jual = round2_(toNumber_(hargaJual));
  const modal = round2_(toNumber_(hargaModal));
  if (!jual) {
    return 0;
  }
  return round2_(((jual - modal) / jual) * 100);
}

/**
 * Kumpulkan alert stok minimum dan warning harga/margin.
 */
function buildProductAlerts_(products) {
  const lowStockItems = (products || [])
    .filter((item) => item && item.stokMinimum > 0 && item.stockCalculated <= item.stokMinimum)
    .map((item) => ({
      sku: item.sku,
      namaProduk: item.namaProduk,
      kategori: item.kategori,
      stokTersedia: round2_(item.stockCalculated),
      stokMinimum: round2_(item.stokMinimum),
    }))
    .sort((a, b) => a.stokTersedia - b.stokTersedia);

  const pricingItems = (products || [])
    .map((item) => {
      const marginPercent = calcMarginPercent_(item.hargaJual, item.hargaModal);
      const belowCost = round2_(item.hargaJual) < round2_(item.hargaModal);
      const lowMargin = !belowCost && marginPercent < APP_CONFIG.minMarginPercent;
      if (!belowCost && !lowMargin) {
        return null;
      }

      return {
        sku: item.sku,
        namaProduk: item.namaProduk,
        hargaModal: round2_(item.hargaModal),
        hargaJual: round2_(item.hargaJual),
        marginPercent: marginPercent,
        type: belowCost ? 'below-cost' : 'low-margin',
      };
    })
    .filter((item) => !!item)
    .sort((a, b) => a.marginPercent - b.marginPercent);

  const lowStockWarnings = lowStockItems.slice(0, APP_CONFIG.maxAlertItems).map((item) => {
    return (
      'Stok minimum terlewati: ' + item.sku + ' - ' + item.namaProduk +
      ' (Stok ' + formatPlainNumber_(item.stokTersedia) +
      ' / Min ' + formatPlainNumber_(item.stokMinimum) + ')'
    );
  });

  const pricingWarnings = pricingItems.slice(0, APP_CONFIG.maxAlertItems).map((item) => {
    if (item.type === 'below-cost') {
      return (
        'Harga jual di bawah modal: ' + item.sku + ' - ' + item.namaProduk +
        ' (Modal ' + formatPlainNumber_(item.hargaModal) +
        ', Jual ' + formatPlainNumber_(item.hargaJual) + ')'
      );
    }
    return (
      'Margin rendah (< ' + formatPlainNumber_(APP_CONFIG.minMarginPercent) + '%): ' + item.sku + ' - ' + item.namaProduk +
      ' (Margin ' + formatPlainNumber_(item.marginPercent) + '%)'
    );
  });

  return {
    lowStockItems: lowStockItems,
    pricingItems: pricingItems,
    lowStockWarnings: lowStockWarnings,
    pricingWarnings: pricingWarnings,
  };
}

/**
 * Normalisasi key invoice.
 */
function normalizeInvoice_(value) {
  return safeText_(value).toUpperCase();
}

/**
 * Ambil key transaksi untuk agregasi.
 * Prioritas: invoice, fallback ke nomor baris agar data lama tetap valid.
 */
function getSalesTransactionKey_(saleRow) {
  const invoiceKey = normalizeInvoice_(saleRow && saleRow.invoiceKey ? saleRow.invoiceKey : saleRow && saleRow.invoice);
  if (invoiceKey) {
    return 'INV:' + invoiceKey;
  }

  const rowNumber = Number(saleRow && saleRow.rowNumber);
  if (rowNumber > 0) {
    return 'ROW:' + rowNumber;
  }

  return 'ROW:UNKNOWN';
}

/**
 * Hitung transaksi unik dari daftar baris penjualan.
 */
function countDistinctTransactions_(salesRows) {
  const seen = {};
  let total = 0;

  (salesRows || []).forEach((row) => {
    const key = getSalesTransactionKey_(row);
    if (!seen[key]) {
      seen[key] = true;
      total += 1;
    }
  });

  return total;
}

/**
 * Normalisasi payload transaksi agar menjadi list item.
 */
function normalizeTransactionItems_(payload) {
  const items = Array.isArray(payload && payload.items)
    ? payload.items
    : [{ sku: payload && payload.sku, qty: payload && payload.qty, note: payload && payload.note }];

  const normalized = [];
  items.forEach((item, idx) => {
    const sku = safeText_(item && item.sku);
    const qty = round2_(toNumber_(item && item.qty));
    const note = safeText_(item && item.note);

    if (!sku && !qty) {
      return;
    }
    if (!sku) {
      throw new Error('SKU item ke-' + (idx + 1) + ' wajib dipilih.');
    }
    if (!qty || qty <= 0) {
      throw new Error('Qty item ke-' + (idx + 1) + ' harus lebih dari 0.');
    }

    normalized.push({
      sku: sku,
      qty: qty,
      note: note,
    });
  });

  return normalized;
}

/**
 * Membuat nomor invoice unik per hari.
 */
function generateInvoiceNumber_(timezone) {
  const tz = timezone || Session.getScriptTimeZone();
  const props = PropertiesService.getDocumentProperties();
  const datePart = Utilities.formatDate(new Date(), tz, 'yyyyMMdd');
  const seqKey = 'invoice_seq_' + datePart;
  const nextSeq = Number(props.getProperty(seqKey) || '0') + 1;
  props.setProperty(seqKey, String(nextSeq));
  return 'INV-' + datePart + '-' + Utilities.formatString('%04d', nextSeq);
}

/**
 * Assign value ke row jika index valid.
 */
function assignRowValue_(row, idx, value) {
  if (idx > -1 && idx < row.length) {
    row[idx] = value;
  }
}

/**
 * Kompres baris top N (sisa jadi "Lainnya").
 */
function compressTopRows_(rows, limit, otherLabel) {
  if (!rows.length || rows.length <= limit) {
    return rows;
  }

  const topRows = rows.slice(0, limit);
  const other = rows.slice(limit).reduce(
    (acc, row) => {
      acc.belanja += row.belanja || 0;
      acc.profit += row.profit || 0;
      return acc;
    },
    { pelanggan: otherLabel, belanja: 0, profit: 0 }
  );

  if (other.profit) {
    topRows.push(other);
  }

  return topRows;
}

/**
 * Format angka ke string singkat untuk pesan.
 */
function formatPlainNumber_(value) {
  return Number(value || 0).toLocaleString('id-ID', { maximumFractionDigits: 2 });
}

/**
 * Konversi ke number aman (ID/EN format).
 */
function toNumber_(value) {
  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  if (value === null || value === undefined || value === '') {
    return 0;
  }

  let text = String(value)
    .replace(/\s+/g, '')
    .replace(/[^0-9,.-]/g, '');

  if (!text) {
    return 0;
  }

  const hasComma = text.indexOf(',') > -1;
  const hasDot = text.indexOf('.') > -1;

  if (hasComma && hasDot) {
    if (text.lastIndexOf(',') > text.lastIndexOf('.')) {
      text = text.replace(/\./g, '').replace(',', '.');
    } else {
      text = text.replace(/,/g, '');
    }
  } else if (hasComma && !hasDot) {
    const parts = text.split(',');
    if (parts.length === 2 && parts[1].length <= 2) {
      text = parts[0].replace(/\./g, '') + '.' + parts[1];
    } else {
      text = text.replace(/,/g, '');
    }
  } else {
    text = text.replace(/,/g, '');
  }

  const num = parseFloat(text);
  return Number.isFinite(num) ? num : 0;
}

/**
 * Konversi ke Date (tanpa jam).
 */
function toDateOnly_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  if (!value) {
    return null;
  }

  const text = String(value).trim();
  const parsed = new Date(text);
  if (!Number.isNaN(parsed.getTime())) {
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
  }

  const dmY = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (dmY) {
    let year = parseInt(dmY[3], 10);
    if (year < 100) {
      year += 2000;
    }
    const month = parseInt(dmY[2], 10) - 1;
    const day = parseInt(dmY[1], 10);
    const dt = new Date(year, month, day);
    if (!Number.isNaN(dt.getTime())) {
      return dt;
    }
  }

  return null;
}

/**
 * Format yyyy-MM-dd menjadi label tanggal ramah.
 */
function formatDateKey_(dateKey, timezone) {
  const dt = new Date(dateKey + 'T00:00:00');
  if (Number.isNaN(dt.getTime())) {
    return dateKey;
  }
  return Utilities.formatDate(dt, timezone, 'dd MMM yyyy');
}

/**
 * Ambil versi cache dashboard.
 */
function getDashboardVersion_() {
  const props = PropertiesService.getDocumentProperties();
  return Number(props.getProperty(APP_CONFIG.cacheVersionKey) || '0');
}

/**
 * Naikkan versi cache dashboard.
 */
function bumpDashboardVersion_() {
  const props = PropertiesService.getDocumentProperties();
  const currentVersion = Number(props.getProperty(APP_CONFIG.cacheVersionKey) || '0');
  props.setProperty(APP_CONFIG.cacheVersionKey, String(currentVersion + 1));
}

/**
 * Cek apakah cell memiliki nilai (termasuk angka 0).
 */
function hasCellValue_(value) {
  return !(value === '' || value === null || value === undefined);
}

/**
 * Ambil string aman.
 */
function safeText_(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

/**
 * Pembulatan 2 desimal.
 */
function round2_(num) {
  return Math.round((Number(num) + Number.EPSILON) * 100) / 100;
}

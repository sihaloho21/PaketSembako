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
  titleDashboard: 'Dashboard Penjualan',
  cachePrefix: 'dashboard:data:v',
  cacheVersionKey: 'dashboard_version',
  cacheTtlSeconds: 45,
  chartItemLimit: 20,
  pieItemLimit: 10,
  sheets: {
    PRODUK: 'PRODUK',
    PENJUALAN: 'PENJUALAN',
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
  ],
  PENJUALAN: [
    'Tanggal',
    'Nama Pelanggan',
    'SKU',
    'Nama Produk',
    'Satuan',
    'Harga Satuan (Rp)',
    'Qty',
    'Total (Rp)',
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
  modal: ['modal', 'nilai modal'],

  tanggal: ['tanggal', 'tgl', 'date', 'waktu', 'timestamp'],
  pelanggan: ['nama pelanggan', 'pelanggan', 'customer', 'pembeli', 'client'],
  hargaSatuan: ['harga satuan (rp)', 'harga satuan', 'harga', 'price'],
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
 * Endpoint Web App:
 * - default: index
 * - ?view=kasir: kasir
 * - ?view=dashboard: dashboard
 */
function doGet(e) {
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
  } else if (view === 'dashboard') {
    template = 'Dashboard';
  }

  let title = APP_CONFIG.titleIndex;
  if (template === 'Kasir') {
    title = APP_CONFIG.titleKasir;
  } else if (template === 'Dashboard') {
    title = APP_CONFIG.titleDashboard;
  }

  return HtmlService.createHtmlOutputFromFile(template)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
    'Data pada sheet PRODUK, PENJUALAN, REKAP_PRODUK, dan REKAP_PELANGGAN akan diganti data contoh.\n\nLanjutkan?',
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
    ]));

    const salesRows = buildDummySalesRows_(dummyProducts);

    overwriteSheetWithRows_(produkSheet, SHEET_HEADERS.PRODUK, productRows);
    overwriteSheetWithRows_(penjualanSheet, SHEET_HEADERS.PENJUALAN, salesRows);

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

  const products = state.products
    .map((item) => ({
      sku: item.sku,
      namaProduk: item.namaProduk,
      kategori: item.kategori,
      satuan: item.satuan,
      hargaJual: round2_(item.hargaJual),
      hargaModal: round2_(item.hargaModal),
      stokTersedia: round2_(item.stockCalculated),
    }))
    .sort((a, b) => a.namaProduk.localeCompare(b.namaProduk));

  const todayKey = Utilities.formatDate(new Date(), state.timezone, 'yyyy-MM-dd');
  const todaySales = state.sales.filter((row) => row.dateKey === todayKey);

  return {
    products: products,
    summary: {
      transaksiHariIni: todaySales.length,
      omzetHariIni: round2_(todaySales.reduce((sum, row) => sum + row.total, 0)),
      labaHariIni: round2_(todaySales.reduce((sum, row) => sum + row.profit, 0)),
    },
    meta: {
      generatedAt: new Date().toISOString(),
      warnings: autoSetupMessages.concat(warnings),
    },
  };
}

/**
 * Simpan transaksi dari form kasir.
 * @param {{customerName:string, sku:string, qty:number|string, note:string}} payload
 */
function saveKasirTransaction(payload) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(20000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureCoreSheets_(ss);

    const customerName = safeText_(payload && payload.customerName);
    const skuInput = safeText_(payload && payload.sku);
    const qty = toNumber_(payload && payload.qty);
    const note = safeText_(payload && payload.note);

    if (!customerName) {
      throw new Error('Nama pelanggan wajib diisi.');
    }
    if (!skuInput) {
      throw new Error('SKU wajib dipilih.');
    }
    if (!qty || qty <= 0) {
      throw new Error('Qty harus lebih dari 0.');
    }

    const warnings = [];
    const stateBefore = collectBusinessState_(ss, warnings);
    const skuKey = normalizeSku_(skuInput);
    const product = stateBefore.productsBySku[skuKey];

    if (!product) {
      throw new Error('SKU tidak ditemukan pada sheet PRODUK.');
    }

    const availableStock = round2_(product.stockCalculated);
    if (qty > availableStock) {
      throw new Error(
        'Stok tidak cukup. Stok tersedia untuk ' + product.namaProduk + ': ' + formatPlainNumber_(availableStock)
      );
    }

    const salesCtx = stateBefore.salesContext;
    if (!salesCtx.sheet) {
      throw new Error('Sheet PENJUALAN tidak ditemukan.');
    }

    const salesCols = resolveColumns_(salesCtx.headers, {
      tanggal: FIELD_ALIASES.tanggal,
      pelanggan: FIELD_ALIASES.pelanggan,
      sku: FIELD_ALIASES.sku,
      namaProduk: FIELD_ALIASES.namaProduk,
      satuan: FIELD_ALIASES.satuan,
      hargaSatuan: FIELD_ALIASES.hargaSatuan,
      qty: FIELD_ALIASES.qty,
      total: FIELD_ALIASES.total,
      catatan: FIELD_ALIASES.catatan,
    });

    ensureRequiredColumns_(salesCols, ['tanggal', 'pelanggan', 'sku', 'qty', 'total'], APP_CONFIG.sheets.PENJUALAN);

    const hargaSatuan = round2_(product.hargaJual);
    const total = round2_(hargaSatuan * qty);
    const rowLength = Math.max(salesCtx.headers.length, SHEET_HEADERS.PENJUALAN.length);
    const newRow = new Array(rowLength).fill('');

    assignRowValue_(newRow, salesCols.tanggal, new Date());
    assignRowValue_(newRow, salesCols.pelanggan, customerName);
    assignRowValue_(newRow, salesCols.sku, product.sku);
    assignRowValue_(newRow, salesCols.namaProduk, product.namaProduk);
    assignRowValue_(newRow, salesCols.satuan, product.satuan);
    assignRowValue_(newRow, salesCols.hargaSatuan, hargaSatuan);
    assignRowValue_(newRow, salesCols.qty, qty);
    assignRowValue_(newRow, salesCols.total, total);
    assignRowValue_(newRow, salesCols.catatan, note);

    salesCtx.sheet.appendRow(newRow);

    const syncResult = rebuildRecapAndStock_(ss);
    bumpDashboardVersion_();

    return {
      success: true,
      message: 'Transaksi berhasil disimpan.',
      transaksi: {
        tanggal: new Date().toISOString(),
        pelanggan: customerName,
        sku: product.sku,
        namaProduk: product.namaProduk,
        qty: qty,
        hargaSatuan: hargaSatuan,
        total: total,
        stokSisa: round2_(availableStock - qty),
      },
      recap: syncResult,
      warnings: warnings,
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
        'Transaksi: ' + result.totalTransaksi
    );

    return result;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Rebuild stok + rekap berdasarkan data PRODUK & PENJUALAN.
 */
function rebuildRecapAndStock_(ss) {
  const warnings = [];
  const state = collectBusinessState_(ss, warnings);

  syncProductStockColumns_(state);

  const rekapProdukRows = state.products
    .map((item) => {
      const qtyTerjual = round2_(state.soldBySku[item.skuKey] || 0);
      const omzet = round2_(qtyTerjual * item.hargaJual);
      const hpp = round2_(qtyTerjual * item.hargaModal);
      const labaKotor = round2_(omzet - hpp);
      return [
        item.sku,
        item.namaProduk,
        item.satuan,
        round2_(item.hargaModal),
        round2_(item.hargaJual),
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
    totalTransaksi: state.sales.length,
    warnings: warnings,
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
  const todayKey = Utilities.formatDate(new Date(), state.timezone, 'yyyy-MM-dd');

  const omzetPerHari = {};
  const qtyPerProduk = {};
  const pelangganMap = {};

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

    if (!pelangganMap[row.pelanggan]) {
      pelangganMap[row.pelanggan] = { pelanggan: row.pelanggan, belanja: 0, profit: 0, qty: 0, transaksi: 0 };
    }

    pelangganMap[row.pelanggan].belanja += row.total;
    pelangganMap[row.pelanggan].profit += row.profit;
    pelangganMap[row.pelanggan].qty += row.qty;
    pelangganMap[row.pelanggan].transaksi += 1;

    totalLabaKotor += row.profit;
  });

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
      warnings: warnings,
      sourceSheetStatus: {
        PRODUK: !!state.productsContext.sheet,
        PENJUALAN: !!state.salesContext.sheet,
        REKAP_PRODUK: !!findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PRODUK),
        REKAP_PELANGGAN: !!findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PELANGGAN),
      },
    },
    kpi: {
      omzetHariIni: round2_(omzetHariIni),
      totalLabaKotor: round2_(totalLabaKotor),
      totalTransaksi: state.sales.length,
      totalPelanggan: pelangganRows.length,
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
    },
    tables: {
      pelanggan: pelangganRows.slice(0, 100),
      produk: qtyProdukRows,
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
  ensureSheetExistsWithHeader_(ss, APP_CONFIG.sheets.PENJUALAN, SHEET_HEADERS.PENJUALAN, messages);
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
 * Status koneksi aplikasi terhadap sheet sumber.
 * Dipakai oleh Index (halaman utama Web App).
 * @return {Object}
 */
function getAppConnectionStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const autoSetupMessages = ensureCoreSheets_(ss);
  const warnings = [];
  const state = collectBusinessState_(ss, warnings);

  const rekapProdukSheet = findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PRODUK);
  const rekapPelangganSheet = findSheetByName_(ss, APP_CONFIG.sheets.REKAP_PELANGGAN);

  return {
    meta: {
      spreadsheetName: ss.getName(),
      spreadsheetId: ss.getId(),
      generatedAt: new Date().toISOString(),
      timezone: ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone(),
      warnings: autoSetupMessages.concat(warnings),
    },
    sheets: {
      PRODUK: buildSheetStatusInfo_(state.productsContext.sheet),
      PENJUALAN: buildSheetStatusInfo_(state.salesContext.sheet),
      REKAP_PRODUK: buildSheetStatusInfo_(rekapProdukSheet),
      REKAP_PELANGGAN: buildSheetStatusInfo_(rekapPelangganSheet),
    },
    stats: {
      totalProduk: state.products.length,
      totalTransaksi: state.sales.length,
      totalPelanggan: Object.keys(state.customerAgg).length,
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
    { sku: 'BRS-001', kategori: 'Beras', namaProduk: 'Beras Premium 5kg', hargaModal: 62000, satuan: 'sak', hargaJual: 75000, stokAwal: 120 },
    { sku: 'BRS-002', kategori: 'Beras', namaProduk: 'Beras Medium 5kg', hargaModal: 56000, satuan: 'sak', hargaJual: 69000, stokAwal: 110 },
    { sku: 'MYK-001', kategori: 'Minyak', namaProduk: 'Minyak Goreng 1L', hargaModal: 14500, satuan: 'pcs', hargaJual: 17500, stokAwal: 240 },
    { sku: 'GLA-001', kategori: 'Gula', namaProduk: 'Gula Pasir 1kg', hargaModal: 14500, satuan: 'pcs', hargaJual: 18000, stokAwal: 180 },
    { sku: 'TLP-001', kategori: 'Telur', namaProduk: 'Telur Ayam 1kg', hargaModal: 24000, satuan: 'kg', hargaJual: 29000, stokAwal: 130 },
    { sku: 'MIE-001', kategori: 'Mie', namaProduk: 'Mie Instan Goreng', hargaModal: 2600, satuan: 'pcs', hargaJual: 3500, stokAwal: 520 },
    { sku: 'TEH-001', kategori: 'Minuman', namaProduk: 'Teh Celup 25s', hargaModal: 7800, satuan: 'box', hargaJual: 10500, stokAwal: 95 },
    { sku: 'SUS-001', kategori: 'Susu', namaProduk: 'Susu Kental Manis', hargaModal: 9800, satuan: 'kaleng', hargaJual: 12500, stokAwal: 140 },
    { sku: 'KCP-001', kategori: 'Bumbu', namaProduk: 'Kecap Manis 600ml', hargaModal: 13500, satuan: 'botol', hargaJual: 17000, stokAwal: 85 },
    { sku: 'GRM-001', kategori: 'Bumbu', namaProduk: 'Garam 500gr', hargaModal: 3200, satuan: 'pcs', hargaJual: 5000, stokAwal: 210 },
    { sku: 'KOP-001', kategori: 'Minuman', namaProduk: 'Kopi Bubuk 200gr', hargaModal: 14200, satuan: 'pack', hargaJual: 18500, stokAwal: 120 },
    { sku: 'SBN-001', kategori: 'Sabun', namaProduk: 'Sabun Mandi Batang', hargaModal: 3400, satuan: 'pcs', hargaJual: 5000, stokAwal: 260 },
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
    const note = notes[Math.floor(Math.random() * notes.length)];

    rows.push([
      d,
      customer,
      product.sku,
      product.namaProduk,
      product.satuan,
      product.hargaJual,
      qty,
      total,
      note,
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

  const productsInfo = parseProducts_(productsContext, warnings);
  const sales = parseSales_(salesContext, productsInfo.bySku, timezone, warnings);

  const soldBySku = {};
  const customerAgg = {};

  sales.forEach((row) => {
    soldBySku[row.skuKey] = (soldBySku[row.skuKey] || 0) + row.qty;

    if (!customerAgg[row.pelanggan]) {
      customerAgg[row.pelanggan] = {
        transaksi: 0,
        qty: 0,
        belanja: 0,
        hpp: 0,
        profit: 0,
      };
    }

    customerAgg[row.pelanggan].transaksi += 1;
    customerAgg[row.pelanggan].qty += row.qty;
    customerAgg[row.pelanggan].belanja += row.total;
    customerAgg[row.pelanggan].hpp += row.hpp;
    customerAgg[row.pelanggan].profit += row.profit;
  });

  productsInfo.rows.forEach((item) => {
    item.qtySold = round2_(soldBySku[item.skuKey] || 0);
    item.stockCalculated = round2_(item.stokAwal - item.qtySold);
  });

  return {
    timezone: timezone,
    productsContext: productsContext,
    salesContext: salesContext,
    productColumns: productsInfo.columns,
    products: productsInfo.rows,
    productsBySku: productsInfo.bySku,
    sales: sales,
    soldBySku: soldBySku,
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
      modal: round2_(toNumber_(raw[columns.modal])),
      qtySold: 0,
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
    pelanggan: FIELD_ALIASES.pelanggan,
    sku: FIELD_ALIASES.sku,
    namaProduk: FIELD_ALIASES.namaProduk,
    satuan: FIELD_ALIASES.satuan,
    hargaSatuan: FIELD_ALIASES.hargaSatuan,
    qty: FIELD_ALIASES.qty,
    total: FIELD_ALIASES.total,
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

    let hargaSatuan = round2_(toNumber_(raw[columns.hargaSatuan]));
    if (!hargaSatuan && product) {
      hargaSatuan = round2_(product.hargaJual);
    }

    let total = round2_(toNumber_(raw[columns.total]));
    if (!total && hargaSatuan) {
      total = round2_(hargaSatuan * qty);
    }

    const tanggal = toDateOnly_(raw[fallbackDateCol]);
    const dateKey = tanggal ? Utilities.formatDate(tanggal, timezone, 'yyyy-MM-dd') : '';

    const namaProduk = safeText_(raw[columns.namaProduk]) || (product ? product.namaProduk : skuRaw || 'Tanpa SKU');
    const pelanggan = safeText_(raw[columns.pelanggan]) || 'Tanpa Nama';
    const hpp = round2_(qty * (product ? product.hargaModal : 0));
    const profit = round2_(total - hpp);

    result.push({
      rowNumber: item.rowNumber,
      date: tanggal,
      dateKey: dateKey,
      pelanggan: pelanggan,
      sku: skuRaw,
      skuKey: skuKey,
      namaProduk: namaProduk,
      qty: qty,
      hargaSatuan: hargaSatuan,
      total: total,
      hpp: hpp,
      profit: profit,
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
    const soldQty = round2_(state.soldBySku[skuKey] || 0);

    const stokAkhir = round2_(stokAwal - soldQty);
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

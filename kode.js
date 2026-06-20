/**
 * SISTEM INVENTARIS PREMIUM - BACKEND API CONTROLLER
 * File: code.gs
 * Deskripsi: Mengatur routing halaman web, otentikasi user, RPC calls,
 * serta mengolah data master, mutasi stok, laporan instalasi, kalkulasi BOM, & grafik analitik.
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("Almerass Premium Inventory")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Helper to get Sheet cleanly
function getSheetByNameSecure(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    throw new Error("Sheet '" + name + "' tidak ditemukan. Harap jalankan setupSpreadsheet() terlebih dahulu.");
  }
  return sheet;
}

/**
 * Otentikasi User (RBAC)
 */
function loginUser(username, password) {
  try {
    const sheet = getSheetByNameSecure("Users");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const dbUser = data[i][0].toString().trim();
      const dbPass = data[i][1].toString().trim();
      const dbRole = data[i][2].toString().toUpperCase().trim();
      
      if (dbUser.toLowerCase() === username.toLowerCase().trim() && dbPass === password) {
        return {
          success: true,
          user: {
            username: dbUser,
            role: dbRole
          }
        };
      }
    }
    return { success: false, message: "Username atau Password salah." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Registrasi User Baru (ADMIN Only)
 */
function registerUser(username, password, role) {
  try {
    const sheet = getSheetByNameSecure("Users");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === username.toLowerCase().trim()) {
        return { success: false, message: "Username sudah terdaftar!" };
      }
    }
    
    sheet.appendRow([username.trim(), password, role.toUpperCase()]);
    SpreadsheetApp.flush();
    return { success: true, message: `User ${username} berhasil didaftarkan.` };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Dashboard stats aggregators
 */
function getDashboardStats() {
  try {
    const masterSheet = getSheetByNameSecure("Master_Barang");
    const transSheet = getSheetByNameSecure("Riwayat_Transaksi");
    
    let totalItems = 0;
    let lowStockCount = 0;
    let todayTrxCount = 0;
    
    if (masterSheet) {
      const masterData = masterSheet.getDataRange().getValues();
      totalItems = masterData.length - 1;
      for (let i = 1; i < masterData.length; i++) {
        const minStok = Number(masterData[i][3]) || 0;
        const stokSekarang = Number(masterData[i][4]) || 0;
        if (stokSekarang <= minStok) {
          lowStockCount++;
        }
      }
    }
    
    if (transSheet) {
      const transData = transSheet.getDataRange().getValues();
      for (let i = 1; i < transData.length; i++) {
        const tipe = transData[i][4].toString().toUpperCase().trim();
        const qty = Number(transData[i][6]) || 0;
        if (tipe === "KELUAR") {
          todayTrxCount += qty; // Volume Out
        }
      }
    }
    
    return {
      success: true,
      totalItems: totalItems,
      lowStockCount: lowStockCount,
      todayTrxCount: todayTrxCount
    };
  } catch (err) {
    return { success: false, totalItems: 0, lowStockCount: 0, todayTrxCount: 0, message: err.toString() };
  }
}

/**
 * Server-Side Pagination & Filtering: MASTER BARANG
 */
function getInventoryData(page, pageSize, searchVal, categoryVal) {
  try {
    const sheet = getSheetByNameSecure("Master_Barang");
    const values = sheet.getDataRange().getValues();
    let categoriesSet = {};
    let filtered = [];
    
    searchVal = searchVal ? searchVal.toLowerCase().trim() : "";
    categoryVal = categoryVal ? categoryVal.toLowerCase().trim() : "";
    
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const kode = row[0].toString();
      const nama = row[1].toString();
      const kategori = row[2].toString();
      const minStok = Number(row[3]) || 0;
      const stokNow = Number(row[4]) || 0;
      const satuan = row[5].toString();
      const hargaBeli = Number(row[6]) || 0;
      const hargaJual = Number(row[7]) || 0;
      
      categoriesSet[kategori] = true;
      
      const matchSearch = !searchVal || kode.toLowerCase().includes(searchVal) || nama.toLowerCase().includes(searchVal);
      const matchCat = !categoryVal || kategori.toLowerCase().trim() === categoryVal;
      
      if (matchSearch && matchCat) {
        filtered.push({
          "Kode Barang": kode,
          "Nama Barang": nama,
          "Kategori": kategori,
          "Stok Minimal": minStok,
          "Stok Sekarang": stokNow,
          "Satuan": satuan,
          "Harga Beli": hargaBeli,
          "Harga Jual": hargaJual
        });
      }
    }
    
    const totalCount = filtered.length;
    const startIndex = (page - 1) * pageSize;
    const pageData = filtered.slice(startIndex, startIndex + pageSize);
    const categoriesList = Object.keys(categoriesSet).sort();
    
    return {
      success: true,
      data: pageData,
      totalCount: totalCount,
      categories: categoriesList
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * CRUD: Menambahkan Item Baru (ADMIN Only)
 */
function generateNextItemCode() {
  const sheet = getSheetByNameSecure("Master_Barang");
  const data = sheet.getDataRange().getValues();
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const code = data[i][0].toString().trim();
    if (code.startsWith("BRG-")) {
      const num = parseInt(code.substring(4), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  return "BRG-" + (maxNum + 1).toString().padStart(4, '0');
}

function addBarang(payload) {
  try {
    const sheet = getSheetByNameSecure("Master_Barang");
    const code = generateNextItemCode();
    
    sheet.appendRow([
      code,
      payload.nama,
      payload.kategori,
      payload.stokMinim,
      payload.stokAwal,
      payload.satuan,
      payload.hargaBeli,
      payload.hargaJual
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: `Barang ${code} berhasil didaftarkan.` };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * CRUD: Mengedit Master Barang (ADMIN Only)
 */
function updateBarang(payload) {
  try {
    const sheet = getSheetByNameSecure("Master_Barang");
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === payload.kode.trim()) {
        rowIdx = i + 1;
        break;
      }
    }
    
    if (rowIdx === -1) return { success: false, message: "Barang tidak ditemukan." };
    
    sheet.getRange(rowIdx, 2).setValue(payload.nama);
    sheet.getRange(rowIdx, 3).setValue(payload.kategori);
    sheet.getRange(rowIdx, 4).setValue(payload.stokMinim);
    sheet.getRange(rowIdx, 5).setValue(payload.stokAwal); 
    sheet.getRange(rowIdx, 6).setValue(payload.satuan);
    sheet.getRange(rowIdx, 7).setValue(payload.hargaBeli);
    sheet.getRange(rowIdx, 8).setValue(payload.hargaJual);
    
    SpreadsheetApp.flush();
    return { success: true, message: "Barang berhasil diperbarui." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * CRUD: Menghapus Master Barang (ADMIN Only)
 */
function deleteBarang(kode) {
  try {
    const sheet = getSheetByNameSecure("Master_Barang");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === kode.trim()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: `Barang ${kode} berhasil dihapus.` };
      }
    }
    return { success: false, message: "Barang tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Server-Side Pagination: RIWAYAT TRANSAKSI
 */
function getTransaksiHistory(page, pageSize) {
  try {
    const sheet = getSheetByNameSecure("Riwayat_Transaksi");
    const values = sheet.getDataRange().getValues();
    let list = [];
    
    for (let i = values.length - 1; i >= 1; i--) {
      const row = values[i];
      let tglStr = "";
      if (row[1] instanceof Date) {
        tglStr = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      } else {
        tglStr = row[1].toString();
      }
      
      list.push({
        "ID Transaksi": row[0].toString(),
        "Tanggal": tglStr,
        "Kode Barang": row[2].toString(),
        "Nama Barang": row[3].toString(),
        "Tipe": row[4].toString().toUpperCase(),
        "User": row[5].toString(),
        "Qty": Number(row[6]) || 0,
        "Catatan": row[7].toString()
      });
    }
    
    const totalCount = list.length;
    const startIndex = (page - 1) * pageSize;
    const pageData = list.slice(startIndex, startIndex + pageSize);
    
    return {
      success: true,
      data: pageData,
      totalCount: totalCount
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Safe Stock Mutation Engine (With ScriptLock)
 */
function recordTransaksi(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 
    
    const masterSheet = getSheetByNameSecure("Master_Barang");
    const transSheet = getSheetByNameSecure("Riwayat_Transaksi");
    
    const masterData = masterSheet.getDataRange().getValues();
    let rowIdx = -1;
    let namaBarang = "";
    let stokSekarang = 0;
    
    for (let i = 1; i < masterData.length; i++) {
      if (masterData[i][0].toString().trim() === payload.kodeBarang.trim()) {
        rowIdx = i + 1;
        namaBarang = masterData[i][1].toString();
        stokSekarang = Number(masterData[i][4]) || 0;
        break;
      }
    }
    
    if (rowIdx === -1) {
      return { success: false, message: "Kode barang tidak ditemukan." };
    }
    
    const qty = Number(payload.jumlah) || 0;
    let stokBaru = stokSekarang;
    
    if (payload.tipe === "MASUK") {
      stokBaru += qty;
    } else if (payload.tipe === "KELUAR") {
      if (stokSekarang < qty) {
        return { success: false, message: "Stok tidak mencukupi untuk pengeluaran." };
      }
      stokBaru -= qty;
    }
    
    // Update Master_Barang
    masterSheet.getRange(rowIdx, 5).setValue(stokBaru);
    
    // Generate TRX ID
    const trxId = generateTransactionId();
    const now = new Date();
    
    // Append to Riwayat_Transaksi
    transSheet.appendRow([
      trxId,
      now,
      payload.kodeBarang,
      namaBarang,
      payload.tipe.toUpperCase(),
      payload.user,
      qty,
      payload.catatan
    ]);
    
    SpreadsheetApp.flush();
    return { success: true, trxId: trxId, message: `Transaksi ${trxId} berhasil direkam.` };
    
  } catch (err) {
    return { success: false, message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Server-Side Analytical Trends (Revenue, Profit, and Volume)
 */
function getSalesAnalytics(timeframe) {
  try {
    const transSheet = getSheetByNameSecure("Riwayat_Transaksi");
    const masterSheet = getSheetByNameSecure("Master_Barang");
    
    const masterData = masterSheet.getDataRange().getValues();
    const masterMap = {};
    for (let i = 1; i < masterData.length; i++) {
      const code = masterData[i][0].toString().trim();
      masterMap[code] = {
        nama: masterData[i][1].toString(),
        beli: Number(masterData[i][6]) || 0,
        jual: Number(masterData[i][7]) || 0
      };
    }
    
    const transData = transSheet.getDataRange().getValues();
    let totalQtySold = 0;
    let totalRevenue = 0;
    let totalProfit = 0;
    
    let salesMap = {};
    let timeGroups = {};
    
    let sortedTrans = [];
    for (let i = 1; i < transData.length; i++) {
      const row = transData[i];
      if (row[4].toString().toUpperCase().trim() === "KELUAR") {
        let tglDate;
        if (row[1] instanceof Date) {
          tglDate = row[1];
        } else {
          tglDate = new Date(row[1]);
        }
        sortedTrans.push({
          row: row,
          date: tglDate
        });
      }
    }
    sortedTrans.sort((a, b) => a.date - b.date);
    
    sortedTrans.forEach(item => {
      const row = item.row;
      const code = row[2].toString().trim();
      const nama = row[3].toString();
      const qty = Number(row[6]) || 0;
      
      const mItem = masterMap[code] || { nama: nama, beli: 0, jual: 0 };
      const rev = qty * mItem.jual;
      const prof = qty * (mItem.jual - mItem.beli);
      
      totalQtySold += qty;
      totalRevenue += rev;
      totalProfit += prof;
      
      if (!salesMap[code]) {
        salesMap[code] = {
          kode: code,
          nama: mItem.nama,
          qty: 0,
          revenue: 0,
          profit: 0
        };
      }
      salesMap[code].qty += qty;
      salesMap[code].revenue += rev;
      salesMap[code].profit += prof;
      
      const dObj = item.date;
      const yy = dObj.getFullYear();
      const mm = (dObj.getMonth() + 1).toString().padStart(2, '0');
      const dd = dObj.getDate().toString().padStart(2, '0');
      const cleanDateStr = `${yy}-${mm}-${dd}`;
      let key = cleanDateStr;
      
      if (timeframe === 'weekly') {
        const day = dObj.getDay();
        const diff = dObj.getDate() - day + (day === 0 ? -6 : 1); 
        const startOfWeek = new Date(dObj.setDate(diff));
        const yr = startOfWeek.getFullYear();
        const mo = (startOfWeek.getMonth() + 1).toString().padStart(2, '0');
        const dy = startOfWeek.getDate().toString().padStart(2, '0');
        key = `W/C ${yr}-${mo}-${dy}`;
      } else if (timeframe === 'monthly') {
        key = `${yy}-${mm}`;
      }
      
      if (!timeGroups[key]) {
        timeGroups[key] = { revenue: 0, profit: 0, qty: 0 };
      }
      timeGroups[key].revenue += rev;
      timeGroups[key].profit += prof;
      timeGroups[key].qty += qty;
    });
    
    const performanceData = Object.values(salesMap).sort((a, b) => b.qty - a.qty);
    const chartLabels = Object.keys(timeGroups).sort();
    const chartRevenue = [];
    const chartProfit = [];
    const chartQty = [];
    
    chartLabels.forEach(k => {
      chartRevenue.push(timeGroups[k].revenue);
      chartProfit.push(timeGroups[k].profit);
      chartQty.push(timeGroups[k].qty);
    });
    
    return {
      success: true,
      totalQtySold: totalQtySold,
      totalRevenue: totalRevenue,
      totalProfit: totalProfit,
      performanceData: performanceData,
      chartLabels: chartLabels,
      chartRevenue: chartRevenue,
      chartProfit: chartProfit,
      chartQty: chartQty
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Server-Side Pagination: LAPORAN INSTALASI
 */
function getLaporanInstalasi(page, pageSize, searchInsVal) {
  try {
    const sheet = getSheetByNameSecure("Laporan_Instalasi");
    const values = sheet.getDataRange().getValues();
    let list = [];
    searchInsVal = searchInsVal ? searchInsVal.toLowerCase().trim() : "";
    
    for (let i = values.length - 1; i >= 1; i--) {
      const row = values[i];
      const idIns = row[0].toString();
      const dateVal = row[1];
      let tglStr = "";
      if (dateVal instanceof Date) {
        tglStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      } else {
        tglStr = dateVal.toString();
      }
      const noPo = row[2].toString();
      const customer = row[3].toString();
      const teknisi = row[4].toString();
      const status = row[5].toString();
      const snUnit = row[6].toString();
      const itemDipasang = row[7].toString();
      const catatan = row[8].toString();
      
      const match = !searchInsVal || 
                    noPo.toLowerCase().includes(searchInsVal) || 
                    customer.toLowerCase().includes(searchInsVal) || 
                    teknisi.toLowerCase().includes(searchInsVal) ||
                    idIns.toLowerCase().includes(searchInsVal);
      
      if (match) {
        list.push({
          "ID Ins": idIns,
          "Tanggal": tglStr,
          "No. PO": noPo,
          "Customer": customer,
          "Teknisi": teknisi,
          "Status": status,
          "SN Unit": snUnit,
          "Item Dipasang": itemDipasang,
          "Catatan": catatan
        });
      }
    }
    
    const totalCount = list.length;
    const startIndex = (page - 1) * pageSize;
    const pageData = list.slice(startIndex, startIndex + pageSize);
    
    return {
      success: true,
      data: pageData,
      totalCount: totalCount
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveLaporanInstalasi(payload) {
  try {
    const sheet = getSheetByNameSecure("Laporan_Instalasi");
    const idIns = generateInstalasiId();
    const tgl = new Date();
    
    sheet.appendRow([
      idIns,
      tgl,
      payload.noPo,
      payload.customer,
      payload.teknisi,
      payload.status,
      payload.snUnit,
      payload.itemDipasang,
      payload.catatan
    ]);
    SpreadsheetApp.flush();
    return { success: true, idIns: idIns, message: "Laporan instalasi berhasil disimpan." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function updateLaporanInstalasi(payload) {
  try {
    const sheet = getSheetByNameSecure("Laporan_Instalasi");
    const data = sheet.getDataRange().getValues();
    let rowIdx = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === payload.idIns.trim()) {
        rowIdx = i + 1;
        break;
      }
    }
    
    if (rowIdx === -1) return { success: false, message: "Data tidak ditemukan." };
    
    sheet.getRange(rowIdx, 3).setValue(payload.noPo);
    sheet.getRange(rowIdx, 4).setValue(payload.customer);
    sheet.getRange(rowIdx, 5).setValue(payload.teknisi);
    sheet.getRange(rowIdx, 6).setValue(payload.status);
    sheet.getRange(rowIdx, 7).setValue(payload.snUnit);
    sheet.getRange(rowIdx, 8).setValue(payload.itemDipasang);
    sheet.getRange(rowIdx, 9).setValue(payload.catatan);
    
    SpreadsheetApp.flush();
    return { success: true, message: "Laporan instalasi berhasil diperbarui." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function deleteLaporanInstalasi(idIns) {
  try {
    const sheet = getSheetByNameSecure("Laporan_Instalasi");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === idIns.trim()) {
        sheet.deleteRow(i + 1);
        SpreadsheetApp.flush();
        return { success: true, message: `Laporan ${idIns} berhasil dihapus.` };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Server-Side: REQUEST ORDER LOGS
 */
function getRequestOrderData(page, limit) {
  try {
    const sheet = getSheetByNameSecure("Request_Order");
    const values = sheet.getDataRange().getValues();
    let list = [];
    
    for (let i = values.length - 1; i >= 1; i--) {
      const row = values[i];
      let tglStr = "";
      if (row[1] instanceof Date) {
        tglStr = Utilities.formatDate(row[1], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm");
      } else {
        tglStr = row[1].toString();
      }
      
      list.push({
        "ID Request": row[0].toString(),
        "Tanggal": tglStr,
        "Target Produk": row[2].toString(),
        "Qty": Number(row[3]) || 0,
        "Pemohon": row[4].toString(),
        "No. Referensi": row[5].toString(),
        "Status": row[6].toString()
      });
    }
    
    const totalCount = list.length;
    const startIndex = (page - 1) * limit;
    const pageData = list.slice(startIndex, startIndex + limit);
    
    return {
      success: true,
      data: pageData,
      totalCount: totalCount
    };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

function saveRequestOrder(payload) {
  try {
    const sheet = getSheetByNameSecure("Request_Order");
    const idReq = generateRequestId();
    const tgl = new Date();
    
    sheet.appendRow([
      idReq,
      tgl,
      payload.targetProduk,
      payload.qty,
      payload.pemohon,
      payload.noReferensi,
      "PROSES"
    ]);
    SpreadsheetApp.flush();
    return { success: true, message: `Request order ${idReq} berhasil disimpan.` };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Real-Time Material cost & availability calculation based on BOM setup
 */
function calculateBOMServer(targetProduk, qty) {
  try {
    const bomSheet = getSheetByNameSecure("BOM_Settings");
    const masterSheet = getSheetByNameSecure("Master_Barang");
    
    const bomValues = bomSheet.getDataRange().getValues();
    const masterValues = masterSheet.getDataRange().getValues();
    
    const masterMap = {};
    for (let i = 1; i < masterValues.length; i++) {
      const code = masterValues[i][0].toString().trim();
      const name = masterValues[i][1].toString().trim();
      const cat = masterValues[i][2].toString();
      const stock = Number(masterValues[i][4]) || 0;
      const unit = masterValues[i][5].toString();
      const price = Number(masterValues[i][6]) || 0;
      
      masterMap[name.toLowerCase()] = {
        kode: code,
        deskripsi: cat,
        stok: stock,
        satuan: unit,
        harga: price
      };
    }
    
    let resData = [];
    for (let i = 1; i < bomValues.length; i++) {
      const prodJadi = bomValues[i][0].toString().trim();
      const komponen = bomValues[i][1].toString().trim();
      const rasio = Number(bomValues[i][2]) || 0;
      const unit = bomValues[i][3].toString();
      
      if (prodJadi.toLowerCase() === targetProduk.toLowerCase()) {
        const needed = rasio * qty;
        const whItem = masterMap[komponen.toLowerCase()];
        
        const code = whItem ? whItem.kode : "N/A";
        const spec = whItem ? whItem.deskripsi : "Bahan Baku";
        const stock = whItem ? whItem.stok : 0;
        const actualUnit = unit || (whItem ? whItem.satuan : "Pcs");
        const price = whItem ? whItem.harga : 0;
        
        resData.push({
          komponen: komponen,
          kode: code,
          deskripsi: spec,
          rasio: rasio,
          totalKebutuhan: needed,
          stokGudang: stock,
          satuan: actualUnit,
          hargaSatuan: price,
          totalBiaya: needed * price,
          tersedia: stock >= needed
        });
      }
    }
    
    return { success: true, data: resData };
  } catch (err) {
    return { success: false, message: err.toString() };
  }
}

/**
 * Sync external BOM Source Simulation
 */
function syncBOMFromExternalSource() {
  return { success: true, message: "Sinkronisasi berhasil! 6 Komponen BOM ter-update dari server pusat." };
}
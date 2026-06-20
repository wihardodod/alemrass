/**
 * SETUP SPREADSHEET FOR ALMERASS INVENTORY
 * File: setup.gs
 * Deskripsi: Menyiapkan semua sheet yang diperlukan dengan kolom dan data awal.
 */

function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let masterSheet = ss.getSheetByName("Master_Barang");
  if (!masterSheet) {
    masterSheet = ss.insertSheet("Master_Barang");
  }
  masterSheet.clear();
  const masterHeaders = [
    "Kode Barang", 
    "Nama Barang", 
    "Kategori", 
    "Stok Minimal", 
    "Stok Sekarang", 
    "Satuan", 
    "Harga Beli", 
    "Harga Jual"
  ];
  masterSheet.getRange(1, 1, 1, masterHeaders.length).setValues([masterHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");
  
  // Dummy Data Master Barang (3 Baris Utama + Tambahan)
  const dummyBarang = [
    ["BRG-0001", "Safety Dump Alert", "Safety Device", 5, 3, "Pcs", 1000000, 4500000],
    ["BRG-0002", "Safety Dump Alert + Lcd", "Safety Device", 5, 10, "Pcs", 2000000, 5500000],
    ["BRG-0003", "Fatigue Warning System", "Safety Device", 5, 7, "Pcs", 2000000, 5500000],
    ["BRG-0004", "Kabel AWG 16", "Suku Cadang", 100, 150, "Meter", 15000, 25000],
    ["BRG-0005", "Box Panel Aluminium", "Suku Cadang", 10, 8, "Pcs", 250000, 350000],
    ["BRG-0006", "Sensor Proximity", "Suku Cadang", 20, 12, "Pcs", 80000, 120000],
    ["BRG-0007", "LCD Display 7 Inch", "Suku Cadang", 5, 3, "Pcs", 450000, 600000]
  ];
  masterSheet.getRange(2, 1, dummyBarang.length, dummyBarang[0].length).setValues(dummyBarang);

  let transSheet = ss.getSheetByName("Riwayat_Transaksi");
  if (!transSheet) {
    transSheet = ss.insertSheet("Riwayat_Transaksi");
  }
  transSheet.clear();
  const transHeaders = [
    "ID Transaksi", 
    "Tanggal", 
    "Kode Barang", 
    "Nama Barang", 
    "Tipe", 
    "User", 
    "Qty", 
    "Catatan"
  ];
  transSheet.getRange(1, 1, 1, transHeaders.length).setValues([transHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");
  
  const dummyTrans = [
    ["TRX-260615-0001", new Date(), "BRG-0001", "Safety Dump Alert", "MASUK", "admin", 3, "Stok awal setup"]
  ];
  transSheet.getRange(2, 1, dummyTrans.length, dummyTrans[0].length).setValues(dummyTrans);

  let insSheet = ss.getSheetByName("Laporan_Instalasi");
  if (!insSheet) {
    insSheet = ss.insertSheet("Laporan_Instalasi");
  }
  insSheet.clear();
  const insHeaders = [
    "ID Ins", 
    "Tanggal", 
    "No. PO", 
    "Customer", 
    "Teknisi", 
    "Status", 
    "SN Unit", 
    "Item Dipasang", 
    "Catatan"
  ];
  insSheet.getRange(1, 1, 1, insHeaders.length).setValues([insHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");
  
  const dummyIns = [
    ["INS-260615-0001", new Date(), "PO-99281", "PT SIS", "Ahmad", "SELESAI", "SN-SDA-9988", "Safety Dump Alert", "Pemasangan di HD785"]
  ];
  insSheet.getRange(2, 1, dummyIns.length, dummyIns[0].length).setValues(dummyIns);

  let reqSheet = ss.getSheetByName("Request_Order");
  if (!reqSheet) {
    reqSheet = ss.insertSheet("Request_Order");
  }
  reqSheet.clear();
  const reqHeaders = [
    "ID Request", 
    "Tanggal", 
    "Target Produk", 
    "Qty", 
    "Pemohon", 
    "No. Referensi", 
    "Status"
  ];
  reqSheet.getRange(1, 1, 1, reqHeaders.length).setValues([reqHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");
  
  const dummyReq = [
    ["REQ-260615-0001", new Date(), "Safety Dump Alert", 2, "warehouse", "REF-001", "PROSES"]
  ];
  reqSheet.getRange(2, 1, dummyReq.length, dummyReq[0].length).setValues(dummyReq);

  let bomSheet = ss.getSheetByName("BOM_Settings");
  if (!bomSheet) {
    bomSheet = ss.insertSheet("BOM_Settings");
  }
  bomSheet.clear();
  const bomHeaders = ["Produk Jadi", "Komponen", "Rasio", "Satuan"];
  bomSheet.getRange(1, 1, 1, bomHeaders.length).setValues([bomHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");
  
  const dummyBOM = [
    ["Safety Dump Alert", "Kabel AWG 16", 5, "Meter"],
    ["Safety Dump Alert", "Box Panel Aluminium", 1, "Pcs"],
    ["Safety Dump Alert", "Sensor Proximity", 2, "Pcs"],
    ["Safety Dump Alert + Lcd", "Kabel AWG 16", 5, "Meter"],
    ["Safety Dump Alert + Lcd", "Box Panel Aluminium", 1, "Pcs"],
    ["Safety Dump Alert + Lcd", "LCD Display 7 Inch", 1, "Pcs"]
  ];
  bomSheet.getRange(2, 1, dummyBOM.length, dummyBOM[0].length).setValues(dummyBOM);

  let usersSheet = ss.getSheetByName("Users");
  if (!usersSheet) {
    usersSheet = ss.insertSheet("Users");
  }
  usersSheet.clear();
  const usersHeaders = ["Username", "Password", "Role"];
  usersSheet.getRange(1, 1, 1, usersHeaders.length).setValues([usersHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");
  
  const dummyUsers = [
    ["admin", "admin123", "ADMIN"],
    ["warehouse", "wh123", "WAREHOUSE"],
    ["sales", "sales123", "SALES"],
    ["teknisi", "tek123", "TEKNISI"],
    ["purchase", "purchase123", "PURCHASE"]
  ];
  usersSheet.getRange(2, 1, dummyUsers.length, dummyUsers[0].length).setValues(dummyUsers);

  let brandingSheet = ss.getSheetByName("Branding_Settings");
  if (!brandingSheet) {
    brandingSheet = ss.insertSheet("Branding_Settings");
  }
  brandingSheet.clear();
  const brandingHeaders = ["Setting Key", "Setting Value"];
  brandingSheet.getRange(1, 1, 1, brandingHeaders.length).setValues([brandingHeaders])
    .setFontWeight("bold")
    .setBackground("#0f172a")
    .setFontColor("#ffffff");

  const defaultBranding = [
    ["app_title", "ALMERASS"],
    ["app_subtitle", "INVENTORY"],
    ["footer_text", "© 2026 Almerass Premium Inventory. Crafted for Ultimate Efficiency and Performance."],
    ["accent_color", "#00ff66"],
    ["logo_text", "A"]
  ];
  brandingSheet.getRange(2, 1, defaultBranding.length, defaultBranding[0].length).setValues(defaultBranding);

  // Auto resize columns for all sheets
  ss.getSheets().forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      sheet.autoResizeColumns(1, lastCol);
    }
  });

  SpreadsheetApp.flush();
  return "Spreadsheet Setup Berhasil!";
}

/**
 * ID Transaksi harian (format: TRX-YYMMDD-XXXX)
 */
function generateTransactionId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Riwayat_Transaksi");
  if (!sheet) {
    sheet = ss.insertSheet("Riwayat_Transaksi");
    sheet.appendRow(["ID Transaksi", "Tanggal", "Kode Barang", "Nama Barang", "Tipe", "User", "Qty", "Catatan"]);
    SpreadsheetApp.flush();
  }
  
  const d = new Date();
  const yy = d.getFullYear().toString().slice(-2);
  const mm = (d.getMonth() + 1).toString().padStart(2, '0');
  const dd = d.getDate().toString().padStart(2, '0');
  const todayPrefix = yy + mm + dd; 
  
  const lastRow = sheet.getLastRow();
  let lastNum = 0;
  
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const id = data[i][0].toString();
      if (id.startsWith("TRX-" + todayPrefix)) {
        const parts = id.split("-");
        if (parts.length === 3) {
          const num = parseInt(parts[2], 10);
          if (!isNaN(num)) {
            lastNum = num;
            break;
          }
        }
      }
    }
  }
  
  const nextNum = (lastNum + 1).toString().padStart(4, '0');
  return "TRX-" + todayPrefix + "-" + nextNum;
}

/**
 * ID Laporan Instalasi harian (format: INS-YYMMDD-XXXX)
 */
function generateInstalasiId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Laporan_Instalasi");
  if (!sheet) {
    sheet = ss.insertSheet("Laporan_Instalasi");
    sheet.appendRow(["ID Ins", "Tanggal", "No. PO", "Customer", "Teknisi", "Status", "SN Unit", "Item Dipasang", "Catatan"]);
    SpreadsheetApp.flush();
  }
  
  const d = new Date();
  const yy = d.getFullYear().toString().slice(-2);
  const mm = (d.getMonth() + 1).toString().padStart(2, '0');
  const dd = d.getDate().toString().padStart(2, '0');
  const todayPrefix = yy + mm + dd;
  
  const lastRow = sheet.getLastRow();
  let lastNum = 0;
  
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const id = data[i][0].toString();
      if (id.startsWith("INS-" + todayPrefix)) {
        const parts = id.split("-");
        if (parts.length === 3) {
          const num = parseInt(parts[2], 10);
          if (!isNaN(num)) {
            lastNum = num;
            break;
          }
        }
      }
    }
  }
  
  const nextNum = (lastNum + 1).toString().padStart(4, '0');
  return "INS-" + todayPrefix + "-" + nextNum;
}

/**
 * ID Request Order harian (format: REQ-YYMMDD-XXXX)
 */
function generateRequestId() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Request_Order");
  if (!sheet) {
    sheet = ss.insertSheet("Request_Order");
    sheet.appendRow(["ID Request", "Tanggal", "Target Produk", "Qty", "Pemohon", "No. Referensi", "Status"]);
    SpreadsheetApp.flush();
  }
  
  const d = new Date();
  const yy = d.getFullYear().toString().slice(-2);
  const mm = (d.getMonth() + 1).toString().padStart(2, '0');
  const dd = d.getDate().toString().padStart(2, '0');
  const todayPrefix = yy + mm + dd;
  
  const lastRow = sheet.getLastRow();
  let lastNum = 0;
  
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      const id = data[i][0].toString();
      if (id.startsWith("REQ-" + todayPrefix)) {
        const parts = id.split("-");
        if (parts.length === 3) {
          const num = parseInt(parts[2], 10);
          if (!isNaN(num)) {
            lastNum = num;
            break;
          }
        }
      }
    }
  }
  
  const nextNum = (lastNum + 1).toString().padStart(4, '0');
  return "REQ-" + todayPrefix + "-" + nextNum;
}
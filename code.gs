// KONFIGURASI NAMA SHEET - DPMPTSP KABUPATEN MOROWALI
const SHEET_PEGAWAI = "Data_Pegawai";
const SHEET_REKAP = "Rekap_Absen";
const SHEET_AKUN = "Data_Akun";
const SHEET_LOG = "Log_Absen_Harian";

// --- FUNGSI GET (MENARIK DATA) ---
function doGet(e) {
  const op = e.parameter.action;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  if (op == "get_pegawai") {
    const sheet = ss.getSheetByName(SHEET_PEGAWAI);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    return response(data.map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }));
  }

  if (op == "get_rekap") {
    const sheet = ss.getSheetByName(SHEET_REKAP);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    return response(data.map(row => {
      let obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }));
  }

  if (op == "get_log_harian") {
    const sheet = ss.getSheetByName(SHEET_LOG);
    if (!sheet) return response([]);
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    const hasil = data.map(row => {
      let obj = {};
      headers.forEach((h, i) => {
        obj[h.replace(/\s+/g, '_')] = row[i];
      });
      return obj;
    });
    return response(hasil);
  }
}

// --- FUNGSI POST (SIMPAN/UPDATE/HAPUS) ---
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = JSON.parse(e.postData.contents);
    const tz = "GMT+08:00"; // WITA (Morowali)
    
    if (data.action === "login") {
      const sheetAkun = ss.getSheetByName(SHEET_AKUN);
      const dbAkun = sheetAkun.getDataRange().getValues();
      for (let i = 1; i < dbAkun.length; i++) {
        if (dbAkun[i][0] == data.username && dbAkun[i][1] == data.password) {
          return response({ status: "success", role: dbAkun[i][2], nama: dbAkun[i][3] });
        }
      }
      return response({ status: "error", message: "Username/Password salah!" });
    }
    
    if (data.action === "simpan_absen") {
      const sheet = ss.getSheetByName(SHEET_REKAP);
      const idRekap = data.nip + "-" + data.bulan + "-" + data.tahun;
      const existingData = sheet.getDataRange().getValues();
      let rowIdx = -1;
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] == idRekap) { rowIdx = i + 1; break; }
      }
      const rowData = [idRekap, data.nip, data.bulan, data.tahun, data.hadir, data.sakit, data.izin, data.alpa, data.keterangan];
      if (rowIdx > -1) {
        sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      return response({status: "success", message: "Data rekap tersimpan"});
    }

    if (data.action === "simpan_pegawai") {
      const sheet = ss.getSheetByName(SHEET_PEGAWAI);
      const existingData = sheet.getDataRange().getValues();
      let rowIdx = -1;
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] == data.nip) { rowIdx = i + 1; break; }
      }
      const rowData = [data.nip, data.nama, data.jabatan, data.kategori, data.status];
      if (rowIdx > -1) {
        sheet.getRange(rowIdx, 1, 1, rowData.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      return response({status: "success", message: "Master pegawai diperbarui"});
    }

    if (data.action === "absen_digital") {
      const sheetLog = ss.getSheetByName(SHEET_LOG);
      const nip = String(data.nip).trim();
      const nama = data.nama;
      let aksi = data.aksi;
      
      const waktuServer = new Date();
      const tglHariIni = Utilities.formatDate(waktuServer, tz, "yyyy-MM-dd");
      // PERBAIKAN: Menggunakan pengunci petik (') agar Sheets tidak merusak format tanggal
      const waktuLengkap = "'" + Utilities.formatDate(waktuServer, tz, "yyyy-MM-dd HH:mm:ss"); 
      const jamServer = Utilities.formatDate(waktuServer, tz, "HH:mm");
      
      if (aksi === "Datang") {
        if (jamServer < "07:30" || jamServer > "09:00") return response({status: "error", message: "Di luar jam Absen Pagi."});
        if (jamServer > "08:00") aksi = "Datang (Terlambat)";
      } else if (aksi === "Pulang") {
        if (jamServer < "16:00" || jamServer > "17:00") return response({status: "error", message: "Di luar jam Absen Pulang."});
      }

      const dataLog = sheetLog.getDataRange().getValues();
      for(let i = 1; i < dataLog.length; i++) {
        let rowTglStr = "";
        try {
          let cellDate = dataLog[i][0];
          if (cellDate instanceof Date) { 
            rowTglStr = Utilities.formatDate(cellDate, tz, "yyyy-MM-dd"); 
          } else {
            let strVal = String(cellDate).replace("'", "").split(" ")[0];
            if(strVal.includes("/")) {
              let p = strVal.split("/");
              rowTglStr = (p[2].length === 4) ? p[2] + "-" + p[1].padStart(2, '0') + "-" + p[0].padStart(2, '0') : p[0] + "-" + p[1].padStart(2, '0') + "-" + p[2].padStart(2, '0');
            } else if(strVal.includes("-")) {
              let p = strVal.split("-");
              rowTglStr = (p[0].length === 4) ? p[0] + "-" + p[1].padStart(2, '0') + "-" + p[2].padStart(2, '0') : p[2] + "-" + p[1].padStart(2, '0') + "-" + p[0].padStart(2, '0');
            }
          }
        } catch(e) { continue; }
        
        if(String(dataLog[i][1]).trim() === nip && rowTglStr === tglHariIni) {
          let rowAksi = dataLog[i][3];
          if((rowAksi.includes("Datang") || rowAksi === "Tugas Luar") && (aksi.includes("Datang") || aksi === "Tugas Luar")) return response({ status: "error", message: "Akses Ditolak: Anda sudah absen masuk!" });
          if(rowAksi === "Pulang" && aksi === "Pulang") return response({status: "error", message: "Akses Ditolak: Anda sudah absen pulang!"});
        }
      }
      
      sheetLog.appendRow([waktuLengkap, "'" + nip, nama, aksi, data.statusJarak, data.koordinat]);
      return response({ status: "success", message: "Berhasil! Absen " + aksi + " terekam." });
    }

// LOGIKA 5: GENERATE REKAP OTOMATIS (VERSI ANTI-SPASI & ANTI-FORMAT TERBALIK)
    if (data.action === "generate_rekap") {
      const sheetLog = ss.getSheetByName(SHEET_LOG);
      const sheetRekap = ss.getSheetByName(SHEET_REKAP);
      const sheetPegawai = ss.getSheetByName(SHEET_PEGAWAI);

      const blnTujuan = data.bulan; 
      const thnTujuan = String(data.tahun);
      const totalHariKerja = parseInt(data.hariKerja) || 0;
      const mapBulan = {"Januari":"01","Februari":"02","Maret":"03","April":"04","Mei":"05","Juni":"06","Juli":"07","Agustus":"08","September":"09","Oktober":"10","November":"11","Desember":"12"};
      const blnAngka = mapBulan[blnTujuan];

      const dataPegawai = sheetPegawai.getDataRange().getValues();
      let listPegawai = [];
      for(let i=1; i<dataPegawai.length; i++){
        if(dataPegawai[i][4] === "Aktif") {
          // PERBAIKAN: Hapus semua spasi dan tanda petik dari NIP Master
          let cleanNipMaster = String(dataPegawai[i][0]).replace(/\s+/g, '').replace(/'/g, "");
          listPegawai.push({ nipAsli: String(dataPegawai[i][0]), nipClean: cleanNipMaster, nama: dataPegawai[i][1] });
        }
      }

      const dataLog = sheetLog.getDataRange().getValues();
      let rekapHarian = {}; 

      for(let i=1; i<dataLog.length; i++) {
        let cellDate = dataLog[i][0];
        let rowTglStr = "";
        try {
          if (cellDate instanceof Date) { 
            rowTglStr = Utilities.formatDate(cellDate, tz, "yyyy-MM-dd"); 
          } else {
            let strVal = String(cellDate).replace(/'/g, "").split(" ")[0];
            let p = strVal.includes("-") ? strVal.split("-") : strVal.split("/");
            if(p[0].length === 4) rowTglStr = p[0] + "-" + p[1].padStart(2, '0') + "-" + p[2].padStart(2, '0');
            else rowTglStr = p[2] + "-" + p[1].padStart(2, '0') + "-" + p[0].padStart(2, '0');
          }
        } catch(e) { continue; }

        if(rowTglStr && rowTglStr.split("-")[0] === thnTujuan && rowTglStr.split("-")[1] === blnAngka) {
            // PERBAIKAN: Hapus semua spasi dan tanda petik dari NIP di Log
            let nipLogClean = String(dataLog[i][1]).replace(/\s+/g, '').replace(/'/g, "");
            
            if(!rekapHarian[nipLogClean]) rekapHarian[nipLogClean] = {};
            if(!rekapHarian[nipLogClean][rowTglStr]) rekapHarian[nipLogClean][rowTglStr] = { masuk: false, pulang: false, tl: false };
            
            let aksi = dataLog[i][3];
            if(aksi.includes("Datang")) rekapHarian[nipLogClean][rowTglStr].masuk = true;
            if(aksi === "Pulang") rekapHarian[nipLogClean][rowTglStr].pulang = true;
            if(aksi === "Tugas Luar") rekapHarian[nipLogClean][rowTglStr].tl = true;
        }
      }

      const sheetRekapValues = sheetRekap.getDataRange().getValues();

      for(let p=0; p<listPegawai.length; p++){
         let nipC = listPegawai[p].nipClean;
         let nipA = listPegawai[p].nipAsli;
         let jmlHadir = 0, jmlTL = 0, jmlLupaPulang = 0;
         
         if(rekapHarian[nipC]) {
            for(let tgl in rekapHarian[nipC]) {
               let h = rekapHarian[nipC][tgl];
               if(h.tl) jmlTL++;
               else if(h.masuk) { jmlHadir++; if(!h.pulang) jmlLupaPulang++; }
            }
         }

         // Cari data existing berdasarkan NIP Bersih
         let existing = { rowIdx: -1, s: 0, i: 0, ket: "" };
         for(let r=1; r<sheetRekapValues.length; r++) {
           let nipRekapClean = String(sheetRekapValues[r][1]).replace(/\s+/g, '').replace(/'/g, "");
           if(nipRekapClean === nipC && sheetRekapValues[r][2] === blnTujuan) {
             existing = { rowIdx: r + 1, s: parseInt(sheetRekapValues[r][5])||0, i: parseInt(sheetRekapValues[r][6])||0, ket: sheetRekapValues[r][8]||"" };
             break;
           }
         }

         let alpa = totalHariKerja - (jmlHadir + jmlTL + existing.s + existing.i);
         if(alpa < 0) alpa = 0; 
         let ketAuto = (jmlTL > 0 ? `TL:${jmlTL}x ` : "") + (jmlLupaPulang > 0 ? `LupaPulang:${jmlLupaPulang}x ` : "");
         let ketFinal = existing.ket ? (ketAuto + "| " + existing.ket) : ketAuto;

         let rowData = [nipA + "-" + blnTujuan + "-" + thnTujuan, "'" + nipA, blnTujuan, thnTujuan, jmlHadir, existing.s, existing.i, alpa, ketFinal];
         if(existing.rowIdx > -1) sheetRekap.getRange(existing.rowIdx, 1, 1, rowData.length).setValues([rowData]);
         else sheetRekap.appendRow(rowData);
      }
      return response({status: "success", message: "Sinkronisasi Berhasil! NIP telah disesuaikan otomatis."});
    }
  } finally { lock.releaseLock(); }
}

function response(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }

// FUNGSI KHUSUS UNTUK MEMPERBAIKI TANGGAL BULAN MEI YANG TERLANJUR TERBALIK
// FUNGSI KHUSUS UNTUK MEMPERBAIKI TANGGAL (VERSI SUPER CEPAT / BATCH PROCESSING)
function perbaikiTanggalTerbalik() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log_Absen_Harian");
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return; // Jika tidak ada data, hentikan
  
  // Ambil KHUSUS kolom A saja sekaligus agar memori tidak berat
  const range = sheet.getRange(2, 1, lastRow - 1, 1);
  const data = range.getDisplayValues(); 
  
  let jumlah = 0;
  let adaPerubahan = false;

  for(let i = 0; i < data.length; i++) {
    let tgl = String(data[i][0]);
    
    // Cari tanggal yang terdeteksi sebagai November
    if(tgl.includes("2026-11-") || tgl.includes("/11/2026")) {
      let jam = "";
      try { jam = tgl.split(" ")[1]; } catch(e) { jam = "00:00:00"; }
      
      let p = tgl.includes("-") ? tgl.split(" ")[0].split("-") : tgl.split(" ")[0].split("/");
      
      // Kembalikan ke format baku terkunci (Mei)
      let fix = "";
      if(p[0].length === 4) {
        fix = "'2026-05-" + p[1].padStart(2,'0') + " " + jam;
      } else {
        fix = "'2026-05-" + p[0].padStart(2,'0') + " " + jam;
      }
      
      data[i][0] = fix; // Ubah di dalam memori
      jumlah++;
      adaPerubahan = true;
    }
  }
  
  // Timpa/simpan ke Spreadsheet SEKALI JALAN (Ini yang bikin cepat 1 detik)
  if(adaPerubahan) {
    range.setValues(data);
    SpreadsheetApp.getUi().alert("Selesai! " + jumlah + " data bulan Mei berhasil diperbaiki dalam sekejap.");
  } else {
    SpreadsheetApp.getUi().alert("Sudah bersih! Tidak ditemukan lagi tanggal yang terbalik.");
  }
}
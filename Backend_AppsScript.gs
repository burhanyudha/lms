function doGet(e) {
  var sheetName = e.parameter.sheet;
  if (!sheetName) {
    return ContentService.createTextOutput(JSON.stringify({error: "Sheet tidak valid"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  if (sheetName === "public_guru") {
    var sheetAdmin = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("admin");
    if (!sheetAdmin) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    
    var rawAdmin = sheetAdmin.getDataRange().getValues();
    var result = [];
    if (rawAdmin.length > 1) {
      var headersAdmin = rawAdmin[0].map(function(h) { return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); });
      var idxNama = headersAdmin.indexOf("nama");
      var idxMapel = headersAdmin.indexOf("mapel");
      
      for (var i = 1; i < rawAdmin.length; i++) {
        if (idxNama !== -1 && idxMapel !== -1) {
           result.push({
             nama: rawAdmin[i][idxNama],
             mapel: rawAdmin[i][idxMapel]
           });
        }
      }
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) {
    if (sheetName === "admin") {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("admin");
      sheet.appendRow(["nama", "username", "password", "role", "mapel", "kelas"]);
      sheet.appendRow(["Admin Utama", "admin", "admin123", "Administrator", "", ""]);
    } else {
      return ContentService.createTextOutput(JSON.stringify([]))
                           .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  var rawData = sheet.getDataRange().getValues();
  var displayData = sheet.getDataRange().getDisplayValues();
  
  if (rawData.length <= 1) {
    return ContentService.createTextOutput(JSON.stringify([]))
                         .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Baca nama header secara dinamis & kebal geser posisi
  var headers = rawData[0].map(function(h) {
    return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); 
  });
  
  var result = [];
  for (var i = 1; i < rawData.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      if (headers[j]) {
        // Ambil displayData khusus untuk waktu agar format (00:00) aman
        if (headers[j] === "waktu" || headers[j] === "timestamp") {
           obj[headers[j]] = displayData[i][j];
        } else {
           obj[headers[j]] = rawData[i][j];
        }
        
        // Simpan key dengan huruf asli juga untuk kompatibilitas Admin Panel (CRUD)
        var originalHeader = rawData[0][j].toString().trim();
        if (originalHeader) {
           if (originalHeader.toLowerCase() === "waktu" || originalHeader.toLowerCase() === "timestamp") {
               obj[originalHeader] = displayData[i][j];
           } else {
               obj[originalHeader] = rawData[i][j];
           }
        }
      }
    }
    
    // TAMBAHAN UNTUK ADMIN PANEL: Menyimpan nomor baris untuk Edit/Hapus
    obj.rowNumber = i + 1; 
    
    result.push(obj);
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
                       .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var postData = JSON.parse(e.postData.contents);
    var action = postData.action;
    var sheetName = postData.sheet;
    var data = postData.data; // Dipakai di fungsi admin jika ada JSON string
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Beberapa action lama tidak mengirim sheetName, jadi didefinisikan secara lokal jika diperlukan.
    var sheet = sheetName ? ss.getSheetByName(sheetName) : null;
    
    // ----------------------------------------------------
    // ACTION 1: MENYIMPAN NILAI KUIS/PRAKTIK (FUNGSI LAMA)
    // ----------------------------------------------------
    if (action === 'saveScore') {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error("Sheet tidak ditemukan");
      
      var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Ambil data langsung dari postData (sesuai script lama)
      var pData = postData.data || postData;
      
      // Otomatis buat header jawaban jika belum ada di sheet
      var normalizedHeaders = headersRow.map(function(h) {
        return h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
      });
      
      if (pData.jawaban && pData.jawaban.length > 0) {
        var addedNewHeader = false;
        for (var k = 0; k < pData.jawaban.length; k++) {
          var expectedHeaderName = "jawaban" + (k + 1);
          if (normalizedHeaders.indexOf(expectedHeaderName) === -1) {
            var newColIdx = headersRow.length + 1;
            sheet.getRange(1, newColIdx).setValue("Jawaban " + (k + 1));
            headersRow.push("Jawaban " + (k + 1));
            normalizedHeaders.push(expectedHeaderName);
            addedNewHeader = true;
          }
        }
      }
      
      // Otomatis buat header urutan_soal jika belum ada
      if (pData.urutan_soal && normalizedHeaders.indexOf("urutansoal") === -1) {
        var newColIdx = headersRow.length + 1;
        sheet.getRange(1, newColIdx).setValue("Urutan Soal");
        headersRow.push("Urutan Soal");
        normalizedHeaders.push("urutansoal");
      }
      
      var newRow = new Array(headersRow.length).fill("");
      
      for (var i = 0; i < headersRow.length; i++) {
        var headerName = headersRow[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        
        if (headerName === "username") newRow[i] = pData.username;
        else if (headerName === "nilai" || headerName === "skor") newRow[i] = pData.nilai;
        else if (headerName === "timestamp") newRow[i] = pData.timestamp;
        else if (headerName === "nama") newRow[i] = pData.nama;
        else if (headerName === "tipe") newRow[i] = pData.tipe;
        else if (headerName === "idmateri" || headerName === "id_materi") newRow[i] = pData.id_materi;
        else if (headerName === "waktu") newRow[i] = pData.waktu;
        else if (headerName === "urutansoal" || headerName === "urutan_soal") newRow[i] = pData.urutan_soal;
        else if (headerName.indexOf("jawaban") === 0) {
          var idx = parseInt(headerName.replace("jawaban", "")) - 1;
          if (pData.jawaban && pData.jawaban[idx] !== undefined) {
             newRow[i] = pData.jawaban[idx];
          }
        }
      }
      
      sheet.appendRow(newRow);
      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // ACTION 2: MENGAMBIL WAKTU TIMER KUIS (FUNGSI LAMA)
    // ----------------------------------------------------
    if (action === "getTimer") {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error("Sheet tidak ditemukan");
      var pData = postData.data || postData;
      
      var rawData = sheet.getDataRange().getValues();
      if (rawData.length > 0) {
        var headers = rawData[0].map(function(h) { return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); });
        
        var idxUser = headers.indexOf("username");
        var idxMateri = headers.indexOf("idmateri");
        if(idxMateri === -1) idxMateri = headers.indexOf("id_materi");
        var idxWaktu = headers.indexOf("waktuaktif");
        
        if(idxUser !== -1 && idxMateri !== -1 && idxWaktu !== -1) {
          for (var i = rawData.length - 1; i >= 1; i--) {
            if (rawData[i][idxUser] == pData.username && rawData[i][idxMateri] == pData.id_materi) {
              return ContentService.createTextOutput(JSON.stringify({status: "success", waktu_aktif: rawData[i][idxWaktu]}))
                                   .setMimeType(ContentService.MimeType.JSON);
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success", waktu_aktif: 0}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // ACTION 3: MENYIMPAN/SYNC WAKTU TIMER KUIS (FUNGSI LAMA)
    // ----------------------------------------------------
    if (action === "syncTimer") {
      sheet = ss.getSheetByName(sheetName);
      if (!sheet) throw new Error("Sheet tidak ditemukan");
      var pData = postData.data || postData;
      
      var rawData = sheet.getDataRange().getValues();
      var headers = [];
      
      // Jika sheet masih benar-benar kosong, buatkan header otomatis
      if (rawData.length === 0 || rawData[0].length === 0 || rawData[0][0] === "") {
        sheet.getRange(1, 1, 1, 3).setValues([["username", "idmateri", "waktuaktif"]]);
        headers = ["username", "idmateri", "waktuaktif"];
        rawData = sheet.getDataRange().getValues();
      } else {
        headers = rawData[0].map(function(h) { return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); });
      }
      
      var idxUser = headers.indexOf("username");
      var idxMateri = headers.indexOf("idmateri");
      if (idxMateri === -1) idxMateri = headers.indexOf("id_materi");
      var idxWaktu = headers.indexOf("waktuaktif");
      
      // Jika ternyata headernya tidak sesuai format, paksa saja ke indeks 0, 1, dan 2
      if (idxUser === -1) idxUser = 0;
      if (idxMateri === -1) idxMateri = 1;
      if (idxWaktu === -1) idxWaktu = 2;
      
      var found = false;
      for (var i = rawData.length - 1; i >= 1; i--) {
        if (rawData[i][idxUser] == pData.username && rawData[i][idxMateri] == pData.id_materi) {
          sheet.getRange(i + 1, idxWaktu + 1).setValue(pData.waktu_aktif);
          found = true;
          break;
        }
      }
      
      if (!found) {
        var newRow = new Array(Math.max(headers.length, 3)).fill("");
        newRow[idxUser] = pData.username;
        newRow[idxMateri] = pData.id_materi;
        newRow[idxWaktu] = pData.waktu_aktif;
        sheet.appendRow(newRow);
      }
      
      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION: SAVE LOGIN DATA
    // ===============================================
    if (action === "saveLogin") {
      sheet = ss.getSheetByName("logLogin");
      if (!sheet) {
        sheet = ss.insertSheet("logLogin");
        sheet.appendRow(["username", "nama", "kelas", "waktu", "status"]);
      }
      
      var pData = postData.data || postData;
      sheet.appendRow([
         pData.username || "-",
         pData.nama || "-",
         pData.kelas || "-",
         pData.waktu || new Date().toISOString(),
         pData.status || "Sukses"
      ]);
      
      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION: UPLOAD TUGAS
    // ===============================================
    if (action === "uploadTugas") {
      var pData = postData.data || postData;
      
      // Ambil folder ID dari settings
      var settingsSheet = ss.getSheetByName("settings");
      var folderId = "";
      if (settingsSheet) {
         var setObj = settingsSheet.getDataRange().getValues();
         for(var s=1; s<setObj.length; s++){
            if(setObj[s][0] === "FOLDER_TUGAS") { folderId = setObj[s][1]; break; }
         }
      }
      if (!folderId) {
         throw new Error("Folder ID belum diatur oleh Admin. Hubungi Admin.");
      }
      
      try {
         var folder = DriveApp.getFolderById(folderId);
         var file;
         if (pData.htmlContent) {
             // Konversi HTML murni ke PDF via GAS
             var htmlTemplate = `
                 <!DOCTYPE html>
                 <html>
                 <head>
                     <meta charset="UTF-8">
                     <style>
                         body { font-family: Arial, sans-serif; padding: 20px; line-height: 1.5; color: #111; word-wrap: break-word; overflow-wrap: break-word; }
                         .kop { border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
                         .kop h1 { margin: 0; font-size: 22px; font-weight: bold; }
                         .kop p { margin: 5px 0 0 0; font-size: 14px; color: #333; }
                         img { max-width: 100%; height: auto; display: block; margin: 10px 0; }
                         p { margin-bottom: 1em; word-break: break-word; white-space: pre-wrap; }
                         ul, ol { margin-bottom: 1em; padding-left: 20px; word-break: break-word; }
                     </style>
                 </head>
                 <body>
                     <div class="kop">
                         <h1>${pData.judul || "Tugas"}</h1>
                         <p><b>Nama:</b> ${pData.nama || "-"}</p>
                         <p><b>Kelas:</b> ${pData.kelas || "-"}</p>
                         <p><b>Waktu Pengumpulan:</b> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}</p>
                     </div>
                     <div class="content">
                         ${pData.htmlContent}
                     </div>
                 </body>
                 </html>
             `;
             var blobHtml = Utilities.newBlob(htmlTemplate, MimeType.HTML, "tugas.html");
             var pdfBlob = blobHtml.getAs(MimeType.PDF);
             pdfBlob.setName(pData.filename);
             file = folder.createFile(pdfBlob);
         } else if (pData.base64) {
             // Decode Base64 (Untuk backward compatibility jika ada fitur lain yang masih pakai base64)
             var splitBase = pData.base64.split(',');
             var type = splitBase[0].split(';')[0].replace('data:', '');
             var byteCharacters = Utilities.base64Decode(splitBase[1]);
             var blob = Utilities.newBlob(byteCharacters, type, pData.filename);
             file = folder.createFile(blob);
         } else {
             throw new Error("Tidak ada data file yang dikirim.");
         }
         
         // Atur izin file agar bisa dilihat oleh siapa saja yang memiliki link (tanpa harus login Google)
         try {
             file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         } catch (sharingError) {
             // Abaikan jika akun Google Workspace sekolah melarang berbagi file publik
             Logger.log("Gagal set izin sharing: " + sharingError.toString());
         }
         
         var fileUrl = file.getUrl();
         
         // Simpan record ke logTugas
         var sheetTugas = ss.getSheetByName("logTugas");
         if (!sheetTugas) {
            sheetTugas = ss.insertSheet("logTugas");
            sheetTugas.appendRow(["username", "nama", "kelas", "id_materi", "materi", "tipe", "link_tugas", "waktu"]);
         }
         
         var now = new Date();
         var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
         
         sheetTugas.appendRow([
            pData.username || "-",
            pData.nama || "-",
            pData.kelas || "-",
            pData.id_materi || "-",
            pData.judul || "Tugas",
            pData.tipe || "Tugas",
            fileUrl,
            timeStr
         ]);
         
         return ContentService.createTextOutput(JSON.stringify({status: "success", url: fileUrl}))
                              .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
         throw new Error("Gagal mengupload file: " + err.message);
      }
    }

    // ===============================================
    // ACTION: REORDER DATA (DRAG AND DROP)
    // ===============================================
    if (action === "reorderData") {
      var sheet = ss.getSheetByName(postData.sheetName);
      if (!sheet) throw new Error("Sheet tidak ditemukan");
      
      var newOrder = postData.newOrder; // Array of row numbers (1-based index from sheet)
      if (!newOrder || !Array.isArray(newOrder)) throw new Error("Invalid newOrder format");
      
      var allData = sheet.getDataRange().getValues();
      var headers = allData[0];
      var newDataRows = [];
      
      var processedRows = new Set();
      
      // Masukkan data berdasarkan urutan baru
      for (var i = 0; i < newOrder.length; i++) {
        var rIdx = newOrder[i] - 1; // Konversi rowNumber (Google Sheet) ke array index (0-based)
        if (rIdx > 0 && rIdx < allData.length) {
          newDataRows.push(allData[rIdx]);
          processedRows.add(rIdx);
        }
      }
      
      // Masukkan baris lain yang mungkin tidak ikut di-reorder (misal disembunyikan filter)
      // Taruh di bawah (atau di posisinya) agar data tidak hilang
      for (var i = 1; i < allData.length; i++) {
        if (!processedRows.has(i)) {
          newDataRows.push(allData[i]);
        }
      }
      
      if (newDataRows.length > 0) {
        // Timpa data mulai dari baris ke-2
        sheet.getRange(2, 1, newDataRows.length, headers.length).setValues(newDataRows);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Urutan berhasil diperbarui" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION 4: ADD / EDIT DATA (ADMIN PANEL)
    // ===============================================
    if (action === "saveData") {
      if (!sheet) {
        if (sheetName === "settings") {
          sheet = ss.insertSheet("settings");
          sheet.appendRow(["key", "value"]);
        } else {
          throw new Error("Sheet " + sheetName + " tidak ditemukan");
        }
      }

      var dataToSave;
      try {
        // Karena di admin.html data dikirim sebagai JSON string di dalam JSON,
        // kita parse lagi properti data-nya.
        dataToSave = (typeof data === "string") ? JSON.parse(data) : data;
      } catch (err) {
        throw new Error("Format data tidak valid (bukan JSON)");
      }

      var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      if (headers.length === 1 && headers[0] === "") headers = []; // Empty sheet edge case
      
      var normalizedHeaders = headers.map(function(h) { return h.toString().trim().toLowerCase(); });
      
      // Otomatis tambahkan kolom baru jika belum ada di sheet
      for (var key in dataToSave) {
        if (dataToSave.hasOwnProperty(key)) {
           var keyLower = key.toString().trim().toLowerCase();
           if (normalizedHeaders.indexOf(keyLower) === -1) {
              headers.push(key);
              normalizedHeaders.push(keyLower);
              sheet.getRange(1, headers.length).setValue(key);
           }
        }
      }
      
      // JIKA MENGEDIT BARIS YANG ADA
      if (postData.rowNumber && postData.rowNumber !== "") {
        var rowNum = parseInt(postData.rowNumber);
        var updateRow = new Array(headers.length).fill("");
        var existingData = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
        
        for (var i = 0; i < headers.length; i++) {
          var h = headers[i].toString().trim(); 
          var hLower = h.toLowerCase();
          
          if (dataToSave[h] !== undefined) {
             updateRow[i] = dataToSave[h];
          } else if (dataToSave[hLower] !== undefined) {
             updateRow[i] = dataToSave[hLower];
          } else {
             updateRow[i] = existingData[i]; 
          }
        }
        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updateRow]);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data berhasil diperbarui" }))
                             .setMimeType(ContentService.MimeType.JSON);
      } 
      // JIKA MENAMBAH BARIS BARU
      else {
        var newRow = new Array(headers.length).fill("");
        for (var i = 0; i < headers.length; i++) {
          var h = headers[i].toString().trim();
          var hLower = h.toLowerCase();
          
          if (dataToSave[h] !== undefined) {
             newRow[i] = dataToSave[h];
          } else if (dataToSave[hLower] !== undefined) {
             newRow[i] = dataToSave[hLower];
          }
        }
        sheet.appendRow(newRow);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data baru berhasil ditambahkan" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ===============================================
    // ACTION 5: DELETE DATA (ADMIN PANEL)
    // ===============================================
    if (action === "deleteData") {
      if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan");

      var rowNum = parseInt(postData.rowNumber);
      if (rowNum && rowNum > 1) {
        sheet.deleteRow(rowNum);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data berhasil dihapus" }))
                             .setMimeType(ContentService.MimeType.JSON);
      } else {
        throw new Error("Nomor baris tidak valid");
      }
    }
    
    // ===============================================
    // ACTION 6: UPDATE NILAI LANGSUNG (ADMIN PANEL)
    // ===============================================
    if (action === "updateNilaiByUsernameMateri") {
       var sheetNilai = ss.getSheetByName("listNilai");
       if (!sheetNilai) throw new Error("Sheet listNilai tidak ditemukan");
       
       var targetUser = postData.username;
       var targetMateri = postData.id_materi;
       var newNilai = postData.nilai;
       var newWaktu = postData.waktu; // Durasi baru
       var targetNama = postData.nama || "";
       var targetKelas = postData.kelas || "";
       
       var sheetData = sheetNilai.getDataRange().getValues();
       var headers = sheetData[0];
       var userCol = -1, materiCol = -1, nilaiCol = -1, waktuCol = -1, namaCol = -1, kelasCol = -1;
       
       for(var i = 0; i < headers.length; i++){
          var h = headers[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
          if(h === "username" || h === "nisn") userCol = i;
          if(h === "id_materi" || h === "idmateri" || h === "modul") materiCol = i;
          if(h === "nilai" || h === "skor") nilaiCol = i;
          if(h === "waktu" || h === "durasi") waktuCol = i;
          if(h === "nama" || h === "namasiswa" || h === "namalengkap") namaCol = i;
          if(h === "kelas" || h === "tingkat") kelasCol = i;
       }
       
       if(userCol === -1 || materiCol === -1 || nilaiCol === -1) {
         throw new Error("Struktur kolom listNilai tidak lengkap (butuh username, id_materi, nilai)");
       }
       
       var updated = false;
       for(var j = sheetData.length - 1; j > 0; j--) {
          var rowUsername = sheetData[j][userCol].toString().trim();
          var rowMateri = sheetData[j][materiCol].toString().trim();
          
          if(rowUsername === targetUser && rowMateri === targetMateri) {
             sheetNilai.getRange(j + 1, nilaiCol + 1).setValue(newNilai);
             if (waktuCol !== -1 && newWaktu !== undefined) {
                 sheetNilai.getRange(j + 1, waktuCol + 1).setValue(newWaktu);
             }
             if (namaCol !== -1 && targetNama !== "") {
                 sheetNilai.getRange(j + 1, namaCol + 1).setValue(targetNama);
             }
             if (kelasCol !== -1 && targetKelas !== "") {
                 sheetNilai.getRange(j + 1, kelasCol + 1).setValue(targetKelas);
             }
             
             // Update timestamp
             var now = new Date();
             var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
             for (var colIdx = 0; colIdx < headers.length; colIdx++) {
                 var hT = headers[colIdx].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                 if ((hT === "timestamp" || hT === "waktusubmit" || hT === "tanggal" || hT === "waktu") && colIdx !== waktuCol) {
                     sheetNilai.getRange(j + 1, colIdx + 1).setValue(timeStr);
                 }
             }
             
             updated = true;
             break;
          }
       }
       
       if(updated) {
         return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Nilai & Durasi berhasil diperbarui" }))
                              .setMimeType(ContentService.MimeType.JSON);
       } else {
         var newRow = new Array(headers.length).fill("");
         newRow[userCol] = targetUser;
         newRow[materiCol] = targetMateri;
         newRow[nilaiCol] = newNilai;
         if(namaCol !== -1) newRow[namaCol] = targetNama;
         if(kelasCol !== -1) newRow[kelasCol] = targetKelas;
         
         for(var i = 0; i < headers.length; i++){
           var h = headers[i].toString().toLowerCase();
           if(h === "waktu" || h === "timestamp") {
              var now = new Date();
              newRow[i] = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
           }
           if (waktuCol !== -1 && i === waktuCol && newWaktu !== undefined) {
              newRow[i] = newWaktu;
           }
           if(h === "tipe" || h === "jenis") newRow[i] = "Kuis";
         }
         sheetNilai.appendRow(newRow);
         return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Nilai baru berhasil ditambahkan" }))
                              .setMimeType(ContentService.MimeType.JSON);
       }
    }
    if (action === "updateNilaiBulk") {
         var sheetNilai = ss.getSheetByName("listNilai");
         if (!sheetNilai) throw new Error("Sheet listNilai tidak ditemukan");
         
         var bulkData = postData.data; // Array of objects
         if(!bulkData || !Array.isArray(bulkData)) throw new Error("Format data bulk tidak valid");
         
         var sheetData = sheetNilai.getDataRange().getValues();
         var headers = sheetData[0];
         var userCol = -1, materiCol = -1, nilaiCol = -1, waktuCol = -1, namaCol = -1, kelasCol = -1;
         
         for(var i = 0; i < headers.length; i++){
            var h = headers[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
            if(h === "username" || h === "nisn") userCol = i;
            if(h === "id_materi" || h === "idmateri" || h === "modul") materiCol = i;
            if(h === "nilai" || h === "skor") nilaiCol = i;
            if(h === "waktu" || h === "durasi") waktuCol = i;
            if(h === "nama" || h === "namasiswa" || h === "namalengkap") namaCol = i;
            if(h === "kelas" || h === "tingkat") kelasCol = i;
         }
         
         if(userCol === -1 || materiCol === -1 || nilaiCol === -1) {
           throw new Error("Struktur kolom listNilai tidak lengkap (butuh username, id_materi, nilai)");
         }
         
         var newRows = [];
         
         for (var d = 0; d < bulkData.length; d++) {
             var item = bulkData[d];
             var targetUser = item.username;
             var targetMateri = item.id_materi;
             var newNilai = item.nilai;
             var newWaktu = item.waktu;
             var targetNama = item.nama || "";
             var targetKelas = item.kelas || "";
             
             var updated = false;
             for(var j = sheetData.length - 1; j > 0; j--) {
                var rowUsername = sheetData[j][userCol].toString().trim();
                var rowMateri = sheetData[j][materiCol].toString().trim();
                
                if(rowUsername === targetUser && rowMateri === targetMateri) {
                   sheetNilai.getRange(j + 1, nilaiCol + 1).setValue(newNilai);
                   if (waktuCol !== -1 && newWaktu !== undefined) {
                       sheetNilai.getRange(j + 1, waktuCol + 1).setValue(newWaktu);
                   }
                   if (namaCol !== -1 && targetNama !== "") {
                       sheetNilai.getRange(j + 1, namaCol + 1).setValue(targetNama);
                   }
                   if (kelasCol !== -1 && targetKelas !== "") {
                       sheetNilai.getRange(j + 1, kelasCol + 1).setValue(targetKelas);
                   }
                   
                   var now = new Date();
                   var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                   for (var colIdx = 0; colIdx < headers.length; colIdx++) {
                       var hT = headers[colIdx].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                       if ((hT === "timestamp" || hT === "waktusubmit" || hT === "tanggal" || hT === "waktu") && colIdx !== waktuCol) {
                           sheetNilai.getRange(j + 1, colIdx + 1).setValue(timeStr);
                       }
                   }
                   
                   updated = true;
                   break;
                }
             }
             
             if (!updated) {
                 var newRow = new Array(headers.length).fill("");
                 newRow[userCol] = targetUser;
                 newRow[materiCol] = targetMateri;
                 newRow[nilaiCol] = newNilai;
                 if(namaCol !== -1) newRow[namaCol] = targetNama;
                 if(kelasCol !== -1) newRow[kelasCol] = targetKelas;
                 
                 for(var i = 0; i < headers.length; i++){
                   var h = headers[i].toString().toLowerCase();
                   if(h === "waktu" || h === "timestamp") {
                      var now = new Date();
                      newRow[i] = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                   }
                   if (waktuCol !== -1 && i === waktuCol && newWaktu !== undefined) {
                      newRow[i] = newWaktu;
                   }
                   if(h === "tipe" || h === "jenis") newRow[i] = "Kuis";
                 }
                 newRows.push(newRow);
             }
         }
         
         if (newRows.length > 0) {
             var lastRow = sheetNilai.getLastRow();
             sheetNilai.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
         }
         
         return ContentService.createTextOutput(JSON.stringify({ status: "success", message: bulkData.length + " data nilai berhasil disimpan" }))
                              .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION 7: MARK TUGAS READ (ADMIN PANEL)
    // ===============================================
    if (action === "markTugasRead") {
      var sheetTugas = ss.getSheetByName("logTugas");
      if (!sheetTugas) throw new Error("Sheet logTugas tidak ditemukan");
      
      var rowNum = parseInt(postData.rowNumber);
      if (!rowNum || rowNum < 2) throw new Error("Nomor baris tidak valid");
      
      var newStatus = postData.status; // "Sudah" atau "Belum"
      
      var headers = sheetTugas.getRange(1, 1, 1, sheetTugas.getLastColumn()).getValues()[0];
      var statusCol = -1;
      
      for (var i = 0; i < headers.length; i++) {
         if (headers[i].toString().toLowerCase() === "status_periksa") {
            statusCol = i + 1;
            break;
         }
      }
      
      // Jika kolom status_periksa belum ada, buat otomatis di kanan
      if (statusCol === -1) {
         statusCol = headers.length + 1;
         sheetTugas.getRange(1, statusCol).setValue("status_periksa");
      }
      
      sheetTugas.getRange(rowNum, statusCol).setValue(newStatus);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status tugas diperbarui" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION: UPDATE MATERI ID & SYNC SOAL
    // ===============================================
    if (action === "updateMateriId") {
      var rowNum = parseInt(postData.rowNumber);
      var oldId = (postData.oldId || "").toString().trim();
      var newId = (postData.newId || "").toString().trim();
      var tingkatMateri = (postData.tingkat || "X").toString().trim();
      var dataToSave = JSON.parse(postData.data);
      var sheetMateri = ss.getSheetByName("materi");
      var sheetSoal = ss.getSheetByName("soal");
      
      if (newId !== "") {
        var allMateri = sheetMateri.getDataRange().getValues();
        var headersM = allMateri[0];
        var idxIdM = -1, idxTingkatM = -1;
        for(var i=0; i<headersM.length; i++){
          var h = headersM[i] ? headersM[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '') : "";
          if(h === "idmateri" || h === "id_materi" || h === "materiid" || h === "id") idxIdM = i;
          if(h === "tingkat" || h === "kelas") idxTingkatM = i;
        }
        for (var i = 1; i < allMateri.length; i++) {
          if (i + 1 === rowNum) continue;
          var rowIdRaw = allMateri[i][idxIdM];
          var rowId = rowIdRaw ? rowIdRaw.toString().trim() : "";
          
          var rowTingkatRaw = allMateri[i][idxTingkatM];
          var rowTingkat = rowTingkatRaw ? rowTingkatRaw.toString().trim() : "";
          
          function getTingkatPure(t) {
            if(!t) return "X";
            t = t.toString().toUpperCase();
            if(t.includes("XII")) return "XII";
            if(t.includes("XI")) return "XI";
            if(t.includes("X")) return "X";
            return t;
          }
          if (rowId === newId && getTingkatPure(rowTingkat) === getTingkatPure(tingkatMateri)) {
            return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ID Materi '" + newId + "' sudah digunakan di tingkat " + rowTingkat }))
                                 .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var headers = sheetMateri.getDataRange().getValues()[0];
      var updateRow = new Array(headers.length).fill("");
      var existingData = sheetMateri.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      for (var i = 0; i < headers.length; i++) {
        var h = headers[i].toString().trim();
        var hLower = h.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (dataToSave[h] !== undefined) updateRow[i] = dataToSave[h];
        else if (dataToSave[hLower] !== undefined) updateRow[i] = dataToSave[hLower];
        else if (hLower === "idmateri" || hLower === "id_materi") updateRow[i] = newId;
        else updateRow[i] = existingData[i];
      }
      sheetMateri.getRange(rowNum, 1, 1, headers.length).setValues([updateRow]);

      // 3. Sinkronisasi ke Sheet Soal
      var syncCount = 0;
      var foundCount = 0;
      if (oldId !== "" && oldId !== newId && sheetSoal) {
        var soalData = sheetSoal.getDataRange().getValues();
        var headersSoal = soalData[0];
        var idxIdSoal = -1;
        
        // Cari kolom ID di Soal dengan sangat teliti
        for(var k=0; k<headersSoal.length; k++){
           var hs = headersSoal[k] ? headersSoal[k].toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim() : "";
           if(hs === "idmateri" || hs === "id_materi" || hs === "materiid" || hs === "modul" || hs === "id") { 
             idxIdSoal = k; 
             break; 
           }
        }

        if (idxIdSoal !== -1) {
          // Ambil seluruh kolom ID soal untuk diproses di memori (lebih cepat)
          var rangeSoal = sheetSoal.getRange(1, idxIdSoal + 1, sheetSoal.getLastRow(), 1);
          var colValues = rangeSoal.getValues();
          
          for (var j = 1; j < colValues.length; j++) {
            var valSoal = colValues[j][0] ? colValues[j][0].toString().trim() : "";
            if (valSoal == oldId) { // Gunakan == untuk pencocokan angka/teks yang longgar
              colValues[j][0] = newId;
              syncCount++;
            }
          }
          
          // Simpan kembali sekaligus jika ada yang berubah
          if (syncCount > 0) {
            rangeSoal.setValues(colValues);
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "Materi diperbarui. " + syncCount + " baris di sheet soal disinkronkan." 
      })).setMimeType(ContentService.MimeType.JSON);
    }
    
    // ========================================================
    // ACTION: BATCH CREATE EXAM FROM SELECTED QUESTIONS
    // ========================================================
    if (action === "createExamFromSelectedQuestions") {
      var sheetMateri = ss.getSheetByName("materi");
      var sheetSoal = ss.getSheetByName("soal");
      if (!sheetMateri || !sheetSoal) throw new Error("Sheet kuis/materi tidak ditemukan");
      
      // 1. Tambahkan materi kuis baru
      var headersM = sheetMateri.getRange(1, 1, 1, sheetMateri.getLastColumn()).getValues()[0];
      var newRowM = new Array(headersM.length).fill("");
      for (var i = 0; i < headersM.length; i++) {
        var h = headersM[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === "idmateri" || h === "id_materi" || h === "materiid" || h === "id") newRowM[i] = postData.newExamId;
        else if (h === "judul" || h === "materi" || h === "namamateri") newRowM[i] = postData.newExamTitle;
        else if (h === "mapel" || h === "matapelajaran") newRowM[i] = postData.newExamMapel;
        else if (h === "tingkat" || h === "kelas") newRowM[i] = postData.newExamKelas;
        else if (h === "tipe") newRowM[i] = "Kuis";
        else if (h === "status") newRowM[i] = "Aktif";
      }
      sheetMateri.appendRow(newRowM);
      
      // 2. Salin soal terpilih dengan ID Materi Ujian baru
      var headersS = sheetSoal.getRange(1, 1, 1, sheetSoal.getLastColumn()).getValues()[0];
      var idxIdMateriS = -1;
      var idxIdSoalS = -1;
      
      for (var i = 0; i < headersS.length; i++) {
        var h = headersS[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === "idmateri" || h === "id_materi" || h === "materiid") idxIdMateriS = i;
        if (h === "id" || h === "idsoal") idxIdSoalS = i;
      }
      
      var selectedRows = postData.selectedQuestionRows; // Array of row numbers
      var newSoalRows = [];
      
      for (var r = 0; r < selectedRows.length; r++) {
        var rowNum = parseInt(selectedRows[r]);
        var rowData = sheetSoal.getRange(rowNum, 1, 1, headersS.length).getValues()[0];
        
        // Ubah ID Materi ke ID Ujian Baru
        if (idxIdMateriS !== -1) {
          rowData[idxIdMateriS] = postData.newExamId;
        }
        // Perbarui ID Soal agar unik (misal: UA-XI-1_1, UA-XI-1_2)
        if (idxIdSoalS !== -1) {
          rowData[idxIdSoalS] = postData.newExamId + "_" + (r + 1);
        }
        
        newSoalRows.push(rowData);
      }
      
      if (newSoalRows.length > 0) {
        var lastRowS = sheetSoal.getLastRow();
        sheetSoal.getRange(lastRowS + 1, 1, newSoalRows.length, headersS.length).setValues(newSoalRows);
      }
      
      return ContentService.createTextOutput(JSON.stringify({ 
        status: "success", 
        message: "Ujian Akhir berhasil dibuat dengan " + newSoalRows.length + " soal." 
      })).setMimeType(ContentService.MimeType.JSON);
    }
 
    throw new Error("Aksi tidak dikenali: " + action);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.message || error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// FUNGSI UNTUK MENGAKTIFKAN OTORISASI DRIVE
// ============================================
function setupAwal() {
  // Jalankan fungsi ini SATU KALI dari editor Apps Script 
  // untuk memunculkan popup izin akses penuh ke Google Drive.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getRootFolder();
  var dummyFile = folder.createFile("dummy_auth_file.txt", "dummy", MimeType.PLAIN_TEXT);
  dummyFile.setTrashed(true);
  Logger.log("Otorisasi berhasil. Script kini memiliki akses penuh untuk menulis file ke Drive.");
}

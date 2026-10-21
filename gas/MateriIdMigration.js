function normalizedHeader_(value) {
  return value.toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function findIdColumns_(headers, includeGenericId) {
  var columns = [];
  headers.forEach(function(header, index) {
    var normalized = normalizedHeader_(header);
    if (["idmateri", "materiid", "modul"].indexOf(normalized) !== -1 || (includeGenericId && normalized === "id")) columns.push(index);
  });
  return columns;
}

function planMateriIdMigration_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var materiSheet = ss.getSheetByName("materi");
  if (!materiSheet) throw new Error("Tab materi tidak ditemukan.");
  var materiValues = materiSheet.getDataRange().getValues();
  if (!materiValues.length) throw new Error("Tab materi belum memiliki header.");

  var headers = materiValues[0];
  var idColumns = findIdColumns_(headers, true);
  var idColumn = headers.findIndex(function(header) { return header.toString().trim().toLowerCase() === "id_materi"; });
  var canonicalColumnMissing = idColumn === -1;
  var titleColumn = headers.findIndex(function(header) { return ["judul", "materi", "namamateri"].indexOf(normalizedHeader_(header)) !== -1; });
  var gradeColumn = headers.findIndex(function(header) { return ["kelas", "tingkat"].indexOf(normalizedHeader_(header)) !== -1; });
  if (idColumn === -1 && idColumns.length) idColumn = idColumns[0];

  var ambiguous = [];
  var materials = materiValues.slice(1).map(function(row, offset) {
    var id = "";
    var ids = [];
    for (var i = 0; i < idColumns.length; i++) {
      if (row[idColumns[i]] != null && row[idColumns[i]].toString().trim()) {
        ids.push(row[idColumns[i]].toString().trim());
      }
    }
    id = ids[0] || "";
    if (ids.some(function(value) { return value.toLowerCase() !== id.toLowerCase(); })) {
      ambiguous.push({type: "material_id_alias_mismatch", rowNumber: offset + 2});
    }
    return {
      rowNumber: offset + 2,
      id: id,
      title: titleColumn === -1 ? "" : (row[titleColumn] || "").toString().trim(),
      grade: gradeColumn === -1 ? "" : (row[gradeColumn] || "").toString().trim()
    };
  }).filter(function(item) { return item.id || item.title || item.grade; });

  var unresolved = [];
  var existingIds = {};
  materials.forEach(function(item) {
    if (!item.id) return;
    var key = item.id.toLowerCase();
    if (existingIds[key]) {
      ambiguous.push({type: "duplicate_id_materi", rows: [existingIds[key], item.rowNumber]});
    } else {
      existingIds[key] = item.rowNumber;
    }
  });

  var missingByTitleGrade = {};
  materials.forEach(function(item) {
    if (item.id) return;
    var key = item.title.toLowerCase() + "|" + item.grade.toLowerCase();
    if (!item.title || !item.grade) {
      ambiguous.push({type: "material_missing_id_or_match_fields", rowNumber: item.rowNumber});
      return;
    }
    if (!missingByTitleGrade[key]) missingByTitleGrade[key] = [];
    missingByTitleGrade[key].push(item);
  });

  var materialUpdates = [];
  if (canonicalColumnMissing) {
    materials.forEach(function(item) {
      if (item.id) materialUpdates.push({rowNumber: item.rowNumber, id_materi: item.id});
    });
  }
  Object.keys(missingByTitleGrade).forEach(function(key) {
    var matches = missingByTitleGrade[key];
    if (matches.length !== 1) {
      ambiguous.push({type: "duplicate_title_grade", rows: matches.map(function(item) { return item.rowNumber; })});
      return;
    }
    matches[0].proposedId = "MAT-" + Utilities.getUuid().replace(/-/g, "").slice(0, 16).toUpperCase();
    materialUpdates.push({rowNumber: matches[0].rowNumber, id_materi: matches[0].proposedId});
  });

  var plannedMaterials = materials.map(function(item) {
    var proposed = materialUpdates.find(function(update) { return update.rowNumber === item.rowNumber; });
    return {rowNumber: item.rowNumber, id: item.id || (proposed && proposed.id_materi), title: item.title, grade: item.grade};
  });
  var referenceUpdates = [];
  var referenceSnapshot = [];
  ["soal", "ListNilai", "logTugas", "sesiKuis"].forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    var values = sheet.getDataRange().getValues();
    var refHeaders = values[0];
    var refIdColumns = findIdColumns_(refHeaders);
    var refIdColumn = refHeaders.findIndex(function(header) { return header.toString().trim().toLowerCase() === "id_materi"; });
    if (refIdColumn === -1) refIdColumn = refIdColumns[0];
    var refTitleColumn = refHeaders.findIndex(function(header) { return ["materi", "judul", "judulmateri", "namamateri"].indexOf(normalizedHeader_(header)) !== -1; });
    var refGradeColumn = refHeaders.findIndex(function(header) { return ["kelas", "tingkat"].indexOf(normalizedHeader_(header)) !== -1; });
    if (refIdColumn === -1) {
      unresolved.push({sheet: name, type: "missing_id_column"});
      return;
    }
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      var currentId = "";
      var currentIds = [];
      for (var c = 0; c < refIdColumns.length; c++) {
        if (row[refIdColumns[c]] != null && row[refIdColumns[c]].toString().trim()) {
          currentIds.push(row[refIdColumns[c]].toString().trim());
        }
      }
      currentId = currentIds[0] || "";
      if (currentIds.some(function(value) { return value.toLowerCase() !== currentId.toLowerCase(); })) {
        ambiguous.push({sheet: name, rowNumber: r + 1, type: "reference_id_alias_mismatch"});
      }
      var title = refTitleColumn === -1 ? "" : (row[refTitleColumn] || "").toString().trim().toLowerCase();
      var grade = refGradeColumn === -1 ? "" : (row[refGradeColumn] || "").toString().trim().toLowerCase();
      referenceSnapshot.push([name, r + 1, currentId.toLowerCase(), title, grade]);
      if (currentId && existingIds[currentId.toLowerCase()]) continue;
      var matches = title ? plannedMaterials.filter(function(item) {
        return item.title.toLowerCase() === title && (!grade || item.grade.toLowerCase() === grade);
      }) : [];
      if (matches.length === 1 && matches[0].id) {
        referenceUpdates.push({sheet: name, rowNumber: r + 1, id_materi: matches[0].id});
      } else if (currentId || title || grade) {
        ambiguous.push({sheet: name, rowNumber: r + 1, type: matches.length > 1 ? "reference_matches_multiple_materials" : "reference_has_no_unique_material"});
      }
    }
  });

  var fingerprintSource = {
    materials: materials.map(function(item) { return [item.rowNumber, item.id, item.title, item.grade]; }),
    references: referenceSnapshot
  };
  var fingerprint = Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, JSON.stringify(fingerprintSource)));
  return {
    safeToApply: ambiguous.length === 0 && unresolved.length === 0,
    materialCount: materials.length,
    materialUpdates: materialUpdates,
    referenceUpdates: referenceUpdates,
    ambiguous: ambiguous,
    unresolved: unresolved,
    fingerprint: fingerprint,
    idColumn: idColumn,
    materialSheetId: materiSheet.getSheetId(),
    canonicalColumnMissing: canonicalColumnMissing
  };
}

function dryRunIdMateriMigration() {
  var plan = planMateriIdMigration_();
  PropertiesService.getScriptProperties().setProperty("ID_MATERI_MIGRATION_PLAN", JSON.stringify(plan));
  return plan;
}

function applyIdMateriMigration() {
  var saved = PropertiesService.getScriptProperties().getProperty("ID_MATERI_MIGRATION_PLAN");
  if (!saved) throw new Error("Jalankan dryRunIdMateriMigration terlebih dahulu.");
  var approvedPlan = JSON.parse(saved);
  if (!approvedPlan.safeToApply) throw new Error("Migrasi dibatalkan: dry-run menemukan pemetaan ambigu.");
  var currentPlan = planMateriIdMigration_();
  if (currentPlan.fingerprint !== approvedPlan.fingerprint || currentPlan.materialSheetId !== approvedPlan.materialSheetId) {
    throw new Error("Migrasi dibatalkan: data berubah sejak dry-run. Jalankan dry-run ulang.");
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var materia = ss.getSheetByName("materi");
  var headers = materia.getRange(1, 1, 1, materia.getLastColumn()).getValues()[0];
  var idColumn = headers.findIndex(function(header) { return header.toString().trim().toLowerCase() === "id_materi"; });
  if (idColumn === -1) {
    idColumn = headers.length;
    materia.getRange(1, idColumn + 1).setValue("id_materi");
  }
  approvedPlan.materialUpdates.forEach(function(update) {
    materia.getRange(update.rowNumber, idColumn + 1).setValue(update.id_materi);
  });
  approvedPlan.referenceUpdates.forEach(function(update) {
    var sheet = ss.getSheetByName(update.sheet);
    var refHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var refColumn = refHeaders.findIndex(function(header) { return header.toString().trim().toLowerCase() === "id_materi"; });
    if (refColumn === -1) refColumn = findIdColumns_(refHeaders)[0];
    sheet.getRange(update.rowNumber, refColumn + 1).setValue(update.id_materi);
  });
  PropertiesService.getScriptProperties().deleteProperty("ID_MATERI_MIGRATION_PLAN");
  return {status: "success", materialsUpdated: approvedPlan.materialUpdates.length, referencesUpdated: approvedPlan.referenceUpdates.length};
}

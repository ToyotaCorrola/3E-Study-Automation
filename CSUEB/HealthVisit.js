function transferHealthVisitData(){
  let scriptProperties = PropertiesService.getScriptProperties();
  // Master Sheet and Linking Destination details.
  var masterSheetId             = scriptProperties.getProperty("MasterSheetKEY");  
  var linkingDestinationSheetId = scriptProperties.getProperty("HealthVisitKEY");
  
  var masterSS = SpreadsheetApp.openById(masterSheetId);
  var sheetNames = {1: 'W1', 2: 'W2', 3: 'W3', 4: 'W4', 5: 'W5'};
  var masters = {};

  // Master headers: A: StudyID (0), D: School Email (3)
  var STUDYID_COL_MASTER = 0;
  var EMAIL_COL_MASTER   = 3;

  // Preload each master sheet’s data, backgrounds, and build map by StudyID and by Email.
  for (var w = 1; w <= 5; w++) {
    var sheet   = masterSS.getSheetByName(sheetNames[w]);
    var numRows = sheet.getLastRow();
    var numCols = sheet.getLastColumn();
    var range = sheet.getRange(1, 1, numRows, numCols);
    var data = range.getValues();
    var bg   = range.getBackgrounds();

    var mapByStudyId = {};
    var mapByEmail   = {};

    for (var i = 1; i < data.length; i++) { // skip header row
      var sid   = data[i][STUDYID_COL_MASTER];
      var email = data[i][EMAIL_COL_MASTER];
      if (sid)   mapByStudyId[String(sid).trim()] = i;
      if (email) mapByEmail[String(email).trim().toLowerCase()] = i;
    }
    masters[w] = { sheet: sheet, data: data, bg: bg, numCols: numCols, mapByStudyId: mapByStudyId, mapByEmail: mapByEmail };
  }
  
  var linkingSS    = SpreadsheetApp.openById(linkingDestinationSheetId);
  var linkingSheet = linkingSS.getSheetByName('Sheet1');
  var destData     = linkingSheet.getDataRange().getValues();

  // linkingSS headers:
  // A: StudyID (0), B: FirstName (1), C: (unused), D: VisitDate (3), E: Wave (4), F: Name/Email (5)

  // Process each destination row (skip header).
  for (var i = 1; i < destData.length; i++) {
    var row        = destData[i],
        studyIdRaw = row[0],
        firstName  = row[1],
        visitDate  = new Date(row[3]),
        waveStr    = row[4] != null ? row[4].toString().toLowerCase() : '',
        nameEmail  = row[5];

    if (isNaN(visitDate)) continue; // skip invalid dates

    var visitMonth = visitDate.getMonth() + 1;
    var visitYear  = visitDate.getFullYear();

    var wNum;
    if (waveStr.includes('1')) wNum = 1;
    else if (waveStr.includes('2')) wNum = 2;
    else if (waveStr.includes('3')) wNum = 3;
    else if (waveStr.includes('4')) wNum = 4;
    else if (waveStr.includes('5')) wNum = 5;
    else continue;
    
    var master = masters[wNum];

    // Resolve master index: StudyID first, then fallback to Email from linkingSS col F.
    var idx;
    var studyId = studyIdRaw != null ? String(studyIdRaw).trim() : '';
    if (studyId.length >= 1) {
      idx = master.mapByStudyId[studyId];
      if(idx == undefined) console.log(firstName + " " + studyId)
    }
    if (studyId.length === 0) {
      if (nameEmail) {
        idx = master.mapByEmail[nameEmail];
        console.log(master.data[idx] + " " + visitMonth + " " + visitYear);
      }
    }
    if (idx === undefined) continue; // no match found

    var mRow = master.data[idx];

    if (wNum === 1) {
      const indexes = {
        healthCompletedIdx : 12,
        healthMonthIdx     : 13,
        healthYearIdx      : 14,
        notesIdx           : 20
      };
      if (!mRow[indexes.healthCompletedIdx] || String(mRow[indexes.healthCompletedIdx]).toLowerCase() === "no") {
        mRow[indexes.healthCompletedIdx] = "YES"; master.bg[idx][indexes.healthCompletedIdx] = "red";
      }
      if (String(mRow[1]).toLowerCase() !== String(firstName).toLowerCase()) {
        appendToNotes(mRow, indexes.notesIdx, "Different first name in health visit");
      }
      if (!mRow[indexes.healthMonthIdx]) { mRow[indexes.healthMonthIdx] = visitMonth; master.bg[idx][indexes.healthMonthIdx] = "red"; }
      if (!mRow[indexes.healthYearIdx])  { mRow[indexes.healthYearIdx]  = visitYear;  master.bg[idx][indexes.healthYearIdx]  = "red"; }
    } 
    else if (wNum === 2) {
      const indexes = {
        healthCompletedIdx : 10,
        healthMonthIdx     : 11,
        healthYearIdx      : 12,
        notesIdx           : 18
      };
      if (!mRow[indexes.healthCompletedIdx] || String(mRow[indexes.healthCompletedIdx]).toLowerCase() === "no") {
        mRow[indexes.healthCompletedIdx] = "YES"; master.bg[idx][indexes.healthCompletedIdx] = "red";
      }
      if (String(mRow[1]).toLowerCase() !== String(firstName).toLowerCase()) {
        appendToNotes(mRow, indexes.notesIdx, "Different first name in health visit");
      }
      if (!mRow[indexes.healthMonthIdx]) { mRow[indexes.healthMonthIdx] = visitMonth; master.bg[idx][indexes.healthMonthIdx] = "red"; }
      if (!mRow[indexes.healthYearIdx])  { mRow[indexes.healthYearIdx]  = visitYear;  master.bg[idx][indexes.healthYearIdx]  = "red"; }
    } 
    else if (wNum === 3) {
      const indexes = {
        healthCompletedIdx : 10,
        healthMonthIdx     : 11,
        healthYearIdx      : 12,
        notesIdx           : 19
      };
      if (!mRow[indexes.healthCompletedIdx] || String(mRow[indexes.healthCompletedIdx]).toLowerCase() === "no") {
        mRow[indexes.healthCompletedIdx] = "YES"; master.bg[idx][indexes.healthCompletedIdx] = "red";
      }
      if (String(mRow[1]).toLowerCase() !== String(firstName).toLowerCase()) {
        appendToNotes(mRow, indexes.notesIdx, "Different first name in health visit");
      }
      if (!mRow[indexes.healthMonthIdx]) { mRow[indexes.healthMonthIdx] = visitMonth; master.bg[idx][indexes.healthMonthIdx] = "red"; }
      if (!mRow[indexes.healthYearIdx])  { mRow[indexes.healthYearIdx]  = visitYear;  master.bg[idx][indexes.healthYearIdx]  = "red"; }
    } 
    else if (wNum === 4) {
      const indexes = {
        healthCompletedIdx : 10,
        healthMonthIdx     : 11,
        healthYearIdx      : 12,
        notesIdx           : 20
      };
      if (!mRow[indexes.healthCompletedIdx] || String(mRow[indexes.healthCompletedIdx]).toLowerCase() === "no") {
        mRow[indexes.healthCompletedIdx] = "YES"; master.bg[idx][indexes.healthCompletedIdx] = "red";
      }
      if (String(mRow[1]).toLowerCase() !== String(firstName).toLowerCase()) {
        appendToNotes(mRow, indexes.notesIdx, "Different first name in health visit");
      }
      if (!mRow[indexes.healthMonthIdx]) { mRow[indexes.healthMonthIdx] = visitMonth; master.bg[idx][indexes.healthMonthIdx] = "red"; }
      if (!mRow[indexes.healthYearIdx])  { mRow[indexes.healthYearIdx]  = visitYear;  master.bg[idx][indexes.healthYearIdx]  = "red"; }
    }
    else if (wNum === 5) {
      const indexes = {
        healthCompletedIdx : 10,
        healthMonthIdx     : 11,
        healthYearIdx      : 12,
        notesIdx           : 21
      };
      if (!mRow[indexes.healthCompletedIdx] || String(mRow[indexes.healthCompletedIdx]).toLowerCase() === "no") {
        mRow[indexes.healthCompletedIdx] = "YES"; master.bg[idx][indexes.healthCompletedIdx] = "red";
      }
      if (String(mRow[1]).toLowerCase() !== String(firstName).toLowerCase()) {
        appendToNotes(mRow, indexes.notesIdx, "Different first name in health visit");
      }
      if (!mRow[indexes.healthMonthIdx]) { mRow[indexes.healthMonthIdx] = visitMonth; master.bg[idx][indexes.healthMonthIdx] = "red"; }
      if (!mRow[indexes.healthYearIdx])  { mRow[indexes.healthYearIdx]  = visitYear;  master.bg[idx][indexes.healthYearIdx]  = "red"; }
    }
  }
  
  // Write back updates for all master sheets.
  for (var w = 1; w <= 5; w++) {
    var master = masters[w];
    master.sheet.getRange(1, 1, master.data.length, master.numCols).setValues(master.data);
    master.sheet.getRange(1, 1, master.bg.length, master.bg[0].length).setBackgrounds(master.bg);
  }
}

function appendToNotes(row, index, note) {
  if (!row[index]) row[index] = note;
  else if (row[index].toString().indexOf(note) === -1) row[index] += "\n" + note;
}

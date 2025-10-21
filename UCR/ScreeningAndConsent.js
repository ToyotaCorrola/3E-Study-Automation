function transferDataToMasterSheet() {
  let scriptProperties = PropertiesService.getScriptProperties();
  // Master Sheet details.
  const masterSheetId = scriptProperties.getProperty("MasterSheetKEY");;
  
  // Define configuration for each master sheet.
  // Note: updateMap maps field names to the 1-indexed column to update.
  var sheetConfigs = {
    "W1": { notesCol: 20, updateMap: { firstName:2, lastName:3, ucrEmail:4, personalEmail:5, phone:6, classStatus:7, source:8 } },
    "W2": { notesCol: 18, updateMap: { firstName:2, lastName:3, ucrEmail:4, personalEmail:5, phone:6 } },
    "W3": { notesCol: 19, updateMap: { firstName:2, lastName:3, ucrEmail:4, personalEmail:5, phone:6 } },
    "W4": { notesCol: 20, updateMap: { firstName:2, lastName:3, ucrEmail:4, personalEmail:5, phone:6 } },
    "W5": { notesCol: 21, updateMap: { firstName:2, lastName:3, ucrEmail:4, personalEmail:5, phone:6 } }
  };

  // Open the master spreadsheet and load each sheet’s data and backgrounds into memory.
  var masterSpreadsheet = SpreadsheetApp.openById(masterSheetId);
  var masterSheets = {};
  for (var sheetName in sheetConfigs) {
    var sheet = masterSpreadsheet.getSheetByName(sheetName);
    var numCols = sheet.getLastColumn();
    var dataRange = sheet.getDataRange();
    var data = dataRange.getValues();
    var backgrounds = sheet.getRange(1, 1, data.length, numCols).getBackgrounds();
    // Build a mapping from StudyID to row index (0-indexed; row0 assumed header)
    var mapping = {};
    for (var i = 1; i < data.length; i++) {
      var studyId = data[i][0];
      if (studyId) {
        mapping[studyId] = i;
      }
    }
    masterSheets[sheetName] = {
      sheet: sheet,
      data: data,
      backgrounds: backgrounds,
      mapping: mapping,
      numCols: numCols,
      config: sheetConfigs[sheetName]
    };
  }
  
  // Data sources (you may leave these as-is)
  var sources = [
    { fileId: scriptProperties.getProperty("ScreeningAndConsent-Email-KEY"), 
      sheetName: 'Email', 
      source: 'EMAIL' },
    { fileId: scriptProperties.getProperty("ScreeningAndConsent-IGWEB-KEY"),  
      sheetName: 'Sheet1', 
      source: 'IG/WEB' },
    { fileId: scriptProperties.getProperty("ScreeningAndConsent-REFERRAL-KEY"), 
      sheetName: 'Sheet1', 
      source: 'REFERRAL' },
    { fileId: scriptProperties.getProperty("ScreeningAndConsent-FLYER-KEY"), 
      sheetName: 'Sheet1', 
      source: 'FLYER' }
  ];
  
  // Process each source file
  sources.forEach(function(sourceInfo) {
    var ss = SpreadsheetApp.openById(sourceInfo.fileId);
    var srcSheet = ss.getSheetByName(sourceInfo.sheetName);
    var srcData = srcSheet.getDataRange().getValues();
    for (var i = 1; i < srcData.length; i++) {
      var row = srcData[i];
      if (!row[1]) continue; // Skip if first name is empty
      
      // Extract values from the source row
      var studyId = row[0];
      var firstName = row[1];
      var lastName = row[2];
      var ucrEmail = row[3];
      var personalEmail = row[4];
      var phone = formatPhoneNumber(row[5]);
      var classStatus = mapClassStatus(row[6]);
      var srcName = sourceInfo.source;
      var referralNote = "";
      if (srcName === "REFERRAL" && row[7]) {
        referralNote = "referred by " + row[7];
      }
      
      // For each master sheet, update or append the row.
      updateMasterRow(masterSheets["W1"], studyId, {
        firstName: firstName,
        lastName: lastName,
        ucrEmail: ucrEmail,
        personalEmail: personalEmail,
        phone: phone,
        classStatus: classStatus,
        source: srcName
      }, referralNote);
      
      updateMasterRow(masterSheets["W2"], studyId, {
        firstName: firstName,
        lastName: lastName,
        ucrEmail: ucrEmail,
        personalEmail: personalEmail,
        phone: phone
      }, referralNote);
      
      updateMasterRow(masterSheets["W3"], studyId, {
        firstName: firstName,
        lastName: lastName,
        ucrEmail: ucrEmail,
        personalEmail: personalEmail,
        phone: phone
      }, referralNote);
      
      updateMasterRow(masterSheets["W4"], studyId, {
        firstName: firstName,
        lastName: lastName,
        ucrEmail: ucrEmail,
        personalEmail: personalEmail,
        phone: phone
      }, referralNote);

      updateMasterRow(masterSheets["W5"], studyId, {
        firstName: firstName,
        lastName: lastName,
        ucrEmail: ucrEmail,
        personalEmail: personalEmail,
        phone: phone
      }, referralNote);
    }
  });
  
  // Write all updates back in bulk.
  for (var sheetName in masterSheets) {
    var m = masterSheets[sheetName];
    m.sheet.getRange(1, 1, m.data.length, m.numCols).setValues(m.data);
    m.sheet.getRange(1, 1, m.backgrounds.length, m.numCols).setBackgrounds(m.backgrounds);
  }
}

// This helper function updates (or appends) a StudyID row in a master sheet’s in‑memory arrays.
function updateMasterRow(masterObj, studyId, newValues, additionalNote) {
  var mapping = masterObj.mapping;
  var data = masterObj.data;
  var backgrounds = masterObj.backgrounds;
  var config = masterObj.config;
  var notesCol = config.notesCol; // 1-indexed
  var updateMap = config.updateMap; // field name -> 1-indexed column
  
  if (mapping.hasOwnProperty(studyId)) {
    var rowIndex = mapping[studyId];
    // Update each field if it has changed.
    for (var field in updateMap) {
      var col = updateMap[field] - 1;
      var oldVal = data[rowIndex][col];
      var newVal = newValues[field];
      if (newVal !== undefined && newVal != oldVal) {
        data[rowIndex][col] = newVal;
        backgrounds[rowIndex][col] = "red";
        // Append change note if not already present
        var noteCell = data[rowIndex][notesCol] || "";
        var changeNote = "New " + field + ": " + newVal;
        if (noteCell.indexOf(changeNote) === -1) {
          data[rowIndex][notesCol] = noteCell ? noteCell + "\n" + changeNote : changeNote;
          backgrounds[rowIndex][notesCol] = "red";
        }
      }
    }
    // Also add any additional (e.g. referral) note.
    if (additionalNote) {
      var noteCell = data[rowIndex][notesCol] || "";
      if (noteCell.indexOf(additionalNote) === -1) {
        data[rowIndex][notesCol] = noteCell ? noteCell + "\n" + additionalNote : additionalNote;
        backgrounds[rowIndex][notesCol] = "red";
      }
    }
  } else {
    // If the row does not exist, create a new row.
    var newRow = new Array(masterObj.numCols).fill("");
    newRow[0] = studyId;
    for (var field in updateMap) {
      var col = updateMap[field] - 1;
      if (newValues[field] !== undefined) {
        newRow[col] = newValues[field];
      }
    }
    // Set the notes cell.
    newRow[notesCol] = additionalNote || "";
    data.push(newRow);
    // Update the mapping with the new row’s index.
    mapping[studyId] = data.length - 1;
    // Build a corresponding backgrounds row: all nonempty cells are red.
    var newBgRow = [];
    for (var i = 0; i < masterObj.numCols; i++) {
      newBgRow.push(newRow[i] !== "" ? "red" : "");
    }
    backgrounds.push(newBgRow);
  }
}

// Utility: Map class status as before.
function mapClassStatus(status) {
  switch(status) {
    case "Freshman":
      return "Frosh";
    case "Sophomore":
      return "Sophomore";
    case "Yes":
      return "First-year transfer";
    default:
      return status;
  }
}

// Utility: Format phone numbers.
function formatPhoneNumber(phone) {
  phone = phone.toString();
  return phone.startsWith("1") ? phone : "1" + phone;
}

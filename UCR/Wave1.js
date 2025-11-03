function transferWave1SurveyData() {
  let scriptProperties = PropertiesService.getScriptProperties();
  // Master Sheet and Linking Destination details.
  var masterSheetId               = scriptProperties.getProperty("MasterSheetKEY"); 
  var linkingDestinationSheetId   = scriptProperties.getProperty("Wave1Key");
  var w1SheetName                 = 'W1';
  var linkingDestinationSheetName = 'Sheet1';
  
  var masterSS = SpreadsheetApp.openById(masterSheetId);
  var w1Sheet = masterSS.getSheetByName(w1SheetName);
  var destSS = SpreadsheetApp.openById(linkingDestinationSheetId);
  var destSheet = destSS.getSheetByName(linkingDestinationSheetName);
  
  var masterRange = w1Sheet.getDataRange();
  var masterData = masterRange.getValues();
  var masterBg = masterRange.getBackgrounds();
  var destData = destSheet.getDataRange().getValues();
  

  var schoolEmailMap = {};
  var personalEmailMap = {};
  var duplicateStudyIds = [];
  var notesIdx = 20; 



  for (var j = 1; j < masterData.length; j++) {
    var mRow = masterData[j];
    var studyId = mRow[0];
    var schoolEmail = mRow[3];   // Column D (index 3)
    var personalEmail = mRow[4]; // Column E (index 4)
    
    // Check School Email duplicates
    if (studyId && schoolEmail && schoolEmail.toString().trim() !== "") {
      var emailLower = schoolEmail.toString().toLowerCase().trim();
      
      if (schoolEmailMap[emailLower]) {
        // Found a duplicate - this is the NEWER one, keep this
        var previousDuplicates = schoolEmailMap[emailLower].allOldIds || [];
        var allOldIds = previousDuplicates.concat([schoolEmailMap[emailLower].studyId]);
        var oldRow = schoolEmailMap[emailLower].row - 1; // Convert to 0-indexed
        var oldStudyId = schoolEmailMap[emailLower].studyId;
        
        
        // Mark the most recent OLD row yellow and update its Study ID
        masterBg[oldRow][0] = "yellow";
        appendToNotes(masterData[oldRow], notesIdx, "Old Study ID: " + oldStudyId + " (replaced by " + studyId + " due to duplicate email)");
        masterData[oldRow][0] = studyId; // Replace old Study ID with new one
        
        // Add note to the NEW/CURRENT row about ALL old duplicates
        appendToNotes(mRow, notesIdx, "Previous duplicate Study IDs: " + allOldIds.join(", "));
        
        duplicateStudyIds.push(oldStudyId);
        
        // Update the map to track this new one as the latest, along with all previous old IDs
        schoolEmailMap[emailLower] = {
          studyId: studyId,
          row: j + 1,
          allOldIds: allOldIds  // Keep track of all previous duplicates
        };
        
      } else {
        schoolEmailMap[emailLower] = {
          studyId: studyId,
          row: j + 1,
          allOldIds: []  // No previous duplicates yet
        };
      }
    }
    
    // Check Personal Email duplicates
    if (studyId && personalEmail && personalEmail.toString().trim() !== "") {
      var emailLower = personalEmail.toString().toLowerCase().trim();
      
      if (personalEmailMap[emailLower]) {
        // Found a duplicate - this is the NEWER one, keep this
        var previousDuplicates = personalEmailMap[emailLower].allOldIds || [];
        var allOldIds = previousDuplicates.concat([personalEmailMap[emailLower].studyId]);
        var oldRow = personalEmailMap[emailLower].row - 1; // Convert to 0-indexed
        var oldStudyId = personalEmailMap[emailLower].studyId;
        
        
        // Mark the most recent OLD row yellow and update its Study ID
        masterBg[oldRow][0] = "yellow";
        appendToNotes(masterData[oldRow], notesIdx, "Old Study ID: " + oldStudyId + " (replaced by " + studyId + " due to duplicate email)");
        masterData[oldRow][0] = studyId; // Replace old Study ID with new one
        
        // Add note to the NEW/CURRENT row about ALL old duplicates
        appendToNotes(mRow, notesIdx, "Previous duplicate Study IDs: " + allOldIds.join(", "));
        
        if (duplicateStudyIds.indexOf(oldStudyId) === -1) {
          duplicateStudyIds.push(oldStudyId);
        }
        
        // Update the map to track this new one as the latest, along with all previous old IDs
        personalEmailMap[emailLower] = {
          studyId: studyId,
          row: j + 1,
          allOldIds: allOldIds  // Keep track of all previous duplicates
        };
        
      } else {
        personalEmailMap[emailLower] = {
          studyId: studyId,
          row: j + 1,
          allOldIds: []  // No previous duplicates yet
        };
      }
    }
  }

  Logger.log("\nSUMMARY:");
  Logger.log("Total unique school emails: " + Object.keys(schoolEmailMap).length);
  Logger.log("Total unique personal emails: " + Object.keys(personalEmailMap).length);
  Logger.log("Old Study IDs replaced: " + duplicateStudyIds.join(", "));
  Logger.log("Total duplicates found: " + duplicateStudyIds.length);

  // Build destination map keyed by Study ID from column A.
  var destMap = {};
  for (var i = 1; i < destData.length; i++) {
    var row = destData[i];
    if (row[0]) destMap[row[0]] = row;
  }
  

  var surveyCompletedIdx          = 8;
  var healthScheduledIdx          = 11;
  var healthCompletedIdx          = 12;
  var interestedSubstudyIdx       = 16;
  var preferredContactMethodIdx   = 19;
  var notesIdx                    = 20;
  for (var j = 1; j < masterData.length; j++) {
    var mRow = masterData[j],
        studyId = mRow[0];
    if (!destMap[studyId]) {
      //By Default, the "Survey Completed" column is set to value "NO"
      if (!mRow[surveyCompletedIdx]) { mRow[surveyCompletedIdx] = "NO"; masterBg[j][surveyCompletedIdx] = "red"; }
/*
      //By Default, the "Health Visit Scheduled" column is set to value "NO" IF AND ONLY IF (IFF) the "Survey Completed" column is set to "YES"
      if (!mRow[healthScheduledIdx]) { 
        if (mRow[surveyCompletedIdx] === "YES"){
          mRow[healthScheduledIdx] = "NO"; masterBg[j][healthScheduledIdx] = "red"; 
        }
      }
      
      //By Default, the "Interested in Substudy" column is set to value "NO" IF AND ONLY IF (IFF) the "Survey Completed" column is set to "YES"
      if (!mRow[interestedSubstudyIdx]) {
        if (mRow[surveyCompletedIdx] === "YES"){
          mRow[interestedSubstudyIdx] = "NO"; masterBg[j][interestedSubstudyIdx] = "red"; 
        }
      }
      
      if (!mRow[healthCompletedIdx]) { 
        if (mRow[healthScheduledIdx] === "YES") { 
          mRow[healthCompletedIdx] = "NO"; masterBg[j][healthCompletedIdx] = "red"; 
        }
      }
*/
    } else {
      var surveyMonthIdx  = 9;
      var surveyYearIdx   = 10;
      var dRow = destMap[studyId];
      if (!mRow[surveyCompletedIdx] || String(mRow[surveyCompletedIdx]).toLowerCase() === "no") { 
        mRow[surveyCompletedIdx] = "YES"; 
        masterBg[j][surveyCompletedIdx] = "red"; 

        //By Default, the "Health Visit Scheduled" column is set to value "NO" IF AND ONLY IF (IFF) the "Survey Completed" column is set to "YES"
        mRow[healthScheduledIdx] = "NO"; 
        masterBg[j][healthScheduledIdx] = "red"; 

        //By Default, the "Interested in Substudy" column is set to value "NO" IF AND ONLY IF (IFF) the "Survey Completed" column is set to "YES"
        mRow[interestedSubstudyIdx] = "NO"; 
        masterBg[j][interestedSubstudyIdx] = "red"; 
      }

      //HEALTH VISIT DATE
      if ((!mRow[surveyMonthIdx] || !mRow[surveyYearIdx]) && dRow[surveyMonthIdx]) {
        var visitDate = new Date(dRow[9]);
        mRow[surveyMonthIdx] = visitDate.getMonth() + 1;
        mRow[surveyYearIdx] = visitDate.getFullYear();
        masterBg[j][surveyMonthIdx] = "red"; masterBg[j][surveyYearIdx] = "red";
      }

      if ((!mRow[healthScheduledIdx] || String(mRow[healthScheduledIdx]).toLowerCase() === "no") && (dRow[2].toLowerCase() === "i scheduled my appointment" || dRow[2].toLowerCase() === "i have already completed my health visit")) {
        mRow[healthScheduledIdx] = "YES"; 
        masterBg[j][healthScheduledIdx] = "red";
        if(String(mRow[healthCompletedIdx].toLowerCase() !== "yes")){
          mRow[healthCompletedIdx] = "NO"; 
          masterBg[j][healthCompletedIdx] = "red"; 
        }
      }
      //INTERESTED IN SUBSTUDY
      if (mRow[interestedSubstudyIdx] !== dRow[3].toUpperCase()) { mRow[interestedSubstudyIdx] = dRow[3].toUpperCase(); masterBg[j][interestedSubstudyIdx] = "red"; }
      
      var PCM = dRow[4];
      if (!["personal email", "school email", "text message", "phone call"].includes(PCM.toLowerCase())) {
        mRow[preferredContactMethodIdx] = "ERROR. CHECK NOTES";
        appendToNotes(mRow, notesIdx, "Preferred Contact Method listed as " + PCM);
      } else if (String(mRow[preferredContactMethodIdx]).toLowerCase() !== PCM.toLowerCase()) {
        mRow[preferredContactMethodIdx] = PCM; masterBg[j][preferredContactMethodIdx] = "red";
      }
      if (PCM.toLowerCase() === "school email" && mRow[3] !== dRow[5]) {
        appendToNotes(mRow, notesIdx, "different contact: " + dRow[5]);
      } else if (PCM.toLowerCase() === "personal email" && mRow[4] !== dRow[6]) {
        appendToNotes(mRow, notesIdx, "different contact: " + dRow[6]);
      } else if (PCM.toLowerCase() === "text message") {
        var fmtText = formatPhoneNumber(dRow[7]);
        if (mRow[5] !== fmtText) appendToNotes(mRow, notesIdx, "different contact: " + fmtText);
      } else if (PCM.toLowerCase() === "phone call") {
        var fmtCall = formatPhoneNumber(dRow[8]);
        if (mRow[5] !== fmtCall) appendToNotes(mRow, notesIdx, "different contact: " + fmtCall);
      }
    }
  }
  
  w1Sheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
  w1Sheet.getRange(1, 1, masterBg.length, masterBg[0].length).setBackgrounds(masterBg);
}

function appendToNotes(row, index, note) {
  if (!row[index]) row[index] = note;
  else if (row[index].indexOf(note) === -1) row[index] += "\n" + note;
}

function formatPhoneNumber(phone) {
  phone = phone.toString();
  return phone.startsWith("1") ? phone : "1" + phone;
}

function cleanupNotes() {
  var masterSheetId = '1rLIUGPVviSjAP9P9PwvjDVWyEwfyUxkH3kgN0sn7xkI';
  var w1SheetName = 'W1';
  var notesIdx = 20; // Column U (index 20)
  
  var masterSS = SpreadsheetApp.openById(masterSheetId);
  var w1Sheet = masterSS.getSheetByName(w1SheetName);
  
  var masterRange = w1Sheet.getDataRange();
  var masterData = masterRange.getValues();
  
  for (var j = 1; j < masterData.length; j++) {
    var mRow = masterData[j];
    var notes = mRow[notesIdx];
    
    if (notes && notes.toString().trim() !== "") {
      var notesStr = notes.toString();
      
      // Find the first occurrence of either phrase
      var prevDupIndex = notesStr.indexOf("Previous duplicate");
      var oldIdIndex = notesStr.indexOf("Old Study ID:");
      
      // Find which one comes first (or if neither exists)
      var cutoffIndex = -1;
      if (prevDupIndex !== -1 && oldIdIndex !== -1) {
        cutoffIndex = Math.min(prevDupIndex, oldIdIndex);
      } else if (prevDupIndex !== -1) {
        cutoffIndex = prevDupIndex;
      } else if (oldIdIndex !== -1) {
        cutoffIndex = oldIdIndex;
      }
      
      // If we found one of the phrases, cut everything from that point
      if (cutoffIndex !== -1) {
        mRow[notesIdx] = notesStr.substring(0, cutoffIndex).trim();
      }
    }
  }
  
  w1Sheet.getRange(1, 1, masterData.length, masterData[0].length).setValues(masterData);
  Logger.log("Notes cleaned up successfully! Removed 'Previous duplicate:' and 'Old Study ID:' and everything after.");
}

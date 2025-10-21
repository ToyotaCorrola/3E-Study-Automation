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

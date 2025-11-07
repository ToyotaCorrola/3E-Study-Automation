function oralHealth() {
  let scriptProperties = PropertiesService.getScriptProperties();
  // Master Sheet and Linking Destination details.
  var masterSheetId             = scriptProperties.getProperty("MasterSheetKEY"); 
  var linkingDestinationSheetId = scriptProperties.getProperty("OralHealthKEY");

  const masterSpreadsheet     = SpreadsheetApp.openById(masterSheetId);
  const oralHealthSpreadsheet = SpreadsheetApp.openById(linkingDestinationSheetId);
  const oralHealthSheet = oralHealthSpreadsheet.getSheetByName("Sheet1");
  const oralHealthData = oralHealthSheet.getDataRange().getValues();

  // Master headers:
  // A: StudyID (index 0)
  // D: School Email (index 3)
  // Oral Health headers:
  // A: StudyID (index 0)
  // F: Email (index 5: "[School Email Address]")

  const EMAIL_COL_MASTER = 3; 
  const STUDYID_COL_MASTER = 0; 
  const NAME_EMAIL_COL_ORAL = 5;
  const STUDYID_COL_ORAL = 0;

  const sheetsInfo = {
    1: {sheet: masterSpreadsheet.getSheetByName('W1'), oralColumn: 15, healthCompletedIdx: 12},
    2: {sheet: masterSpreadsheet.getSheetByName('W2'), oralColumn: 13, healthCompletedIdx: 10},
    3: {sheet: masterSpreadsheet.getSheetByName('W3'), oralColumn: 13, healthCompletedIdx: 10},
    4: {sheet: masterSpreadsheet.getSheetByName('W4'), oralColumn: 13, healthCompletedIdx: 10},
    5: {sheet: masterSpreadsheet.getSheetByName('W5'), oralColumn: 13, healthCompletedIdx: 10}
  };

  // Preload data and create maps for efficient lookup (by StudyID and by Email)
  const masterDataMaps = {};
  Object.keys(sheetsInfo).forEach(waveKey => {
    const sheet = sheetsInfo[waveKey].sheet;
    const range = sheet.getDataRange();
    const data = range.getValues();
    const backgrounds = range.getBackgrounds();

    const mapByStudyId = {};
    const mapByEmail = {};

    for (let i = 1; i < data.length; i++) { // skip header
      const row = data[i];
      const studyIdRaw = row[STUDYID_COL_MASTER];
      const emailRaw = row[EMAIL_COL_MASTER];

      const studyId = studyIdRaw !== null && studyIdRaw !== undefined ? String(studyIdRaw).trim() : '';
      const email = emailRaw ? String(emailRaw).trim().toLowerCase() : '';

      const rec = {row: i, data: row, backgrounds: backgrounds[i]};
      if (studyId) mapByStudyId[studyId] = rec;
      if (email) mapByEmail[email] = rec;
    }

    masterDataMaps[waveKey] = {sheet, data, backgrounds, mapByStudyId, mapByEmail};
  });

  // Update Oral Health Data
  // Oral health rows: need wave in column B (index 1) as in your original
  for (let r = 1; r < oralHealthData.length; r++) {
    const row = oralHealthData[r];
    const oralStudyId = row[STUDYID_COL_ORAL] != null ? String(row[STUDYID_COL_ORAL]).trim() : '';
    const wave = row[1]; // assumes Wave number in col B as in your code
    const sheetInfo = sheetsInfo[wave];
    if (!sheetInfo) continue;

    const {oralColumn} = sheetInfo;
    const masterSheetData = masterDataMaps[String(wave)];
    if (!masterSheetData) continue;

    let participant = null;

    // 1) Try StudyID
    if (oralStudyId) {
      participant = masterSheetData.mapByStudyId[oralStudyId];
    }

    // 2) Fallback by Email (parsed from Oral Health "Name/Email" col F)
    if (!participant) {
      const emailFromOral = row[NAME_EMAIL_COL_ORAL];
      if (emailFromOral) {
        participant = masterSheetData.mapByEmail[emailFromOral] || null;
      }
    }

    if (participant) {
      const currentVal = String(participant.data[oralColumn] || '').toUpperCase();
      if (currentVal !== 'YES') {
        participant.data[oralColumn] = 'YES';
        participant.backgrounds[oralColumn] = 'red';
      }
    }
  }

  // Default all empty or non-YES cells to NO (only if Health Completed = YES)
  Object.entries(masterDataMaps).forEach(([waveKey, {data, backgrounds}]) => {
    const {oralColumn, healthCompletedIdx} = sheetsInfo[waveKey];
    for (let i = 1; i < data.length; i++) {
      const completed = String(data[i][healthCompletedIdx] || '').toUpperCase() === 'YES';
      if (!completed) continue;

      const val = String(data[i][oralColumn] || '').toUpperCase();
      if (val !== 'YES') {
        if (val !== 'NO') {
          data[i][oralColumn] = 'NO';
          backgrounds[i][oralColumn] = 'red';
        }
      }
    }
  });

  // Write updates back to sheets in bulk
  Object.values(masterDataMaps).forEach(({sheet, data, backgrounds}) => {
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    sheet.getRange(1, 1, backgrounds.length, backgrounds[0].length).setBackgrounds(backgrounds);
  });
}

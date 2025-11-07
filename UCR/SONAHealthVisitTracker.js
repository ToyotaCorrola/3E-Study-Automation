function SONA_HEALTH_VISIT_TRACKER() {


  add_Health_Visit_Study_ID_If_Missing();


  var scriptProperties = PropertiesService.getScriptProperties();


  const TEMP_LINKING_KEY                = SpreadsheetApp.openById(scriptProperties.getProperty("MasterSheetKEY"));  
  const HV_LINKING_DEST                 = SpreadsheetApp.openById(scriptProperties.getProperty("HealthVisitKEY"));
  const SONA_TRACKER_DEST               = SpreadsheetApp.openById(scriptProperties.getProperty("SONA_TRACKER_KEY"));


  const TEMP_LINKING_KEY_SHEET          = TEMP_LINKING_KEY.getSheetByName("W1"); // We assume we are only working in Wave 1
  const SONA_TRACKER_DEST_SHEET         = SONA_TRACKER_DEST.getSheetByName("Main Sheet")
  const HV_LINKING_DEST_SHEET           = HV_LINKING_DEST.getSheetByName("Sheet1");


  const MASTER_INDICES = {
    STUDYID_COL_MASTER                  : 0,
    SCHOOL_EMAIL_COL_MASTER             : 3,
    PERSONAL_EMAIL_COL_MASTER           : 4,
    SOURCE_COL_MASTER                   : 7
  };


  const HV_INDICES = {
        STUDYID_COL_HV                  : 0,
        NAME_COL_HV                     : 1,
        DATE_COL_HV                     : 2,
        WAVE_COL_HV                     : 3,
        EMAIL_COL_HV                    : 4
  };


  var TEMP_LINKING_KEY_MAP        = createLinkingKeyMap(TEMP_LINKING_KEY_SHEET,
                                                        MASTER_INDICES.STUDYID_COL_MASTER, 
                                                        MASTER_INDICES.SCHOOL_EMAIL_COL_MASTER, 
                                                        MASTER_INDICES.PERSONAL_EMAIL_COL_MASTER);
  var TEMP_LK_SOURCES_MAP         = getLinkingKeySources(TEMP_LINKING_KEY_SHEET,
                                                         MASTER_INDICES.STUDYID_COL_MASTER,
                                                         MASTER_INDICES.SOURCE_COL_MASTER);
  var HV_LINKING_DEST_MAP         = {};
  var SONA_TRACKER_DEST_MAP       = {};


  const startingIdx = 881;
  const lastHVRow = HV_LINKING_DEST_SHEET.getLastRow();
  const lastHVCol = HV_LINKING_DEST_SHEET.getLastColumn();

  for(let index = startingIdx + (SONA_TRACKER_DEST_SHEET.getLastRow()-1); 
      index <= lastHVRow; 
      index++)
  /*------------ 
  INIT: startingIdx == 881 because Dustin built this system 
  on 11/7/2025 and that is when the SONA tracking was first launched.

  let index = startingIdx + (SONA_TRACKER_DEST_SHEET.getLastRow()-1); 
  gets the newest elements that are not already in SONA.
  -------------*/
  { 
    const HV_ROW        = HV_LINKING_DEST_SHEET
                          .getRange(index, 1, 1, lastHVCol)
                          .getValues()[0];
    const HV_ROW_DATA   = {
      STUDYID           : HV_ROW[HV_INDICES.STUDYID_COL_HV],
      NAME              : HV_ROW[HV_INDICES.NAME_COL_HV],
      DATE              : HV_ROW[HV_INDICES.DATE_COL_HV],
      WAVE              : HV_ROW[HV_INDICES.WAVE_COL_HV],
      EMAIL             : HV_ROW[HV_INDICES.EMAIL_COL_HV].toString().toLowerCase().trim(),
    };
    if(HV_ROW_DATA.WAVE == 1){
      if(TEMP_LINKING_KEY_MAP.STUDYID_KEY.has(String(HV_ROW_DATA.STUDYID))      || 
         TEMP_LINKING_KEY_MAP.SCHOOL_EMAIL_KEY.has(HV_ROW_DATA.EMAIL)   ||
         TEMP_LINKING_KEY_MAP.PERSONAL_EMAIL_KEY.has(HV_ROW_DATA.EMAIL)){
        const src = (TEMP_LK_SOURCES_MAP.get(HV_ROW_DATA.STUDYID) || '').toString().trim();
        const background_color = src === 'SONA' ? 'yellow' : 'red';

        const nextRow = SONA_TRACKER_DEST_SHEET.getLastRow() + 1;

        const values = [[
          HV_ROW_DATA.STUDYID,
          HV_ROW_DATA.NAME,
          HV_ROW_DATA.DATE,   // keep as Date; will display per sheet format
          HV_ROW_DATA.WAVE,
          HV_ROW_DATA.EMAIL
        ]];

        const rng = SONA_TRACKER_DEST_SHEET.getRange(nextRow, 1, 1, 5);
        rng.setValues(values);
        rng.setBackgrounds([[background_color, background_color, background_color, background_color, background_color]]);
      }
    }
  }
}

function createLinkingKeyMap(
  TEMP_LINKING_KEY_SHEET,
  STUDYID_COL_MASTER,
  SCHOOL_EMAIL_COL_MASTER,
  PERSONAL_EMAIL_COL_MASTER,
) {
  const numRows = TEMP_LINKING_KEY_SHEET.getLastRow();
  const numCols = TEMP_LINKING_KEY_SHEET.getLastColumn();
  const data    = TEMP_LINKING_KEY_SHEET.getRange(1, 1, numRows, numCols).getValues();

  const result = {
    STUDYID_KEY:        new Set(),
    SCHOOL_EMAIL_KEY:   new Set(),
    PERSONAL_EMAIL_KEY: new Set(),
  };

  const add = (k, v) => {
    if (v == null) return;
    const s = String(v).trim();
    if (s) result[k].add(s);
  };

  for (let i = 1; i < data.length; i++) { // skip header row at index 0
    add("STUDYID_KEY",        data[i][STUDYID_COL_MASTER]);
    add("SCHOOL_EMAIL_KEY",   data[i][SCHOOL_EMAIL_COL_MASTER]?.toString().toLowerCase().trim());
    add("PERSONAL_EMAIL_KEY", data[i][PERSONAL_EMAIL_COL_MASTER]?.toString().toLowerCase().trim());
  }

  return result; // { STUDYID_KEY: Set, SCHOOL_EMAIL_KEY: Set, PERSONAL_EMAIL_KEY: Set}
}

function getLinkingKeySources(
  TEMP_LINKING_KEY_SHEET,
  STUDYID_COL_MASTER,
  SOURCE_COL_MASTER
) {
  const numRows = TEMP_LINKING_KEY_SHEET.getLastRow();
  const numCols = TEMP_LINKING_KEY_SHEET.getLastColumn();
  const data    = TEMP_LINKING_KEY_SHEET.getRange(1, 1, numRows, numCols).getValues();

  const byStudyId = new Map(); // Map<number, string>

  for (let i = 1; i < data.length; i++) { // skip header row 0
    /** @type {number} */
    const id = data[i][STUDYID_COL_MASTER];
    /** @type {string} */
    const source = data[i][SOURCE_COL_MASTER];

    if (source && source.trim()) {
      byStudyId.set(id, source.trim());
    }
  }

  return byStudyId;
}

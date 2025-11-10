const scriptProperties = PropertiesService.getScriptProperties();
const TEMP_LINKING_KEY = SpreadsheetApp.openById(
  scriptProperties.getProperty("MasterSheetKEY")
);

class EmailToStudyIDsClass {
  constructor(keepStudyID = null, studyIDs = []) {
    this.keepStudyID = Number.isInteger(keepStudyID) ? keepStudyID : null;
    this.studyIDs = new Set();
    for (const id of studyIDs) {
      if (Number.isInteger(id)) this.studyIDs.add(id);
    }
  }
}

function removeDuplicateStudyIDsByEmail() {
  const sheetNames = ["W1", "W2", "W3", "W4", "W5"];

  const notesColumnBySheet = {
    "W1": 21,
    "W2": 19,
    "W3": 20,
    "W4": 21,
    "W5": 22
  };

  // Only this color will be re-applied to cells
  const HIGHLIGHT_COLOR = "#ff0000";

  // Column indices (0-based) in the W* sheets
  const STUDY_ID_COL_INDEX = 0; // A
  const EMAIL_COL_INDEX = 3;    // D
  const SOURCE_COL_INDEX = 7;   // H (only used for W1)

  // PHASE 1: detect emails that have multiple StudyIDs with extra data across all waves
  const emailToDataStudyIDsGlobal = new Map(); // email -> Set(studyID)

  sheetNames.forEach(sheetName => {
    const sheet = TEMP_LINKING_KEY.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Phase 1 - sheet not found: %s", sheetName);
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return;

    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();

    let extraDataCols;
    if (sheetName === "W1") {
      // J, K, M, N, O -> 10,11,13,14,15 -> indices 9,10,12,13,14
      extraDataCols = [9, 10, 12, 13, 14];
    } else {
      // H, I, J, K, L, M -> 8–13 -> indices 7–12
      extraDataCols = [7, 8, 9, 10, 11, 12];
    }

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const studyIDRaw = row[STUDY_ID_COL_INDEX];
      const email = row[EMAIL_COL_INDEX];

      if (!email) continue;
      if (studyIDRaw === "" || studyIDRaw == null) continue;

      const studyID = Number(studyIDRaw);
      if (!Number.isInteger(studyID)) continue;

      let hasExtraData = false;
      for (const c of extraDataCols) {
        if (c >= row.length) continue;
        const val = row[c];
        if (val !== "" && val != null) {
          hasExtraData = true;
          break;
        }
      }

      if (!hasExtraData) continue;

      let set = emailToDataStudyIDsGlobal.get(email);
      if (!set) {
        set = new Set();
        emailToDataStudyIDsGlobal.set(email, set);
      }
      set.add(studyID);
    }
  });

  // Emails with >1 StudyID that contain data somewhere across the waves
  const multiDataEmails = new Set();
  emailToDataStudyIDsGlobal.forEach((set, email) => {
    if (set.size > 1) {
      multiDataEmails.add(email);
    }
  });

  // PHASE 2: main logic per sheet, respecting multiDataEmails
  sheetNames.forEach(sheetName => {
    const sheet = TEMP_LINKING_KEY.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log("Sheet not found: %s", sheetName);
      return;
    }

    let lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < 2) return; // only header or empty

    // Read all data and backgrounds (for highlight preservation)
    const dataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = dataRange.getValues();
    const backgrounds = dataRange.getBackgrounds();

    const emailMap = new Map();      // email -> EmailToStudyIDsClass
    const indexLookup = new Map();   // "email|studyID" -> row index in values[]
    const highlightMap = new Map();  // "email|studyID" -> array of original backgrounds for that row

    // Extra-data columns per sheet (0-based indices)
    let extraDataCols;
    if (sheetName === "W1") {
      extraDataCols = [9, 10, 12, 13, 14];
    } else {
      extraDataCols = [7, 8, 9, 10, 11, 12];
    }

    // Build structures by iterating backwards
    for (let i = values.length - 1; i >= 0; i--) {
      const row = values[i];
      const studyIDRaw = row[STUDY_ID_COL_INDEX];
      const email = row[EMAIL_COL_INDEX];

      if (!email) continue;
      if (studyIDRaw === "" || studyIDRaw == null) continue;

      const studyID = Number(studyIDRaw);
      if (!Number.isInteger(studyID)) continue;

      const key = email + "|" + studyID;

      // Save original backgrounds for the entire row (for later cell-level restore)
      highlightMap.set(key, backgrounds[i].slice());

      // Check extra-data columns for this row
      let hasExtraData = false;
      for (const c of extraDataCols) {
        if (c >= row.length) continue;
        const val = row[c];
        if (val !== "" && val != null) {
          hasExtraData = true;
          break;
        }
      }

      indexLookup.set(key, i);

      let entry = emailMap.get(email);
      if (!entry) {
        entry = new EmailToStudyIDsClass();
        emailMap.set(email, entry);
      }

      entry.studyIDs.add(studyID);

      // If any extra-data columns have data and we haven't chosen keepStudyID yet,
      // use this row's StudyID as keepStudyID (only matters for non-multiDataEmails).
      if (hasExtraData && entry.keepStudyID == null) {
        entry.keepStudyID = studyID;
      }
    }

    const notesCol = notesColumnBySheet[sheetName];
    if (typeof notesCol !== "number") {
      Logger.log("Notes column not defined for sheet %s", sheetName);
      return;
    }

    const indicesToDelete = [];
    const deletedForSources = []; // only used for W1

    // First pass: notes + which rows to delete
    for (const [email, entry] of emailMap.entries()) {
      if (entry.studyIDs.size <= 1) continue;

      const allIdsSorted = Array.from(entry.studyIDs).sort((a, b) => a - b);

      // SPECIAL CASE: email has multiple StudyIDs with data somewhere across waves
      if (multiDataEmails.has(email)) {
        // Keep all StudyIDs in all waves for this email.
        // Optionally annotate each row with the full set of StudyIDs.
        const fullSetText = allIdsSorted.join(", ");
        const noteToAdd =
          `Also has duplicate StudyIDs of [${fullSetText}]`;

        for (const id of entry.studyIDs) {
          const idx = indexLookup.get(email + "|" + id);
          if (idx == null) continue;
          const rowNumber = idx + 2;
          const notesCell = sheet.getRange(rowNumber, notesCol);
          const existingNote = notesCell.getValue();

          if (existingNote && existingNote.indexOf(noteToAdd) !== -1) {
            continue;
          }

          const newNote = existingNote
            ? existingNote + "\n" + noteToAdd
            : noteToAdd;
          notesCell.setValue(newNote);

          Logger.log(
            "Sheet %s - Email %s is multi-data; keeping StudyID=%s; fullSet=[%s]",
            sheetName,
            email,
            id,
            fullSetText
          );
        }

        // Do NOT mark anything for deletion for this email.
        continue;
      }

      // Normal behavior for non-multiData emails:
      if (entry.keepStudyID != null) {
        // CASE 1: Some row has extra data -> we have a keepStudyID
        const idsToDelete = [];

        for (const id of entry.studyIDs) {
          if (id !== entry.keepStudyID) {
            idsToDelete.push(id);
            const idx = indexLookup.get(email + "|" + id);
            if (idx != null) {
              indicesToDelete.push(idx);

              // If this is W1, capture StudyID + Source from column H
              if (sheetName === "W1") {
                const src = values[idx][SOURCE_COL_INDEX];
                if (src) {
                  deletedForSources.push({
                    studyID: id,
                    source: String(src).trim()
                  });
                }
              }
            }
          }
        }

        const keepIndex = indexLookup.get(email + "|" + entry.keepStudyID);
        if (keepIndex != null && idsToDelete.length > 0) {
          const keepRowNumber = keepIndex + 2; // + header row
          const notesCell = sheet.getRange(keepRowNumber, notesCol);

          const existingNote = notesCell.getValue();
          const duplicateText = idsToDelete.sort((a, b) => a - b).join(", ");
          const noteToAdd =
            `Also has duplicate StudyIDs of [${duplicateText}]`;

          // Only add if this exact note does not already exist
          if (!existingNote || existingNote.indexOf(noteToAdd) === -1) {
            const newNote = existingNote
              ? existingNote + "\n" + noteToAdd
              : noteToAdd;
            notesCell.setValue(newNote);
          }

          Logger.log(
            "Sheet %s - Email %s: keepStudyID=%s; duplicates=[%s]",
            sheetName,
            email,
            entry.keepStudyID,
            duplicateText
          );
        }
      } else {
        // CASE 2: No row with extra data for this email -> keep all rows
        const fullSetText = allIdsSorted.join(", ");
        const noteToAdd =
          `Also has duplicate StudyIDs of [${fullSetText}]`;

        for (const id of entry.studyIDs) {
          const idx = indexLookup.get(email + "|" + id);
          if (idx == null) continue;
          const rowNumber = idx + 2;
          const notesCell = sheet.getRange(rowNumber, notesCol);
          const existingNote = notesCell.getValue();

          if (existingNote && existingNote.indexOf(noteToAdd) !== -1) {
            continue;
          }

          const newNote = existingNote
            ? existingNote + "\n" + noteToAdd
            : noteToAdd;
          notesCell.setValue(newNote);

          Logger.log(
            "Sheet %s - Email %s: (no extra-data) StudyID=%s; fullSet=[%s]",
            sheetName,
            email,
            id,
            fullSetText
          );
        }
      }
    }

    // For W1, delete these StudyIDs from the Screening/Consent source sheets as well
    if (sheetName === "W1" && deletedForSources.length > 0) {
      deleteStudyIDsFromSources_(deletedForSources);
    }

    // Delete rows for CASE 1 emails, bottom-up in the current W* sheet
    const rowsToDelete = [...new Set(indicesToDelete)]
      .map(idx => idx + 2)
      .sort((a, b) => b - a);

    rowsToDelete.forEach(rowNum => {
      // Safety: never delete a keepStudyID row
      const rowValues = sheet.getRange(rowNum, 1, 1, 4).getValues()[0];
      const studyIDRaw = rowValues[0];
      const email = rowValues[3];
      const studyID = Number(studyIDRaw);

      const entry = emailMap.get(email);
      if (entry && entry.keepStudyID != null && studyID === entry.keepStudyID) {
        Logger.log(
          "Skipping deletion of keepStudyID row: Sheet %s, Email %s, StudyID %s, Row %s",
          sheetName,
          email,
          studyID,
          rowNum
        );
        return;
      }

      sheet.deleteRow(rowNum);
    });

    // After deletions, restore highlight state per remaining row,
    // but ONLY for cells that were originally HIGHLIGHT_COLOR
    lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const newDataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const newValues = newDataRange.getValues();
    const newBackgrounds = newDataRange.getBackgrounds();

    for (let i = 0; i < newValues.length; i++) {
      const row = newValues[i];
      const studyIDRaw = row[STUDY_ID_COL_INDEX];
      const email = row[EMAIL_COL_INDEX];

      if (!email) continue;
      if (studyIDRaw === "" || studyIDRaw == null) continue;

      const studyID = Number(studyIDRaw);
      if (!Number.isInteger(studyID)) continue;

      const key = email + "|" + studyID;
      if (!highlightMap.has(key)) continue;

      const originalRowBgs = highlightMap.get(key);

      for (let c = 0; c < lastCol && c < originalRowBgs.length; c++) {
        const origColor = originalRowBgs[c];
        // Only reapply the highlight color (red), leave everything else alone
        if (origColor && origColor.toLowerCase() === HIGHLIGHT_COLOR) {
          newBackgrounds[i][c] = origColor;
        }
      }
    }

    newDataRange.setBackgrounds(newBackgrounds);
  });
}

/**
 * Delete rows with given StudyIDs from the source Screening/Consent spreadsheets.
 * Only called for W1.
 * deletedItems: array of { studyID: number, source: string }
 */
function deleteStudyIDsFromSources_(deletedItems) {
  var sources = [
    {
      fileId: scriptProperties.getProperty("ScreeningAndConsent-Email-KEY"),
      sheetName: "Email",
      source: "EMAIL"
    },
    {
      fileId: scriptProperties.getProperty("ScreeningAndConsent-IGWEB-KEY"),
      sheetName: "Sheet1",
      source: "IG/WEB"
    },
    {
      fileId: scriptProperties.getProperty("ScreeningAndConsent-REFERRAL-KEY"),
      sheetName: "Sheet1",
      source: "REFERRAL"
    },
    {
      fileId: scriptProperties.getProperty("ScreeningAndConsent-FLYER-KEY"),
      sheetName: "Sheet1",
      source: "FLYER"
    },
    {
      fileId: scriptProperties.getProperty("ScreeningAndConsent-SONA"),
      sheetName: "Sheet1",
      source: "SONA"
    }
  ];

  const bySource = new Map();
  deletedItems.forEach(item => {
    const src = item.source;
    const id = Number(item.studyID);
    if (!Number.isInteger(id)) return;
    if (!bySource.has(src)) bySource.set(src, new Set());
    bySource.get(src).add(id);
  });

  bySource.forEach((idSet, srcLabel) => {
    const config = sources.find(s => s.source === srcLabel);
    if (!config || !config.fileId) {
      Logger.log(
        "No source config found for source '%s'; StudyIDs: %s",
        srcLabel,
        Array.from(idSet).join(", ")
      );
      return;
    }

    const ss = SpreadsheetApp.openById(config.fileId);
    const sh = ss.getSheetByName(config.sheetName);
    if (!sh) {
      Logger.log(
        "Sheet '%s' not found in source file for '%s'",
        config.sheetName,
        srcLabel
      );
      return;
    }

    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;

    // Assume StudyID is in column A
    const idRange = sh.getRange(2, 1, lastRow - 1, 1);
    const idValues = idRange.getValues();

    const rowsToDelete = [];
    for (let i = 0; i < idValues.length; i++) {
      const cellVal = Number(idValues[i][0]);
      if (!Number.isInteger(cellVal)) continue;
      if (idSet.has(cellVal)) {
        rowsToDelete.push(i + 2); // data starts at row 2
      }
    }

    rowsToDelete.sort((a, b) => b - a);
    rowsToDelete.forEach(r => {
      Logger.log("Deleting StudyID row in source '%s': row %s", srcLabel, r);
      sh.deleteRow(r);
    });
  });
}
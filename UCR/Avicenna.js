function updateSubstudyEnrollment() {
  const avicennaFileId = "1HwUFMNkQHDypZ5Uz0eL-XWPwTymRQUMNZqW1JnB7SW4";
  const linkingKeyFileId = "1rLIUGPVviSjAP9P9PwvjDVWyEwfyUxkH3kgN0sn7xkI";

  const avicennaFile = SpreadsheetApp.openById(avicennaFileId);
  const linkingKeyFile = SpreadsheetApp.openById(linkingKeyFileId);

  processSheet("W1", "W1", "R", "S", avicennaFile, linkingKeyFile);
  processSheet("W2", "W2", "P", "Q", avicennaFile, linkingKeyFile);
  processSheet("W3", "W3", "Q", "R", avicennaFile, linkingKeyFile);
  processSheet("W4", "W4", "R", "S", avicennaFile, linkingKeyFile);
  processSheet("W5", "W5", "S", "T", avicennaFile, linkingKeyFile);
}

function processSheet(avicennaSheetName, linkingKeySheetName, enrolledCol, idCol, avicennaFile, linkingKeyFile) {
  const avicennaSheet = avicennaFile.getSheetByName(avicennaSheetName);
  const linkingKeySheet = linkingKeyFile.getSheetByName(linkingKeySheetName);
  const avicennaLastRow = avicennaSheet.getLastRow();
  const linkingKeyLastRow = linkingKeySheet.getLastRow();
  if (avicennaLastRow < 2 || linkingKeyLastRow < 2) return; // no data

  // Ensure we load enough columns so that the update columns (enrolled and id) are included.
  const maxCol = Math.max(columnLetterToNumber(enrolledCol), columnLetterToNumber(idCol));
  const numCols = maxCol - 1; // reading from column B onward
  const linkingKeyData = linkingKeySheet.getRange(2, 2, linkingKeyLastRow - 1, numCols).getValues();
  const enrolledIdx = columnLetterToNumber(enrolledCol) - 2; // index in linkingKeyData
  const idIdx = columnLetterToNumber(idCol) - 2;

  // Load backgrounds for the update columns
  const enrolledBg = linkingKeySheet.getRange(2, columnLetterToNumber(enrolledCol), linkingKeyLastRow - 1, 1).getBackgrounds();
  const idBg = linkingKeySheet.getRange(2, columnLetterToNumber(idCol), linkingKeyLastRow - 1, 1).getBackgrounds();

  // Build a lookup: map each email (from ucrEmail & personalEmail) to an array of row indices.
  let emailMap = {};
  for (let j = 0; j < linkingKeyData.length; j++) {
    const ucrEmail = linkingKeyData[j][2] ? linkingKeyData[j][2].toString().toLowerCase() : "";
    const personalEmail = linkingKeyData[j][3] ? linkingKeyData[j][3].toString().toLowerCase() : "";
    if (ucrEmail) {
      if (!emailMap[ucrEmail]) emailMap[ucrEmail] = [];
      emailMap[ucrEmail].push(j);
    }
    if (personalEmail) {
      if (!emailMap[personalEmail]) emailMap[personalEmail] = [];
      emailMap[personalEmail].push(j);
    }
  }

  // Process avicenna data
  const avicennaData = avicennaSheet.getRange(2, 1, avicennaLastRow - 1, 12).getValues();
  for (let i = 0; i < avicennaData.length; i++) {
    const row = avicennaData[i];
    const id = row[0];
    const email = row[11] ? row[11].toString().toLowerCase() : "";
    if (email && emailMap[email]) {
      emailMap[email].forEach(function(j) {
        if (linkingKeyData[j][enrolledIdx] !== "YES" || linkingKeyData[j][idIdx] !== id) {
          linkingKeyData[j][enrolledIdx] = "YES";
          linkingKeyData[j][idIdx] = id;
          enrolledBg[j][0] = "red";
          idBg[j][0] = "red";
        }
      });
    }
  }

  // Write back updated columns in one go
  linkingKeySheet
    .getRange(2, columnLetterToNumber(enrolledCol), linkingKeyData.length, 1)
    .setValues(linkingKeyData.map(r => [r[enrolledIdx]]));
  linkingKeySheet
    .getRange(2, columnLetterToNumber(enrolledCol), linkingKeyData.length, 1)
    .setBackgrounds(enrolledBg);
  linkingKeySheet
    .getRange(2, columnLetterToNumber(idCol), linkingKeyData.length, 1)
    .setValues(linkingKeyData.map(r => [r[idIdx]]));
  linkingKeySheet
    .getRange(2, columnLetterToNumber(idCol), linkingKeyData.length, 1)
    .setBackgrounds(idBg);
}

function columnLetterToNumber(letter) {
  return letter.charCodeAt(0) - 'A'.charCodeAt(0) + 1;
}

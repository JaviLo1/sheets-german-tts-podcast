/**
 * Main function to run the process:
 * 1. Read the table name (named range) from Dashboard!E5.
 * 2. Retrieve texts (English, German word, German sentence) from that named range.
 * 3. Authenticate with Google Cloud TTS using service account key from Drive (File ID from Script Properties).
 * 4. Generate MP3 for each row (English -> German word -> German sentence), merge them row-by-row.
 * 5. Concatenate all row MP3s into a single MP3.
 * 6. Save final MP3 to Drive, naming it <tableName>.mp3 in the same folder as the key file (fallback: root).
 */
function runTextToSpeechProcess() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1) Read table name from Dashboard!E5
  const dashboardSheet = ss.getSheetByName("Dashboard");
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('No sheet named "Dashboard" found.');
    return;
  }

  const tableName = String(dashboardSheet.getRange("E5").getValue()).trim();
  if (!tableName) {
    SpreadsheetApp.getUi().alert("Please enter a valid named range in cell E5.");
    return;
  }

  // 2) Get named range (3 columns: English, German word, German sentence)
  const namedRange = ss.getRangeByName(tableName);
  if (!namedRange) {
    SpreadsheetApp.getUi().alert(`No named range found with the name "${tableName}".`);
    return;
  }
  const rows = namedRange.getValues();

  // 3) Get Service Account Key file from Drive (File ID stored in Script Properties)
  const props = PropertiesService.getScriptProperties();
  const SERVICE_ACCOUNT_KEY_FILE_ID = props.getProperty('SERVICE_ACCOUNT_KEY_FILE_ID');
  if (!SERVICE_ACCOUNT_KEY_FILE_ID) {
    SpreadsheetApp.getUi().alert('Missing Script Property "SERVICE_ACCOUNT_KEY_FILE_ID". Please set it in Project Settings → Script properties.');
    return;
  }

  let serviceAccountKey;
  try {
    serviceAccountKey = getServiceAccountKeyFromDrive(SERVICE_ACCOUNT_KEY_FILE_ID);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Could not read service account key from Drive. " + e);
    return;
  }

  let accessToken;
  try {
    accessToken = getAccessToken(serviceAccountKey, ['https://www.googleapis.com/auth/cloud-platform']);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Could not obtain access token. " + e);
    return;
  }

  // 4) Generate audio per row and merge EN -> word -> sentence
  const mergedBlobs = [];
  rows.forEach(row => {
    const englishText = row[0];
    const germanWord = row[1];
    const germanSentence = row[2];

    if (!englishText || !germanWord || !germanSentence) return;

    const englishBlob = generateSpeechBlob(accessToken, englishText, 'en-US');
    const germanWordBlob = generateSpeechBlob(accessToken, germanWord, 'de-DE');
    const germanSentenceBlob = generateSpeechBlob(accessToken, germanSentence, 'de-DE');

    let mergedRowBlob = mergeMp3Blobs(englishBlob, germanWordBlob);
    mergedRowBlob = mergeMp3Blobs(mergedRowBlob, germanSentenceBlob);
    mergedBlobs.push(mergedRowBlob);
  });

  if (mergedBlobs.length === 0) {
    SpreadsheetApp.getUi().alert("No valid rows found or data missing in the named range.");
    return;
  }

  // 5) Concatenate all rows
  let finalMp3Blob = mergedBlobs[0];
  for (let i = 1; i < mergedBlobs.length; i++) {
    finalMp3Blob = mergeMp3Blobs(finalMp3Blob, mergedBlobs[i]);
  }

  // 6) Save final MP3 alongside the key file (or Drive root)
  const keyFile = DriveApp.getFileById(SERVICE_ACCOUNT_KEY_FILE_ID);
  const parents = keyFile.getParents();
  const targetFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const finalFileName = `${tableName}.mp3`;
  const finalFile = targetFolder.createFile(finalMp3Blob.setName(finalFileName));

  Logger.log(`Final MP3 file created: ${finalFile.getName()} (ID: ${finalFile.getId()})`);
  SpreadsheetApp.getUi().alert(`Final MP3 file created: ${finalFile.getName()}`);
}

/**
 * Retrieve the service account JSON key from Drive and parse it.
 */
function getServiceAccountKeyFromDrive(fileId) {
  const file = DriveApp.getFileById(fileId);
  const content = file.getBlob().getDataAsString();
  return JSON.parse(content);
}

/**
 * Get an OAuth access token for the given scopes using a service account JSON key (JWT flow).
 */
function getAccessToken(serviceAccountKey, scopes) {
  const iat = Math.floor(Date.now() / 1000);
  const exp = iat + 3600;

  const header = { alg: "RS256", typ: "JWT" };
  const claimSet = {
    iss: serviceAccountKey.client_email,
    scope: scopes.join(' '),
    aud: "https://oauth2.googleapis.com/token",
    iat: iat,
    exp: exp
  };

  const encodedHeader = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=+$/, '');
  const encodedClaim = Utilities.base64EncodeWebSafe(JSON.stringify(claimSet)).replace(/=+$/, '');
  const toSign = `${encodedHeader}.${encodedClaim}`;
  const signatureBytes = Utilities.computeRsaSha256Signature(toSign, serviceAccountKey.private_key);
  const encodedSignature = Utilities.base64EncodeWebSafe(signatureBytes).replace(/=+$/, '');
  const jwt = `${toSign}.${encodedSignature}`;

  const response = UrlFetchApp.fetch("https://oauth2.googleapis.com/token", {
    method: "post",
    muteHttpExceptions: true,
    payload: {
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: jwt
    }
  });

  const result = JSON.parse(response.getContentText());
  if (!result.access_token) {
    throw new Error("Access token response: " + response.getContentText());
  }
  return result.access_token;
}

/**
 * Call Google TTS to generate an MP3 blob (SSML with a brief pause; slower speaking rate).
 */
function generateSpeechBlob(accessToken, text, languageCode) {
  const ssmlText = `<speak>${text}<break time="0.5s"/></speak>`;
  const payload = {
    input: { ssml: ssmlText },
    voice: { languageCode: languageCode, ssmlGender: "NEUTRAL" },
    audioConfig: { audioEncoding: "MP3", speakingRate: 0.8 }
  };

  const response = UrlFetchApp.fetch("https://texttospeech.googleapis.com/v1/text:synthesize", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: { Authorization: "Bearer " + accessToken }
  });

  const result = JSON.parse(response.getContentText());
  if (!result.audioContent) {
    throw new Error("Text-to-Speech API call failed: " + response.getContentText());
  }

  return Utilities.newBlob(
    Utilities.base64Decode(result.audioContent),
    "audio/mpeg",
    "temp.mp3"
  );
}

/**
 * Naively merge two MP3 blobs by concatenating bytes (may duplicate ID3 tags; usually fine for players).
 */
function mergeMp3Blobs(blob1, blob2) {
  const bytes1 = blob1.getBytes();
  const bytes2 = blob2.getBytes();
  const merged = new Uint8Array(bytes1.length + bytes2.length);
  merged.set(bytes1, 0);
  merged.set(bytes2, bytes1.length);
  return Utilities.newBlob(merged, "audio/mpeg", "merged.mp3");
}

/**
 * Fill missing sentences (Column 3) in the named range specified by Dashboard!E5.
 * Columns: [0] English, [1] German word, [2] German sentence
 */
function fillColumn3FromTableName() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = spreadsheet.getSheetByName("Dashboard");
  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('No sheet named "Dashboard" found.');
    return;
  }

  const tableName = String(dashboardSheet.getRange("E5").getValue()).trim();
  if (!tableName) {
    SpreadsheetApp.getUi().alert("Please enter a valid table name in cell E5.");
    return;
  }

  const namedRange = spreadsheet.getRangeByName(tableName);
  if (!namedRange) {
    SpreadsheetApp.getUi().alert(`No named range found with the name "${tableName}".`);
    return;
  }

  const values = namedRange.getValues();
  let updated = false;

  for (let i = 0; i < values.length; i++) {
    const germanWord = values[i][1];
    const exampleSentence = values[i][2];

    if (germanWord && (!exampleSentence || String(exampleSentence).trim() === '')) {
      values[i][2] = generateGermanSentence(germanWord);
      updated = true;
    }
  }

  if (updated) {
    namedRange.setValues(values);
    SpreadsheetApp.getUi().alert("Successfully filled missing example sentences in Column 3!");
  } else {
    SpreadsheetApp.getUi().alert("No missing sentences found to fill.");
  }
}

/***********************************
 * OpenAI Helper
 ***********************************/
function generateGermanSentence(germanWord) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('OPENAI_API_KEY');
  if (!apiKey) {
    throw new Error('Missing Script Property "OPENAI_API_KEY". Set it in Project Settings → Script properties.');
  }

  const prompt = `
You are a helpful assistant.
Generate a short, simple German sentence using only the word: "${germanWord}".
Provide ONLY the sentence (no extra text).
Examples:
Word: "Haus" -> "Ich gehe in das Haus."
Word: "Auto" -> "Das Auto ist rot."
----
Word: "${germanWord}" ->
`.trim();

  const url = "https://api.openai.com/v1/chat/completions";
  const payload = {
    model: "gpt-4",
    messages: [
      { role: "system", content: "You are ChatGPT, a German language assistant." },
      { role: "user", content: prompt }
    ],
    max_tokens: 50,
    temperature: 0.7
  };

  const response = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());
  if (!data.choices || !data.choices[0] || !data.choices[0].message) {
    throw new Error("OpenAI API error: " + response.getContentText());
  }
  return String(data.choices[0].message.content).trim();
}

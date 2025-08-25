# Sheets → German TTS Podcast

Generate short German-learning audio “mini-lessons” straight from Google Sheets.

The workflow:
1) Read a **named range** (English, German word, German sentence) selected via `Dashboard!E5`.
2) For each row, synthesize speech:
   - English text (en-US)
   - German word (de-DE)
   - German sentence (de-DE)
3) Merge each row’s clips (EN → word → sentence), then concatenate all rows into one MP3.
4) Save `<tableName>.mp3` to Google Drive.

Optional helper: auto-fill missing German example sentences for a list of words using OpenAI.

---

## ✨ Features
- **Spreadsheet-driven** content (easy to edit and version by non-devs).
- **Google Cloud Text-to-Speech** for high-quality EN/DE voices.
- **Row-by-row merging** into a single lesson MP3.
- **Sentence generator**: fill Column 3 with simple German sentences from a given word.

---

## 📦 What’s inside
- `apps-script/Code.gs`: main runner `runTextToSpeechProcess()`, TTS helpers, naive MP3 concatenation.
- `apps-script/fillSentences.gs`: `fillColumn3FromTableName()` + OpenAI helper.

> Note: You can keep a single `Code.gs` if you prefer; files are split here for readability.

---

## 🧱 Prerequisites
- A Google account with access to **Google Sheets**, **Apps Script**, and **Drive**.
- A Google Cloud project with **Text-to-Speech API** enabled.
- A **service account** with a **JSON key** (upload the key file to Drive).
- An **OpenAI API key** (for the optional sentence filler).

---

## 🔐 Security & Secrets
- **Do not commit API keys or service account JSONs** to GitHub.  
- This project expects secrets to be stored in **Apps Script Properties**:  
  - `OPENAI_API_KEY`  
  - `SERVICE_ACCOUNT_KEY_FILE_ID` 

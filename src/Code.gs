/**
 * ATTENDANCE_SHEET_URL_
 * URL of the Sierra Speakers attendance Google Sheet.
 * Defined at top level so buildFancyHtml_ can inject it as a hyperlink.
 */
const ATTENDANCE_SHEET_URL_ = "https://docs.google.com/spreadsheets/d/1Jpn0lGoO_XtPsbbkb2b9pLelgNyoKnB9950oMgTHMOY/edit?gid=287681997#gid=287681997";

/**
 * SCHEDULING_SHEET_URL_
 * URL used for the "scheduling sheet" hyperlink in all confirmation emails.
 *
 * To use the real (live) sheet:      set to null
 * To use a test/staging copy:        paste the test sheet URL as a string
 *
 * Currently pointing to the TEST sheet. Flip to null before deploying to production.
 */
const SCHEDULING_SHEET_URL_ = "https://docs.google.com/spreadsheets/d/1BAQPmjadcvjnPOsj16H6y22Ls-t34EeY63jlGPHPGU4/edit?gid=1919006288#gid=1919006288";
// const SCHEDULING_SHEET_URL_ = null; // ← uncomment this line (and comment the line above) to use the live sheet URL

// ================================================================
// WOD MEMORY — persistent cache for Word of the Day selections
// ================================================================

/**
 * WOD_MEMORY_HEADERS_
 * Column headers for the hidden WOD_Memory sheet.
 * @type {string[]}
 */
const WOD_MEMORY_HEADERS_ = [
  "Date", "Word", "Definition", "Pronunciation", "Part of Speech",
  "Example", "Source", "Theme Used", "Gemini Model Used"
];

/**
 * GEMINI_MODEL_RANK_
 * Maps Gemini model labels to a numeric rank for strength comparison.
 * Higher rank = stronger model. Used by shouldRefreshGemini_().
 * @type {Object<string, number>}
 */
const GEMINI_MODEL_RANK_ = {
  "gemma":        1,
  "gemini-lite":  2,
  "gemini-flash": 3
};

/**
 * getOrCreateWodMemorySheet_
 * Returns the hidden "WOD_Memory" sheet, creating it with headers and
 * hiding it if it does not yet exist.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The WOD_Memory sheet.
 */
function getOrCreateWodMemorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("WOD_Memory");
  if (!sheet) {
    sheet = ss.insertSheet("WOD_Memory");
    sheet.getRange(1, 1, 1, WOD_MEMORY_HEADERS_.length)
      .setValues([WOD_MEMORY_HEADERS_])
      .setFontWeight("bold");
    sheet.setColumnWidths(1, WOD_MEMORY_HEADERS_.length, 140);
        ss.setActiveSheet(ss.getSheetByName('SCHED 2026') || ss.getSheets()[0]);
    sheet.hideSheet();
  }
  return sheet;
}
/**
 * onOpen
 * Runs automatically when the spreadsheet opens.
/**
 * lookupWodCache_
 * Searches the WOD_Memory sheet for a row matching the given date string.
 * @param {string} dateStr - The formatted meeting date (e.g. "4/17/2026").
 * @return {Object|null} Cached WOD data object, or null if not found.
 */
function lookupWodCache_(dateStr) {
  const sheet = getOrCreateWodMemorySheet_();
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowDate = data[i][0] ? data[i][0].toString().trim() : "";
    if (rowDate === dateStr) {
      return {
        row:              i + 1,
        date:             rowDate,
        word:             (data[i][1] || "").toString().trim(),
        definition:       (data[i][2] || "").toString().trim(),
        pronunciation:    (data[i][3] || "").toString().trim(),
        partOfSpeech:     (data[i][4] || "").toString().trim(),
        example:          (data[i][5] || "").toString().trim(),
        source:           (data[i][6] || "").toString().trim(),
        themeUsed:        (data[i][7] || "").toString().trim(),
        geminiModelUsed:  (data[i][8] || "").toString().trim()
      };
    }
  }
  return null;
}
/**
 * saveWodToCache_
 * Writes or updates a WOD_Memory row for the given date.
 * If a row already exists for this date it is updated in place;
 * otherwise a new row is appended.
 * @param {string} dateStr          - Formatted meeting date.
 * @param {string} word             - The selected word.
 * @param {string} definition       - The chosen definition.
 * @param {string} pronunciation    - Phonetic pronunciation.
 * @param {string} partOfSpeech     - Part of speech.
 * @param {string} example          - Example sentence.
 * @param {string} source           - "mw" or a Gemini model label.
 * @param {string} themeUsed        - Meeting theme at generation time.
 * @param {string} geminiModelUsed  - Specific Gemini model label used.
 * @return {void}
 */
function saveWodToCache_(dateStr, word, definition, pronunciation,
                         partOfSpeech, example, source, themeUsed,
                         geminiModelUsed) {
  const sheet = getOrCreateWodMemorySheet_();
  const existing = lookupWodCache_(dateStr);
  const rowData = [
    dateStr, word || "", definition || "", pronunciation || "",
    partOfSpeech || "", example || "", source || "",
    themeUsed || "", geminiModelUsed || ""
  ];
  if (existing) {
    sheet.getRange(existing.row, 1, 1, rowData.length).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
}

/**
 * onOpen
 * Runs automatically when the spreadsheet opens.
 * Adds the "Toastmasters" menu with "Start Role Confirmations" and "Generate Meeting Agenda".
 * @return {void}
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
/**
 * shouldRefreshGemini_
 * Determines whether the Gemini API should be re-called based on
 * changes to the meeting theme or availability of a stronger model.
 * @param {Object} cachedRow        - Object from lookupWodCache_().
 * @param {string} currentTheme     - The current meeting theme.
 * @param {string} currentModelLabel - Label from getAiModel_().label.
 * @return {boolean} True if Gemini should be re-pinged.
 */
function shouldRefreshGemini_(cachedRow, currentTheme, currentModelLabel) {
  if (!cachedRow) return true;

  // Theme changed since last cache → re-ping Gemini
  if (cachedRow.themeUsed && currentTheme &&
      cachedRow.themeUsed.toLowerCase() !== currentTheme.toLowerCase()) {
    return true;
  }

  // Current model is stronger than the one used last time → re-ping
  const cachedRank  = GEMINI_MODEL_RANK_[cachedRow.geminiModelUsed] || 0;
  const currentRank = GEMINI_MODEL_RANK_[currentModelLabel] || 0;
  if (currentRank > cachedRank && cachedRow.geminiModelUsed) {
    return true;
  }

  return false;
}

  ui.createMenu('Toastmasters')
    .addItem('Start Role Confirmations', 'startRoleConfirmations')
    .addItem('Generate Meeting Agenda', 'generateAgenda')
    .addItem('Draft Club Meeting Email', 'sendClubHypeEmail')
    .addSeparator()
    .addItem('Deploy Code to Another Sheet', 'deployToAnotherSheet')
    .addSeparator()
    .addItem('Update AI Models', 'refreshModelRegistry_')
    .addToUi();
}

/**
 * hexToRgb
 * Converts a hex color string to an {r, g, b} object.
 * @param {string} hex - Hex color string (e.g. "#ff0000" or "ff0000").
 * @return {{r:number, g:number, b:number}|null} RGB object, or null if invalid.
 */
function hexToRgb(hex) {
  const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result
    ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16),
      }
    : null;
}

/**
 * hexToHsl
 * Converts a hex color string to an {h, s, l} object.
 * @param {string} hex - Hex color string.
 * @return {{h:number, s:number, l:number}|null} HSL object (h: 0-360, s/l: 0-100), or null if invalid.
 */
function hexToHsl(hex) {
  const rgb = hexToRgb(hex);
  if (!rgb) return null;

  let { r, g, b } = rgb;
  r /= 255;
  g /= 255;
  b /= 255;

  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h,
    s,
    l = (max + min) / 2;

  if (max === min) {
    h = s = 0;
  } else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h *= 60;
  }

  return { h, s: s * 100, l: l * 100 };
}

/**
 * colorDistance
 * Calculates Euclidean distance between two RGB colors.
 * @param {{r:number, g:number, b:number}} c1 - First RGB color.
 * @param {{r:number, g:number, b:number}} c2 - Second RGB color.
 * @return {number} Distance value (0 = identical, higher = more different).
 */
function colorDistance(c1, c2) {
  return Math.sqrt(
    Math.pow(c1.r - c2.r, 2) + Math.pow(c1.g - c2.g, 2) + Math.pow(c1.b - c2.b, 2)
  );
}

/**
 * isColorMatch
 * Returns true if a hex color is within a threshold distance of a target RGB.
 * @param {string} hex - Hex color string to test.
 * @param {{r:number, g:number, b:number}} targetRgb - Target RGB color.
 * @param {number} [threshold=60] - Max allowed color distance.
 * @return {boolean}
 */
function isColorMatch(hex, targetRgb, threshold = 60) {
  const color = hexToRgb(hex);
  if (!color) return false;
  return colorDistance(color, targetRgb) < threshold;
}

/**
 * isRoughlyGreen
 * Returns true if the hex color falls in the green hue range (used to detect confirmed cells).
 * @param {string} hex - Hex color string.
 * @return {boolean}
 */
function isRoughlyGreen(hex) {
  const hsl = hexToHsl(hex);
  if (!hsl) return false;
  const { h, s, l } = hsl;
  return h >= 75 && h <= 155 && s >= 20 && l >= 10 && l <= 90;
}

/**
 * isRoughlyRed
 * Returns true if the hex color falls in the red hue range (used to detect "unable" cells).
 * Explicitly excludes peach (#fce5cd) which is not a role-status color.
 * @param {string} hex - Hex color string.
 * @return {boolean}
 */
function isRoughlyRed(hex) {
  const hsl = hexToHsl(hex);
  if (!hsl) return false;
  const { h, s, l } = hsl;

  if (hex.toLowerCase() === "#fce5cd") return false; // Excluded peach tone

  return (
    ((h >= 340 || h <= 20) && s >= 30 && l >= 20 && l <= 85) ||
    (h >= 10 && h <= 25 && s >= 50 && l < 60) ||
    (h >= 330 && h <= 350 && s >= 20 && l >= 80) ||
    (h >= 325 && h <= 345 && s >= 20 && l >= 60 && l <= 80) ||
    ((h >= 325 || h <= 10) && s >= 25 && l >= 15 && l <= 35)
  );
}

/**
 * isRoughlyYellow
 * Returns true if the hex color falls in the yellow hue range (used to detect "email sent" cells).
 * @param {string} hex - Hex color string.
 * @return {boolean}
 */
function isRoughlyYellow(hex) {
  const hsl = hexToHsl(hex);
  if (!hsl) return false;
  const { h, s, l } = hsl;

  return (
    (h >= 40 && h <= 65 && s >= 40 && l >= 40 && l <= 90) ||
    (h >= 30 && h <= 70 && s >= 30 && l >= 85) ||
    (h >= 30 && h <= 45 && s >= 50 && l >= 70 && l <= 85)
  );
}

/**
 * isAlreadyColored
 * Returns true if a cell background is green, red, or yellow (i.e. has a role-status color).
 * White and empty backgrounds return false.
 * @param {string} hex - Hex color string from cell background.
 * @return {boolean}
 */
function isAlreadyColored(hex) {
  if (!hex || hex === "#ffffff" || hex === "#fff" || hex === "white") return false;
  return isRoughlyGreen(hex) || isRoughlyRed(hex) || isRoughlyYellow(hex);
}

// ============================================================
// ROLE CONFIRMATIONS — continuation-based flow (no polling loops)
//
// Each step ends by showing a dialog and returning. When the user
// clicks a button in the dialog, google.script.run calls the next
// function directly. This resets the 6-minute execution clock at
// every dialog, so the flow never times out no matter how long
// you take reviewing emails.
//
// Flow:
//   startRoleConfirmations()
//     → shows status summary dialog
//     → user clicks OK  → proceedToEmails()
//       → does all pre-flight prompts (theme, WOTD, sender, intro)
//       → saves state to ScriptProperties
//       → showNextEmail(0)
//         → shows email N dialog
//         → user clicks Send  → handleEmailAction("send", body)
//         → user clicks Skip  → handleEmailAction("skip", "")
//           → sends/colors if needed, advances index
//           → showNextEmail(index+1)  ...repeat...
//             → when done → finishConfirmations()
// ============================================================

// ── Helpers: save/load the session state in ScriptProperties ──
/**
 * normalizeDateParts_
 * Parses a "M/D/YYYY" or "M/D/YY" date string into its component parts,
 * stripping leading zeros so comparisons work regardless of formatting.
 * Shared by startRoleConfirmations and generateAgenda.
 * @param {string} input - Date string like "3/5/2026" or "03/05/26".
 * @return {{m:string, d:string, y:string}} Month, day, and 4-digit year strings.
 */
function normalizeDateParts_(input) {
  const parts = input.split("/");
  const m = parts[0]?.replace(/^0/, "");
  const d = parts[1]?.replace(/^0/, "");
  let y = parts[2];
  if (y) y = y.length === 2 ? "20" + y.replace(/^0/, "") : y.replace(/^0/, "");
  return { m, d, y };
}

/**
 * saveConfirmationState_
 * Persists the role confirmation session state to Script Properties as JSON.
 * @param {Object} state - The full session state object to serialize.
 * @return {void}
 */
function saveConfirmationState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    "_confirmState", JSON.stringify(state)
  );
}

/**
 * loadConfirmationState_
 * Retrieves and deserializes the role confirmation session state from Script Properties.
 * @return {Object|null} The session state object, or null if none is saved.
 */
function loadConfirmationState_() {
  const raw = PropertiesService.getScriptProperties().getProperty("_confirmState");
  return raw ? JSON.parse(raw) : null;
}

/**
 * clearConfirmationState_
 * Deletes the saved role confirmation session state from Script Properties.
 * @return {void}
 */
function clearConfirmationState_() {
  PropertiesService.getScriptProperties().deleteProperty("_confirmState");
}

// ── Step 1: build data, show status summary dialog, then stop ──
/**
 * startRoleConfirmations
 * Entry point for the role confirmation flow. Reads the SCHED sheet, builds a
 * list of assigned roles and their confirmation status, saves state, then shows
 * a modal dialog summarizing the status and collecting meeting metadata (theme,
 * word of the day, sender, intro question). The user clicks "Send Emails" to
 * trigger proceedToEmails() via google.script.run.
 * @return {void}
 */
function startRoleConfirmations() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const schedSheets = sheets
    .map((s) => s.getName())
    .filter((name) => /^SCHED\s\d{4}$/.test(name))
    .sort((a, b) => {
      const yearA = parseInt(a.split(" ")[1]);
      const yearB = parseInt(b.split(" ")[1]);
      return yearB - yearA;
    });

  let selectedSheetName = schedSheets[0];
  let sheet = spreadsheet.getSheetByName(selectedSheetName);

  const confirm = ui.alert(
    `Continue with "${selectedSheetName}" for role confirmations?`,
    ui.ButtonSet.YES_NO
  );

  if (confirm === ui.Button.NO) {
    const input = ui.prompt(
      "Enter the sheet name you want to use:",
      `Available sheets:\n${sheets.map((s) => s.getName()).join("\n")}`,
      ui.ButtonSet.OK_CANCEL
    );
    if (input.getSelectedButton() !== ui.Button.OK) return;
    selectedSheetName = input.getResponseText().trim();
    sheet = spreadsheet.getSheetByName(selectedSheetName);
    if (!sheet) {
      ui.alert(`Sheet "${selectedSheetName}" not found. Please try again.`);
      return;
    }
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();

  // ── Name → Email map ──
  const nameToEmail = {};
  let row = 1;
  while (row < data.length) {
    const firstName = data[row][0];
    const lastName  = data[row][1];
    const email     = data[row][4];
    const bgColor   = backgrounds[row][0];
    if (bgColor === "#cfe2f3" || (!firstName && !email)) break;
    if (firstName && lastName && email) {
      nameToEmail[`${firstName.trim()} ${lastName.trim()}`] = email.trim();
    }
    row++;
  }

  // ── Find Roles header row ──
  let rolesHeaderRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0]?.toString().trim().toLowerCase() === "roles") { rolesHeaderRow = r; break; }
  }
  if (rolesHeaderRow === -1) {
    for (let r = 0; r < data.length; r++) {
      if (data[r][0]?.toString().trim().toLowerCase() === "toastmaster") { rolesHeaderRow = r - 1; break; }
    }
  }
  if (rolesHeaderRow < 0) {
    ui.alert('Could not find the roles header row.');
    return;
  }

  // ── Find first date column ──
  let firstDateCol = -1;
  for (let c = 0; c < data[rolesHeaderRow].length; c++) {
    if (data[rolesHeaderRow][c] instanceof Date) { firstDateCol = c; break; }
  }
  if (firstDateCol === -1) {
    ui.alert('Could not find any dates in the "Roles" row.');
    return;
  }

  // ── Build header map ──
  const headerMap = data[rolesHeaderRow].slice(firstDateCol).map((h, index) => {
    let formatted = h instanceof Date
      ? Utilities.formatDate(h, Session.getScriptTimeZone(), "M/d/yyyy")
      : (h ? h.toString().trim() : "");
    return { isDate: h instanceof Date, formatted, colIndex: firstDateCol + index };
  });

  const now = new Date();
  const formattedDates = headerMap.filter(h => h.isDate && new Date(h.formatted) >= now).map(h => h.formatted);
  const exampleDate = formattedDates[0] || "";

  const response = ui.prompt(
    "Which meeting date?",
    `Available options:\n${formattedDates.join("\n")}\n\nSelect OK to use "${exampleDate}", or enter a different date below:`,
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return;

  let selectedDateRaw = response.getResponseText().trim() || exampleDate;

  const inputParts = normalizeDateParts_(selectedDateRaw);
  const match = headerMap.find(h => {
    const hp = normalizeDateParts_(h.formatted);
    if (inputParts.y) return inputParts.m === hp.m && inputParts.d === hp.d && inputParts.y === hp.y;
    return inputParts.m === hp.m && inputParts.d === hp.d;
  });

  if (!match) {
    ui.alert(`No meeting found for "${selectedDateRaw}".`);
    return;
  }

  const colIndex  = match.colIndex;
  const startRow  = rolesHeaderRow + 1;
  const meetingTheme = rolesHeaderRow > 0
    ? (data[rolesHeaderRow - 1][colIndex]?.toString().trim() || "")
    : "";

  // ── Build confirmations list ──
  const confirmations = [];
  let speechCounter = 1, evaluatorCounter = 1;

  for (let r = startRow; r < data.length; r++) {
    let role = data[r][0]?.toString().trim();
    const assignedName = data[r][colIndex]?.toString().trim();
    if (!role || !assignedName) continue;

    const rawBg = backgrounds[r][colIndex]?.toLowerCase() || "";
    let status = "❓ Needs confirmation";
    if (!rawBg || rawBg === "#ffffff" || rawBg === "#fff" || rawBg === "white") status = "❓ Needs confirmation";
    else if (isRoughlyRed(rawBg))    status = "⚠️ Unable to Attend";
    else if (isRoughlyGreen(rawBg))  status = "✅ Confirmed";
    else if (isRoughlyYellow(rawBg)) status = "📩 Email sent";

    let roleLabel = role, roleType = "general";
    const roleLower = role.toLowerCase();
    if      (roleLower.startsWith("speech"))    { roleLabel = `Speech ${speechCounter++}`;    roleType = "speech"; }
    else if (roleLower.startsWith("evaluator")) { roleLabel = `Evaluator ${evaluatorCounter++}`; roleType = "evaluator"; }
    else if (roleLower === "toastmaster")                                   roleType = "toastmaster";
    else if (roleLower.includes("table topics master"))                     roleType = "tabletopics";
    else if (roleLower.includes("grammarian"))                              roleType = "grammarian";
    else if (roleLower === "timer")                                        roleType = "timer";
    else if (roleLower.includes("joke master"))                             roleType = "jokemaster";
    else if (roleLower.includes("2-minute special") || roleLower.includes("2 minute special")) roleType = "twominute";

    let email = nameToEmail[assignedName];
    let note = "", fuzzyMatched = false;
    if (!email) {
      const [first, last] = assignedName.split(" ");
      const possibleMatch = Object.keys(nameToEmail).find(fn => {
        const [fn1, ln1] = fn.split(" ");
        return fn1 === first && ln1 && last && ln1.includes(last);
      });
      if (possibleMatch) { email = nameToEmail[possibleMatch]; note = `Suggested match: "${possibleMatch}"`; fuzzyMatched = true; }
      else note = "No email found";
    }

    confirmations.push({ role: roleLabel, roleType, name: assignedName, email: email || "",
      note, status, fuzzyMatched, rowIndex: r, colIndex, currentBg: rawBg });
  }

  if (confirmations.length === 0) { ui.alert("No roles assigned for that date."); return; }

  // ── Save state for next steps ──
  saveConfirmationState_({
    sheetName: selectedSheetName,
    selectedDateRaw,
    meetingTheme,
    confirmations,
    themeRowIndex: rolesHeaderRow - 1,  // row where theme lives (one above Roles header)
    themeColIndex: colIndex,             // same column as selected meeting date
  });

  // ── Build status summary ──
  let summaryHtml = "<ul style='padding-left:18px;margin:0 0 12px 0;'>";
  confirmations.forEach(entry => {
    summaryHtml += `<li style="margin-bottom:6px;"><strong>${entry.name}</strong> — ${entry.role}<br>
      <small style="color:#555;">Status: ${entry.status} | Email: ${entry.email || "(No email found)"}</small>`;
    if (entry.note) summaryHtml += `<br><small style="color:#888;font-style:italic;">${entry.note}</small>`;
    summaryHtml += "</li>";
  });
  summaryHtml += "</ul>";

  // TBD warnings inline
  const tbdWarnings = confirmations
    .filter(e => e.name.toUpperCase() === "TBD")
    .map(e => `<div style="color:#b94a48;font-size:12px;margin-bottom:4px;">⚠️ "${e.role}" still needs to be filled.</div>`)
    .join("");

  // Detect whether grammarian / speeches exist for conditional fields
  const hasGrammarian = confirmations.some(e => e.roleType === "grammarian" && e.name.toUpperCase() !== "TBD");
  const hasSpeeches   = confirmations.some(e => e.roleType === "speech");
  const toastmasterEntry = confirmations.find(e => e.roleType === "toastmaster");
  const toastmasterName  = toastmasterEntry ? toastmasterEntry.name : "";

  // ── WOD Memory: pre-fill from cache ──
  const wodCacheConf_ = lookupWodCache_(selectedDateRaw);
  const wodPrefillConf_ = wodCacheConf_ ? wodCacheConf_.word : "";
  const wotdField = hasGrammarian ? `
    <div style="margin-bottom:10px;">
      <label style="font-weight:bold;display:block;margin-bottom:3px;">Word of the Day</label>
      <input id="wotd" type="text" placeholder="Enter Word of the Day" value="${wodPrefillConf_}" style="width:100%;box-sizing:border-box;padding:5px;font-size:13px;border:1px solid #ccc;border-radius:3px;">
    </div>` : "";

  const introField = hasSpeeches ? `
    <div style="margin-bottom:10px;">
      <label style="font-weight:bold;display:block;margin-bottom:3px;">Speaker Intro Question <span style="font-weight:normal;color:#888;">(optional)</span></label>
      <input id="introQ" type="text" placeholder="e.g. What is something most people don&#39;t know about you?" style="width:100%;box-sizing:border-box;padding:5px;font-size:13px;border:1px solid #ccc;border-radius:3px;">
    </div>` : "";

  const senderField = toastmasterName ? `
    <div style="margin-bottom:10px;">
      <label style="font-weight:bold;display:block;margin-bottom:3px;">Who is sending these emails?</label>
      <label style="display:block;margin-bottom:4px;">
        <input type="radio" name="senderType" value="tm" checked onchange="toggleSenderName(this)"> ${toastmasterName} (Toastmaster)
      </label>
      <label style="display:block;margin-bottom:4px;">
        <input type="radio" name="senderType" value="other" onchange="toggleSenderName(this)"> Someone else
      </label>
      <input id="senderName" type="text" placeholder="Enter sender name" style="width:100%;box-sizing:border-box;padding:5px;font-size:13px;border:1px solid #ccc;border-radius:3px;display:none;margin-top:4px;">
    </div>` : "";

  const formatField = `
    <div style="margin-bottom:10px;">
      <label style="font-weight:bold;display:block;margin-bottom:5px;">Meeting Format</label>
      <label style="display:inline-flex;align-items:center;margin-right:14px;cursor:pointer;">
        <input type="radio" name="meetingFormat" value="hybrid" onchange="toggleFormatDetails()"> &nbsp;Hybrid
      </label>
      <label style="display:inline-flex;align-items:center;margin-right:14px;cursor:pointer;">
        <input type="radio" name="meetingFormat" value="in_person" onchange="toggleFormatDetails()"> &nbsp;In Person <!-- BUG-3 FIX: Added In Person option -->
      </label>
      <label style="display:inline-flex;align-items:center;margin-right:14px;cursor:pointer;">
        <input type="radio" name="meetingFormat" value="virtual" onchange="toggleFormatDetails()"> &nbsp;Virtual
      </label>
      <label style="display:inline-flex;align-items:center;cursor:pointer;">
        <input type="radio" name="meetingFormat" value="undecided" checked onchange="toggleFormatDetails()"> &nbsp;Undecided
      </label>
      <div id="addressField" style="display:none;margin-top:7px;">
        <input id="meetingAddress" type="text" value="633 Folsom Street, San Francisco CA" style="width:100%;box-sizing:border-box;padding:5px;font-size:13px;border:1px solid #ccc;border-radius:3px;">
      </div>
    </div>`;

  const dialogHtml = `
    <div style="font-family:Arial,sans-serif;font-size:13px;padding:8px 12px;max-height:520px;overflow-y:auto;">
      <p style="font-weight:bold;margin:0 0 6px 0;">Confirmation status:</p>
      ${tbdWarnings}
      ${summaryHtml}
      <hr style="margin:10px 0;border:none;border-top:1px solid #ddd;">
      <div style="margin-bottom:10px;">
        <label style="font-weight:bold;display:block;margin-bottom:3px;">Meeting Theme</label>
        <input id="theme" type="text" value="${meetingTheme.replace(/"/g, '&quot;')}" style="width:100%;box-sizing:border-box;padding:5px;font-size:13px;border:1px solid #ccc;border-radius:3px;">
      </div>
      ${formatField}
      ${wotdField}
      ${senderField}
      ${introField}
      <div style="text-align:right;margin-top:14px;">
        <button onclick="cancel()" style="padding:6px 14px;margin-right:8px;cursor:pointer;border:1px solid #ccc;border-radius:3px;background:#fff;">Cancel</button>
        <button id="sendBtn" onclick="proceed()" style="padding:6px 16px;background:#4a86e8;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;">Send Emails →</button>
      </div>
    </div>
    <script>
      function toggleSenderName(radio) {
        document.getElementById('senderName').style.display = radio.value === 'other' ? 'block' : 'none';
      }
      function toggleFormatDetails() {
        var fmt = document.querySelector('input[name=\"meetingFormat\"]:checked');
        var val = fmt ? fmt.value : 'undecided';
        document.getElementById('addressField').style.display = (val === 'hybrid' || val === 'in_person') ? 'block' : 'none' // BUG-3 FIX: show address for In Person too;
      }
      function proceed() {
        document.getElementById('sendBtn').disabled = true;
        document.getElementById('sendBtn').textContent = 'Loading...';
        const senderRadio = document.querySelector('input[name="senderType"]:checked');
        const senderType  = senderRadio ? senderRadio.value : 'tm';
        var fmtRadio = document.querySelector('input[name="meetingFormat"]:checked');
        var meetingFormat = fmtRadio ? fmtRadio.value : 'undecided';
        const formData = {
          theme:          (document.getElementById('theme')          || {value:''}).value.trim(),
          wotd:           (document.getElementById('wotd')           || {value:''}).value.trim(),
          introQ:         (document.getElementById('introQ')         || {value:''}).value.trim(),
          senderType:     senderType,
          senderName:     (document.getElementById('senderName')     || {value:''}).value.trim(),
          meetingFormat:  meetingFormat,
          meetingAddress: (document.getElementById('meetingAddress') || {value:''}).value.trim()
        };
        google.script.run
          .withSuccessHandler(function(data) {
            if (data && data.error) { alert(data.error); return; }
            if (data && data.done) { google.script.host.close(); return; }
            renderEmailDialog(data);
          })
          .withFailureHandler(function(err) {
            document.getElementById('sendBtn').disabled = false;
            document.getElementById('sendBtn').textContent = 'Send Emails →';
            alert('Error: ' + (err.message || JSON.stringify(err)));
          })
          .proceedToEmails(formData);
      }
      function cancel() {
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).cancelConfirmations();
      }

      // ── Renders email review dialog inside THIS dialog (no new window) ──
      function renderEmailDialog(data) {
        document.title = data.title || 'Email Draft';
        window._standardBody  = data.body;
        window._thankYouBody  = data.thankYouBody || '';
        var toggleHtml = '';
        if (data.thankYouBody) {
          toggleHtml = '<div style="margin-bottom:8px;background:#e8f5e9;border:1px solid #a8d5b5;border-radius:4px;padding:7px 10px;font-size:12px;">' +
            '<strong style="color:#2d6a3f;">Email type:</strong>&nbsp;&nbsp;' +
            '<label style="cursor:pointer;margin-right:12px;"><input type="radio" name="emailType" value="standard" checked onchange="switchEmailType()"> Standard confirmation</label>' +
            '<label style="cursor:pointer;"><input type="radio" name="emailType" value="thankyou" onchange="switchEmailType()"> Thank you (already confirmed)</label>' +
            '</div>';
        }
        document.body.innerHTML = \`
          <div style="font-family:Arial,sans-serif;font-size:13px;padding:8px;max-height:520px;overflow-y:auto;">
            <p style="margin:0 0 2px;color:#888;font-size:11px;">Email \${data.current} of \${data.total}</p>
            \${data.banner || ''}
            \${toggleHtml}
            <p style="margin:0 0 4px;"><strong>To:</strong> \${data.to}</p>
            <p style="margin:0 0 12px;"><strong>Subject:</strong> \${data.subject}</p>
            <p style="margin:0 0 4px;">Edit the message below if needed:</p>
            <textarea id="emailBody" style="width:100%;height:220px;font-size:12px;font-family:Arial,sans-serif;padding:6px;box-sizing:border-box;">\${data.body}</textarea>
            <div style="text-align:right;margin-top:10px;">
              <button id="skipBtn" onclick="doAction('skip')" style="padding:6px 14px;margin-right:8px;cursor:pointer;border:1px solid #ccc;border-radius:3px;">Skip</button>
              <button id="sendBtn" onclick="doAction('send')" style="padding:6px 16px;background:#4a86e8;color:white;border:none;border-radius:4px;cursor:pointer;">Store as Draft</button>
            </div>
          </div>\`;

        window.switchEmailType = function() {
          var sel = document.querySelector('input[name="emailType"]:checked');
          document.getElementById('emailBody').value = (sel && sel.value === 'thankyou') ? window._thankYouBody : window._standardBody;
        };

        window.doAction = function(action) {
          document.getElementById('sendBtn').disabled = true;
          document.getElementById('skipBtn').disabled = true;
          document.getElementById('sendBtn').textContent = action === 'send' ? 'Creating draft...' : 'Skipping...';
          // Pre-open a blank tab synchronously to avoid popup blocker
          const body = action === 'send' ? document.getElementById('emailBody').value : '';
          google.script.run
            .withSuccessHandler(function(next) {
              if (!next || next.done) {
                showDoneScreen(null, null);
                finishUp();
              } else {
                renderEmailDialog(next);
              }
            })
  .withFailureHandler(function(err) {
    alert('Error: ' + (err.message || JSON.stringify(err)));
    document.getElementById('sendBtn').disabled = false;
    document.getElementById('skipBtn').disabled = false;
  })
  .handleEmailAction(action, body);
        };
      }

     function finishUp() {
        google.script.run
          .withSuccessHandler(function(result) {
            if (result && result.tmInfo) {
              google.script.run
                .withSuccessHandler(function(tmDraftUrl) {
                  if (tmDraftUrl) showDoneScreen(tmDraftUrl, result.tmInfo.name);
                })
                .withFailureHandler(function() {})
                .createTmDraft_(result.tmInfo);
            }
          })
          .withFailureHandler(function() {})
          .finishConfirmations_();
      }

      function showDoneScreen(tmDraftUrl, tmName) {
        const tmSection = tmDraftUrl
          ? \`<p>A reminder draft for <strong>\${tmName}</strong> is ready.</p>
             <p><a href="\${tmDraftUrl}" target="_blank" style="color:#4a86e8;font-weight:bold;">Open TM Reminder Draft →</a></p>\`
          : '';
        document.body.innerHTML = \`
          <div style="font-family:Arial,sans-serif;padding:20px;font-size:13px;text-align:center;">
            <p style="font-size:18px;margin-bottom:8px;">✅ All done!</p>
            <p>All drafts have been saved.</p>
            <p><a href="https://mail.google.com/mail/#drafts" target="_blank" style="color:#4a86e8;font-weight:bold;font-size:14px;">Open Gmail Drafts →</a></p>
            \${tmSection}
            <button onclick="google.script.host.close()" style="margin-top:14px;padding:8px 20px;background:#4a86e8;color:white;border:none;border-radius:4px;cursor:pointer;font-size:13px;">Close</button>
          </div>\`;
      }
    <\/script>`;

  ui.showModalDialog(
    HtmlService.createHtmlOutput(dialogHtml).setWidth(460).setHeight(560),
    "Pending Role Confirmations"
  );
  // Script ends here — proceedToEmails(formData) will be called by the dialog
}

// ── Step 2: pre-flight prompts, save full state, show first email ──
/**
 * proceedToEmails
 * Step 2 of the confirmation flow. Called by the status dialog via google.script.run.
 * Reads form data (theme, WOTD, sender, intro question), groups confirmations by email
 * address, builds email bodies, saves updated state, and returns the first email dialog content.
 * @param {{theme:string, wotd:string, introQ:string, senderType:string, senderName:string, meetingFormat:string, zoomLink:string}} formData
 * @return {Object} First email dialog content object from getEmailDialogContent_(0).
 */
function proceedToEmails(formData) {
  const state = loadConfirmationState_();
  if (!state) return;

  const { confirmations, meetingTheme, selectedDateRaw, themeRowIndex, themeColIndex } = state;

  // Extract values from form
  // Normalize theme and WOTD to sentence case (first letter capitalized, rest lowercase)
  // so that regardless of what the user types (ALL CAPS, all lower, Mixed), the text
  // appears consistently in emails.
  function toSentenceCase_(s) {
    if (!s) return s;
    return s.charAt(0).toUpperCase() + s.slice(1).toLowerCase();
  }
  const confirmedTheme = (formData && formData.theme) ? toSentenceCase_(formData.theme.trim()) : toSentenceCase_(meetingTheme);
  const wordOfTheDay   = (formData && formData.wotd)  ? toSentenceCase_(formData.wotd.trim())  : "";
  const introQuestion        = (formData && formData.introQ)     ? formData.introQ     : "";
  // Persist so generateAgenda can retrieve it without asking again
  if (introQuestion) {
    PropertiesService.getScriptProperties().setProperty("LAST_INTRO_QUESTION", introQuestion);
  }
  const senderType           = (formData && formData.senderType) ? formData.senderType : "tm";
  const meetingFormat        = (formData && formData.meetingFormat)  ? formData.meetingFormat  : "undecided";
  const meetingAddress       = (formData && formData.meetingAddress) ? formData.meetingAddress : "633 Folsom Street, San Francisco CA";

  // Build location line and attendance note based on meeting format.
  // For hybrid, the address is folded into the attendance note rather than
  // being a standalone location block, so recipients see it alongside the
  // attendance sheet prompt.
  let locationLine;
  let attendanceNote;
  if (meetingFormat === "hybrid") {
    locationLine   = ""; // hybrid address appears in attendanceNote instead
    attendanceNote = "We'll have a hybrid meeting with our in-person location at " + meetingAddress +
      ". Please update the attendance sheet if you haven't already. This will let us know whether you'll be joining in person or virtually.";
  } else if (meetingFormat === "virtual") {
    // BUG-4 FIX: virtual meetings don't need the attendance sheet prompt (everyone is on Zoom).
    locationLine   = "We'll be meeting virtually on Zoom.";
    attendanceNote = "";
  } else if (meetingFormat === "in_person") { // BUG-3 FIX: Added In Person meeting format
    // BUG-4 FIX: in-person meetings don't need the attendance sheet prompt (everyone is in the room).
    locationLine  = "We'll be meeting in person at " + meetingAddress + ".";
    attendanceNote = "";
  } else {
    // BUG-4 FIX: undecided — acknowledge format is TBD, no attendance prompt.
    locationLine   = "Heads up: the meeting format (in person, virtual, or hybrid) is still being decided — we'll confirm before the meeting.";
    attendanceNote = "";
  }

  const toastmasterEntry     = confirmations.find(e => e.roleType === "toastmaster");
  const toastmasterName      = toastmasterEntry ? toastmasterEntry.name : "the Toastmaster";

  let senderIsToastmaster = (senderType === "tm");
  let senderName = senderIsToastmaster
    ? toastmasterName
    : ((formData && formData.senderName) ? formData.senderName : toastmasterName) || toastmasterName;

  // Write theme back to sheet if changed
  if (confirmedTheme && confirmedTheme !== meetingTheme && themeRowIndex >= 0) {
    try {
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(state.sheetName);
      if (sheet) sheet.getRange(themeRowIndex + 1, themeColIndex + 1).setValue(confirmedTheme);
    } catch(e) {}
  }


  // ── Build email groups: one group per unique email address ──
  // Entries sharing an email address are merged into one combined email.
  const eligibleEntries = confirmations.filter(e =>
    e.roleType !== "toastmaster" && e.name.toUpperCase() !== "TBD" && e.email
  );

  const groupMap = {};
  eligibleEntries.forEach(entry => {
    if (!groupMap[entry.email]) groupMap[entry.email] = [];
    groupMap[entry.email].push(entry);
  });

  const emailGroups = Object.values(groupMap).map(entries => {
    const roles = entries.map(e => e.role);

    // Subject line: list both roles for 2, fall back to "Multiple Roles" for 3+
    let subject;
    if (roles.length === 1) {
      subject = `Sierra Speakers Toastmasters: ${roles[0]} Confirmation for ${selectedDateRaw}`;
    } else if (roles.length === 2) {
      subject = `Sierra Speakers Toastmasters: ${roles[0]} & ${roles[1]} Confirmation for ${selectedDateRaw}`;
    } else {
      subject = `Sierra Speakers Toastmasters: Multiple Roles Confirmation for ${selectedDateRaw}`;
    }

    // Build body: single-role uses existing builder; multi-role uses combined builder
    let body;
    if (entries.length === 1) {
      // BUG-1 FIX: was toastmasterName - use senderName for email sign-off
      body = buildEmailBody(entries[0], confirmedTheme, wordOfTheDay, {}, confirmations, senderName, selectedDateRaw, introQuestion, locationLine, attendanceNote);
    } else {
      // BUG-1 FIX: was toastmasterName - use senderName for email sign-off
      body = buildCombinedEmailBody_(entries, confirmedTheme, wordOfTheDay, confirmations, senderName, selectedDateRaw, introQuestion, locationLine, attendanceNote);
    }

    // Build the thank-you alternative (shown in dialog when all entries are already green)
    // BUG-1 FIX: was toastmasterName - use senderName for email sign-off
    const thankYouBody = buildThankYouEmailBody_(entries, confirmedTheme, wordOfTheDay, senderName, selectedDateRaw, locationLine, attendanceNote);

    return {
      email: entries[0].email,
      name: entries[0].name,
      entries,
      subject,
      body,
      thankYouBody,
    };
  });

  // Save full state and start email loop
  saveConfirmationState_({
    sheetName: state.sheetName,
    selectedDateRaw,
    confirmedTheme,
    wordOfTheDay,
    toastmasterName,
    senderIsToastmaster,
    senderName,
    introQuestion,
    meetingFormat,
    meetingAddress,
    locationLine,
    attendanceNote,
    confirmations,
    emailGroups,
    emailIndex: 0,
  });

  // Return first email dialog content to client
  return getEmailDialogContent_(0);
}

// ── Returns email dialog HTML for group at index i (no ui calls) ──
/**
 * getEmailDialogContent_
 * Builds and returns a plain data object representing the email dialog for a given group index.
 * Called by proceedToEmails and handleEmailAction to feed the client-side renderEmailDialog().
 * @param {number} index - Zero-based index into emailGroups array.
 * @return {{done:boolean, index:number, total:number, current:number, title:string,
 *           to:string, subject:string, body:string, banner:string}|{done:true, error?:string}}
 */
function getEmailDialogContent_(index) {
  const state = loadConfirmationState_();
  if (!state) return { done: true, error: "Session expired." };

  const { emailGroups, selectedDateRaw } = state;

  if (index >= emailGroups.length) {
    return { done: true };
  }

  const group   = emailGroups[index];
  const total   = emailGroups.length;
  const current = index + 1;

  const safeBody = group.body
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  const allConfirmed  = group.entries.every(e => isRoughlyGreen(e.currentBg));
  const someConfirmed = !allConfirmed && group.entries.some(e => isRoughlyGreen(e.currentBg));

  // Offer the thank-you alternative only when all entries are already green
  const safeThankYouBody = allConfirmed && group.thankYouBody
    ? group.thankYouBody.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    : null;

  const banner = allConfirmed
    ? `<div style="background:#e6f4ea;border:1px solid #a8d5b5;border-radius:4px;padding:7px 10px;margin-bottom:10px;font-size:12px;color:#2d6a3f;">✅ <strong>${group.name}</strong> has already confirmed all roles — their cell(s) are green. You can still send or skip.</div>`
    : someConfirmed
    ? `<div style="background:#fff8e1;border:1px solid #ffe082;border-radius:4px;padding:7px 10px;margin-bottom:10px;font-size:12px;color:#7a5800;">⚠️ <strong>${group.name}</strong> has confirmed some roles but not all.</div>`
    : "";

  const title = group.entries.length > 1
    ? `📧 ${group.name} (${group.entries.map(e => e.role).join(" & ")})`
    : `📧 ${group.name} (${group.entries[0].role})`;

  return {
    done:         false,
    index:        index,
    total:        total,
    current:      current,
    title:        title,
    to:           group.email,
    subject:      group.subject,
    body:         safeBody,
    thankYouBody: safeThankYouBody,
    banner:       banner
  };
}

// ── Step 3: show email dialog for group at index i ──
/**
 * showNextEmail_
 * Legacy version of the email dialog flow that opened a new modal per email.
 * Kept for safety — the active flow uses renderEmailDialog() on the client side instead.
 * @param {number} startIndex - Zero-based index into emailGroups.
 * @param {string|null} openDraftUrl - URL of a just-created draft to auto-open, or null.
 * @return {void}
 */
function showNextEmail_(startIndex, openDraftUrl) {
  const ui    = SpreadsheetApp.getUi();
  const state = loadConfirmationState_();
  if (!state) { ui.alert("Session expired. Please restart."); return; }


  const { emailGroups, selectedDateRaw } = state;

  // Warn about any skipped entries with no email (only on first pass)
  if (startIndex === 0) {
    const { confirmations } = state;
    confirmations.forEach(e => {
      if (e.roleType !== "toastmaster" && e.name.toUpperCase() !== "TBD" && !e.email) {
        ui.alert(`⚠️ Skipping ${e.role} — no email found for "${e.name}".`);
      }
    });
  }

  // If we've exhausted all groups, finish up
  if (startIndex >= emailGroups.length) {
    finishConfirmations_();
    return;
  }

  state.emailIndex = startIndex;
  saveConfirmationState_(state);

  const group   = emailGroups[startIndex];
  const total   = emailGroups.length;
  const current = startIndex + 1;

  const safeBody = group.body
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

  // Green cell banner: show if ALL entries in group are already green
  const allConfirmed = group.entries.every(e => isRoughlyGreen(e.currentBg));
  const someConfirmed = !allConfirmed && group.entries.some(e => isRoughlyGreen(e.currentBg));
  const confirmedBanner = allConfirmed
    ? `<div style="background:#e6f4ea;border:1px solid #a8d5b5;border-radius:4px;padding:7px 10px;margin-bottom:10px;font-size:12px;color:#2d6a3f;">
        ✅ <strong>${group.name}</strong> has already confirmed all roles — their cell(s) are green. You can still send or skip.
       </div>`
    : someConfirmed
    ? `<div style="background:#fff8e1;border:1px solid #ffe082;border-radius:4px;padding:7px 10px;margin-bottom:10px;font-size:12px;color:#7a5800;">
        ⚠️ <strong>${group.name}</strong> has confirmed some roles (green) but not all. You can still send or skip.
       </div>`
    : "";

  const dialogTitle = group.entries.length > 1
    ? `📧 Draft Email — ${group.name} (${group.entries.map(e => e.role).join(" & ")})`
    : `📧 Draft Email — ${group.name} (${group.entries[0].role})`;

  const emailHtml = `
    <div style="font-family:Arial,sans-serif;font-size:13px;padding:8px;">
      <p style="margin:0 0 2px;color:#888;font-size:11px;">Email ${current} of ${total}</p>
      ${confirmedBanner}
      <p style="margin:0 0 4px;"><strong>To:</strong> ${group.email}</p>
      <p style="margin:0 0 12px;"><strong>Subject:</strong> ${group.subject}</p>
      <p style="margin:0 0 4px;">Edit the message below if needed:</p>
      <textarea id="emailBody" style="width:100%;height:260px;font-size:12px;font-family:Arial,sans-serif;padding:6px;box-sizing:border-box;">${safeBody}</textarea>
      <div style="text-align:right;margin-top:10px;">
        <button onclick="doSkip()" style="padding:6px 14px;margin-right:8px;cursor:pointer;">Skip</button>
        <button onclick="doSend()" style="padding:6px 16px;background:#4a86e8;color:white;border:none;border-radius:4px;cursor:pointer;">Send</button>
      </div>
    </div>
    <script>
      function doSend() {
        const body = document.getElementById("emailBody").value;
        document.querySelectorAll("button").forEach(b => b.disabled = true);
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).handleEmailAction("send", body);
      }
      function doSkip() {
        document.querySelectorAll("button").forEach(b => b.disabled = true);
        google.script.run.withSuccessHandler(function() {
          google.script.host.close();
        }).handleEmailAction("skip", "");
      }
    <\/script>`;

  // If a draft was just created, inject a script to open it
  const draftScript = openDraftUrl
    ? `<script>window.open(${JSON.stringify(openDraftUrl)}, '_blank');<\/script>`
    : "";

  ui.showModalDialog(
    HtmlService.createHtmlOutput(draftScript + emailHtml).setWidth(500).setHeight(430),
    dialogTitle
  );
}

// ── Step 4: called by dialog Send/Skip buttons ──
/**
 * handleEmailAction
 * Processes the user's Send or Skip action for the current email group.
 * On "send": builds the fancy HTML body, creates a Gmail draft, and colors cells yellow.
 * On "skip": advances the index without sending.
 * Returns the next email dialog content object (or {done:true} when finished).
 * @param {"send"|"skip"} action - The user's chosen action.
 * @param {string} body - The (possibly edited) plain-text email body.
 * @return {Object} Next email dialog content from getEmailDialogContent_, with draftUrl added.
 */
function handleEmailAction(action, body) {
  const state = loadConfirmationState_();
  if (!state) return;

  const { emailGroups, emailIndex, sheetName } = state;
  const group = emailGroups[emailIndex];

  let draftUrl = null;

  if (action === "send") {
    const allRoles = group.entries.map(e => e.role);
    const htmlBody = buildFancyHtml_(body, allRoles);

    // Create Gmail draft via API
    draftUrl = createGmailDraft_(group.email, group.subject, body, htmlBody);

    // Color all cells yellow immediately
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (sheet) {
      group.entries.forEach(entry => {
        if (!isAlreadyColored(entry.currentBg)) {
          sheet.getRange(entry.rowIndex + 1, entry.colIndex + 1).setBackground("#ffff00");
        }
      });
    }
  }

  // Save next index
  state.emailIndex = emailIndex + 1;
  saveConfirmationState_(state);

  // Return next email data + draft URL to client
  const nextData = getEmailDialogContent_(emailIndex + 1);
  nextData.draftUrl = draftUrl;
  return nextData;
}

// ── Create a Gmail draft via the Gmail REST API ──
/**
 * createGmailDraft_
 * Creates a Gmail draft with both plain text and HTML bodies.
 * Returns the Gmail drafts URL for the active user, or a fallback URL on error.
 * @param {string} to - Recipient email address.
 * @param {string} subject - Email subject line.
 * @param {string} plainText - Plain text body.
 * @param {string} htmlBody - HTML body.
 * @return {string} URL to Gmail drafts folder.
 */
function createGmailDraft_(to, subject, plainText, htmlBody) {
  const fallback = "https://mail.google.com/mail/#drafts";
  try {
    GmailApp.createDraft(to, subject, plainText, { htmlBody: htmlBody });
    const userEmail = Session.getActiveUser().getEmail();
    return userEmail
      ? "https://mail.google.com/mail/?authuser=" + encodeURIComponent(userEmail) + "#drafts"
      : fallback;
  } catch(e) {
    Logger.log("createGmailDraft_ error: " + e.toString());
    return fallback;  // ← always returns a URL now
  }
}

// ── Step 5: clear state, return TM info so client can create draft separately ──
/**
 * finishConfirmations_
 * Called after all email groups have been processed.
 * Clears session state and, if the sender is not the Toastmaster, returns a tmInfo
 * object so the client can separately call createTmDraft_() to notify the TM.
 * @return {{done:true, tmInfo: Object|null}}
 */
function finishConfirmations_() {
  const state = loadConfirmationState_();
  if (!state) return { done: true, needsTmDraft: false };

  const { confirmations, confirmedTheme, selectedDateRaw,
          senderIsToastmaster, senderName, toastmasterName } = state;

  let tmInfo = null;
  if (!senderIsToastmaster) {
    const toastmasterEntry = confirmations.find(e => e.roleType === "toastmaster");
    if (toastmasterEntry && toastmasterEntry.email) {
      const tmFirstName = toastmasterName.split(" ")[0];
      const tmBody =
        `Hi ${tmFirstName},\n\n` +
        `This is a quick heads-up that ${senderName} has just sent out role confirmation emails to the team on your behalf for the ${selectedDateRaw} meeting.\n\n` +
        (confirmedTheme ? `The meeting theme is: "${confirmedTheme}"\n\n` : "") +
        `If you have any questions or need to make changes, please reach out to your team directly.\n\n` +
        `See you at the meeting!\n\n` +
        `— ${senderName}`;
      tmInfo = {
        email:   toastmasterEntry.email,
        subject: `Sierra Speakers Toastmasters: Role Confirmations Sent for ${selectedDateRaw}`,
        body:    tmBody,
        name:    tmFirstName
      };
    }
  }

  clearConfirmationState_();
  return { done: true, tmInfo: tmInfo };
}

// ── Creates TM reminder draft — called separately by client after finishConfirmations_ ──
/**
 * createTmDraft_
 * Builds a fancy HTML version of the TM notification email and creates a Gmail draft.
 * Called by the client-side finishUp() after finishConfirmations_() returns tmInfo.
 * @param {{email:string, subject:string, body:string, name:string}|null} tmInfo
 * @return {string|null} Gmail drafts URL, or null if tmInfo is missing.
 */
function createTmDraft_(tmInfo) {
  if (!tmInfo) return null;
  const tmHtml = buildFancyHtml_(tmInfo.body, []);
  return createGmailDraft_(tmInfo.email, tmInfo.subject, tmInfo.body, tmHtml);
}

// ── Called by Cancel button in status dialog ──
/**
 * cancelConfirmations
 * Clears saved session state when the user clicks Cancel in the status dialog.
 * @return {void}
 */
function cancelConfirmations() {
  clearConfirmationState_();
}

// ── Legacy callbacks kept for safety ──
function acknowledgePendingRoles() {}
function setEmailResponse(action, body) {}

// =============================

// =============================
// Fancy HTML email with Sierra Speakers branding
// =============================
/**
 * buildFancyHtml_
 * Wraps a plain-text email body in Sierra Speakers branded HTML.
 * Bolds any role label strings found in the body and hyperlinks "scheduling sheet".
 * @param {string} plainText - The plain-text email body.
 * @param {string[]} roleLabels - Array of role label strings to bold in the output.
 * @return {string} Complete HTML email string.
 */
function buildFancyHtml_(plainText, roleLabels) {
  const NAVY  = "#1B2A4A";
  const GREEN = "#1E5631";

  let escaped = plainText
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  const roles = Array.isArray(roleLabels) ? roleLabels : (roleLabels ? [roleLabels] : []);
  roles.forEach(role => {
    if (role) {
      const safeRole = role.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      escaped = escaped.replace(
        new RegExp("(" + safeRole + ")", "g"),
        '<strong style="color:' + GREEN + ';">$1</strong>'
      );
    }
  });

  const sheetUrl = SCHEDULING_SHEET_URL_ || SpreadsheetApp.getActiveSpreadsheet().getUrl();
  let bodyHtml = escaped.replace(/\n/g, "<br>");
  bodyHtml = bodyHtml.replace(
    "scheduling sheet",
    "<a href=\"" + sheetUrl + "\" style=\"color:#1E5631;\">scheduling sheet</a>"
  );
  bodyHtml = bodyHtml.replace(
    "attendance sheet",
    "<a href=\"" + ATTENDANCE_SHEET_URL_ + "\" style=\"color:#1E5631;\">attendance sheet</a>"
  );

  return "<!DOCTYPE html>" +
    "<html><head><meta charset=\"UTF-8\"></head>" +
    "<body style=\"margin:0;padding:0;font-family:Arial,sans-serif;background:#ffffff;\">" +
    "<table width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background:#ffffff;\">" +
    "<tr><td style=\"background:" + NAVY + ";padding:20px 24px;text-align:center;\">" +
    "<div style=\"color:#ffffff;font-size:22px;font-weight:bold;letter-spacing:1px;\">Sierra Speakers</div>" +
    "<div style=\"color:#A8D8F0;font-size:12px;margin-top:5px;letter-spacing:3px;\">TOASTMASTERS</div>" +
    "<div style=\"width:48px;height:3px;background:" + GREEN + ";margin:12px auto 0;border-radius:2px;\"></div>" +
    "</td></tr>" +
    "<tr><td style=\"padding:24px;color:#222222;font-size:15px;line-height:1.8;\">" + bodyHtml + "</td></tr>" +
    "<tr><td style=\"background:#f0f4f8;padding:16px 24px;text-align:center;border-top:3px solid " + GREEN + ";\">" +
    "<div style=\"color:" + NAVY + ";font-size:12px;font-weight:bold;letter-spacing:1px;\">SIERRA SPEAKERS TOASTMASTERS</div>" +
    "<div style=\"color:#888888;font-size:11px;margin-top:4px;\">Building confident communicators</div>" +
    "</td></tr>" +
    "</table></body></html>";
}

// =============================
// Wrap role names in asterisks for plain-text emphasis
// e.g. "Evaluator 1" → "*Evaluator 1*"
// =============================
/**
 * boldRolesWithAsterisks_
 * Wraps each role label string in asterisks for plain-text emphasis.
 * @param {string} text - The text to process.
 * @param {string[]} roleLabels - Role label strings to wrap.
 * @return {string} Text with role labels wrapped in asterisks.
 */
function boldRolesWithAsterisks_(text, roleLabels) {
  let result = text;
  const roles = Array.isArray(roleLabels) ? roleLabels : (roleLabels ? [roleLabels] : []);
  roles.forEach(role => {
    if (role) {
      const safeRole = role.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      result = result.replace(new RegExp(`(${safeRole})`, "g"), "*$1*");
    }
  });
  return result;
}
// =============================
// Combined Email Body Builder (multiple roles, same person)
// =============================
/**
 * buildCombinedEmailBody_
 * Builds a single flowing plain-text email for a member with multiple roles.
 * Each role gets a short labelled blurb — no repeated greeting/location/sign-off.
 * @param {Object[]} entries - Array of confirmation entry objects for this member.
 * @param {string} theme - Meeting theme.
 * @param {string} wordOfTheDay - Word of the Day (for Grammarian).
 * @param {Object[]} allConfirmations - Full confirmations list (used for evaluator matching).
 * @param {string} toastmasterName - Full name of the Toastmaster (used as sign-off).
 * @param {string} meetingDate - Meeting date string (e.g. "3/6/2026").
 * @param {string} introQuestion - Optional speaker intro question.
 * @param {string} [locationLine] - Pre-built location string.
 * @param {string} [attendanceNote] - Attendance sheet note for hybrid/virtual.
 * @return {string} Combined plain-text email body.
 */
function buildCombinedEmailBody_(entries, theme, wordOfTheDay, allConfirmations, toastmasterName, meetingDate, introQuestion, locationLine, attendanceNote) {
  const firstName    = entries[0].name.split(" ")[0];
  const roleList     = entries.map(e => e.role);
  const sheetUrl     = SCHEDULING_SHEET_URL_ || SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const locationStr  = locationLine  || "";
  const attendanceStr = attendanceNote || "";

  const weekLabel = (() => {
    try {
      const today = new Date(); today.setHours(0,0,0,0);
      const md = new Date(meetingDate); md.setHours(0,0,0,0);
      const days = Math.round((md - today) / (1000 * 60 * 60 * 24));
      if (days <= 6)  return "this week";
      if (days <= 13) return "next week";
      return "in " + Math.round(days / 7) + " weeks";
    } catch(e) { return "this week"; }
  })();

  // Format role list naturally: "A & B" or "A, B, & C"
  const rolesFormatted = roleList.length === 2
    ? roleList[0] + " & " + roleList[1]
    : roleList.slice(0, -1).join(", ") + ", & " + roleList[roleList.length - 1];

  // ── One-time context line: theme and/or WOTD mentioned once after the intro ──
  const themeRoleTypes = ["tabletopics", "jokemaster", "twominute"];
  const hasThemeRole = entries.some(e => themeRoleTypes.includes(e.roleType));
  const contextParts = [];
  if (theme && hasThemeRole) contextParts.push("the theme " + weekLabel + " is \"" + theme + "\"");
  if (wordOfTheDay)          contextParts.push("the Word of the Day is \"" + wordOfTheDay + "\"");
  const contextLine = contextParts.length > 0
    ? contextParts.join(" and ").replace(/^./, c => c.toUpperCase()) +
      ", in case you want to tie it to your roles, but it's completely optional."
    : "";

  // ── Per-role blurb: essential details only — theme/WOTD NOT repeated ──
  function getRoleBlurb(entry) {
    switch (entry.roleType) {

      case "tabletopics":
        return "You'll be leading the Table Topics segment. Aim for 20-30 minutes total, and arrive a few minutes early so we can sync up.";

      case "grammarian":
        return "You'll be our Grammarian." +
          (wordOfTheDay
            ? " Please introduce \"" + wordOfTheDay + "\" at the start of the meeting with its definition and an example sentence, track language use throughout, and give a brief report at the end."
            : " Please track language use throughout the meeting and give a brief report at the end.");


      case "timer":
        return "You\'ll be our Timer. Please use green, yellow, and red signals to help speakers stay within their allotted time, and give a brief report at the end.";

      case "jokemaster":
        return "You'll be opening the meeting with a bit of humor. Aim for about 1-2 minutes. Just keep it fun and engaging!";

      case "twominute":
        return "You'll be delivering a 2-Minute Special at the start of the meeting. Keep it prepared or impromptu, whatever works best for you!";

      case "speech": {
        const introLine = introQuestion ? "\n  - " + introQuestion + " (for your introduction)" : "";
        return "You'll be delivering " + entry.role + ". When you get a chance, could you reply with:\n" +
          "  - Speech Title\n" +
          "  - Speech Purpose (the objective from your pathway)\n" +
          "  - Pathway Name and Speech Number within the pathway\n" +
          "  - Allotted Time (e.g. 5-7 minutes)" + introLine + "\n" +
          "Please also arrive a few minutes early so we can go over the speaker order.";
      }

      case "evaluator": {
        const evalNum = entry.role.replace("Evaluator ", "").trim();
        const matchedSpeech = allConfirmations.find(e => e.role === "Speech " + evalNum);
        const speakerName  = matchedSpeech ? matchedSpeech.name  : "your assigned speaker";
        const speakerFirst = speakerName.split(" ")[0];
        const speakerEmail = matchedSpeech ? matchedSpeech.email : "";
        return "You'll be evaluating " + speakerName + ". I'd encourage you to reach out to " +
          speakerFirst + (speakerEmail ? " at " + speakerEmail : "") +
          " before the meeting to introduce yourself and ask if there's anything specific they'd like you to focus on. Your evaluation should run about 2-3 minutes, and I'll share their speech details as soon as I have them.";
      }

      default:
        return "Please reply to confirm you're all set!";
    }
  }

  // Build each role block as "*Role*: blurb"
  const roleBlocks = entries.map(entry => "*" + entry.role + "*: " + getRoleBlurb(entry)).join("\n\n");

  const greenCellNote = "When you get a moment, highlight your name green in our scheduling sheet if you haven't already. This helps me keep track of who's attending.";
  const footerNotes = [greenCellNote, attendanceStr].filter(Boolean).join("\n\n");

  return "Hi " + firstName + ",\n\n" +
    "I'm reaching out to confirm your roles for our " + meetingDate + " meeting! You're signed up for " + rolesFormatted + " " + weekLabel + ", and I wanted to make sure you have everything you need." +
    (contextLine ? " " + contextLine : "") + "\n\n" +
    roleBlocks +
    (locationStr ? "\n\n" + locationStr : "") +
    "\n\n" + footerNotes + "\n\n" +
    "Looking forward to seeing you " + weekLabel + "!\n\n" +
    "Warm regards,\n" + toastmasterName + "\nSierra Speakers Toastmasters";
}

// =============================
// Thank-You Email Builder (for already-confirmed roles)
// =============================
/**
 * buildThankYouEmailBody_
 * Builds a short thank-you email for a member whose role(s) are already confirmed (green).
 * Includes location info if hybrid/virtual, and role-specific context (theme, WOTD)
 * for roles that need it. Always includes an attendance sheet link.
 * @param {Object[]} entries - Confirmation entry objects for this group.
 * @param {string} confirmedTheme - Meeting theme.
 * @param {string} wordOfTheDay - Word of the Day.
 * @param {string} toastmasterName - Full name of the Toastmaster.
 * @param {string} meetingDate - Meeting date string (e.g. "3/6/2026").
 * @param {string} locationLine - Pre-built location string (or empty string).
 * @param {string} attendanceNote - Pre-built attendance note (includes address for hybrid).
 * @return {string} Plain-text thank-you email body.
 */
function buildThankYouEmailBody_(entries, confirmedTheme, wordOfTheDay, toastmasterName, meetingDate, locationLine, attendanceNote) {
  const firstName = entries[0].name.split(" ")[0];
  const roleList  = entries.map(e => e.role);
  const rolesFormatted = roleList.length === 1
    ? roleList[0]
    : roleList.length === 2
    ? roleList[0] + " & " + roleList[1]
    : roleList.slice(0, -1).join(", ") + ", & " + roleList[roleList.length - 1];

  const weekLabel = (() => {
    try {
      const today = new Date(); today.setHours(0,0,0,0);
      const md = new Date(meetingDate); md.setHours(0,0,0,0);
      const days = Math.round((md - today) / (1000 * 60 * 60 * 24));
      if (days <= 6)  return "this week";
      if (days <= 13) return "next week";
      return "in " + Math.round(days / 7) + " weeks";
    } catch(e) { return "this week"; }
  })();

  // Role-specific context reminders
  const themeRoles = ["tabletopics", "jokemaster", "twominute"];
  const needsTheme = confirmedTheme && entries.some(e => themeRoles.includes(e.roleType));
  const needsWotd  = wordOfTheDay   && entries.some(e => e.roleType === "grammarian");

  const contextLines = [];
  if (needsTheme) contextLines.push("Quick reminder, our meeting theme " + weekLabel + " is \"" + confirmedTheme + "\".");
  if (needsWotd)  contextLines.push("And just a heads-up, the Word of the Day is \"" + wordOfTheDay + "\".");
  const contextBlock = contextLines.length > 0 ? "\n\n" + contextLines.join(" ") : "";

  const locationBlock  = locationLine  ? "\n\n" + locationLine  : "";
  const attendanceBlock = attendanceNote ? "\n\n" + attendanceNote : "";

  return "Hi " + firstName + ",\n\n" +
    "Thanks so much for confirming your " + rolesFormatted + " for our " + meetingDate + " meeting!" +
    locationBlock +
    contextBlock +
    attendanceBlock + "\n\n" +
    "See you " + weekLabel + "!\n\n" +
    "Warm regards,\n" + toastmasterName + "\nSierra Speakers Toastmasters";
}

// =============================
// Email Body Builder
// =============================
/**
 * buildEmailBody
 * Builds the plain-text email body for a single role confirmation.
 * Switches on entry.roleType to produce role-specific copy.
 * @param {Object} entry - Confirmation entry object with roleType, role, name, email, etc.
 * @param {string} theme - Meeting theme string.
 * @param {string} wordOfTheDay - Word of the Day (used for Grammarian role).
 * @param {Object} speechDetails - Reserved for future use (currently unused).
 * @param {Object[]} allConfirmations - Full confirmations list (used to match evaluator to speaker).
 * @param {string} toastmasterName - Full name of the Toastmaster, used as sign-off.
 * @param {string} meetingDate_ - Meeting date string (e.g. "3/6/2026").
 * @param {string} introQuestion - Optional intro question to include in speech confirmation.
 * @param {string} [locationLine] - Pre-built location string (e.g. "We meet at 633 Folsom...").
 * @param {string} [attendanceNote] - Attendance sheet note for hybrid/virtual.
 * @return {string} Plain-text email body.
 */
function buildEmailBody(entry, theme, wordOfTheDay, speechDetails, allConfirmations, toastmasterName, meetingDate_, introQuestion, locationLine, attendanceNote) {
  const meetingDate = meetingDate_;
  const firstName = entry.name.split(" ")[0];
  const sheetUrl = SCHEDULING_SHEET_URL_ || SpreadsheetApp.getActiveSpreadsheet().getUrl();
  const locationStr = locationLine || "";
  const greenCellNote = "When you get a moment, highlight your name green in our scheduling sheet if you haven't already. This helps me keep track of who's attending.";
  const footerNotes = [greenCellNote, (attendanceNote || "")].filter(Boolean).join("\n\n");

  // Compute relative week label based on days until meeting Thursday
  const weekLabel = (() => {
    try {
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const meetingDate = new Date(meetingDate_);
      meetingDate.setHours(0, 0, 0, 0);
      const days = Math.round((meetingDate - today) / (1000 * 60 * 60 * 24));
      if (days <= 6)  return "this week";
      if (days <= 13) return "next week";
      const weeks = Math.round(days / 7);
      return "in " + weeks + " weeks";
    } catch(e) { return "this week"; }
  })();

  switch (entry.roleType) {

    case "tabletopics": {
      return `Hi ${firstName},

I'm reaching out to confirm roles for our Toastmasters meeting ${weekLabel}! You are currently signed up to be the Table Topics Master on ${meetingDate}.

${theme ? `As Table Topics Master, you'll be leading the impromptu speaking segment. Our theme ${weekLabel} is "${theme}", so feel free to draw from it when crafting your questions, or keep things open-ended, whatever makes for an engaging session!` : `As Table Topics Master, you'll be leading the impromptu speaking segment. Feel free to keep things open-ended, whatever makes for an engaging session!`}

A few reminders:
- Aim for 20-30 minutes total for the Table Topics segment.
- Arrive a few minutes early so we can sync up before the meeting starts.
${locationStr ? `- ${locationStr}` : ""}

${footerNotes}

Looking forward to a great meeting!

Warm regards,
${toastmasterName}
Sierra Speakers Toastmasters`;
    }

    case "grammarian": {
      return `Hi ${firstName},

I'm reaching out to confirm roles for our Toastmasters meeting ${weekLabel}! You are currently signed up to be the Grammarian on ${meetingDate}.

As Grammarian, you'll introduce the Word of the Day at the start of the meeting and track language use throughout. This week's Word of the Day is: "${wordOfTheDay}"

In the meeting, you will share its definition and use it in a sentence when you introduce it. At the end of the meeting, give a brief report on notable language highlights and any correct uses of the word.

${locationStr ? locationStr + "\n\n" : ""}${footerNotes}

Let me know if you have any questions!

Warm regards,
${toastmasterName}
Sierra Speakers Toastmasters`;
    }

    case "jokemaster": {
      return `Hi ${firstName},

I'm reaching out to confirm roles for our Toastmasters meeting ${weekLabel}! You are currently signed up to be the Joke Master on ${meetingDate}.

${theme ? `You'll be opening the meeting with a bit of humor. Aim for about 1-2 minutes. Our theme ${weekLabel} is "${theme}", so feel free to tie your joke in if inspiration strikes, but it's totally optional. Just keep it engaging and fun!` : `You'll be opening the meeting with a bit of humor. Aim for about 1-2 minutes. Just keep it engaging and fun!`}

${locationStr ? locationStr + "\n\n" : ""}${footerNotes}

See you there!

Warm regards,
${toastmasterName}
Sierra Speakers Toastmasters`;
    }

    case "twominute": {
      return `Hi ${firstName},

I'm reaching out to confirm roles for our Toastmasters meeting ${weekLabel}! You are currently signed up to deliver a 2-Minute Special in the beginning of our meeting on ${meetingDate}.

${theme ? `You'll have about 2 minutes to deliver a short prepared or impromptu piece. Our theme ${weekLabel} is "${theme}", so feel free to draw from it if you'd like, though it's not required!` : `You'll have about 2 minutes to deliver a short prepared or impromptu piece. Keep it prepared or impromptu, whatever works best for you!`}

${locationStr ? locationStr + "\n\n" : ""}${footerNotes}

Let me know if you have any questions. Looking forward to it!

Warm regards,
${toastmasterName}
Sierra Speakers Toastmasters`;
    }

    case "speech": {
      const introLine = introQuestion
        ? `\n- Answer the question to be used for your speech introduction: ${introQuestion}\n`
        : "";
      return `Hi ${firstName},

I'm reaching out to confirm roles for our Toastmasters meeting ${weekLabel}! You are currently signed up to deliver a prepared speech, ${entry.role}, at our meeting on ${meetingDate}.

I'd love to share a few details with your evaluator ahead of time so they can prepare a more personalized evaluation for you. Could you reply with the following?

- Speech Title
- Speech Purpose (the objective from your pathway)
- Pathway Name and Speech Number within the pathway
- Allotted Time (e.g. 5-7 minutes)${introLine}
Arrive a few minutes early if you can so we can go over the speaker order together.

${locationStr ? locationStr + "\n\n" : ""}${footerNotes}

Looking forward to hearing you speak!

Warm regards,
${toastmasterName}
Sierra Speakers Toastmasters`;
    }

    case "evaluator": {
      const evalNum = entry.role.replace("Evaluator ", "").trim();
      const matchedSpeech = allConfirmations.find((e) => e.role === `Speech ${evalNum}`);
      const speakerName = matchedSpeech ? matchedSpeech.name : "your assigned speaker";
      const speakerFirst = speakerName.split(" ")[0];
      const speakerEmail = matchedSpeech ? matchedSpeech.email : "";

      return `Hi ${firstName},

I'm reaching out to confirm roles for our Toastmasters meeting ${weekLabel}! You are currently signed up as ${entry.role} on ${meetingDate}.

You'll be evaluating ${speakerName}. I'd encourage you to reach out to ${speakerFirst}${speakerEmail ? ` at ${speakerEmail}` : ""} ahead of the meeting to introduce yourself and ask if there's anything specific they'd like you to focus on. That could be delivery, structure, vocal variety, or something else. It makes a big difference for the speaker!

Your evaluation should run about 2-3 minutes. I'll share ${speakerFirst}'s speech details with you as soon as I have them.

${locationStr ? locationStr + "\n\n" : ""}${footerNotes}

Let me know if you have any questions!

Warm regards,
${toastmasterName}
Sierra Speakers Toastmasters`;
    }

    default: {
      const parts = [
        "Hi " + firstName + ",",
        "",
        "I'm reaching out to confirm roles for our Toastmasters meeting " + weekLabel + "! You are currently signed up for the role of " + entry.role + " on " + meetingDate + ".",
        "",
        "Please reply to this email and highlight your name green in our scheduling sheet to confirm your attendance.",
      ];
      if (locationStr) { parts.push(""); parts.push(locationStr); }
      if (attendanceNote) { parts.push(""); parts.push(attendanceNote); }
      parts.push("", "Don't hesitate to reach out if you have any questions!", "", "Warm regards,", toastmasterName, "Sierra Speakers Toastmasters");
      return parts.join("\n");
    }
  }
}


// ============================================================
// AGENDA GENERATOR — Apps Script function
// ============================================================
/**
 * fmtTime_
 * Converts a total-minutes value (from midnight) into a 12-hour time string.
 * Shared by generateAgenda and buildAgendaDocx.
 * @param {number} totalMins - Minutes since midnight (e.g. 6*60+20 = 380 → "6:20 am").
 * @return {string} Formatted time string like "6:20 am" or "7:30 pm".
 */
function fmtTime_(totalMins) {
  const h = Math.floor(totalMins/60), m = totalMins%60;
  const h12 = h % 12 === 0 ? 12 : h % 12;
  const ampm = h >= 12 ? "pm" : "am";
  return `${h12}:${m.toString().padStart(2,"0")} ${ampm}`;
}

/**
 * generateAgenda
 * Entry point for the agenda generation flow. Prompts the user to pick a sheet and
 * meeting date, collects Word of the Day (with optional Merriam-Webster auto-lookup),
 * lets the user confirm which speeches to include, then calls buildAgendaDocx() to
 * produce a .docx file saved to Google Drive.
 * @return {void}
 */
function generateAgenda() {
  const ui = SpreadsheetApp.getUi();

  // Always wipe temp keys at the start so a previous cancelled/crashed run
  // never causes dialogs to be skipped on the next run
  const tempKeys = ["_speechSelection", "_evaluatorSelection", "_wotdSelection", "_wotdDefinition", "_agendaMode"];
  tempKeys.forEach(k => PropertiesService.getScriptProperties().deleteProperty(k));

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  const schedSheets = sheets
    .map(s => s.getName())
    .filter(name => /^SCHED\s\d{4}$/.test(name))
    .sort((a, b) => parseInt(b.split(" ")[1]) - parseInt(a.split(" ")[1]));

  let selectedSheetName = schedSheets[0];
  let sheet = spreadsheet.getSheetByName(selectedSheetName);

  const confirmSheet = ui.alert(`Generate agenda using "${selectedSheetName}"?`, ui.ButtonSet.YES_NO);
  if (confirmSheet === ui.Button.NO) {
    const input = ui.prompt("Enter the sheet name to use:",
      `Available sheets:\n${sheets.map(s => s.getName()).join("\n")}`, ui.ButtonSet.OK_CANCEL);
    if (input.getSelectedButton() !== ui.Button.OK) return;
    selectedSheetName = input.getResponseText().trim();
    sheet = spreadsheet.getSheetByName(selectedSheetName);
    if (!sheet) { ui.alert(`Sheet "${selectedSheetName}" not found.`); return; }
  }

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();

  // ── Find Roles header row ──
  let rolesHeaderRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0]?.toString().trim().toLowerCase() === "roles") { rolesHeaderRow = r; break; }
  }
  if (rolesHeaderRow === -1) {
    for (let r = 0; r < data.length; r++) {
      if (data[r][0]?.toString().trim().toLowerCase() === "toastmaster") { rolesHeaderRow = r - 1; break; }
    }
  }
  if (rolesHeaderRow < 0) { ui.alert("Could not find the Roles header row."); return; }

  // ── Find first date column ──
  let firstDateCol = -1;
  for (let c = 0; c < data[rolesHeaderRow].length; c++) {
    if (data[rolesHeaderRow][c] instanceof Date) { firstDateCol = c; break; }
  }
  if (firstDateCol === -1) { ui.alert("Could not find dates in the Roles row."); return; }

  // ── Build header map and pick date ──
  const headerMap = data[rolesHeaderRow].slice(firstDateCol).map((h, index) => {
    let formatted = "";
    if (h instanceof Date) formatted = Utilities.formatDate(h, Session.getScriptTimeZone(), "M/d/yyyy");
    else formatted = h ? h.toString().trim() : "";
    return { original: h, formatted, colIndex: firstDateCol + index };
  });

  const now = new Date();
  now.setHours(0, 0, 0, 0); // compare against midnight so today is included
  const formattedDates = headerMap.filter(h => h.original instanceof Date && h.original >= now).map(h => h.formatted);
  const upcomingHeader = headerMap.find(h => h.original instanceof Date && h.original >= now);
  const exampleDate = upcomingHeader ? Utilities.formatDate(upcomingHeader.original, Session.getScriptTimeZone(), "M/d/yyyy") : "";

  const dateResponse = ui.prompt("Which meeting date?",
    `Available options:\n${formattedDates.join("\n")}\n\nClick OK to use "${exampleDate}", or enter a different date:`,
    ui.ButtonSet.OK_CANCEL);
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  let selectedDateRaw = dateResponse.getResponseText().trim() || exampleDate;

  const inputParts = normalizeDateParts_(selectedDateRaw);
  const match = headerMap.find(h => {
    const hp = normalizeDateParts_(h.formatted);
    if (inputParts.y) return inputParts.m === hp.m && inputParts.d === hp.d && inputParts.y === hp.y;
    return inputParts.m === hp.m && inputParts.d === hp.d;
  });
  if (!match) { ui.alert(`No meeting found for "${selectedDateRaw}".`); return; }

  const colIndex = match.colIndex;
  const meetingDateObj = match.original instanceof Date ? match.original : new Date(match.formatted);
  const formattedLongDate = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "MMMM d, yyyy");

  // ── Meeting theme ──
  let meetingTheme = "";
  if (rolesHeaderRow > 0) meetingTheme = data[rolesHeaderRow - 1][colIndex]?.toString().trim() || "";

  // ── WOD Memory cache lookup ──
      var dateStr = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), 'M/d/yyyy');
  const wodCache_ = lookupWodCache_(dateStr);
  const wodCachedWord_ = wodCache_ ? wodCache_.word : "";
  // ── Word of the Day — single-screen HTML dialog ──
  const mwKeyCheck = PropertiesService.getScriptProperties().getProperty("MW_API_KEY") || "";
  const wotdHtml = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      h3 { margin: 0 0 12px; color: #1B2A4A; font-size: 15px; }
      label { display: block; margin-bottom: 4px; font-weight: bold; color: #333; }
      input[type=text], textarea { width: 100%; box-sizing: border-box; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; font-family: Arial, sans-serif; }
      textarea { height: 60px; resize: vertical; }
      .radios { margin: 10px 0 8px; display: flex; gap: 20px; }
      .radios label { font-weight: normal; display: flex; align-items: center; gap: 6px; cursor: pointer; }
      .manual-fields { display: none; margin-top: 10px; }
      .field { margin-bottom: 10px; }
      .footer { margin-top: 14px; display: flex; justify-content: flex-end; gap: 8px; }
      button { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }
      .btn-cancel { background: #eee; color: #333; }
      .btn-ok { background: #1B2A4A; color: white; }
      .hint { font-size: 11px; color: #888; margin-top: 2px; }
    </style></head><body>
    <h3>Word of the Day</h3>
    <div class="field">
      <label for="word">Word</label>
      <input type="text" id="word" placeholder="Type in Word of the Day here" value="${wodCachedWord_}" autofocus>
    </div>
    <div class="radios">
      <label><input type="radio" name="mode" value="auto" ${mwKeyCheck ? "checked" : "disabled"}> Look up automatically</label>
      <label><input type="radio" name="mode" value="manual" ${!mwKeyCheck ? "checked" : ""}> Enter manually</label>
    </div>
    ${!mwKeyCheck ? '<p class="hint" style="color:#c00;">MW_API_KEY not set — manual entry only.</p>' : ""}
    <div class="manual-fields" id="manualFields">
      <div class="field">
        <label for="def">Definition</label>
        <textarea id="def" placeholder="e.g. a projecting part of a fortification"></textarea>
      </div>
      <div class="field">
        <label for="ex">Example sentence <span style="font-weight:normal;color:#888;">(optional)</span></label>
        <textarea id="ex" placeholder="e.g. The castle's bastion overlooked the valley"></textarea>
      </div>
    </div>
    <div class="footer">
      <button class="btn-cancel" onclick="google.script.run.withSuccessHandler(() => google.script.host.close()).setWotdSelection(null)">Skip</button>
      <button class="btn-ok" onclick="submit()">Continue</button>
    </div>
    <script>
      document.querySelectorAll('input[name=mode]').forEach(r => {
        r.addEventListener('change', toggle);
      });
      function toggle() {
        const manual = document.querySelector('input[name=mode]:checked').value === 'manual';
        document.getElementById('manualFields').style.display = manual ? 'block' : 'none';
      }
      toggle();
      function submit() {
        const word = document.getElementById('word').value.trim();
        const mode = document.querySelector('input[name=mode]:checked').value;
        const def  = document.getElementById('def').value.trim();
        const ex   = document.getElementById('ex').value.trim();
        if (!word) { google.script.run.withSuccessHandler(() => google.script.host.close()).setWotdSelection(null); return; }
        google.script.run.withSuccessHandler(() => google.script.host.close()).setWotdSelection(JSON.stringify({ word, mode, def, ex }));
      }
    </script></body></html>
  `).setWidth(420).setHeight(mwKeyCheck ? 260 : 310);

  PropertiesService.getScriptProperties().deleteProperty("_wotdSelection");
  SpreadsheetApp.getUi().showModalDialog(wotdHtml, "Word of the Day");

  let waited = 0;
  while (!PropertiesService.getScriptProperties().getProperty("_wotdSelection") && waited < 60) {
    Utilities.sleep(500);
    waited += 0.5;
  }
  const wotdRaw = PropertiesService.getScriptProperties().getProperty("_wotdSelection");
  PropertiesService.getScriptProperties().deleteProperty("_wotdSelection");

  let wordOfTheDay = "";
  let wotdPronunciation = "", wotdPartOfSpeech = "", wotdDefinition = "", wotdExample = "";

  if (wotdRaw && wotdRaw !== "null") {
    const wotdData = JSON.parse(wotdRaw);
    wordOfTheDay = wotdData.word || "";

    if (wotdData.mode === "manual") {
      // User entered definition/example directly in the dialog
      wotdDefinition = wotdData.def || "";
      wotdExample    = wotdData.ex  || "";
    } else if (wotdData.mode === "auto" && wordOfTheDay) {

      // ── WOD Memory cache check ──
      const currentAiModel_ = getAiModel_();
      const needsRefresh_ = !wodCache_
        || wodCache_.word.toLowerCase() !== wordOfTheDay.toLowerCase()
        || shouldRefreshGemini_(wodCache_, meetingTheme, currentAiModel_.label);

      if (!needsRefresh_ && wodCache_ && wodCache_.definition) {
        // Cache hit — same word, same theme, same-or-better model
        wotdPronunciation = wodCache_.pronunciation;
        wotdPartOfSpeech  = wodCache_.partOfSpeech;
        wotdDefinition    = wodCache_.definition;
        wotdExample       = wodCache_.example;
        console.log("WOD cache hit for " + wordOfTheDay + " on " + dateStr);
      } else {
        // Cache miss or refresh needed — proceed with API lookups
        console.log("WOD cache miss/refresh for " + wordOfTheDay + " on " + dateStr);
      // ── Merriam-Webster auto-lookup ──
      const mwKey = PropertiesService.getScriptProperties().getProperty("MW_API_KEY") || "";
      let allDefs = [], mwPronunciation = "", apiWorked = false;

      try {
        const url = "https://www.dictionaryapi.com/api/v3/references/collegiate/json/" +
                    encodeURIComponent(wordOfTheDay) + "?key=" + mwKey;
        const resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
        const parsed = JSON.parse(resp.getContentText());
        if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === "object" && parsed[0].meta) {
          apiWorked = true;
          const entry = parsed[0];
          try { mwPronunciation = entry.hwi.prs[0].mw || ""; } catch(pe) {}
          const pos = entry.fl || "";
          (entry.shortdef || []).forEach(d => { if (d && allDefs.length < 5) allDefs.push({ pos, def: d, ex: "", pronunciation: mwPronunciation, source: "mw" }); });
          try {
            let exIdx = 0;
            for (let di = 0; di < entry.def.length && exIdx < allDefs.length; di++) {
              for (let si = 0; si < entry.def[di].sseq.length && exIdx < allDefs.length; si++) {
                for (let ei = 0; ei < entry.def[di].sseq[si].length && exIdx < allDefs.length; ei++) {
                  const sense = entry.def[di].sseq[si][ei];
                  if (sense[0] === "sense" && sense[1] && sense[1].dt) {
                    for (let dti = 0; dti < sense[1].dt.length; dti++) {
                      if (sense[1].dt[dti][0] === "vis" && sense[1].dt[dti][1] && sense[1].dt[dti][1].length > 0) {
                        let ex = (sense[1].dt[dti][1][0].t || "").replace(/\{[^}]+\}/g, "").replace(/\s+/g, " ").trim();
                        if (ex) { allDefs[exIdx].ex = ex; exIdx++; }
                      }
                    }
                  }
                }
              }
            }
          } catch(exErr) {}
        }
      } catch(e) { Logger.log("MW API failed: " + e.toString()); }

      // ── Gemini lookup — fetch before showing the picker so it appears as a choice ──
      const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
      if (geminiKey) {
        const geminiPrompt =
          "You are helping prepare a Toastmasters meeting agenda. " +
          "The Word of the Day is \"" + wordOfTheDay + "\". " +
          (meetingTheme ? "The meeting theme is \"" + meetingTheme + "\". " : "") +
          "Please respond with ONLY a valid JSON object (no markdown, no explanation) in this exact format: " +
          "{\"definition\": \"a concise, clear definition\", " +
          "\"partOfSpeech\": \"the part of speech, e.g. noun, verb, adjective\", " +
          "\"pronunciation\": \"phonetic pronunciation using simple syllable notation like \\\"kon-SISE\\\"\", " +
          "\"example\": \"a vivid example sentence that naturally relates to the meeting theme\"}";

        const WOTD_MAX_ATTEMPTS = 3;
        for (let attempt = 1; attempt <= WOTD_MAX_ATTEMPTS; attempt++) {
          try {
            const aiModel = getAiModel_();
            console.log("WOTD attempt " + attempt + " using model: " + aiModel.model);
            const geminiResp = UrlFetchApp.fetch(
              "https://generativelanguage.googleapis.com/v1beta/models/" + aiModel.model + ":generateContent?key=" + geminiKey,
              {
                method: "post",
                contentType: "application/json",
                muteHttpExceptions: true,
                payload: JSON.stringify({
                  contents: [{ parts: [{ text: geminiPrompt }] }],
                  generationConfig: { temperature: 0.4, maxOutputTokens: 600 }
                })
              }
            );
            if (geminiResp.getResponseCode() === 200) {
              const geminiJson = JSON.parse(geminiResp.getContentText());
              const rawText = geminiJson?.candidates?.[0]?.content?.parts?.[0]?.text || "";
              const cleaned = rawText.replace(/```json|```/gi, "").trim();
              try {
                const gParsed = JSON.parse(cleaned);
                if (gParsed.definition) {
                  recordAiPing_(aiModel.label);
                  allDefs.push({
                    pos: gParsed.partOfSpeech || "",
                    def: gParsed.definition,
                    ex: gParsed.example || "",
                    pronunciation: gParsed.pronunciation || "",
                    source: aiModel.label
                  });
                  apiWorked = true;
                  break; // success — exit retry loop
                } else {
                  console.log("WOTD attempt " + attempt + " (" + aiModel.model + "): JSON ok but no definition field. Raw: " + cleaned.substring(0, 300));
                }
              } catch (parseErr) {
                console.log("WOTD attempt " + attempt + " (" + aiModel.model + "): JSON parse failed — " + parseErr.toString() + " — Raw: " + cleaned.substring(0, 300));
              }
            } else {
              console.log("WOTD attempt " + attempt + " (" + aiModel.model + "): HTTP " + geminiResp.getResponseCode() + " — " + geminiResp.getContentText().substring(0, 300));
            }
            if (attempt < WOTD_MAX_ATTEMPTS) Utilities.sleep(2000);
          } catch (geminiErr) {
            console.log("WOTD attempt " + attempt + " failed: " + geminiErr.toString());
            if (attempt < WOTD_MAX_ATTEMPTS) Utilities.sleep(2000);
          }
        }
      }

      if (apiWorked && allDefs.length > 0) {
        // Show definition picker — all MW defs + Gemini option in one unified dialog
        const defsJson = JSON.stringify(allDefs);
        const pickerHtml = HtmlService.createHtmlOutput(`
          <!DOCTYPE html><html><head>
          <style>
            body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
            h3 { margin: 0 0 14px; color: #1B2A4A; font-size: 15px; }
            .source-header { display: flex; align-items: center; gap: 8px; margin: 14px 0 7px; }
            .source-header:first-child { margin-top: 0; }
            .source-header .source-label {
              font-size: 10px; font-weight: bold; letter-spacing: 0.8px; text-transform: uppercase;
              padding: 2px 8px; border-radius: 10px; white-space: nowrap;
            }
            .source-header.mw .source-label { background: #e8f5e9; color: #1E5631; border: 1px solid #a8d5b5; }
            .source-header.gemini .source-label { background: #e8f0fe; color: #1a56cc; border: 1px solid #a8c0f5; }
            .source-header .source-line { flex: 1; height: 1px; background: #e0e0e0; }
            .def { padding: 8px 10px; margin-bottom: 5px; border: 1px solid #a8d5b5; border-radius: 4px; cursor: pointer; background: #f2faf3; }
            .def:hover { background: #e0f5e4; border-color: #1E5631; }
            .def.selected { background: #e0f5e4; border-color: #1E5631; }
            .def.gemini { border-color: #a8c0f5; background: #f0f4fe; }
            .def.gemini:hover { background: #e0eafd; border-color: #4285F4; }
            .def.gemini.selected { background: #e0eafd; border-color: #4285F4; }
            .source-group { border-radius: 6px; padding: 8px; margin-bottom: 4px; }
            .source-group.mw     { background: #f2faf3; border: 1px solid #c6e6c6; }
            .source-group.gemini { background: #f0f4fe; border: 1px solid #c0d0f8; }
            .source-group .def     { background: #ffffff; border-color: #a8d5b5; }
            .source-group .def:hover,
            .source-group .def.selected              { background: #e0f5e4; border-color: #1E5631; }
            .source-group.gemini .def                { background: #ffffff; border-color: #a8c0f5; }
            .source-group.gemini .def:hover,
            .source-group.gemini .def.selected       { background: #e0eafd; border-color: #4285F4; }
            .pos { color: #888; font-style: italic; font-size: 11px; }
            .ex { color: #555; font-size: 11px; margin-top: 3px; }
            .pron { color: #777; font-size: 11px; margin-top: 2px; }
            .section-divider { border: none; border-top: 1px solid #ddd; margin: 14px 0 10px; }
            .source-header.manual .source-label { background: #f5f5f5; color: #555; border: 1px solid #d0d0d0; }
            .source-group.manual-group { background: #f9f9f9; border: 1px solid #e0e0e0; border-radius: 6px; padding: 8px; margin-bottom: 4px; }
            .manual-section label { font-weight: bold; display: block; margin-bottom: 4px; color: #333; }
            .manual-section textarea { width: 100%; box-sizing: border-box; padding: 6px 8px; border: 1px solid #ccc; border-radius: 4px; font-size: 13px; font-family: Arial, sans-serif; height: 55px; resize: vertical; margin-bottom: 8px; }
            .footer { margin-top: 14px; display: flex; justify-content: flex-end; gap: 8px; }
            button { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }
            .btn-cancel { background: #eee; color: #333; }
            .btn-ok { background: #1B2A4A; color: white; }
          </style></head><body>
          <h3>Choose a definition for "${wordOfTheDay}"</h3>
          <div id="list"></div>
          <hr class="section-divider">
          <div class="source-header manual">
            <div class="source-line"></div>
            <span class="source-label manual-label">✎ Your Own</span>
            <div class="source-line"></div>
          </div>
          <div class="source-group manual-group">
            <div class="manual-section">
              <textarea id="manDef" placeholder="Type definition here"></textarea>
              <label for="manEx">Example sentence <span style="font-weight:normal;color:#888;">(optional)</span></label>
              <textarea id="manEx" placeholder="Type example sentence here"></textarea>
            </div>
          </div>
          <div class="footer">
            <button class="btn-cancel" onclick="google.script.run.withSuccessHandler(()=>google.script.host.close()).setWotdDefinition(null)">Cancel</button>
            <button class="btn-ok" onclick="submit()">Use This</button>
          </div>
          <script>
            const defs = ${defsJson};
            let selected = 0;
            const list = document.getElementById('list');
            const manDef = document.getElementById('manDef');
            const manEx  = document.getElementById('manEx');
            let lastSource = null;
            let mwCounter = 0;
            let currentGroup = null;
            defs.forEach((d, i) => {
              const isGemini = d.source === 'gemini-lite' || d.source === 'gemini-flash' || d.source === 'gemma';
              const source   = isGemini ? 'gemini' : 'mw';
              const sourceLabel = d.source === 'gemma' ? '✦ Gemma' : (d.source === 'gemini-lite' || d.source === 'gemini-flash') ? '✦ Gemini' : 'Merriam-Webster';

              // Inject a source section header and open a new group whenever the source changes
              if (source !== lastSource) {
                const hdr = document.createElement('div');
                hdr.className = 'source-header ' + source;
                hdr.innerHTML = '<div class="source-line"></div>' +
                  '<span class="source-label">' + sourceLabel + '</span>' +
                  '<div class="source-line"></div>';
                list.appendChild(hdr);

                currentGroup = document.createElement('div');
                currentGroup.className = 'source-group ' + source;
                list.appendChild(currentGroup);

                lastSource = source;
                if (!isGemini) mwCounter = 0;
              }

              const div = document.createElement('div');
              div.className = 'def' + (i === 0 ? ' selected' : '');
              const posHtml = '<span class="pos">(' + (d.pos || 'word') + ')</span>';
              const pronHtml = d.pronunciation ? '<div class="pron">' + d.pronunciation + '</div>' : '';
              const num = isGemini ? '' : '<strong>' + (++mwCounter) + '.</strong> ';
              div.innerHTML = num + posHtml + ' ' + d.def +
                pronHtml +
                (d.ex ? '<div class="ex">e.g. "' + d.ex + '"</div>' : '');
              div.onclick = () => {
                document.querySelectorAll('.def').forEach(x => x.classList.remove('selected'));
                div.classList.add('selected');
                selected = i;
                manDef.value = '';
                manEx.value  = '';
              };
              currentGroup.appendChild(div);
            });
            manDef.addEventListener('input', () => {
              if (manDef.value.trim()) {
                document.querySelectorAll('.def').forEach(x=>x.classList.remove('selected'));
                selected = -1;
              }
            });
            function submit() {
              const customDef = manDef.value.trim();
              const customEx  = manEx.value.trim();
              if (customDef) {
                google.script.run.withSuccessHandler(()=>google.script.host.close())
                  .setWotdDefinition(JSON.stringify({ idx: -1, pronunciation: '', manualDef: customDef, manualEx: customEx }));
              } else if (selected >= 0 && selected < defs.length) {
                google.script.run.withSuccessHandler(()=>google.script.host.close())
                  .setWotdDefinition(JSON.stringify({ idx: selected, pronunciation: defs[selected].pronunciation || '' }));
              } else {
                alert('Please select a definition or type your own before continuing.');
              }
            }
          </script></body></html>
        `).setWidth(460).setHeight(500);

        PropertiesService.getScriptProperties().deleteProperty("_wotdDefinition");
        SpreadsheetApp.getUi().showModalDialog(pickerHtml, "Choose Definition");
        let w2 = 0;
        while (!PropertiesService.getScriptProperties().getProperty("_wotdDefinition") && w2 < 60) {
          Utilities.sleep(500); w2 += 0.5;
        }
        const defRaw = PropertiesService.getScriptProperties().getProperty("_wotdDefinition");
        PropertiesService.getScriptProperties().deleteProperty("_wotdDefinition");

        if (defRaw && defRaw !== "null") {
          const defData = JSON.parse(defRaw);
          wotdPronunciation = defData.pronunciation || "";
          if (defData.idx === -1) {
            wotdPartOfSpeech = "";
            wotdDefinition   = defData.manualDef || "";
            wotdExample      = defData.manualEx  || "";
          } else {
            const chosen = allDefs[defData.idx];
            wotdPartOfSpeech = chosen.pos;
            wotdDefinition   = chosen.def;
            wotdExample      = chosen.ex;
          }
        }
      } else {
        ui.alert('Couldn\'t find "' + wordOfTheDay + '" in the dictionary. The word will appear on the agenda without a definition.');
      }
      } // end of cache-miss else block
    }
  }

  // ── WOD Memory write-back ──
  if (wordOfTheDay && wotdDefinition) {
    const wodSource_ = wotdPartOfSpeech ? "mw" : (typeof currentAiModel_ !== "undefined" ? currentAiModel_.label : "");
    saveWodToCache_(
      dateStr, wordOfTheDay, wotdDefinition, wotdPronunciation,
      wotdPartOfSpeech, wotdExample, wodSource_,
      meetingTheme || "", typeof currentAiModel_ !== "undefined" ? currentAiModel_.label : ""
    );
  }
  // ── Read roles ──
  const startRow = rolesHeaderRow + 1;
  const roles = {};
  let speechCounter = 1, evaluatorCounter = 1;
  const speechFlags = {}; // key -> { name, isRed, isTbd }

  for (let r = startRow; r < data.length; r++) {
    const roleRaw = data[r][0]?.toString().trim();
    const assignedRaw = data[r][colIndex]?.toString().trim();
    if (!roleRaw) continue;
    const roleLower = roleRaw.toLowerCase();
    let key = roleRaw;
    if (roleLower.startsWith("speech"))    key = `Speech ${speechCounter++}`;
    else if (roleLower.startsWith("evaluator")) key = `Evaluator ${evaluatorCounter++}`;
    const isTbd = !assignedRaw || assignedRaw.toUpperCase() === "TBD";
    const name = isTbd ? "" : assignedRaw;
    roles[key] = name;
    if (key.startsWith("Speech ")) {
      const bgColor = (backgrounds[r] && backgrounds[r][colIndex]) ? backgrounds[r][colIndex].toLowerCase() : "";
      const isRed = bgColor.startsWith("#ff") || bgColor === "#ea4335" || bgColor === "#e06666" || bgColor === "#cc0000" || bgColor === "#ff0000";
      speechFlags[key] = { name, isRed, isTbd, rowIndex: r };
    }
  }

  const speechKeys    = Object.keys(roles).filter(k => k.startsWith("Speech "));
  const evaluatorKeys = Object.keys(roles).filter(k => k.startsWith("Evaluator "));
  let speechDetails = {}; // populated later after inbox scan; declared here so speech dialog can reference it
  // ── Speech confirmation dialog ──
  const speechItemsJson = JSON.stringify(speechKeys.map(k => ({
    key: k,
    name: speechFlags[k] ? speechFlags[k].name : roles[k] || "",
    flagged: speechFlags[k] ? (speechFlags[k].isRed || speechFlags[k].isTbd) : false
  })));

  const speechHtml = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
      h3 { margin: 0 0 12px; color: #1B2A4A; font-size: 15px; }
      .speech { padding: 8px 10px; margin-bottom: 6px; border-radius: 4px; border: 1px solid #ddd; }
      .speech.flagged { background: #fff3f3; border-color: #f5c6c6; }
      .speech.ok { background: #f3f9f3; border-color: #c6e6c6; }
      .top-row { display: flex; align-items: center; }
      label { flex: 1; cursor: pointer; display: flex; align-items: center; gap: 10px; }
      input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; flex-shrink: 0; }
      .name { font-weight: bold; color: #1B2A4A; }
      .tbd { color: #cc0000; font-style: italic; }
      .badge { font-size: 10px; padding: 2px 6px; border-radius: 10px; margin-left: 6px; }
      .badge.red { background: #fdd; color: #c00; }
      .badge.ok { background: #dfd; color: #080; }
      .footer { margin-top: 14px; display: flex; justify-content: flex-end; gap: 8px; }
      button { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }
      .btn-cancel { background: #eee; color: #333; }
      .btn-ok { background: #1B2A4A; color: white; }
    </style></head><body>
    <h3>Confirm Prepared Speeches</h3>
    <div id="list"></div>
    <div class="footer">
      <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
      <button class="btn-ok" onclick="submit()">Continue</button>
    </div>
    <script>
      const speeches = ${speechItemsJson};
      const list = document.getElementById('list');
      speeches.forEach((s, i) => {
        const checked = !s.flagged;
        const div = document.createElement('div');
        div.className = 'speech ' + (s.flagged ? 'flagged' : 'ok');
        const nameHtml = s.name
          ? '<span class="name">' + s.name + '</span>'
          : '<span class="tbd">No speaker assigned</span>';
        const badge = s.flagged
          ? '<span class="badge red">⚠ flagged</span>'
          : '<span class="badge ok">✓ confirmed</span>';

        div.innerHTML =
          '<div class="top-row"><label>' +
            '<input type="checkbox" id="cb'+i+'" ' + (checked ? 'checked' : '') + '>' +
            s.key + ' — ' + nameHtml + badge +
          '</label></div>';
        list.appendChild(div);
      });
      function submit() {
        const kept = speeches.filter((s,i) => document.getElementById('cb'+i).checked).map(s => s.key);
        google.script.run.withSuccessHandler(() => google.script.host.close())
          .setSpeechSelection(kept, []);
      }
    </script></body></html>
  `).setWidth(420).setHeight(300);

  // Store speech keys temporarily, show dialog, wait for response via script property
  PropertiesService.getScriptProperties().deleteProperty("_speechSelection");
  SpreadsheetApp.getUi().showModalDialog(speechHtml, "Prepared Speeches");

  // Poll for user response (set by setSpeechSelection callback)
  let waited2 = 0;
  while (!PropertiesService.getScriptProperties().getProperty("_speechSelection") && waited2 < 60) {
    Utilities.sleep(500);
    waited2 += 0.5;
  }
  const selectionRaw = PropertiesService.getScriptProperties().getProperty("_speechSelection");
  if (!selectionRaw) return; // cancelled or timed out
  PropertiesService.getScriptProperties().deleteProperty("_speechSelection");

  const selectionData  = JSON.parse(selectionRaw);
  const keptSpeechKeys = selectionData.kept;
  let impromptuKeys  = new Set(); // populated after inbox scan for speakers without details

  // ── Evaluator confirmation dialog (only when fewer than all speeches are selected) ──
  let keptEvaluatorKeys;
  if (keptSpeechKeys.length < speechKeys.length && evaluatorKeys.length > 0) {
    // Pre-check evaluators whose number matches a kept speech number
    // e.g. if Speech 2 and Speech 3 are kept, Evaluator 2 and Evaluator 3 are pre-checked
    const keptSpeechNums = new Set(keptSpeechKeys.map(k => k.replace("Speech ", "").trim()));
    const evalItemsJson = JSON.stringify(evaluatorKeys.map((k) => ({
      key: k,
      name: roles[k] || "",
      preChecked: keptSpeechNums.has(k.replace("Evaluator ", "").trim())
    })));

    const evalHtml = HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><head>
      <style>
        body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
        h3 { margin: 0 0 4px; color: #1B2A4A; font-size: 15px; }
        .subtitle { color: #666; font-size: 12px; margin-bottom: 12px; }
        .evaluator { display: flex; align-items: center; padding: 8px 10px; margin-bottom: 6px; border-radius: 4px; border: 1px solid #ddd; }
        .evaluator.checked { background: #f3f9f3; border-color: #c6e6c6; }
        .evaluator.unchecked { background: #f7f7f7; border-color: #ddd; }
        label { flex: 1; cursor: pointer; display: flex; align-items: center; gap: 10px; }
        input[type=checkbox] { width: 16px; height: 16px; cursor: pointer; }
        .name { font-weight: bold; color: #1B2A4A; }
        .footer { margin-top: 14px; display: flex; justify-content: flex-end; gap: 8px; }
        button { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }
        .btn-cancel { background: #eee; color: #333; }
        .btn-ok { background: #1B2A4A; color: white; }
      </style></head><body>
      <h3>Confirm Evaluators</h3>
      <p class="subtitle">Selected evaluators will be mapped in order to the selected speeches.</p>
      <div id="list"></div>
      <div class="footer">
        <button class="btn-cancel" onclick="google.script.host.close()">Cancel</button>
        <button class="btn-ok" onclick="submit()">Continue</button>
      </div>
      <script>
        const evaluators = ${evalItemsJson};
        const list = document.getElementById('list');
        evaluators.forEach((e, i) => {
          const div = document.createElement('div');
          div.className = 'evaluator ' + (e.preChecked ? 'checked' : 'unchecked');
          div.id = 'row' + i;
          const nameHtml = e.name
            ? '<span class="name">' + e.name + '</span>'
            : '<span style="color:#999;font-style:italic">No evaluator assigned</span>';
          div.innerHTML = '<label><input type="checkbox" id="ecb'+i+'" ' + (e.preChecked ? 'checked' : '') + ' onchange="updateRow('+i+')">' +
            e.key + ' \u2014 ' + nameHtml + '</label>';
          list.appendChild(div);
        });
        function updateRow(i) {
          const cb = document.getElementById('ecb' + i);
          document.getElementById('row' + i).className = 'evaluator ' + (cb.checked ? 'checked' : 'unchecked');
        }
        function submit() {
          const kept = evaluators.filter((e, i) => document.getElementById('ecb' + i).checked).map(e => e.key);
          google.script.run.withSuccessHandler(() => google.script.host.close()).setEvaluatorSelection(kept);
        }
      </script></body></html>
    `).setWidth(420).setHeight(260);

    PropertiesService.getScriptProperties().deleteProperty("_evaluatorSelection");
    SpreadsheetApp.getUi().showModalDialog(evalHtml, "Evaluators");

    // Poll for user response (set by setEvaluatorSelection callback)
    let waited3 = 0;
    while (!PropertiesService.getScriptProperties().getProperty("_evaluatorSelection") && waited3 < 60) {
      Utilities.sleep(500);
      waited3 += 0.5;
    }
    const evalSelectionRaw = PropertiesService.getScriptProperties().getProperty("_evaluatorSelection");
    if (!evalSelectionRaw) return; // cancelled or timed out
    PropertiesService.getScriptProperties().deleteProperty("_evaluatorSelection");

    keptEvaluatorKeys = JSON.parse(evalSelectionRaw);
  } else {
    // All speeches selected — use default 1:1 mapping, no dialog needed
    keptEvaluatorKeys = keptSpeechKeys.map((_, i) => evaluatorKeys[i] || "").filter(Boolean);
  }

  const activeSpeechCount = keptSpeechKeys.length; // may be reduced further if cancelled speakers are removed after inbox scan

  // ── Build nameToEmail map from the member directory at the top of the sheet ──
  // Used to find each speaker's email address for the inbox scan.
  const nameToEmail = {};
  {
    let r = 1;
    while (r < data.length) {
      const firstName = data[r][0]?.toString().trim();
      const lastName  = data[r][1]?.toString().trim();
      const email     = data[r][4]?.toString().trim();
      const bgColor   = backgrounds[r]?.[0];
      if (bgColor === "#cfe2f3" || (!firstName && !email)) break;
      if (firstName && lastName && email) {
        nameToEmail[firstName + " " + lastName] = email;
      }
      r++;
    }
  }

  // ── Retrieve stored intro question (saved by proceedToEmails when emails were sent) ──
  let storedIntroQ = PropertiesService.getScriptProperties().getProperty("LAST_INTRO_QUESTION") || "";

  // ── Agenda mode dialog: skeleton or inbox scan ──
  const agendaModeHtml = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body { font-family: Arial, sans-serif; padding: 20px 20px 16px; font-size: 13px; }
      h3   { margin: 0 0 8px; color: #1B2A4A; font-size: 15px; }
      p    { color: #555; font-size: 12px; margin: 0 0 18px; line-height: 1.55; }
      .btn { display: block; width: 100%; padding: 11px; margin-bottom: 9px; border: none;
             border-radius: 5px; font-size: 13px; cursor: pointer; font-weight: bold; }
      .btn-scan     { background: #1B2A4A; color: white; }
      .btn-skeleton { background: #f0f0f0; color: #333; border: 1px solid #ccc; }
      .btn:disabled { opacity: 0.5; }
    </style></head><body>
    <h3>How would you like to build the agenda?</h3>
    <p>Generate a skeleton now, or scan your inbox for speaker replies from the past 7 days to pre-fill speech details and speaker introductions.</p>
    <button class="btn btn-scan" id="btnScan" onclick="choose('scan')">📬  Scan Inbox for Speaker Replies</button>
    <button class="btn btn-skeleton" id="btnSkel" onclick="choose('skeleton')">📄  Generate Skeleton Agenda</button>
    <script>
      function choose(mode) {
        document.getElementById('btnScan').disabled = true;
        document.getElementById('btnSkel').disabled = true;
        google.script.run
          .withSuccessHandler(() => google.script.host.close())
          .setAgendaMode(mode);
      }
    </script></body></html>
  `).setWidth(390).setHeight(220);

  PropertiesService.getScriptProperties().deleteProperty("_agendaMode");
  SpreadsheetApp.getUi().showModalDialog(agendaModeHtml, "Agenda Mode");

  let waitedMode = 0;
  while (!PropertiesService.getScriptProperties().getProperty("_agendaMode") && waitedMode < 60) {
    Utilities.sleep(500);
    waitedMode += 0.5;
  }
  const agendaMode = PropertiesService.getScriptProperties().getProperty("_agendaMode") || "skeleton";
  PropertiesService.getScriptProperties().deleteProperty("_agendaMode");

  // ── Inbox scan (only if user chose it) ──
  speechDetails = {}; // { [speechKey]: {title, purpose, pathway, speechNum, time, intro} }

  if (agendaMode === "scan") {
    // Map kept speech keys to their speaker's email
    const speakerEntries = keptSpeechKeys
      .map(key => ({
        speechKey: key,
        name:  roles[key] || "",
        email: roles[key] ? (nameToEmail[roles[key]] || "") : ""
      }))
      .filter(e => e.name && e.email);

    if (speakerEntries.length === 0) {
      ui.alert("Inbox Scan", "No email addresses found for the scheduled speakers in your roster. Generating skeleton agenda.", ui.ButtonSet.OK);
    } else {
      // Scan inbox — this may take a few seconds
      const scanResult = scanSpeakerEmails_(speakerEntries, storedIntroQ);
      speechDetails = scanResult.details;
      const emailFoundKeys = scanResult.emailFoundKeys;
      // Backfill intro question from email thread if it wasn't stored from a prior send
      if (!storedIntroQ && scanResult.extractedIntroQuestion) {
        storedIntroQ = scanResult.extractedIntroQuestion;
        PropertiesService.getScriptProperties().setProperty("LAST_INTRO_QUESTION", storedIntroQ);
      }

      // ── Categorise each speaker into one of six states ──
      const fullReplies    = []; // title + at least one of purpose/pathway extracted
      const partialReplies = []; // title found but purpose AND pathway both empty
      const noFields       = []; // email found but Gemini extracted nothing at all
      const unavailable    = []; // speaker said they cannot deliver the speech
      const uncertain      = []; // speaker is unsure / maybe
      const noReply        = []; // no email found in the past 7 days

      speakerEntries.forEach(e => {
        const name = e.name.split(" ")[0]; // first name for natural phrasing
        const det  = speechDetails[e.speechKey];
        const avail = det && det.availability ? det.availability : "";
        if (!emailFoundKeys.has(e.speechKey)) {
          noReply.push(e.name);
        } else if (avail === "unavailable") {
          unavailable.push(name);
        } else if (avail === "uncertain") {
          uncertain.push(name);
        } else if (!det || !det.title) {
          noFields.push(e.name);
        } else if (!det.purpose && !det.pathway) {
          partialReplies.push(e.name);
        } else {
          fullReplies.push(e.name);
        }
      });

      // ── Build per-speaker scan data for the combined dialog ──
      // Also check if all details failed (fall back to skeleton)
      if (fullReplies.length === 0 && partialReplies.length === 0) {
        speechDetails = {};
      }

      const withIntros = speakerEntries.filter(e => speechDetails[e.speechKey] && speechDetails[e.speechKey].intro);
      const hasIntroDoc = storedIntroQ && withIntros.length > 0;

      // ── Handle cancelled speakers (unavailable) before showing dialog ──
      if (unavailable.length > 0) {
        const cancelledEntries = speakerEntries.filter(e => {
          const det = speechDetails[e.speechKey];
          return det && det.availability === "unavailable";
        });

        const cancelNames = cancelledEntries.map(e => e.name).join(", ");
        const confirmMsg = cancelNames + (cancelledEntries.length === 1
          ? " can no longer deliver a speech. Remove them from the agenda and mark their cell red in the sheet?"
          : " can no longer deliver their speeches. Remove them from the agenda and mark their cells red in the sheet?");

        const confirmResponse = ui.alert("Remove Cancelled Speakers", confirmMsg, ui.ButtonSet.YES_NO);

        if (confirmResponse === ui.Button.YES) {
          cancelledEntries.forEach(e => {
            const flag = speechFlags[e.speechKey];
            if (flag && flag.rowIndex !== undefined) {
              sheet.getRange(flag.rowIndex + 1, colIndex + 1).setBackground("#ea4335");
            }
          });

          const cancelledKeys = new Set(cancelledEntries.map(e => e.speechKey));
          const filteredSpeechKeys = keptSpeechKeys.filter(k => !cancelledKeys.has(k));
          const filteredEvaluatorKeys = filteredSpeechKeys.map((_, i) => keptEvaluatorKeys[i] || "").filter(Boolean);

          keptSpeechKeys.length = 0;
          filteredSpeechKeys.forEach(k => keptSpeechKeys.push(k));
          keptEvaluatorKeys.length = 0;
          filteredEvaluatorKeys.forEach(k => keptEvaluatorKeys.push(k));
        }
      }

      // ── Combined scan results + impromptu dialog ──
      const scanItemsJson = JSON.stringify(keptSpeechKeys.map(k => {
        const det = speechDetails && speechDetails[k];
        const avail = det && det.availability ? det.availability : "";
        let status;
        if (!det || !det.title) {
          if (avail === "uncertain") status = "uncertain";
          else status = "missing";
        } else if (!det.purpose && !det.pathway) {
          status = "partial";
        } else {
          status = "full";
        }
        return {
          key: k,
          name: roles[k] || "No speaker assigned",
          status: status,
          title: (det && det.title) ? det.title : ""
        };
      }));

      const footerNote = hasIntroDoc
        ? "<p style=\"margin:10px 0 0;color:#2d6a3f;font-size:11px;\">📝 Intro answers found — a Speaker Introductions document will also be created.</p>"
        : "";

      const scanHtml = HtmlService.createHtmlOutput(`
        <!DOCTYPE html><html><head>
        <style>
          body { font-family: Arial, sans-serif; padding: 16px; font-size: 13px; }
          h3 { margin: 0 0 4px; color: #1B2A4A; font-size: 15px; }
          .subtitle { margin: 0 0 12px; color: #555; font-size: 12px; }
          .row { padding: 8px 10px; margin-bottom: 6px; border-radius: 4px; border: 1px solid #ddd; }
          .row.full    { background: #f3f9f3; border-color: #c6e6c6; }
          .row.partial { background: #fffbe6; border-color: #ffe082; }
          .row.missing { background: #fafafa; }
          .row.uncertain { background: #fff8f0; border-color: #f5c6c6; }
          .top { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }
          .name { font-weight: bold; color: #1B2A4A; }
          .badge { font-size: 10px; padding: 2px 6px; border-radius: 10px; white-space: nowrap; }
          .badge.full    { background: #dfd; color: #080; }
          .badge.partial { background: #fff3cd; color: #856404; }
          .badge.uncertain { background: #ffe5cc; color: #804000; }
          .sub { font-size: 11px; color: #555; margin-top: 3px; }
          .imp-row { margin-top: 5px; }
          .imp-row label { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #555; cursor: pointer; }
          .imp-row input { width: 13px; height: 13px; cursor: pointer; }
          .footer { margin-top: 14px; display: flex; justify-content: flex-end; }
          button { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; background: #1B2A4A; color: white; }
        </style></head><body>
        <h3>Inbox Scan Results</h3>
        <p class="subtitle">Review what was found. For speakers without details, mark as impromptu if applicable.</p>
        <div id="list"></div>
        ${footerNote}
        <div class="footer"><button onclick="submitChecked()">Continue</button></div>
        <script>
          const items = ${scanItemsJson};
          const list = document.getElementById('list');
          items.forEach((s, i) => {
            const div = document.createElement('div');
            div.className = 'row ' + (s.status === 'missing' ? 'missing' : s.status);
            let html = '<div class="top"><span class="name">'+s.key+' — '+s.name+'</span>';
            if (s.status === 'full')     html += '<span class="badge full">✓ details found</span>';
            if (s.status === 'partial')  html += '<span class="badge partial">🟡 partial details</span>';
            if (s.status === 'uncertain') html += '<span class="badge uncertain">🤔 uncertain</span>';
            html += '</div>';
            if (s.title) html += '<div class="sub">📝 '+s.title+'</div>';
            if (s.status === 'missing' || s.status === 'uncertain') {
              html += '<div class="imp-row"><label><input type="checkbox" id="imp'+i+'">🎲 Impromptu speech (no title / prep needed)</label></div>';
            }
            div.innerHTML = html;
            list.appendChild(div);
          });
          function submitChecked() {
            const checked = items.filter((s,i) => {
              const el = document.getElementById('imp'+i);
              return el && el.checked;
            }).map(s => s.key);
            google.script.run.withSuccessHandler(() => google.script.host.close())
              .setImpromptuSelection(checked);
          }
        </script></body></html>
      `).setWidth(430).setHeight(Math.min(140 + keptSpeechKeys.length * 75, 480));

      PropertiesService.getScriptProperties().deleteProperty("_impromptuSelection");
      SpreadsheetApp.getUi().showModalDialog(scanHtml, "Inbox Scan Results");

      let waitedImp = 0;
      while (!PropertiesService.getScriptProperties().getProperty("_impromptuSelection") && waitedImp < 60) {
        Utilities.sleep(500);
        waitedImp += 0.5;
      }
      const impRaw = PropertiesService.getScriptProperties().getProperty("_impromptuSelection");
      PropertiesService.getScriptProperties().deleteProperty("_impromptuSelection");
      if (impRaw) {
        impromptuKeys = new Set(JSON.parse(impRaw));
      }
    }
  }


    // ── Calculate times ──
  // evalMins is anchored to the FULL complement of speeches so the evaluation
  // session always starts at the same time regardless of how many speeches are
  // selected. Table Topics fills the gap: fewer speeches → longer TT segment.
  const totalAvailableSpeeches = evaluatorKeys.length; // all possible speech slots
  const speechStartMins = 6*60+20;
  const activeSpeechCountFinal = keptSpeechKeys.length;
  const evalMins = speechStartMins + totalAvailableSpeeches * 10 + 16; // fixed anchor
  const ttMins   = speechStartMins + activeSpeechCountFinal * 10;      // floats left as speeches drop

  // ── Build the docx as base64 via the embedded builder ──
  const docxBase64 = buildAgendaDocx({
    date: formattedLongDate,
    theme: meetingTheme,
    roles,
    speechKeys: keptSpeechKeys,
    evaluatorKeys: keptEvaluatorKeys,
    wordOfTheDay,
    wotdPronunciation,
    wotdPartOfSpeech,
    wotdDefinition,
    wotdExample,
    speechStartMins,
    ttMins,
    evalMins,
    fmtTime: fmtTime_,
    speechDetails,
    impromptuKeys: [...impromptuKeys],  // array of speech keys marked as impromptu
  });

  // ── Save agenda to Google Drive ──
  const filename = `Toastmasters Agenda ${selectedDateRaw}`;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(docxBase64),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    filename + ".docx"
  );
  const file    = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  const fileUrl = file.getUrl();

  // Persist for hype email auto-population
  PropertiesService.getScriptProperties().setProperties({
    LAST_AGENDA_URL:    fileUrl,
    LAST_AGENDA_DATE:   selectedDateRaw,
    LAST_AGENDA_WOTD:   wordOfTheDay   || "",
    LAST_AGENDA_THEME:  meetingTheme   || "",
    LAST_AGENDA_WOTD_DEF: wotdDefinition || "",
    LAST_AGENDA_WOTD_EX:  wotdExample    || "",
  });

  // ── Build Introductions docx (only if at least one speaker provided an intro answer) ──
  let introFileUrl = null;
  const introData = keptSpeechKeys
    .map(key => ({
      name:  roles[key] || "",
      title: (speechDetails[key] && speechDetails[key].title) ? speechDetails[key].title : "",
      time:  (speechDetails[key] && speechDetails[key].time)  ? speechDetails[key].time  : "",
      intro: (speechDetails[key] && speechDetails[key].intro) ? speechDetails[key].intro : ""
    }))
    .filter(spk => spk.name && spk.intro);

  if (introData.length > 0) {
    const introBase64   = buildIntroductionsDocx_(introData, formattedLongDate, storedIntroQ);
    const introFilename = "Speaker Introductions " + selectedDateRaw;
    const introBlob = Utilities.newBlob(
      Utilities.base64Decode(introBase64),
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      introFilename + ".docx"
    );
    introFileUrl = DriveApp.createFile(introBlob).getUrl();
  }

  // ── Success dialog — shows one or two Drive links ──
  const introLinkHtml = introFileUrl
    ? `<p style="margin:10px 0 0;">
         📝 <a href="${introFileUrl}" target="_blank"
            style="color:#1E5631;font-weight:bold;font-size:14px;">Speaker Introductions ${selectedDateRaw}</a>
       </p>`
    : "";

  const successHtml = HtmlService.createHtmlOutput(
    `<div style="font-family:Arial,sans-serif;padding:14px;">
       <p style="margin:0 0 10px;">✅ <strong>${filename}.docx</strong> saved to your Google Drive.</p>
       <p style="margin:0 0 6px;">
         <a href="${fileUrl}" target="_blank"
            style="color:#1B2A4A;font-weight:bold;font-size:14px;">${filename}</a>
       </p>
       ${introLinkHtml}
     </div>`
  ).setWidth(430).setHeight(introFileUrl ? 140 : 110);

  SpreadsheetApp.getUi().showModalDialog(successHtml, "Agenda Created");
}

// ============================================================
// DOCX BUILDER — pure Apps Script, no external dependencies
// Builds a minimal but well-formatted .docx (OOXML) as base64
// ============================================================
/**
 * buildAgendaDocx
 * Constructs a fully-formatted Toastmasters meeting agenda as a .docx file
 * (OOXML/ZIP) encoded in base64, using only built-in Apps Script utilities —
 * no external dependencies or Drive API calls.
 *
 * The document contains a two-image header row, a WOTD section, and a
 * time-stamped agenda table that reflects the roles read from the schedule
 * sheet and the speeches confirmed by the user.
 *
 * @param {Object} d - Agenda data object.
 * @param {string}   d.date              - Long-form meeting date, e.g. "April 17, 2025".
 * @param {string}   d.theme             - Meeting theme string (may be empty).
 * @param {Object}   d.roles             - Map of role key → assigned member name.
 * @param {string[]} d.speechKeys        - Ordered array of kept speech keys.
 * @param {string[]} d.evaluatorKeys     - Ordered array of evaluator keys matching speechKeys.
 * @param {string}   d.wordOfTheDay      - Word of the Day string (may be empty).
 * @param {string}   d.wotdPronunciation - IPA / MW pronunciation string.
 * @param {string}   d.wotdPartOfSpeech  - Part of speech, e.g. "noun".
 * @param {string}   d.wotdDefinition    - Dictionary definition text.
 * @param {string}   d.wotdExample       - Example sentence (may be empty).
 * @param {number}   d.speechStartMins   - Speeches start time in minutes since midnight.
 * @param {number}   d.ttMins            - Table Topics start time in minutes since midnight.
 * @param {number}   d.evalMins          - Evaluations start time in minutes since midnight.
 * @param {Function} d.fmtTime           - Time formatter: (totalMins: number) => string.
 * @return {string} Base64-encoded .docx file content.
 */
function buildAgendaDocx(d) {
  const DARK_GREEN  = "1E5631";
  const LIGHT_GREEN = "E8F5E9";
  const NAVY        = "1B2A4A";
  const GRAY        = "555555";
  const WHITE       = "FFFFFF";
  const ROLE_COLOR  = "1E5631";

  // Image base64 strings (embedded at build time)
  const IMG_LEFT_B64  = "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";
  const IMG_RIGHT_B64 = "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";

  const CONTENT_W = 10512;
  const TIME_W    = 1050;
  const ROLE_W    = 4500;
  const NAME_W    = CONTENT_W - TIME_W - ROLE_W; // 4962

  // ── XML helpers ──
  function e(text) {
    return (text||"").toString()
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&apos;");
  }

  function rPr(bold, color, size, italic) {
    return `<w:rPr>${bold?`<w:b/><w:bCs/>`:""}${italic?`<w:i/><w:iCs/>`:""}
      <w:color w:val="${color||"000000"}"/>
      <w:sz w:val="${size||19}"/><w:szCs w:val="${size||19}"/>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
    </w:rPr>`;
  }

  function textRun(text, bold, color, size, italic) {
    return `<w:r>${rPr(bold,color,size,italic)}<w:t xml:space="preserve">${e(text)}</w:t></w:r>`;
  }

  function pPr(align, spaceBefore, spaceAfter, line) {
    return `<w:pPr>
      ${align?`<w:jc w:val="${align}"/>`:""}
      <w:spacing w:before="${spaceBefore||0}" w:after="${spaceAfter||0}" w:line="${line||240}" w:lineRule="auto"/>
    </w:pPr>`;
  }

  function cellXml(width, bg, vAlign, mt, mb, children) {
    const shade = bg ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : "";
    return `<w:tc>
      <w:tcPr>
        <w:tcW w:w="${width}" w:type="dxa"/>
        <w:tcBorders>
          <w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
        </w:tcBorders>
        ${shade}
        <w:tcMar>
          <w:top w:w="${mt||60}" w:type="dxa"/><w:bottom w:w="${mb||60}" w:type="dxa"/>
          <w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/>
        </w:tcMar>
        <w:vAlign w:val="${vAlign||"center"}"/>
      </w:tcPr>
      ${children}
    </w:tc>`;
  }

  function agendaRow(time, role, name, bg) {
    const shade = bg||"";
    return `<w:tr>
      ${cellXml(TIME_W, shade, "center", 60, 60,
        `<w:p>${pPr("left")}${textRun(time, false, ROLE_COLOR, 19)}</w:p>`)}
      ${cellXml(ROLE_W, shade, "center", 60, 60,
        `<w:p>${pPr("left")}${textRun(role, true, ROLE_COLOR, 19)}</w:p>`)}
      ${cellXml(NAME_W, shade, "center", 60, 60,
        `<w:p>${pPr("right")}${textRun(name||"", false, "333333", 19)}</w:p>`)}
    </w:tr>`;
  }

  function subRow(label, value, bg, italic) {
    const shade = bg||"";
    const content = value
      ? textRun(label, true, GRAY, 17) + textRun("  " + value, false, GRAY, 17, italic!==false)
      : textRun(label, false, GRAY, 17, true);
    return `<w:tr>
      ${cellXml(TIME_W, shade, "center", 0, 2, `<w:p>${pPr()}</w:p>`)}
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${ROLE_W + NAME_W}" w:type="dxa"/>
          <w:gridSpan w:val="2"/>
          <w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>
          ${shade ? `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>` : ""}
          <w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          <w:vAlign w:val="top"/>
        </w:tcPr>
        <w:p>${pPr()}${content}</w:p>
      </w:tc>
    </w:tr>`;
  }

  // Wide sub-row: spans ROLE_W+NAME_W so text never wraps early
  // If value is null, renders label alone (used for the header line)
  function wotdSubRow(label, value, italic) {
    const content = value !== null && value !== undefined
      ? textRun(label, true, GRAY, 17) + textRun("  " + value, false, GRAY, 17, italic===true)
      : textRun(label, true, ROLE_COLOR, 19); // header line — larger, green, bold
    return `<w:tr>
      ${cellXml(TIME_W, "", "center", 0, 2, `<w:p>${pPr()}</w:p>`)}
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${ROLE_W + NAME_W}" w:type="dxa"/>
          <w:gridSpan w:val="2"/>
          <w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>
          <w:tcMar><w:top w:w="${value !== null && value !== undefined ? 0 : 60}" w:type="dxa"/><w:bottom w:w="50" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          <w:vAlign w:val="top"/>
        </w:tcPr>
        <w:p>${pPr()}${content}</w:p>
      </w:tc>
    </w:tr>`;
  }

  function sectionHeaderRow(text) {
    return `<w:tr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${CONTENT_W}" w:type="dxa"/>
          <w:gridSpan w:val="3"/>
          <w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>
          <w:shd w:val="clear" w:color="auto" w:fill="${DARK_GREEN}"/>
          <w:tcMar>
            <w:top w:w="90" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/>
            <w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/>
          </w:tcMar>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p>${pPr("center")}${textRun(text, true, WHITE, 20)}</w:p>
      </w:tc>
    </w:tr>`;
  }

  function spacerRow(line) {
    return `<w:tr>
      <w:tc>
        <w:tcPr>
          <w:tcW w:w="${CONTENT_W}" w:type="dxa"/>
          <w:gridSpan w:val="3"/>
          <w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>
          <w:tcMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>
        </w:tcPr>
        <w:p>${pPr("", 0, 0, line||180)}</w:p>
      </w:tc>
    </w:tr>`;
  }

  // ── Image drawing XML ──
  function imgDrawing(rId, cx, cy) {
    return `<w:drawing>
      <wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing">
        <wp:extent cx="${cx}" cy="${cy}"/>
        <wp:effectExtent b="0" l="0" r="0" t="0"/>
        <wp:docPr id="1" name="img${rId}"/>
        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:nvPicPr>
                <pic:cNvPr id="0" name="img${rId}"/>
                <pic:cNvPicPr preferRelativeResize="0"/>
              </pic:nvPicPr>
              <pic:blipFill>
                <a:blip r:embed="${rId}"/>
                <a:stretch><a:fillRect/></a:stretch>
              </pic:blipFill>
              <pic:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
                <a:prstGeom prst="rect"/>
              </pic:spPr>
            </pic:pic>
          </a:graphicData>
        </a:graphic>
      </wp:inline>
    </w:drawing>`;
  }

  // ── Header table ──
  const HIMG_W = 1260, HTXT_W = CONTENT_W - HIMG_W*2;
  const navyShade = `<w:shd w:val="clear" w:color="auto" w:fill="${NAVY}"/>`;
  const imgCellBorders = `<w:tcBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/></w:tcBorders>`;

  const headerTableXml = `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${CONTENT_W}" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
        <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="${HIMG_W}"/><w:gridCol w:w="${HTXT_W}"/><w:gridCol w:w="${HIMG_W}"/>
    </w:tblGrid>
    <w:tr>
      <w:tc>
        <w:tcPr><w:tcW w:w="${HIMG_W}" w:type="dxa"/>${imgCellBorders}${navyShade}
          <w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tcMar>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p>${pPr("left")}<w:r>${imgDrawing("rIdImgL", 800000, 699999)}</w:r></w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="${HTXT_W}" w:type="dxa"/>${imgCellBorders}${navyShade}
          <w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p>${pPr("center",0,40)}${textRun("Sierra Speakers Toastmasters", true, WHITE, 30)}</w:p>
        <w:p>${pPr("center",0,30)}${textRun(d.date, false, "C8D8F0", 19)}</w:p>
        <w:p>${pPr("center",0,0)}${textRun(d.theme||"", false, "A8D8A8", 21, true)}</w:p>
      </w:tc>
      <w:tc>
        <w:tcPr><w:tcW w:w="${HIMG_W}" w:type="dxa"/>${imgCellBorders}${navyShade}
          <w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tcMar>
          <w:vAlign w:val="center"/>
        </w:tcPr>
        <w:p>${pPr("right")}<w:r>${imgDrawing("rIdImgR", 800000, 630720)}</w:r></w:p>
      </w:tc>
    </w:tr>
  </w:tbl>`;

  // ── Divider ──
  const dividerXml = `<w:p>
    <w:pPr>
      <w:spacing w:before="0" w:after="140"/>
      <w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="${DARK_GREEN}"/></w:pBdr>
    </w:pPr>
  </w:p>`;

  // ── Main agenda rows ──
  let agendaRows = "";

  agendaRows += agendaRow("6:05 pm", "Meeting Open",        d.roles["Sergeant at Arms"]||"", LIGHT_GREEN);
  agendaRows += agendaRow("",        "Toastmaster Welcome", d.roles["Toastmaster"]||"");
  agendaRows += agendaRow("",        "JokeMaster",          d.roles["Joke Master"]||"",       LIGHT_GREEN);
  agendaRows += agendaRow("",        "2-Minute Special",    d.roles["2-Minute Special"]||"");

  // General Evaluator
  agendaRows += agendaRow("6:15 pm", "General Evaluator", d.roles["General Evaluator"]||"", LIGHT_GREEN);
  agendaRows += spacerRow();

  // Word of the Day — all rows span full width so nothing wraps
  if (d.wordOfTheDay) {
    const wordCapitalized = d.wordOfTheDay.charAt(0).toUpperCase() + d.wordOfTheDay.slice(1);
    const wordWithPos = "Word of the Day:  " + wordCapitalized +
      (d.wotdPartOfSpeech  ? "  |  (" + d.wotdPartOfSpeech + ")" : "") +
      (d.wotdPronunciation ? "  |  " + d.wotdPronunciation        : "");
    agendaRows += wotdSubRow(wordWithPos, null, false);
    agendaRows += wotdSubRow("Definition:", d.wotdDefinition || "[please fill in]", false);
    if (d.wotdExample) agendaRows += wotdSubRow("Example:", d.wotdExample, true);
  } else {
    agendaRows += agendaRow("", "Word of the Day", "");
  }
  agendaRows += spacerRow();

  // Speeches
  agendaRows += sectionHeaderRow("Prepared Speeches");
  d.speechKeys.forEach((key, i) => {
    if (!d.roles[key]) return;
    const bg   = i % 2 === 0 ? LIGHT_GREEN : "";
    const time = i === 0 ? "6:20 pm" : d.fmtTime(d.speechStartMins + i * 10);

    const isImpromptu = d.impromptuKeys && d.impromptuKeys.includes(key);

    if (isImpromptu) {
      // Impromptu speech — show a single row with a brief note, no Purpose/Pathway sub-rows
      agendaRows += agendaRow(time, "Impromptu Speech", d.roles[key], bg);
      agendaRows += subRow("Topics suggested by club members on the day  |  5-7 minutes", null, bg, true);
      agendaRows += spacerRow(160);
    } else {
      // Pull confirmed details from speechDetails if available; fall back to placeholders
      const det      = (d.speechDetails && d.speechDetails[key]) || {};
      const title    = det.title    || "[Speech Title TBD]";
      const purpose  = det.purpose  || "[To be confirmed by speaker]";
      const pathway  = det.pathway  || "[TBD]";
      const speechNo = det.speechNum ? " #" + det.speechNum : "";
      const timeStr  = det.time     || "5-7 minutes";

      agendaRows += agendaRow(time, title, d.roles[key], bg);
      agendaRows += subRow("Purpose:", purpose, bg, false);
      agendaRows += subRow("Pathway:", pathway + speechNo + "  |  " + timeStr, bg, true);
      agendaRows += spacerRow(160);
    }
  });

  // Table Topics
  agendaRows += sectionHeaderRow("Table Topics");
  agendaRows += agendaRow(d.fmtTime(d.ttMins), "Table Topics Master", d.roles["Table Topics Master"]||"", LIGHT_GREEN);
  agendaRows += subRow("Impromptu speaking round  |  1-2 minutes per speaker", null, LIGHT_GREEN);
  agendaRows += spacerRow();

  // Evaluation Session
  agendaRows += sectionHeaderRow("Evaluation Session  (led by General Evaluator)");
  agendaRows += agendaRow(d.fmtTime(d.evalMins), "Evaluation Session", d.roles["General Evaluator"]||"", LIGHT_GREEN);
  d.evaluatorKeys.forEach((key, i) => {
    const bg = i % 2 !== 0 ? LIGHT_GREEN : "";
    agendaRows += agendaRow("", `    Speech Evaluator #${i+1}  (2-3 mins)`, d.roles[key]||"", bg);
  });
    const flipEval = d.evaluatorKeys.length % 2 !== 0; // odd evaluator count flips alternation
  agendaRows += agendaRow("", "    Grammarian Report  (1-2 mins)",         d.roles["Grammarian"]||"", flipEval ? LIGHT_GREEN : "");
  agendaRows += agendaRow("", "    Ah Counter Report  (1-2 mins)",         d.roles['"Ah" Counter']||d.roles["Ah Counter"]||"", flipEval ? "" : LIGHT_GREEN);
  agendaRows += agendaRow("", "    Timer Report  (1-2 mins)",              d.roles["Timer"]||"", flipEval ? LIGHT_GREEN : "");
  agendaRows += agendaRow("", "    General Evaluator Report  (2-3 mins)",  d.roles["General Evaluator"]||"", flipEval ? "" : LIGHT_GREEN);
  agendaRows += spacerRow();

  // Close
  agendaRows += agendaRow("7:30 pm", "Officer Announcements", "",                    LIGHT_GREEN);
  agendaRows += agendaRow("",        "Meeting Close",          d.roles["Toastmaster"]||"");

  // ── Main table ──
  const mainTableXml = `<w:tbl>
    <w:tblPr>
      <w:tblW w:w="${CONTENT_W}" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/>
        <w:insideH w:val="nil"/><w:insideV w:val="nil"/>
      </w:tblBorders>
    </w:tblPr>
    <w:tblGrid>
      <w:gridCol w:w="${TIME_W}"/><w:gridCol w:w="${ROLE_W}"/><w:gridCol w:w="${NAME_W}"/>
    </w:tblGrid>
    ${agendaRows}
  </w:tbl>`;

  // ── Assemble document.xml ──
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    ${headerTableXml}
    ${dividerXml}
    ${mainTableXml}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
      <w:pgMar w:top="576" w:right="864" w:bottom="576" w:left="864" w:header="708" w:footer="708"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  // ── Relationships ──
  const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rIdImgL" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imgL.jpg"/>
  <Relationship Id="rIdImgR" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/imgR.png"/>
</Relationships>`;

  const stylesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:sz w:val="19"/><w:szCs w:val="19"/>
    </w:rPr></w:rPrDefault>
  </w:docDefaults>
</w:styles>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Default Extension="jpg"  ContentType="image/jpeg"/>
  <Default Extension="png"  ContentType="image/png"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"   ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;

  const rootRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

  // ── Build ZIP (docx) using Apps Script's built-in utilities ──
  const files = [
    { name: "[Content_Types].xml", content: contentTypesXml, mime: "text/xml" },
    { name: "_rels/.rels",         content: rootRelsXml,     mime: "text/xml" },
    { name: "word/document.xml",   content: documentXml,     mime: "text/xml" },
    { name: "word/styles.xml",     content: stylesXml,       mime: "text/xml" },
    { name: "word/_rels/document.xml.rels", content: relsXml, mime: "text/xml" },
  ];

  const blobs = files.map(f =>
    Utilities.newBlob(f.content, f.mime, f.name)
  );

  // Add images
  blobs.push(Utilities.newBlob(Utilities.base64Decode(IMG_LEFT_B64),  "image/jpeg", "word/media/imgL.jpg"));
  blobs.push(Utilities.newBlob(Utilities.base64Decode(IMG_RIGHT_B64), "image/png",  "word/media/imgR.png"));


  const zipBlob = Utilities.zip(blobs, "agenda.docx");
  return Utilities.base64Encode(zipBlob.getBytes());
}

// ============================================================
// INBOX SCAN — Gmail search + Gemini extraction for speaker replies
// ============================================================
/**
 * setAgendaMode
 * Dialog callback — stores "skeleton" or "scan" in Script Properties
 * so the polling loop in generateAgenda can read it.
 * @param {string} mode - "skeleton" or "scan"
 * @return {void}
 */
function setAgendaMode(mode) {
  PropertiesService.getScriptProperties().setProperty("_agendaMode", mode || "skeleton");
}

/**
 * scanSpeakerEmails_
 * Searches Gmail for replies from known speakers in the past 7 days.
 * For each thread found, collects ALL messages from the speaker (in
 * chronological order) and sends the combined text to Gemini for
 * structured extraction of speech details — so details buried in an
 * earlier message within a back-and-forth thread are not missed.
 *
 * @param {Object[]} speakerEntries - Array of {speechKey, name, email} for each kept speech.
 * @param {string}   introQuestion  - The intro question posed in the confirmation email (may be empty).
 * @return {Object} Map of { [speechKey]: {title, purpose, pathway, speechNum, time, intro} }
 */
function scanSpeakerEmails_(speakerEntries, introQuestion) {
  const results = {};
  const emailFoundKeys = new Set(); // speechKeys where an email was found, even if extraction was empty
  const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
  if (!geminiKey) {
    Logger.log("scanSpeakerEmails_: GEMINI_API_KEY not set, skipping extraction.");
    return { details: results, emailFoundKeys };
  }

  // Build a map from lowercase email → speaker entry for fast lookup
  const emailToEntry = {};
  speakerEntries.forEach(e => {
    if (e.email) emailToEntry[e.email.toLowerCase().trim()] = e;
  });

  const emailList = speakerEntries.map(e => e.email).filter(Boolean);
  if (emailList.length === 0) return { details: results, emailFoundKeys };

  // One Gmail search covering all speaker addresses, last 7 days
  const fromClause = emailList.map(e => "from:" + e).join(" OR ");
  let threads;
  try {
    threads = GmailApp.search("(" + fromClause + ") newer_than:7d", 0, 50);
  } catch (e) {
    Logger.log("scanSpeakerEmails_: Gmail search failed: " + e.toString());
    return { details: results, emailFoundKeys };
  }

  threads.forEach(thread => {
    const messages = thread.getMessages();

    const speakerBodies = {}; // { speechKey: [body1, body2, ...] }

    messages.forEach(msg => {
      const fromRaw   = msg.getFrom();
      const fromEmail = (fromRaw.match(/<(.+)>/) ? fromRaw.match(/<(.+)>/)[1] : fromRaw).trim().toLowerCase();
      const entry = emailToEntry[fromEmail];
      if (!entry) return;
      const rawBody = msg.getPlainBody() || "";
      // Strip quoted reply lines ("> ...") and "On [date]...wrote:" headers
      const body = rawBody
        .split("\n")
        .filter(line => !line.startsWith(">") && !/^On .{10,}wrote:/.test(line.trim()))
        .join("\n")
        .trim();
      if (!body) return;
      if (!speakerBodies[entry.speechKey]) speakerBodies[entry.speechKey] = [];
      speakerBodies[entry.speechKey].push(body);
    });

    // Extract once per speaker. Sort messages longest-first so Gemini sees
    // the most detail-rich content before any short back-and-forth replies.
    Object.keys(speakerBodies).forEach(speechKey => {
      if (results[speechKey] && results[speechKey].title) return; // already have a good result
      emailFoundKeys.add(speechKey);  // email was found regardless of extraction outcome
      const sorted = speakerBodies[speechKey].slice().sort((a, b) => b.length - a.length);
      const combinedBody = sorted.join("\n\n---\n\n");
      const entry = speakerEntries.find(e => e.speechKey === speechKey);
      if (!entry) return;
      const extracted = callGeminiForSpeechExtraction_(combinedBody, entry.name, introQuestion, geminiKey);
      // Only store if we got something meaningful — a partial result (e.g. just availability)
      // should not block a later thread that contains the actual speech details
      if (extracted && (extracted.title || extracted.availability === "unavailable" || extracted.availability === "uncertain")) {
        results[speechKey] = extracted;
      }
    });
  });

  // Collect any introQuestion Gemini extracted from the thread (used as fallback
  // when LAST_INTRO_QUESTION was not stored from a prior confirmation email send).
  let extractedIntroQuestion = "";
  Object.values(results).forEach(det => {
    if (!extractedIntroQuestion && det && det.introQuestion) {
      extractedIntroQuestion = det.introQuestion;
    }
  });

  return { details: results, emailFoundKeys, extractedIntroQuestion };
}

/**
 * callGeminiForSpeechExtraction_
 * Sends an email body (or combined thread) to Gemini and asks it to extract
 * structured speech data, returning a parsed JSON object.
 *
 * @param {string} emailBody    - Plain-text body of the speaker's reply(s), separated by ---.
 * @param {string} speakerName  - Full name of the speaker (for context).
 * @param {string} introQuestion- The intro question posed (may be empty).
 * @param {string} geminiKey    - Gemini API key.
 * @return {Object|null} Parsed JSON with {title, purpose, pathway, speechNum, time, intro}, or null on error.
 */
/**
 * getAiModel_
 * Returns the model string and label to use for this call.
 * Priority order:
 *   1. gemini-2.5-flash-lite  — up to 18 pings/day
 *   2. gemini-2.5-flash       — up to 18 pings/day (separate quota)
 *   3. gemma-3-27b-it         — fallback, effectively unlimited
 * @return {{ model: string, label: string }}
 */
function getAiModel_() {
  const props    = PropertiesService.getScriptProperties();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  function getCount(dateKey, countKey) {
    const savedDate = props.getProperty(dateKey) || "";
    return savedDate === todayStr ? parseInt(props.getProperty(countKey) || "0", 10) : 0;
  }

  const liteCount  = getCount("GEMINI_LITE_DATE",  "GEMINI_LITE_COUNT");
  const flashCount = getCount("GEMINI_FLASH_DATE", "GEMINI_FLASH_COUNT");

  if (liteCount < 18) {
    return { model: "gemini-2.5-flash-lite", label: "gemini-lite"  };
  } else if (flashCount < 18) {
    return { model: "gemini-2.5-flash",      label: "gemini-flash" };
  } else {
    return { model: "gemma-3-27b-it",        label: "gemma"        };
  }
}

/**
 * recordAiPing_
 * Increments the daily ping counter for whichever Gemini model was just used.
 * Gemma calls are not tracked (unlimited quota).
 * @param {string} label - "gemini-lite", "gemini-flash", or "gemma"
 */
function recordAiPing_(label) {
  if (label === "gemma") return;
  const props    = PropertiesService.getScriptProperties();
  const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  const dateKey  = label === "gemini-flash" ? "GEMINI_FLASH_DATE"  : "GEMINI_LITE_DATE";
  const countKey = label === "gemini-flash" ? "GEMINI_FLASH_COUNT" : "GEMINI_LITE_COUNT";

  const savedDate = props.getProperty(dateKey) || "";
  const count     = savedDate === todayStr ? parseInt(props.getProperty(countKey) || "0", 10) : 0;
  props.setProperty(dateKey,  todayStr);
  props.setProperty(countKey, String(count + 1));
}

function callGeminiForSpeechExtraction_(emailBody, speakerName, introQuestion, geminiKey) {
  const introLine = introQuestion
    ? "\nThe speaker was asked this introduction question: \"" + introQuestion + "\""
    : "";

  const prompt =
    "You are helping prepare a Toastmasters meeting agenda. " +
    "A speaker has replied to a role confirmation email with their speech details.\n" +
    "Speaker name: " + speakerName + introLine + "\n\n" +
    "Email body (may contain multiple messages separated by ---):\n" + emailBody.substring(0, 6000) + "\n\n" +
    "Extract the following from the email. Return ONLY a valid JSON object with no markdown fences " +
    "and no extra text. Leave a field as an empty string \"\" if the information is absent:\n" +
    "{\n" +
    "  \"availability\": \"available | unavailable | uncertain. Use unavailable if they say they cannot attend or deliver. Use uncertain if unsure or maybe. Use available otherwise.\",\n" +

    "  \"title\": \"the speech title\",\n" +
    "  \"purpose\": \"the speech purpose or objective from their Toastmasters pathway\",\n" +
    "  \"pathway\": \"the Toastmasters pathway name\",\n" +
    "  \"speechNum\": \"the speech number within the pathway — digits only, e.g. \\\"3\\\"\",\n" +
    "  \"time\": \"allotted time, e.g. \\\"5-7 minutes\\\" — if not specified use \\\"5-7 minutes\\\"\",\n" +
    "  \"intro\": \"the speaker's verbatim (or near-verbatim) answer to the intro question, " +
                  "or empty string if no intro question was asked or no answer was given\",\n" +
    (introQuestion ? "" :
    "  \"introQuestion\": \"the exact question the speaker was asked to answer for their introduction, " +
                        "found in the body of the original email they are replying to — " +
                        "or empty string if no such question is present\"\n") +
    "}";

  const MAX_ATTEMPTS = 3;
  for (let attempt = 1; attempt <= MAX_ATTEMPTS; attempt++) {
    try {
      const aiModel = getAiModel_();
      const resp = UrlFetchApp.fetch(
        "https://generativelanguage.googleapis.com/v1beta/models/" + aiModel.model + ":generateContent?key=" + geminiKey,
        {
          method: "post",
          contentType: "application/json",
          muteHttpExceptions: true,
          payload: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: { temperature: 0.1, maxOutputTokens: 1200 }
          })
        }
      );
      if (resp.getResponseCode() !== 200) {
        Logger.log("callGeminiForSpeechExtraction_ attempt " + attempt + ": HTTP " + resp.getResponseCode() + " — " + resp.getContentText().substring(0, 200));
        if (attempt < MAX_ATTEMPTS) { Utilities.sleep(2000); continue; }
        return null;
      }
      const json    = JSON.parse(resp.getContentText());
      const rawText = json?.candidates?.[0]?.content?.parts?.[0]?.text || "";
      const cleaned = rawText.replace(/```json|```/gi, "").trim();
      const parsed  = JSON.parse(cleaned);
      recordAiPing_(aiModel.label);
      return parsed;
    } catch (e) {
      Logger.log("callGeminiForSpeechExtraction_ attempt " + attempt + " failed: " + e.toString());
      if (attempt < MAX_ATTEMPTS) { Utilities.sleep(2000); }
    }
  }
  return null;
}

/**
 * buildIntroductionsDocx_
 * Constructs a branded Speaker Introductions .docx as base64, matching the
 * agenda's visual style (navy header, green accents, Arial font). Each speaker
 * gets a name heading, speech title + time line, and their introduction paragraph.
 *
 * @param {Object[]} introData     - Array of {name, title, time, intro} — only speakers with an intro answer.
 * @param {string}   formattedDate - Long-form date string, e.g. "March 20, 2026".
 * @return {string} Base64-encoded .docx file content.
 */
function buildIntroductionsDocx_(introData, formattedDate, introQuestion) {
  const DARK_GREEN = "1E5631";
  const NAVY       = "1B2A4A";
  const GRAY       = "555555";
  const WHITE      = "FFFFFF";
  const CONTENT_W  = 10512;

  // ── XML helpers ──
  function esc(text) {
    return (text || "").toString()
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
  }

  function rPr(bold, color, size, italic) {
    return "<w:rPr>" +
      (bold   ? "<w:b/><w:bCs/>" : "") +
      (italic ? "<w:i/><w:iCs/>" : "") +
      "<w:color w:val=\"" + (color || "000000") + "\"/>" +
      "<w:sz w:val=\"" + (size || 22) + "\"/><w:szCs w:val=\"" + (size || 22) + "\"/>" +
      "<w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>" +
      "</w:rPr>";
  }

  function run(text, bold, color, size, italic) {
    return "<w:r>" + rPr(bold, color, size, italic) +
      "<w:t xml:space=\"preserve\">" + esc(text) + "</w:t></w:r>";
  }

  function para(content, align, spaceBefore, spaceAfter) {
    return "<w:p>" +
      "<w:pPr>" +
        "<w:jc w:val=\"" + (align || "left") + "\"/>" +
        "<w:spacing w:before=\"" + (spaceBefore || 0) + "\" w:after=\"" + (spaceAfter || 100) + "\"/>" +
      "</w:pPr>" +
      (content || "") +
      "</w:p>";
  }

  // ── Navy branded header (text-only — no logo images) ──
  const headerXml =
    "<w:tbl>" +
    "<w:tblPr>" +
      "<w:tblW w:w=\"" + CONTENT_W + "\" w:type=\"dxa\"/>" +
      "<w:tblBorders>" +
        "<w:top w:val=\"nil\"/><w:left w:val=\"nil\"/><w:bottom w:val=\"nil\"/><w:right w:val=\"nil\"/>" +
        "<w:insideH w:val=\"nil\"/><w:insideV w:val=\"nil\"/>" +
      "</w:tblBorders>" +
    "</w:tblPr>" +
    "<w:tblGrid><w:gridCol w:w=\"" + CONTENT_W + "\"/></w:tblGrid>" +
    "<w:tr>" +
      "<w:tc>" +
        "<w:tcPr>" +
          "<w:tcW w:w=\"" + CONTENT_W + "\" w:type=\"dxa\"/>" +
          "<w:tcBorders><w:top w:val=\"nil\"/><w:left w:val=\"nil\"/><w:bottom w:val=\"nil\"/><w:right w:val=\"nil\"/></w:tcBorders>" +
          "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"" + NAVY + "\"/>" +
          "<w:tcMar>" +
            "<w:top w:w=\"180\" w:type=\"dxa\"/><w:bottom w:w=\"180\" w:type=\"dxa\"/>" +
            "<w:left w:w=\"200\" w:type=\"dxa\"/><w:right w:w=\"200\" w:type=\"dxa\"/>" +
          "</w:tcMar>" +
          "<w:vAlign w:val=\"center\"/>" +
        "</w:tcPr>" +
        para(run("Sierra Speakers Toastmasters", true, WHITE, 30), "center", 0, 40) +
        para(run("Speaker Introductions", false, "C8D8F0", 19), "center", 0, 40) +
        para(run(formattedDate, false, "A8D8A8", 17, true), "center", 0, 0) +
      "</w:tc>" +
    "</w:tr>" +
    "</w:tbl>";

  const dividerXml =
    "<w:p><w:pPr>" +
      "<w:spacing w:before=\"0\" w:after=\"140\"/>" +
      "<w:pBdr><w:bottom w:val=\"single\" w:sz=\"8\" w:space=\"1\" w:color=\"" + DARK_GREEN + "\"/></w:pBdr>" +
    "</w:pPr></w:p>";

  const noteXml = para(
    run("Use the introduction below for each speaker before their speech.", false, GRAY, 17, true),
    "left", 0, introQuestion ? 80 : 240
  );

  const introQXml = introQuestion
    ? para(
        run("Intro question asked: ", true, NAVY, 18) +
        run(introQuestion, false, GRAY, 18, true),
        "left", 0, 240
      )
    : "";

  // ── Speaker sections ──
  let speakerXml = "";
  introData.forEach(function(spk, i) {
    const timeStr  = spk.time  || "5\u20137 minutes";
    const titleStr = spk.title || "(title not confirmed)";
    const metaLine = titleStr + "   |   " + timeStr;

    // Name heading with green underline
    speakerXml +=
      "<w:p>" +
        "<w:pPr>" +
          "<w:spacing w:before=\"" + (i === 0 ? 80 : 320) + "\" w:after=\"60\"/>" +
          "<w:pBdr><w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"1\" w:color=\"" + DARK_GREEN + "\"/></w:pBdr>" +
        "</w:pPr>" +
        run(spk.name, true, NAVY, 24) +
      "</w:p>";

    // Speech title | time
    speakerXml += para(run(metaLine, false, GRAY, 18, true), "left", 0, 80);

    // Intro text
    speakerXml += para(run(spk.intro, false, "222222", 21), "left", 40, 40);
  });

  // ── Assemble document XML ──
  const documentXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
    "<w:document xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
    "  xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
    "<w:body>" +
      headerXml +
      dividerXml +
      noteXml +
      introQXml +
      speakerXml +
      "<w:sectPr>" +
        "<w:pgSz w:w=\"12240\" w:h=\"15840\" w:orient=\"portrait\"/>" +
        "<w:pgMar w:top=\"576\" w:right=\"864\" w:bottom=\"576\" w:left=\"864\" w:header=\"708\" w:footer=\"708\"/>" +
      "</w:sectPr>" +
    "</w:body></w:document>";

  const relsXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
    "  <Relationship Id=\"rId1\"" +
    "    Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"" +
    "    Target=\"styles.xml\"/>" +
    "</Relationships>";

  const stylesXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
    "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">" +
    "<w:docDefaults><w:rPrDefault><w:rPr>" +
      "<w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:cs=\"Arial\"/>" +
      "<w:sz w:val=\"22\"/><w:szCs w:val=\"22\"/>" +
    "</w:rPr></w:rPrDefault></w:docDefaults>" +
    "</w:styles>";

  const contentTypesXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
    "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
    "  <Default Extension=\"xml\"  ContentType=\"application/xml\"/>" +
    "  <Override PartName=\"/word/document.xml\"" +
    "    ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
    "  <Override PartName=\"/word/styles.xml\"" +
    "    ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
    "</Types>";

  const rootRelsXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
    "  <Relationship Id=\"rId1\"" +
    "    Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"" +
    "    Target=\"word/document.xml\"/>" +
    "</Relationships>";

  const files = [
    { name: "[Content_Types].xml",           content: contentTypesXml, mime: "text/xml" },
    { name: "_rels/.rels",                   content: rootRelsXml,     mime: "text/xml" },
    { name: "word/document.xml",             content: documentXml,     mime: "text/xml" },
    { name: "word/styles.xml",               content: stylesXml,       mime: "text/xml" },
    { name: "word/_rels/document.xml.rels",  content: relsXml,         mime: "text/xml" },
  ];

  const blobs = files.map(function(f) {
    return Utilities.newBlob(f.content, f.mime, f.name);
  });

  return Utilities.base64Encode(Utilities.zip(blobs, "introductions.docx").getBytes());
}

// ============================================================
// DIALOG CALLBACKS — called by google.script.run from HTML dialogs
// ============================================================
/**
 * setSpeechSelection
 * Dialog callback — stores the kept speech keys (JSON array) in script properties
 * so the polling loop in generateAgenda can read it.
 * @param {string[]|null} keptKeys - Kept speech key strings, or null if dismissed.
 * @return {void}
 */
function setSpeechSelection(keptKeys, impromptuKeys) {
  PropertiesService.getScriptProperties().setProperty(
    "_speechSelection", JSON.stringify({ kept: keptKeys || [], impromptu: impromptuKeys || [] })
  );
}

/**
 * setImpromptuSelection
 * Dialog callback — stores the impromptu speech keys (JSON array) in script properties.
 * @param {string[]} keys - Array of speech keys marked as impromptu.
 */
function setImpromptuSelection(keys) {
  PropertiesService.getScriptProperties().setProperty(
    "_impromptuSelection", JSON.stringify(keys || [])
  );
}

/**
 * setEvaluatorSelection
 * Dialog callback — stores the kept evaluator keys (JSON array) in script properties
 * so the polling loop in generateAgenda can read it.
 * @param {string[]|null} keptKeys - Kept evaluator key strings, or null if dismissed.
 * @return {void}
 */
function setEvaluatorSelection(keptKeys) {
  PropertiesService.getScriptProperties().setProperty(
    "_evaluatorSelection", keptKeys ? JSON.stringify(keptKeys) : "[]"
  );
}

/**
 * setWotdSelection
 * Dialog callback — stores the Word of the Day selection JSON (or "null") in
 * script properties so the polling loop in generateAgenda can read it.
 * @param {string|null} data - JSON string with { word, mode, def, ex }, or null.
 * @return {void}
 */
function setWotdSelection(data) {
  PropertiesService.getScriptProperties().setProperty(
    "_wotdSelection", data !== null ? data : "null"
  );
}

/**
 * setWotdDefinition
 * Dialog callback — stores the chosen definition index and pronunciation JSON
 * (or "null") in script properties so the polling loop in generateAgenda can read it.
 * @param {string|null} data - JSON string with { idx, pronunciation }, or null.
 * @return {void}
 */
function setWotdDefinition(data) {
  PropertiesService.getScriptProperties().setProperty(
    "_wotdDefinition", data !== null ? data : "null"
  );
}


// ============================================================
// DEBUG — tests the Merriam-Webster API key and sees the raw response
// ============================================================
/**
 * debugMWApi
 * Debug utility — logs the first 1 000 characters of the Merriam-Webster
 * Collegiate Dictionary API response for the word "convenience" to validate
 * that MW_API_KEY is correctly set. Run manually from the Apps Script editor.
 * @return {void}
 */
function debugMWApi() {
  const mwKey = PropertiesService.getScriptProperties().getProperty("MW_API_KEY") || "";

  if (!mwKey) {
    console.log("MW_API_KEY is not set.");
    return;
  }

  console.log("Key: " + mwKey.substring(0, 8));

  try {
    const url = "https://www.dictionaryapi.com/api/v3/references/collegiate/json/convenience?key=" + mwKey;
    const resp = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    console.log("HTTP " + resp.getResponseCode());
    console.log(resp.getContentText().substring(0, 1000));
  } catch(e) {
    console.log("Fetch failed: " + e.toString());
  }
}

// ============================================================
// DIAGNOSTIC — inspect Script Properties storage usage
// ============================================================
/**
 * diagnosticScriptProperties
 * Logs all current Script Properties, their sizes, and total storage used.
 * Also clears any leftover temporary polling keys that may be bloating storage.
 * Run manually from the Apps Script editor to diagnose storage issues.
 * @return {void}
 */
function diagnosticScriptProperties() {
  const props = PropertiesService.getScriptProperties().getProperties();
  const tempKeys = ["_speechSelection", "_evaluatorSelection", "_wotdSelection", "_wotdDefinition", "_agendaMode"];

  let totalBytes = 0;
  let report = "=== Script Properties ===\n";
  for (const [k, v] of Object.entries(props)) {
    const bytes = k.length + v.length;
    totalBytes += bytes;
    const flag = tempKeys.includes(k) ? "  ← STALE TEMP KEY" : "";
    report += `  [${bytes}B] ${k} = ${v.substring(0, 80)}${v.length > 80 ? "…" : ""}${flag}\n`;
  }
  report += `\nTotal: ~${totalBytes} bytes of ~9,216 bytes (9KB) used.\n`;

  // Clean up any stale temp keys left by cancelled/crashed runs
  let cleaned = 0;
  for (const k of tempKeys) {
    if (props[k] !== undefined) {
      PropertiesService.getScriptProperties().deleteProperty(k);
      cleaned++;
    }
  }
  if (cleaned > 0) report += `\n🧹 Cleaned up ${cleaned} stale temp key(s). Try setGeminiKey() again.\n`;
  else report += "\nNo stale temp keys found.\n";

  // Show current model and daily ping counts
  const aiModel    = getAiModel_();
  const liteDate   = props["GEMINI_LITE_DATE"]   || "(never)";
  const liteCount  = props["GEMINI_LITE_COUNT"]  || "0";
  const flashDate  = props["GEMINI_FLASH_DATE"]  || "(never)";
  const flashCount = props["GEMINI_FLASH_COUNT"] || "0";
  report += `\n🤖 Active model: ${aiModel.model}`;
  report += `\n📊 Gemini 2.5 Flash Lite pings (${liteDate}): ${liteCount}/20 — falls back to Flash at 18.`;
  report += `\n📊 Gemini 2.5 Flash pings (${flashDate}): ${flashCount}/20 — falls back to Gemma at 18.\n`;

  console.log(report);
}

// ============================================================
// SETUP — run once to permanently store your Gemini API key
// ============================================================
// ============================================================
// CLUB HYPE EMAIL — standalone menu item
// ============================================================

/**
 * sendClubHypeEmail
 * Menu entry point. Reads the sheet for the next meeting's date, theme,
 * and speakers, then shows a dialog to collect WOTD, agenda URL, and
 * optional guest CC/BCC emails before drafting a Gemini-written hype email.
 */
/**
 * resolveMeetingFormatForColumn_
 * Scans a SCHED sheet column for clues about meeting format.
 * Heuristics (mirrors role confirmation format logic):
 *   - Cell text mentions Zoom/virtual/online AND an in-person address or
 *     "in person"/"hybrid"/"Folsom"/"Asana" -> "hybrid"
 *   - Only Zoom/virtual/online mentions                                 -> "virtual"
 *   - Only address/"in person" mentions                                 -> "in_person"
 *   - Nothing recognisable                                              -> "" (caller defaults)
 *
 * Used by sendClubHypeEmail() to pick a default for the Meeting Format radio.
 *
 * @param {any[][]} data      Full sheet values (from Range.getValues()).
 * @param {number}  colIndex  Zero-based column index of the target date.
 * @return {string} "hybrid" | "in_person" | "virtual" | ""
 */
function resolveMeetingFormatForColumn_(data, colIndex) {
  let sawZoom    = false;
  let sawInPerson = false;
  for (let r = 0; r < data.length; r++) {
    const raw = data[r][colIndex];
    if (raw === undefined || raw === null) continue;
    const s = raw.toString().toLowerCase();
    if (!s) continue;
    if (/zoom|virtual|online/.test(s)) sawZoom = true;
    if (/in[\s-]*person|folsom|asana|hybrid|street|\baddress\b/.test(s)) sawInPerson = true;
  }
  if (sawZoom && sawInPerson) return "hybrid";
  if (sawZoom)                return "virtual";
  if (sawInPerson)            return "in_person";
  return "";
}

function sendClubHypeEmail() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();

  // ── Find the SCHED sheet ──
  const sheets = spreadsheet.getSheets();
  const schedSheets = sheets.map(s => s.getName())
    .filter(n => /^SCHED\s\d{4}$/.test(n))
    .sort((a, b) => parseInt(b.split(" ")[1]) - parseInt(a.split(" ")[1]));
  if (!schedSheets.length) { ui.alert("No SCHED sheet found."); return; }
  const sheet = spreadsheet.getSheetByName(schedSheets[0]);

  const dataRange   = sheet.getDataRange();
  const data        = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();

  // ── Find roles header row ──
  let rolesHeaderRow = -1;
  for (let r = 0; r < data.length; r++) {
    const cell = data[r][0]?.toString().trim().toLowerCase();
    if (cell === "roles") { rolesHeaderRow = r; break; }
    if (cell === "toastmaster") { rolesHeaderRow = r - 1; break; }
  }
  if (rolesHeaderRow < 0) { ui.alert("Could not find roles header row."); return; }

  // ── Find upcoming dates ──
  // Compare date-only (strip time) so today's meeting is still included
  // even if the script runs hours after midnight.
  const now = new Date();
  const todayMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const upcomingDates = [];
  for (let c = 0; c < data[rolesHeaderRow].length; c++) {
    const cell = data[rolesHeaderRow][c];
    if (cell instanceof Date && cell >= todayMidnight) {
      upcomingDates.push({
        date: cell,
        colIndex: c,
        formatted: Utilities.formatDate(cell, Session.getScriptTimeZone(), "M/d/yyyy")
      });
    }
  }
  if (!upcomingDates.length) { ui.alert("No upcoming meeting dates found in the sheet."); return; }

  // ── Default to nearest upcoming date ──
  const nearest  = upcomingDates[0];
  const colIndex = nearest.colIndex;
  const dateStr  = nearest.formatted;
  const longDate = Utilities.formatDate(nearest.date, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");

  // ── Pull theme ──
  const theme = rolesHeaderRow > 0
    ? (data[rolesHeaderRow - 1][colIndex]?.toString().trim() || "")
    : "";

  // ── Pull speakers (Speech N rows) ──
  const speakers = [];
  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    const role = data[r][0]?.toString().trim() || "";
    const name = data[r][colIndex]?.toString().trim() || "";
    if (/^Speech\s*\d+$/i.test(role) && name && name.toUpperCase() !== "TBD") {
      speakers.push({ role, name });
    }
  }
  // — BUG-2 FIX: Build per-date data (theme + speakers) for ALL upcoming dates —
  // so the dialog can refresh fields when the user picks a different week.
  const perDateData = {};
  upcomingDates.forEach(ud => {
    const ci = ud.colIndex;
    const t = rolesHeaderRow > 0
      ? (data[rolesHeaderRow - 1][ci]?.toString().trim() || "")
      : "";
    const sp = [];
    for (let r = rolesHeaderRow + 1; r < data.length; r++) {
      const role = data[r][0]?.toString().trim() || "";
      const name = data[r][ci]?.toString().trim() || "";
      if (/^Speech\s*\d+$/i.test(role) && name && name.toUpperCase() !== "TBD") {
        sp.push({ role, name });
      }
    }
    // BUG-5 FIX: Per-date meeting format so the dialog can default the format
    // radio based on whatever is encoded in the sheet column for that date.
    const fmt = resolveMeetingFormatForColumn_(data, ci) || "hybrid";
    perDateData[ud.formatted] = { theme: t, speakers: sp, meetingFormat: fmt };
  });
  const perDateJson = JSON.stringify(perDateData);

  // BUG-5 FIX: Default meeting format for the initially-selected (nearest) date.
  // Dialog radios default to this; onDateChange() updates it if the user picks a different date.
  const meetingFormatDefault = (perDateData[dateStr] && perDateData[dateStr].meetingFormat) || "hybrid";

  // ── Pre-fill from last agenda generation ──
  // Pre-fill regardless of date match — the user can still change the date in the dialog.
  // If they pick a different date, URL/WOTD may not apply, but theme usually does.
  const lastWotd      = props.getProperty("LAST_AGENDA_WOTD")      || "";
  const lastAgendaUrl = props.getProperty("LAST_AGENDA_URL")       || "";
  const lastTheme     = props.getProperty("LAST_AGENDA_THEME")     || "";
  const lastDate      = props.getProperty("LAST_AGENDA_DATE")      || "";
  const lastWotdDef   = props.getProperty("LAST_AGENDA_WOTD_DEF")  || "";
  const lastWotdEx    = props.getProperty("LAST_AGENDA_WOTD_EX")   || "";
  // URL and WOTD are meeting-specific; theme is carried over always as a default
  const agendaUrl     = lastDate === dateStr ? lastAgendaUrl : "";

  // ── WOD Memory: override pre-fill with cache if available ──
  const wodCacheEmail_ = lookupWodCache_(dateStr);
  const wotdPrefill   = wodCacheEmail_ ? wodCacheEmail_.word : (lastDate === dateStr ? lastWotd : "");
  const wotdDef       = wodCacheEmail_ ? wodCacheEmail_.definition : (lastDate === dateStr ? lastWotdDef : "");
  const wotdEx        = lastDate === dateStr ? lastWotdEx  : "";
  const themePrefill  = theme || lastTheme;

  const speakersJson = JSON.stringify(speakers);
  const datesJson    = JSON.stringify(upcomingDates.map(d => d.formatted));

  // ── Show dialog ──
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      body { font-family: Arial, sans-serif; font-size: 13px; padding: 14px; }
      label { font-weight: bold; display: block; margin-bottom: 3px; }
      .sub  { font-weight: normal; color: #888; font-size: 12px; }
      input, textarea, select { width: 100%; box-sizing: border-box; padding: 6px;
        font-size: 13px; border: 1px solid #ccc; border-radius: 3px; margin-bottom: 10px; }
      .row  { display: flex; gap: 8px; }
      .row > div { flex: 1; }
      .btn  { padding: 7px 18px; border: none; border-radius: 4px; cursor: pointer;
              font-size: 13px; font-weight: bold; }
      .primary { background: #1B2A4A; color: white; }
      .cancel  { background: #fff; border: 1px solid #ccc !important; color: #333; }
      .actions { text-align: right; margin-top: 6px; }
    </style></head><body>
    <div class="row">
      <div>
        <label>Meeting Date</label>
        <select id="dateSelect" onchange="onDateChange()">
          ${upcomingDates.map(d => `<option value="${d.formatted}" ${d.formatted === dateStr ? "selected" : ""}>${d.formatted}</option>`).join("")}
        </select>
      </div>
      <div>
        <label>Theme <span class="sub">(editable)</span></label>
        <input id="theme" value="${themePrefill.replace(/"/g, '&quot;')}">
      </div>
    </div>
    <label>Meeting Format</label>
    <div id="formatRow" style="margin-bottom:10px;font-weight:normal;">
      <label style="display:inline-flex;align-items:center;margin-right:14px;cursor:pointer;font-weight:normal;">
        <input type="radio" name="meetingFormat" value="hybrid">&nbsp;Hybrid
      </label>
      <label style="display:inline-flex;align-items:center;margin-right:14px;cursor:pointer;font-weight:normal;">
        <input type="radio" name="meetingFormat" value="in_person">&nbsp;In Person
      </label>
      <label style="display:inline-flex;align-items:center;margin-right:14px;cursor:pointer;font-weight:normal;">
        <input type="radio" name="meetingFormat" value="virtual">&nbsp;Virtual
      </label>
      <label style="display:inline-flex;align-items:center;cursor:pointer;font-weight:normal;">
        <input type="radio" name="meetingFormat" value="undecided">&nbsp;Undecided
      </label>
    </div>
    <label>Word of the Day <span class="sub">(optional — will be highlighted)</span></label>
    <input id="wotd" value="${wotdPrefill.replace(/"/g, '&quot;')}" placeholder="e.g. Convivial">
    <label>Agenda URL <span class="sub">(optional — adds a button to the email)</span></label>
    <input id="agendaUrl" value="${agendaUrl.replace(/"/g, '&quot;')}" placeholder="https://drive.google.com/…">
    <label>Guest CC/BCC <span class="sub">(optional — comma-separated emails)</span></label>
    <input id="guestEmails" placeholder="guest@example.com, another@example.com">
    <div style="margin-bottom:6px;">
      <label style="display:inline;font-weight:normal;">
        <input type="radio" name="guestMode" value="cc" checked> CC
      </label>
      &nbsp;&nbsp;
      <label style="display:inline;font-weight:normal;">
        <input type="radio" name="guestMode" value="bcc"> BCC
      </label>
    </div>
    <div class="actions">
      <button class="btn cancel" onclick="google.script.host.close()">Cancel</button>
      &nbsp;
      <button class="btn primary" id="draftBtn" onclick="submit()">✉️ Draft Meeting Email</button>
    </div>
    <script>
      const allDates   = ${datesJson};
      let allSpeakers = ${speakersJson}; // BUG-2 FIX: changed to let so onDateChange can update
      const perDateMap  = ${perDateJson}; // BUG-2 FIX: per-date theme+speakers data
      function onDateChange() { // BUG-2 FIX: refresh theme + speakers when date changes
        const sel = document.getElementById('dateSelect').value;
        const dd = perDateMap[sel];
        if (dd) {
          document.getElementById('theme').value = dd.theme || '';
          allSpeakers = dd.speakers;  // update speakers for submit
          // BUG-5 FIX: sync the Meeting Format radio with the newly-selected date.
          if (dd.meetingFormat) setMeetingFormat(dd.meetingFormat);
        }
        // Clear URL/WOTD since they are date-specific and not stored in the sheet per-date
        document.getElementById('agendaUrl').value = '';
        document.getElementById('wotd').value = '';
      }
      // BUG-5 FIX: helper to flip the Meeting Format radios to a given value.
      function setMeetingFormat(fmt) {
        const radios = document.querySelectorAll('input[name="meetingFormat"]');
        radios.forEach(function(r) { r.checked = (r.value === fmt); });
      }
      // Initial default sourced from the sheet for the nearest upcoming date.
      setMeetingFormat('${meetingFormatDefault}');
      function submit() {
        const btn = document.getElementById('draftBtn');
        btn.disabled = true;
        btn.textContent = 'Drafting with Gemini…';
        const guestMode = document.querySelector('input[name="guestMode"]:checked')?.value || 'cc';
        // BUG-5 FIX: Pass selected meeting format through to the server-side draft builder.
        const meetingFormat = document.querySelector('input[name="meetingFormat"]:checked')?.value || 'hybrid';
        google.script.run
          .withSuccessHandler(draftsUrl => {
            document.body.innerHTML =
              '<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;padding:24px;text-align:center;">' +
              '<p style="font-size:36px;margin:0 0 8px;">✅</p>' +
              '<p style="font-weight:bold;font-size:15px;margin:0 0 6px;color:#1B2A4A;">Draft saved to Gmail!</p>' +
              '<p style="font-size:12px;color:#666;margin:0 0 16px;">Find it in your Drafts folder.</p>' +
              '<a href="' + draftsUrl + '" target="_blank" ' +
                'style="display:inline-block;background:#1B2A4A;color:white;padding:9px 20px;' +
                       'border-radius:4px;text-decoration:none;font-size:13px;font-weight:bold;margin-bottom:12px;">' +
                'Open Gmail Drafts</a><br>' +
              '<a href="#" onclick="google.script.host.close();return false;" style="font-size:12px;color:#888;">Close</a>' +
              '</div>';
          })
          .withFailureHandler(err => {
            btn.disabled = false;
            btn.textContent = '✉️ Draft Meeting Email';
            alert('Error: ' + (err.message || JSON.stringify(err)));
          })
          .createClubHypeEmailDraftPublic({
            dateStr:     document.getElementById('dateSelect').value,
            theme:       document.getElementById('theme').value.trim(),
            wotd:        document.getElementById('wotd').value.trim(),
            agendaUrl:   document.getElementById('agendaUrl').value.trim(),
            guestEmails: document.getElementById('guestEmails').value.trim(),
            guestMode:   guestMode,
            meetingFormat: meetingFormat,
            speakers:    allSpeakers,
            longDate:    '${longDate}',
            wotdDef:     '${wotdDef.replace(/'/g, "\\'").replace(/\n/g, " ")}',
            wotdEx:      '${wotdEx.replace(/'/g, "\\'").replace(/\n/g, " ")}',
          });
      }
    </script></body></html>
  `).setWidth(480).setHeight(440);

  SpreadsheetApp.getUi().showModalDialog(html, "Draft Club Meeting Email");
}

/**
 * createClubHypeEmailDraftPublic
 * Public wrapper so google.script.run can call createClubHypeEmailDraft_
 * (functions ending in _ are private in Apps Script and cannot be called from HTML dialogs).
 */
function createClubHypeEmailDraftPublic(opts) {
  createClubHypeEmailDraft_(opts);
  return "https://mail.google.com/mail/u/0/#drafts";
}

/**
 * createClubHypeEmailDraft_
 * Called by the hype email dialog. Uses Gemini to write the body, then
 * builds a branded HTML email and saves it as a Gmail draft.
 * @param {Object} opts - { dateStr, theme, wotd, agendaUrl, guestEmails, guestMode, speakers, longDate }
 */
function createClubHypeEmailDraft_(opts) {
  const { dateStr, theme, wotd, agendaUrl, guestEmails, guestMode, speakers, longDate, wotdDef, wotdEx, meetingFormat } = opts;
  // BUG-5 FIX: Meeting format drives whether Zoom details, address, or both appear in the email body.
  const fmt = (meetingFormat === "hybrid" || meetingFormat === "in_person" || meetingFormat === "virtual" || meetingFormat === "undecided")
    ? meetingFormat
    : "hybrid";

  const ZOOM_URL    = "https://us02web.zoom.us/j/648879176";
  const ZOOM_PASS   = "sierra4844";
  const ZOOM_ID     = "648 879 176";
  const ADDRESS     = "633 Folsom Street, San Francisco CA";
  const CLUB_EMAIL  = "sierra-speakers@googlegroups.com";

  // ── Build Gemini prompt ──
  const speakersList = speakers.length
    ? speakers.map(s => `${s.name} (${s.role})`).join(", ")
    : "various members";

  const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
  let hypeBody = "";

  if (geminiKey) {
    const prompt = `You are writing a short, enthusiastic club announcement email for Sierra Speakers Toastmasters.
Meeting date: ${longDate}
${theme ? `Meeting theme: "${theme}" — this theme should be the centerpiece of the email, referenced creatively and repeatedly.` : "No specific theme."}
${wotd ? `Word of the Day: "${wotd}" — weave this word naturally into the email body at least once.` : ""}
${speakers.length ? `Prepared speakers: ${speakersList}` : ""}
${fmt === "hybrid"    ? "Meeting format: HYBRID — some members will join in person at the Asana office, others on Zoom. Invite folks to attend whichever way works." : ""}
${fmt === "in_person" ? "Meeting format: IN PERSON only at the Asana office this week — no Zoom option. Encourage folks to come in person." : ""}
${fmt === "virtual"   ? "Meeting format: VIRTUAL only this week — everyone is on Zoom, no in-person location. Mention we\'re meeting online this week." : ""}
${fmt === "undecided" ? "Meeting format: TBD — we\'re still confirming whether the meeting is in person, virtual, or hybrid; acknowledge this gently and say details will follow." : ""}

Write 2–3 short paragraphs (plain text, no markdown, no bullet points, no headers) that:
- Open with an energetic hook about the meeting theme (if provided), making it feel exciting and relevant
- Mention the prepared speakers by first name only if provided
- ${wotd ? `Naturally use the word "${wotd}" somewhere in the body` : ""}
- Close with a warm, upbeat invitation to attend

Tone: warm, encouraging, community-focused. Max 120 words total. No sign-off line needed.`;

    const { model } = getAiModel_();
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        const resp = UrlFetchApp.fetch(
          `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${geminiKey}`,
          {
            method: "post",
            contentType: "application/json",
            muteHttpExceptions: true,
            payload: JSON.stringify({
              contents: [{ parts: [{ text: prompt }] }],
              generationConfig: { temperature: 0.8, maxOutputTokens: 400 }
            })
          }
        );
        const parsed = JSON.parse(resp.getContentText());
        const text = parsed?.candidates?.[0]?.content?.parts?.[0]?.text?.trim();
        if (text) { hypeBody = text; recordAiPing_(model); break; }
      } catch (e) {
        if (attempt < 3) Utilities.sleep(2000);
      }
    }
  }

  // Fallback if Gemini fails
  if (!hypeBody) {
    hypeBody = `Hey Sierra Speakers! We have an exciting meeting coming up on ${longDate}${theme ? ` with the theme "${theme}"` : ""} and we'd love to see you there.\n\n` +
      (speakers.length ? `Join us as ${speakers.map(s => s.name.split(" ")[0]).join(", ")} take the stage with their prepared speeches.\n\n` : "") +
      `Whether you're a seasoned speaker or just curious about Toastmasters, this is a great opportunity to learn, grow, and connect.`;
  }

  // ── Did You Know? — Gemini generates one fun fact per available input ──
  let didYouKnowFacts = [];
  if (geminiKey) {
    const factParts = [];
    if (theme) factParts.push(
      `One punchy, surprising historical or cultural fun fact (1 sentence, max 25 words) about the concept of "${theme}". ` +
      `Start with the fact itself, no preamble like "Did you know".`
    );
    if (wotd) factParts.push(
      `One punchy etymology or origin fun fact (1 sentence, max 25 words) about the word "${wotd}". ` +
      `Start with the fact itself, no preamble like "Did you know".`
    );
    if (factParts.length) {
      const factPrompt =
        "Return ONLY a valid JSON array of strings with no markdown fences and no extra text. " +
        "Each element is one fun fact as instructed:\n" +
        factParts.map((p, i) => `${i + 1}. ${p}`).join("\n");
      const { model } = getAiModel_();
      for (let attempt = 1; attempt <= 3; attempt++) {
        try {
          const resp = UrlFetchApp.fetch(
            `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${geminiKey}`,
            {
              method: "post",
              contentType: "application/json",
              muteHttpExceptions: true,
              payload: JSON.stringify({
                contents: [{ parts: [{ text: factPrompt }] }],
                generationConfig: { temperature: 0.9, maxOutputTokens: 200 }
              })
            }
          );
          const parsed = JSON.parse(resp.getContentText());
          const raw = parsed?.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";
          const cleaned = raw.replace(/```json|```/gi, "").trim();
          didYouKnowFacts = JSON.parse(cleaned);
          recordAiPing_(model);
          break;
        } catch (e) {
          if (attempt < 3) Utilities.sleep(2000);
        }
      }
    }
  }

  // ── Highlight WOTD in hype body ──
  let hypeBodyHtml = hypeBody
    .split("\n\n")
    .map(para => `<p style="margin:0 0 12px;line-height:1.6;">${para.trim()}</p>`)
    .join("");

  if (wotd) {
    // Bold + highlight every occurrence of the WOTD (case-insensitive)
    const wotdRegex = new RegExp(`(${wotd.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")})`, "gi");
    hypeBodyHtml = hypeBodyHtml.replace(wotdRegex,
      `<strong style="background:#fff9c4;padding:1px 3px;border-radius:2px;">$1</strong>`);
  }

  // ── Speaker callout block ──
  const speakerBlock = speakers.length ? `
    <div style="background:#f0f4f8;border-left:4px solid #1B2A4A;border-radius:0 4px 4px 0;padding:12px 16px;margin:16px 0;">
      <p style="margin:0 0 6px;font-weight:bold;color:#1B2A4A;font-size:13px;text-align:center;text-transform:uppercase;letter-spacing:1px;font-size:11px;">Prepared Speeches</p>
      ${speakers.map(s => `<p style="margin:0 0 4px;font-size:13px;">• <strong>${s.name}</strong></p>`).join("")}
    </div>` : "";

  // ── WOTD callout ──
  const wotdBlock = wotd ? `
    <div style="background:#fff9c4;border:1px solid #f0c040;border-radius:4px;padding:12px 16px;margin:16px 0;text-align:center;">
      <p style="margin:0;font-size:11px;color:#7a6000;text-transform:uppercase;letter-spacing:1px;">Word of the Day</p>
      <p style="margin:4px 0 6px;font-size:22px;font-weight:bold;color:#5a4000;">${wotd}</p>
      ${wotdDef ? `<p style="margin:0 0 4px;font-size:13px;color:#5a4000;">${wotdDef}</p>` : ""}
      ${wotdEx  ? `<p style="margin:0;font-size:12px;color:#7a6000;font-style:italic;">"${wotdEx}"</p>` : ""}
    </div>` : "";

  // ── Did You Know block ──
  // De-duplicate facts to prevent rendering the same fact twice (BUG-FIX)
  const uniqueFacts = Array.from(new Set(didYouKnowFacts));
  const didYouKnowBlock = uniqueFacts.length ? `
    <div style="background:#f0f4f8;border:1px solid #c8d8e8;border-radius:4px;padding:12px 16px;margin:16px 0;">
      <p style="margin:0 0 8px;font-size:11px;font-weight:bold;color:#1B2A4A;text-transform:uppercase;letter-spacing:1px;">Did You Know?</p>
      ${uniqueFacts.map(f => `<p style="margin:0 0 6px;font-size:13px;color:#333;line-height:1.5;">${f}</p>`).join("")}
    </div>` : "";

  // ── Agenda button ──
  const agendaBlock = agendaUrl ? `
    <div style="text-align:center;margin:16px 0;">
      <a href="${agendaUrl}"
         style="display:inline-block;background:#1B2A4A;color:white;padding:10px 22px;
                border-radius:4px;text-decoration:none;font-weight:bold;font-size:14px;">
        View the Meeting Agenda
      </a>
    </div>` : "";

  // ── Format-aware location & attendance copy ──
  // BUG-5 FIX: Hype email body now adapts to the meeting format chosen in the dialog.
  //   hybrid    -> address + Zoom + "attend either way" line
  //   in_person -> address only, no Zoom
  //   virtual   -> Zoom only, explicit "this week the meeting is virtual"
  //   undecided -> no address, no Zoom, format TBD note
  let locationDetails;    // inner HTML of the Meeting Details card
  let closingLine;        // plain closing paragraph under the card
  let plainLocationLines; // plain-text equivalent for the fallback body
  if (fmt === "virtual") {
    locationDetails =
      '<p style="margin:0 0 6px;">This week the meeting is <strong>virtual</strong> — please join us on Zoom.</p>' +
      '<p style="margin:0 0 2px;"><strong>Virtual (Zoom):</strong></p>' +
      '<p style="margin:0 0 2px;"><a href="' + ZOOM_URL + '" style="color:#1B2A4A;">' + ZOOM_URL + '</a></p>' +
      '<p style="margin:0 0 2px;">Meeting ID: ' + ZOOM_ID + '</p>' +
      '<p style="margin:0;">Passcode: <strong>' + ZOOM_PASS + '</strong></p>';
    closingLine        = "We hope to see you on Zoom this week!";
    plainLocationLines = "This week the meeting is virtual.\nZoom: " + ZOOM_URL +
                         "\n   Meeting ID: " + ZOOM_ID + " | Passcode: " + ZOOM_PASS + "\n\n";
  } else if (fmt === "in_person") {
    locationDetails =
      '<p style="margin:0 0 6px;">This week the meeting is <strong>in person only</strong> — no Zoom option.</p>' +
      '<p style="margin:0;"><strong>In-Person (Asana HQ):</strong><br>' + ADDRESS + '</p>';
    closingLine        = "We hope to see you in person at the Asana office this week!";
    plainLocationLines = "This week the meeting is in person only.\nIn-Person (Asana HQ): " + ADDRESS + "\n\n";
  } else if (fmt === "undecided") {
    locationDetails =
      '<p style="margin:0;font-style:italic;color:#666;">Heads up: the meeting format (in person, virtual, or hybrid) is still being decided — we\'ll confirm before the meeting.</p>';
    closingLine        = "We'll share the final meeting format shortly — stay tuned!";
    plainLocationLines = "Heads up: the meeting format (in person, virtual, or hybrid) is still being decided - we'll confirm before the meeting.\n\n";
  } else { // hybrid (default)
    locationDetails =
      '<p style="margin:0 0 6px;"><strong>In-Person (Asana HQ):</strong><br>' + ADDRESS + '</p>' +
      '<p style="margin:0 0 2px;"><strong>Virtual (Zoom):</strong></p>' +
      '<p style="margin:0 0 2px;"><a href="' + ZOOM_URL + '" style="color:#1B2A4A;">' + ZOOM_URL + '</a></p>' +
      '<p style="margin:0 0 2px;">Meeting ID: ' + ZOOM_ID + '</p>' +
      '<p style="margin:0;">Passcode: <strong>' + ZOOM_PASS + '</strong></p>';
    closingLine        = "We hope to see you there — in person at the Asana office or virtually on Zoom!";
    plainLocationLines = "In-Person (Asana HQ): " + ADDRESS + "\nZoom: " + ZOOM_URL +
                         "\n   Meeting ID: " + ZOOM_ID + " | Passcode: " + ZOOM_PASS + "\n\n";
  }

  // ── Location details ──
  const locationBlock = `
    <div style="background:#f9f9f9;border:1px solid #e0e0e0;border-radius:4px;padding:12px 16px;margin:16px 0;font-size:13px;">
      <p style="margin:0 0 10px;font-weight:bold;color:#1B2A4A;">Meeting Details</p>
      <div style="background:#1B2A4A;border-radius:4px;padding:10px 14px;margin:0 0 12px;text-align:center;">
        <p style="margin:0 0 2px;font-size:12px;color:#aac8e0;text-transform:uppercase;letter-spacing:1px;">Doors Open</p>
        <p style="margin:0 0 6px;font-size:20px;font-weight:bold;color:white;">6:00 PM</p>
        <p style="margin:0;font-size:11px;color:#C6E6C6;">Meeting starts promptly at <strong style="color:white;font-size:13px;">6:05 PM</strong></p>
      </div>
      ${locationDetails}
    </div>`;

  // ── Build full HTML email ──
  const themeHeadline = theme
    ? `<p style="margin:0 0 4px;font-size:12px;color:#C6E6C6;text-transform:uppercase;letter-spacing:1px;">Theme</p>
       <p style="margin:0;font-size:18px;font-weight:bold;color:white;">${theme}</p>`
    : "";

  const htmlBody = `
    <div style="font-family:Arial,sans-serif;width:100%;border:1px solid #ddd;">
      <div style="background:#1B2A4A;padding:20px 24px;text-align:center;">
        <p style="margin:0 0 2px;color:white;font-size:22px;font-weight:bold;">Sierra Speakers</p>
        <p style="margin:0 0 12px;color:#C6E6C6;font-size:13px;">Toastmasters International</p>
        <p style="margin:0 0 4px;font-size:13px;color:#aac8e0;">${longDate}</p>
        ${themeHeadline}
      </div>
      <div style="padding:20px 24px;background:#ffffff;">
        ${hypeBodyHtml}
        ${wotdBlock}
        ${didYouKnowBlock}
        ${speakerBlock}
        ${agendaBlock}
        ${locationBlock}
        <p style="margin:16px 0 0;font-size:13px;color:#444;">
          ${closingLine}
        </p>
      </div>
      <div style="background:#1E5631;padding:12px 24px;">
        <p style="margin:0;color:#C6E6C6;font-size:12px;">
          Sierra Speakers Toastmasters &nbsp;·&nbsp; Inspiring communicators since day one
        </p>
      </div>
    </div>`;

  // ── Plain text fallback ──
  const plainText = `Hey Sierra Speakers!\n\n${hypeBody}\n\n` +
    (wotd ? `Word of the Day: ${wotd}\n\n` : "") +
    (speakers.length ? `Prepared Speeches:\n${speakers.map(s => `• ${s.name}`).join("\n")}\n\n` : "") +
    (agendaUrl ? `View the agenda: ${agendaUrl}\n\n` : "") +
    plainLocationLines +
    `See you Thursday!\nSierra Speakers`;

  const subject = `Sierra Speakers — ${longDate}${theme ? " | " + theme : ""}`;

  // ── CC/BCC handling ──
  const guestList = guestEmails
    ? guestEmails.split(",").map(e => e.trim()).filter(e => e.includes("@"))
    : [];

  const draftOptions = { htmlBody };
  if (guestList.length) {
    if (guestMode === "bcc") draftOptions.bcc = guestList.join(",");
    else                     draftOptions.cc  = guestList.join(",");
  }

  GmailApp.createDraft(CLUB_EMAIL, subject, plainText, draftOptions);
}

/**
 * setMeetingDetails
 * Kept for reference — Zoom/address are now hardcoded in createClubHypeEmailDraft_
 * since they rarely change. Update them directly in that function if needed.
 */
function setMeetingDetails() {
  console.log("Zoom and address are now hardcoded in createClubHypeEmailDraft_(). Edit that function directly to update them.");
}



/**
 * setGeminiKey
 * One-time setup: paste your Gemini API key into the string below,
 * run this function once from the Apps Script editor, then delete the key
 * from the source code. The key is stored in Script Properties permanently.
 * @return {void}
 */
function setGeminiKey() {
  const YOUR_KEY_HERE = "PASTE_YOUR_KEY_HERE"; // ← replace, run once, then delete key from here
  if (YOUR_KEY_HERE === "PASTE_YOUR_KEY_HERE") {
    console.log("ERROR: You haven't pasted your key yet. Edit this function first.");
    return;
  }
  PropertiesService.getScriptProperties().setProperty("GEMINI_API_KEY", YOUR_KEY_HERE);
  console.log("✅ GEMINI_API_KEY saved successfully. You can now remove the key from the source code.");
}

// ============================================================
// DEBUG — tests the Gemini API key and sees the raw response
// ============================================================
/**
 * debugGeminiApi
 * Debug utility — sends a test prompt to Gemini for the word "convenience"
 * and logs the raw response to validate that GEMINI_API_KEY is correctly set.
 * Run manually from the Apps Script editor.
 * @return {void}
 */
function debugGeminiApi() {
  const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";

  if (!geminiKey) {
    console.log("GEMINI_API_KEY is not set in Script Properties.");
    return;
  }

  console.log("Key prefix: " + geminiKey.substring(0, 8));

  try {
    const prompt =
      "The Word of the Day is \"convenience\". The meeting theme is \"Modern Life\". " +
      "Respond with ONLY a valid JSON object (no markdown) in this format: " +
      "{\"definition\": \"...\", \"pronunciation\": \"...\", \"example\": \"...\"}";

    const aiModel = getAiModel_();
    console.log("Using model: " + aiModel.model);
    const resp = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/" + aiModel.model + ":generateContent?key=" + geminiKey,
      {
        method: "post",
        contentType: "application/json",
        muteHttpExceptions: true,
        payload: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.4, maxOutputTokens: 300 }
        })
      }
    );
    console.log("HTTP " + resp.getResponseCode());
    console.log(resp.getContentText().substring(0, 1000));
  } catch(e) {
    console.log("Fetch failed: " + e.toString());
  }
}

// ============================================================
// DEBUG — diagnose inbox scan for speaker email replies
// ============================================================
/**
 * debugSpeakerScan
 * Run manually from the Apps Script editor to diagnose why the inbox
 * scan is not finding speaker replies. Logs:
 *   1. The email addresses pulled from the roster for each speaker
 *   2. The exact Gmail search query used
 *   3. Every thread and message found (sender, date, first 200 chars)
 *   4. The raw Gemini extraction result for each speaker found
 *
 * HOW TO USE:
 *   1. Open Apps Script editor
 *   2. Select this function from the dropdown
 *   3. Click Run
 *   4. Open View → Logs to see the output
 * @return {void}
 */
function debugSpeakerScan() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();

  // ── Find most recent SCHED sheet ──
  const schedSheets = sheets
    .map(s => s.getName())
    .filter(name => /^SCHED\s\d{4}$/.test(name))
    .sort((a, b) => parseInt(b.split(" ")[1]) - parseInt(a.split(" ")[1]));

  if (schedSheets.length === 0) { console.log("ERROR: No SCHED sheet found."); return; }
  const sheet = spreadsheet.getSheetByName(schedSheets[0]);
  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();

  // ── Build nameToEmail from member directory ──
  const nameToEmail = {};
  let r = 1;
  while (r < data.length) {
    const firstName = data[r][0]?.toString().trim();
    const lastName  = data[r][1]?.toString().trim();
    const email     = data[r][4]?.toString().trim();
    const bgColor   = backgrounds[r]?.[0];
    if (bgColor === "#cfe2f3" || (!firstName && !email)) break;
    if (firstName && lastName && email) nameToEmail[firstName + " " + lastName] = email;
    r++;
  }

  console.log("=== nameToEmail map ===");
  Object.entries(nameToEmail).forEach(([name, email]) => console.log("  " + name + " -> " + email));

  // ── Find roles header row and most upcoming date column ──
  let rolesHeaderRow = -1;
  for (let i = 0; i < data.length; i++) {
    if (data[i][0]?.toString().trim().toLowerCase() === "roles") { rolesHeaderRow = i; break; }
  }
  if (rolesHeaderRow < 0) { console.log("ERROR: Could not find Roles header row."); return; }

  let firstDateCol = -1;
  for (let c = 0; c < data[rolesHeaderRow].length; c++) {
    if (data[rolesHeaderRow][c] instanceof Date) { firstDateCol = c; break; }
  }
  if (firstDateCol < 0) { console.log("ERROR: No date columns found."); return; }

  const now = new Date(); now.setHours(0,0,0,0);
  let colIndex = -1;
  for (let c = firstDateCol; c < data[rolesHeaderRow].length; c++) {
    const h = data[rolesHeaderRow][c];
    if (h instanceof Date && h >= now) { colIndex = c; break; }
  }
  if (colIndex < 0) { console.log("ERROR: No upcoming meeting date found."); return; }

  const meetingDate = Utilities.formatDate(data[rolesHeaderRow][colIndex], Session.getScriptTimeZone(), "M/d/yyyy");
  console.log("\n=== Meeting date: " + meetingDate + " (col " + colIndex + ") ===");

  // ── Collect speaker names and emails for this meeting ──
  const speakerEntries = [];
  let speechCounter = 1;
  for (let i = rolesHeaderRow + 1; i < data.length; i++) {
    const roleRaw = data[i][0]?.toString().trim().toLowerCase();
    if (!roleRaw) continue;
    if (roleRaw.startsWith("speech")) {
      const name  = data[i][colIndex]?.toString().trim();
      const email = name ? (nameToEmail[name] || "") : "";
      const key   = "Speech " + speechCounter++;
      console.log("  " + key + ": name=\"" + name + "\"  email=\"" + email + "\"");
      if (name && name.toUpperCase() !== "TBD") speakerEntries.push({ speechKey: key, name, email });
    }
  }

  if (speakerEntries.length === 0) { console.log("No speakers found for this meeting."); return; }

  const emailList = speakerEntries.map(e => e.email).filter(Boolean);
  if (emailList.length === 0) {
    console.log("\nWARNING: Speaker names found but none have emails in the roster.");
    console.log("Check that schedule names exactly match the member directory names.");
    return;
  }

  // ── Run the exact same Gmail search used in production ──
  const fromClause = emailList.map(e => "from:" + e).join(" OR ");
  const query = "(" + fromClause + ") newer_than:7d";
  console.log("\n=== Gmail search query ===\n  " + query);

  let threads;
  try {
    threads = GmailApp.search(query, 0, 50);
  } catch(e) {
    console.log("ERROR: Gmail search failed: " + e.toString());
    return;
  }

  console.log("\n=== Threads found: " + threads.length + " ===");
  if (threads.length === 0) {
    console.log("  No threads found. Possible reasons:");
    console.log("  - Reply arrived more than 7 days ago");
    console.log("  - Sender address doesn't match the roster");
    console.log("  - Email is in Spam or an unexpected label");
    return;
  }

  // ── Log every message in every thread ──
  const emailToEntry = {};
  speakerEntries.forEach(e => { if (e.email) emailToEntry[e.email.toLowerCase().trim()] = e; });

  threads.forEach((thread, ti) => {
    console.log("\n  Thread " + (ti+1) + ": \"" + thread.getFirstMessageSubject() + "\"");
    thread.getMessages().forEach((msg, mi) => {
      const fromRaw   = msg.getFrom();
      const fromEmail = (fromRaw.match(/<(.+)>/) ? fromRaw.match(/<(.+)>/)[1] : fromRaw).trim().toLowerCase();
      const date      = Utilities.formatDate(msg.getDate(), Session.getScriptTimeZone(), "M/d/yyyy h:mm a");
      const matched   = emailToEntry[fromEmail] ? " <-- MATCHES " + emailToEntry[fromEmail].speechKey : "";
      console.log("    [" + mi + "] from=" + fromRaw + "  (" + date + ")" + matched);
      if (matched) {
        const body = msg.getPlainBody() || "";
        console.log("        body preview: " + body.substring(0, 200).replace(/\n/g, " "));
      }
    });
  });

  // ── Run Gemini extraction and log raw result ──
  const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY") || "";
  if (!geminiKey) { console.log("\nNOTE: GEMINI_API_KEY not set -- skipping extraction test."); return; }

  const introQ = PropertiesService.getScriptProperties().getProperty("LAST_INTRO_QUESTION") || "";
  console.log("\n=== Running Gemini extraction ===");
  console.log("  Stored intro question: \"" + introQ + "\"");

  const extracted = scanSpeakerEmails_(speakerEntries, introQ);
  console.log("  Email found keys: " + JSON.stringify([...extracted.emailFoundKeys]));
  console.log("  Extraction results: " + JSON.stringify(extracted.details, null, 2));
}

// ============================================================================
// DEPLOY TO ANOTHER WORKBOOK
// Copies Script Properties (and optionally WOD_Memory data) to a target sheet.
// ============================================================================

/**
 * deployToAnotherSheet
 * Entry point triggered from the Toastmasters menu.
 * Shows a dialog to collect the target spreadsheet URL, then orchestrates
 * the property-transfer, WOD_Memory copy, and code-copy helper.
 * @return {void}
 */
function deployToAnotherSheet() {
  const ui = SpreadsheetApp.getUi();

  // --- 1. Ask for the target sheet URL ---
  const urlResponse = ui.prompt(
    'Deploy to Another Sheet',
    'Enter the URL of the TARGET Google Sheet (the master/production sheet):',
    ui.ButtonSet.OK_CANCEL
  );
  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;

  const targetUrl = urlResponse.getResponseText().trim();
  if (!targetUrl) {
    ui.alert('Deploy cancelled', 'No URL provided.', ui.ButtonSet.OK);
    return;
  }

  // --- 2. Open the target spreadsheet ---
  let targetSs;
  try {
    targetSs = SpreadsheetApp.openByUrl(targetUrl);
  } catch (e) {
    ui.alert(
      'Error opening target sheet',
      'Could not open the spreadsheet at:\n' + targetUrl +
      '\n\nMake sure the URL is correct and you have edit access.\n\nError: ' + e.message,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 3. Read all Script Properties from THIS project ---
  const sourceProps = PropertiesService.getScriptProperties().getAll();
  const propKeys = Object.keys(sourceProps).sort();

  if (propKeys.length === 0) {
    ui.alert('No Script Properties', 'This project has no Script Properties to deploy.', ui.ButtonSet.OK);
    return;
  }

  // --- 4. Handle SCHEDULING_SHEET_URL_ (code-level const) ---
  //   SCHEDULING_SHEET_URL_ is a hardcoded const in Code.gs, NOT a Script Property.
  //   We alert the user that they will need to update it manually in the target code.
  //   However, if there is a script property with that name, we handle it.

  // --- 5. Decide what to deploy — show a preview ---
  const maskedSummary = propKeys.map(function(key) {
    return '  ' + key + ' = ' + maskSensitiveValue_(key, sourceProps[key]);
  }).join('\n');

  const confirmResponse = ui.alert(
    'Confirm Property Deployment',
    'The following ' + propKeys.length + ' Script Properties will be copied to:\n' +
    targetSs.getName() + '\n\n' +
    maskedSummary + '\n\n' +
    'Proceed?',
    ui.ButtonSet.YES_NO
  );
  if (confirmResponse !== ui.Button.YES) {
    ui.alert('Deploy cancelled.', '', ui.ButtonSet.OK);
    return;
  }

  // --- 6. Ask about SCHEDULING_SHEET_URL_ script property ---
  //   When deploying to the master sheet, this should typically be null
  //   (self-reference) instead of pointing to the staging copy.
  let schedulingUrlAction = 'keep'; // default
  if (sourceProps['SCHEDULING_SHEET_URL_'] !== undefined) {
    const schedResponse = ui.alert(
      'SCHEDULING_SHEET_URL_ Setting',
      'The source has a SCHEDULING_SHEET_URL_ script property.\n\n' +
      'The target sheet URL is:\n' + targetUrl + '\n\n' +
      'For the MASTER sheet this should typically be null (self-reference).\n\n' +
      'Click YES to set it to null (recommended for master sheet).\n' +
      'Click NO to keep the current value as-is.',
      ui.ButtonSet.YES_NO
    );
    if (schedResponse === ui.Button.YES) {
      schedulingUrlAction = 'null';
    }
  }

  // --- 7. Build properties object for deployment ---
  //   NOTE: PropertiesService is scoped to the *calling* project.
  //   We cannot write to another project's Script Properties directly.
  //   Instead we generate a helper snippet the user runs in the target project.

  const propsToSet = {};
  propKeys.forEach(function(key) {
    if (key === 'SCHEDULING_SHEET_URL_' && schedulingUrlAction === 'null') {
      propsToSet[key] = '';  // empty string represents null
    } else {
      propsToSet[key] = sourceProps[key];
    }
  });

  const propsJson = JSON.stringify(propsToSet, null, 2);

  // --- 8. Ask about WOD_Memory ---
  let wodCopied = false;
  const wodResponse = ui.alert(
    'Copy WOD_Memory Data?',
    'Do you also want to copy the WOD_Memory sheet data to the target spreadsheet?\n\n' +
    'This copies the Word of the Day history so the target sheet has the same memory.',
    ui.ButtonSet.YES_NO
  );
  if (wodResponse === ui.Button.YES) {
    wodCopied = copyWodMemoryToTarget_(targetSs);
  }

  // --- 9. Show the completion dialog with helper snippet + code copy instructions ---
  showDeployCompletionDialog_(targetUrl, targetSs.getName(), propsJson, maskedSummary, wodCopied, propKeys.length);
}

/**
 * maskSensitiveValue_
 * Partially masks values that look like API keys or secrets.
 * Shows first 4 and last 4 characters for keys, full value for short/safe values.
 * @param {string} key   - The property key name.
 * @param {string} value - The property value.
 * @return {string} The masked or original value.
 */
function maskSensitiveValue_(key, value) {
  if (!value) return '(empty)';

  const sensitivePatterns = ['KEY', 'SECRET', 'TOKEN', 'PASSWORD', 'API'];
  const isSensitive = sensitivePatterns.some(function(pattern) {
    return key.toUpperCase().indexOf(pattern) >= 0;
  });

  if (isSensitive && value.length > 12) {
    return value.substring(0, 4) + '****' + value.substring(value.length - 4);
  }

  // Truncate long values
  if (value.length > 60) {
    return value.substring(0, 57) + '...';
  }

  return value;
}

/**
 * copyWodMemoryToTarget_
 * Copies the WOD_Memory sheet data from this spreadsheet to the target.
 * Creates the sheet in the target if it doesn't exist; clears and replaces
 * data if it does.
 * @param {SpreadsheetApp.Spreadsheet} targetSs - The target spreadsheet.
 * @return {boolean} True if the copy succeeded.
 */
function copyWodMemoryToTarget_(targetSs) {
  try {
    const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = sourceSs.getSheetByName('WOD_Memory');
    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert('WOD_Memory sheet not found in source spreadsheet.');
      return false;
    }

    const data = sourceSheet.getDataRange().getValues();
    if (data.length === 0) {
      SpreadsheetApp.getUi().alert('WOD_Memory sheet is empty \u2014 nothing to copy.');
      return false;
    }

    let targetSheet = targetSs.getSheetByName('WOD_Memory');
    if (!targetSheet) {
      targetSheet = targetSs.insertSheet('WOD_Memory');
    } else {
      targetSheet.clear();
    }

    targetSheet.getRange(1, 1, data.length, data[0].length).setValues(data);

    // Hide the sheet in the target to match source behavior
    if (sourceSheet.isSheetHidden()) {
      targetSheet.hideSheet();
    }

    return true;
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error copying WOD_Memory: ' + e.message);
    return false;
  }
}

/**
 * showDeployCompletionDialog_
 * Displays an HTML dialog with:
 *   - A summary of transferred properties
 *   - A helper snippet the user can paste into the target Apps Script to set properties
 *   - Instructions for manual code deployment
 * @param {string}  targetUrl      - URL of the target spreadsheet.
 * @param {string}  targetName     - Name of the target spreadsheet.
 * @param {string}  propsJson      - JSON string of properties to set.
 * @param {string}  maskedSummary  - Human-readable masked property summary.
 * @param {boolean} wodCopied      - Whether WOD_Memory was copied.
 * @param {number}  propCount      - Number of properties transferred.
 * @return {void}
 */
function showDeployCompletionDialog_(targetUrl, targetName, propsJson, maskedSummary, wodCopied, propCount) {
  // Build the helper snippet that the user runs in the TARGET project
  const helperSnippet =
    '/**\n' +
    ' * RUN THIS ONCE in the TARGET project to set Script Properties.\n' +
    ' * After running, you can delete this function.\n' +
    ' */\n' +
    'function _pasteDeployedProperties() {\n' +
    '  var props = ' + propsJson + ';\n' +
    '  var sp = PropertiesService.getScriptProperties();\n' +
    '  Object.keys(props).forEach(function(k) {\n' +
    '    sp.setProperty(k, props[k]);\n' +
    '  });\n' +
    '  SpreadsheetApp.getUi().alert("Done! " + Object.keys(props).length + " properties set.");\n' +
    '}\n';

  // Escape for HTML embedding
  var escapedSnippet = helperSnippet
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

  var escapedSummary = maskedSummary
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
  var html = '<html><head><style>' +
    'body { font-family: Arial, sans-serif; font-size: 13px; padding: 12px; line-height: 1.5; }' +
    'h3 { margin: 12px 0 6px; color: #1a73e8; }' +
    '.summary { background: #f1f3f4; padding: 10px; border-radius: 6px; ' +
    '  white-space: pre-wrap; font-family: monospace; font-size: 11px; max-height: 150px; overflow-y: auto; }' +
    '.snippet { background: #263238; color: #eeffff; padding: 10px; border-radius: 6px; ' +
    '  white-space: pre-wrap; font-family: monospace; font-size: 11px; max-height: 200px; overflow-y: auto; }' +
    '.btn { display: inline-block; margin: 6px 4px 6px 0; padding: 8px 16px; ' +
    '  background: #1a73e8; color: white; border: none; border-radius: 4px; cursor: pointer; font-size: 13px; }' +
    '.btn:hover { background: #1558b0; }' +
    '.success { color: #1e8e3e; font-weight: bold; }' +
    '.warn { color: #e37400; }' +
    '#copyStatus { margin-left: 8px; color: #1e8e3e; font-size: 12px; }' +
    '</style></head><body>' +

    '<h3>Deployment Summary</h3>' +
    '<p>Target: <strong>' + targetName + '</strong></p>' +
    '<p>' + propCount + ' Script Properties prepared for deployment.' +
    (wodCopied ? ' <span class="success">WOD_Memory data copied successfully.</span>' :
                 ' WOD_Memory was not copied.') + '</p>' +

    '<div class="summary">' + escapedSummary + '</div>' +

    '<h3>Step 1 \u2014 Set Script Properties in Target</h3>' +
    '<p>Copy the snippet below, paste it into the <strong>target</strong> project\'s Code.gs ' +
    '(at the bottom), run <code>_pasteDeployedProperties</code>, then delete the snippet.</p>' +
    '<div class="snippet" id="snippetBlock">' + escapedSnippet + '</div>' +
    '<button class="btn" onclick="copySnippet()">Copy Property Snippet</button>' +
    '<span id="copyStatus"></span>' +

    '<h3>Step 2 \u2014 Copy Code.gs Manually</h3>' +
    '<p class="warn">Google Apps Script does not allow programmatic code copying between projects.</p>' +
    '<p>To deploy the code:</p>' +
    '<ol>' +
    '<li>In <strong>this</strong> project, press <strong>Ctrl+A</strong> to select all code in Code.gs</li>' +
    '<li>Press <strong>Ctrl+C</strong> to copy</li>' +
    '<li>Open the <strong>target</strong> sheet\'s Apps Script editor (Extensions &gt; Apps Script)</li>' +
    '<li>Select all code there (<strong>Ctrl+A</strong>) and paste (<strong>Ctrl+V</strong>)</li>' +
    '<li>Save (<strong>Ctrl+S</strong>)</li>' +
    '</ol>' +

    '<h3>Step 3 \u2014 Update SCHEDULING_SHEET_URL_</h3>' +
    '<p>In the <strong>target</strong> Code.gs, find the <code>SCHEDULING_SHEET_URL_</code> constant near the top (around line 17).</p>' +
    '<ul>' +
    '<li>For the <strong>master/production</strong> sheet: set it to <code>null</code></li>' +
    '<li>For a <strong>staging</strong> copy: set it to the staging sheet URL</li>' +
    '</ul>' +
    '<script>' +
    'function copySnippet() {' +
    '  var text = document.getElementById("snippetBlock").textContent;' +
    '  navigator.clipboard.writeText(text).then(function() {' +
    '    document.getElementById("copyStatus").textContent = "Copied!";' +
    '  }).catch(function() {' +
    '    var ta = document.createElement("textarea");' +
    '    ta.value = text; document.body.appendChild(ta);' +
    '    ta.select(); document.execCommand("copy");' +
    '    document.body.removeChild(ta);' +
    '    document.getElementById("copyStatus").textContent = "Copied!";' +
    '  });' +
    '}' +
    '</script>' +

    '</body></html>';

  var htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(620)
    .setHeight(580)
    .setTitle('Deploy to Another Sheet \u2014 Results');

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Deploy to Another Sheet \u2014 Results');
}




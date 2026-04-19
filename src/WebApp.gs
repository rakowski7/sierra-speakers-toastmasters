/**
 * ============================================================
 * Sierra Speakers Toastmasters — Mobile Web App (PWA)
 * Server-Side Google Apps Script
 * ============================================================
 *
 * This file should be ADDED to the existing Apps Script project
 * alongside Code.gs — it does NOT replace it.
 *
 * It exposes a doGet() web app that serves a mobile-friendly
 * schedule viewer, role confirmation tool, and admin actions.
 * ============================================================
 */

// ── Configuration ───────────────────────────────────────────
const WEBAPP_CONFIG_ = {
  SPREADSHEET_ID: "1gLWiPXAzW_LXw-7_GDG7ykaGOB2faa3WcY6h18qH-rw",
  ADMIN_EMAILS: [
    "rakowski.7@gmail.com",
    // Add other admin emails here
  ],
  DEFAULT_MEETING_LOCATION: "633 Folsom Street, San Francisco CA",
  CLUB_NAME: "Sierra Speakers Toastmasters",
};

// ── Web App Entry Point ─────────────────────────────────────
/**
 * doGet
 * Serves the mobile PWA as a single-page application.
 * Routes: ?page=schedule (default), ?page=meeting&date=M/D/YYYY,
 *         ?page=confirm, ?page=admin
 * Also handles ?action=... for API-style JSON responses.
 * @param {Object} e - Event parameter with queryString, parameter, etc.
 * @return {HtmlOutput|TextOutput} HTML page or JSON response.
 */
function doGet(e) {
  // DEBUG: remove once identity confirmed working
  Logger.log('doGet: activeUser=' + Session.getActiveUser().getEmail() + ' effectiveUser=' + Session.getEffectiveUser().getEmail());
  const action = (e && e.parameter && e.parameter.action) || "";

  // ── JSON API endpoints ──
  if (action) {
    const result = handleApiAction_(action, e.parameter);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── Serve the PWA HTML ──
  const template = HtmlService.createTemplateFromFile("Index");
  // The web app is deployed with `Execute as: User accessing the web app`
  // and the Sheet is shared with members via a Google Group, so
  // Session.getActiveUser().getEmail() resolves to the visitor's own email
  // for any signed-in Google account.
  var currentEmail = getCurrentUserEmail_();
  template.userEmail = currentEmail;
  template.isAdmin = isAdmin_();
  template.config = JSON.stringify({
    clubName: WEBAPP_CONFIG_.CLUB_NAME,
    isAdmin: isAdmin_(),
    userEmail: currentEmail,
  });

  return template.evaluate()
    .setTitle("Sierra Speakers")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag("viewport", "width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no");
}

/**
 * include
 * Helper to include HTML partials (Styles.html, JavaScript.html).
 * Called from Index.html via <?!= include('Styles') ?>
 * @param {string} filename - Name of the HTML file (without .html extension).
 * @return {string} Raw HTML content.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── Auth Helpers ────────────────────────────────────────────

/**
 * normalizeEmail_
 * Canonical email normalizer used for ALL email comparisons and for
 * ingesting emails into the member directory. Lowercases and trims
 * surrounding whitespace. Intentionally does NOT strip dots or "+tags"
 * because those carry user-facing meaning.
 * @param {*} s - Any value (string, null, undefined).
 * @return {string} Normalized email, or "" if input was empty.
 */
function normalizeEmail_(s) {
  return String(s || "").trim().toLowerCase();
}

/**
 * getCurrentUserEmail_
 * Returns the identified user's email (lowercased).
 *
 * Identity source is `ScriptApp.getIdentityToken()`, which returns a
 * signed JWT (header.payload.signature) whose `email` claim is the
 * visitor's verified Google account email. This works for any signed-in
 * Google account (including consumer @gmail.com), unlike
 * `Session.getActiveUser().getEmail()`, which returns "" for consumer
 * Gmail users when the deployment is configured as "Execute as: User
 * accessing the web app" + "Who has access: Anyone with Google account".
 *
 * Requires the `openid` OAuth scope (declared in appsscript.json) and
 * the `userinfo.email` scope for the `email` claim to be populated.
 *
 * We deliberately do NOT fall back to Session.getEffectiveUser(): that
 * returns the deployer's email (Mateusz) for every visitor and would
 * silently grant admin to everyone.
 *
 * @return {string} The visitor's email (lowercase), or "" if unidentified.
 */
function getCurrentUserEmail_() {
  try {
    var token = ScriptApp.getIdentityToken();
    if (!token) {
      // DEBUG: remove once identity confirmed working
      Logger.log('getCurrentUserEmail_: no identity token returned');
      return "";
    }
    var parts = token.split('.');
    if (parts.length < 2) {
      Logger.log('getCurrentUserEmail_: malformed identity token (parts=' + parts.length + ')');
      return "";
    }
    // base64url-decode the JWT payload. Utilities.base64DecodeWebSafe
    // handles the URL-safe alphabet and tolerates missing padding.
    var payloadBytes = Utilities.base64DecodeWebSafe(parts[1]);
    var payloadJson = Utilities.newBlob(payloadBytes).getDataAsString();
    var payload = JSON.parse(payloadJson);
    var email = normalizeEmail_(payload && payload.email);
    // DEBUG: remove once identity confirmed working
    Logger.log('getCurrentUserEmail_: tokenObtained=true email=' + email +
               ' email_verified=' + (payload && payload.email_verified));
    return email;
  } catch (err) {
    Logger.log('getIdentityToken failed: ' + err);
    return "";
  }
}

/**
 * isAdmin_
 * Returns true if the current user's email is in the ADMIN_EMAILS list.
 * @return {boolean}
 */
function isAdmin_() {
  const email = getCurrentUserEmail_();
  if (!email) return false;
  return WEBAPP_CONFIG_.ADMIN_EMAILS.some(a => normalizeEmail_(a) === email);
}

/**
 * isMember_
 * Returns true if the current user's email appears in the member directory
 * at the top of any SCHED sheet.
 * @return {boolean}
 */
function isMember_() {
  const email = getCurrentUserEmail_();
  if (!email) return false;
  const ss = SpreadsheetApp.openById(WEBAPP_CONFIG_.SPREADSHEET_ID);
  const sheet = getLatestSchedSheet_(ss);
  if (!sheet) return false;
  const data = sheet.getDataRange().getValues();
  const backgrounds = sheet.getDataRange().getBackgrounds();
  for (let r = 1; r < data.length; r++) {
    if (backgrounds[r][0] === "#cfe2f3" || (!data[r][0] && !data[r][4])) break;
    const memberEmail = normalizeEmail_(data[r][4]);
    if (memberEmail === email) return true;
  }
  return isAdmin_(); // admins always count as members
}

// ── API Action Router ───────────────────────────────────────

/**
 * handleApiAction_
 * Routes API actions to their handler functions.
 * @param {string} action - The action name.
 * @param {Object} params - URL parameters.
 * @return {Object} JSON-serializable result.
 */
function handleApiAction_(action, params) {
  try {
    switch (action) {
      case "getSchedule":       return api_getSchedule_();
      case "getMeetingDetails": return api_getMeetingDetails_(params.date);
      case "confirmRole":       return api_confirmRole_(params.date, params.role);
      case "getMyRoles":        return api_getMyRoles_();
      case "triggerConfirmations": return isAdmin_() ? api_triggerConfirmations_(params.date) : { error: "Admin only" };
      case "triggerAgenda":     return isAdmin_() ? api_triggerAgenda_(params.date) : { error: "Admin only" };
      case "getServiceWorker":  return { sw: HtmlService.createHtmlOutputFromFile("ServiceWorker").getContent() };
      default: return { error: "Unknown action: " + action };
    }
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Sheet Helpers ───────────────────────────────────────────

/**
 * getLatestSchedSheet_
 * Returns the most recent SCHED YYYY sheet from the spreadsheet.
 * @param {Spreadsheet} ss - The spreadsheet object.
 * @return {Sheet|null}
 */
function getLatestSchedSheet_(ss) {
  const sheets = ss.getSheets();
  const schedSheets = sheets
    .filter(s => /^SCHED\s\d{4}$/.test(s.getName()))
    .sort((a, b) => {
      const yA = parseInt(a.getName().split(" ")[1]);
      const yB = parseInt(b.getName().split(" ")[1]);
      return yB - yA;
    });
  return schedSheets[0] || null;
}

/**
 * parseScheduleData_
 * Reads and parses the SCHED sheet into structured data.
 * Returns members, roles header row info, date columns, and role assignments.
 * @param {Sheet} [sheetOverride] - Optional specific sheet to use.
 * @return {Object} Parsed schedule data.
 */
function parseScheduleData_(sheetOverride) {
  const ss = SpreadsheetApp.openById(WEBAPP_CONFIG_.SPREADSHEET_ID);
  const sheet = sheetOverride || getLatestSchedSheet_(ss);
  if (!sheet) return { error: "No SCHED sheet found." };

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const backgrounds = dataRange.getBackgrounds();

  // ── Member directory ──
  const members = {};
  for (let r = 1; r < data.length; r++) {
    const bgColor = backgrounds[r][0];
    if (bgColor === "#cfe2f3" || (!data[r][0] && !data[r][4])) break;
    const firstName = (data[r][0] || "").toString().trim();
    const lastName  = (data[r][1] || "").toString().trim();
    const email     = normalizeEmail_(data[r][4]);
    if (firstName && lastName && email) {
      members[`${firstName} ${lastName}`] = email;
    }
  }

  // ── Roles header row ──
  let rolesHeaderRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (data[r][0]?.toString().trim().toLowerCase() === "roles") { rolesHeaderRow = r; break; }
  }
  if (rolesHeaderRow === -1) {
    for (let r = 0; r < data.length; r++) {
      if (data[r][0]?.toString().trim().toLowerCase() === "toastmaster") { rolesHeaderRow = r - 1; break; }
    }
  }
  if (rolesHeaderRow < 0) return { error: "Could not find Roles header row." };

  // ── Date columns ──
  let firstDateCol = -1;
  for (let c = 0; c < data[rolesHeaderRow].length; c++) {
    if (data[rolesHeaderRow][c] instanceof Date) { firstDateCol = c; break; }
  }
  if (firstDateCol === -1) return { error: "No dates found in Roles row." };

  const dateColumns = [];
  for (let c = firstDateCol; c < data[rolesHeaderRow].length; c++) {
    const h = data[rolesHeaderRow][c];
    if (h instanceof Date) {
      dateColumns.push({
        date: Utilities.formatDate(h, Session.getScriptTimeZone(), "M/d/yyyy"),
        dateObj: h,
        colIndex: c,
        theme: rolesHeaderRow > 0 ? (data[rolesHeaderRow - 1][c]?.toString().trim() || "") : "",
        // Meeting Location row is typically 2-3 rows above roles
        location: "",
      });
    }
  }

  // Try to find meeting location row
  for (let r = Math.max(0, rolesHeaderRow - 5); r < rolesHeaderRow; r++) {
    const label = (data[r][0] || "").toString().trim().toLowerCase();
    if (label.includes("meeting location") || label.includes("location")) {
      dateColumns.forEach(dc => {
        const loc = (data[r][dc.colIndex] || "").toString().trim();
        if (loc) dc.location = loc;
      });
      break;
    }
  }

  // Try to find meeting format row
  for (let r = Math.max(0, rolesHeaderRow - 5); r < rolesHeaderRow; r++) {
    const label = (data[r][0] || "").toString().trim().toLowerCase();
    if (label.includes("meeting format") || label.includes("format")) {
      dateColumns.forEach(dc => {
        const fmt = (data[r][dc.colIndex] || "").toString().trim();
        if (fmt) dc.format = fmt;
      });
      break;
    }
  }

  return {
    data,
    backgrounds,
    members,
    rolesHeaderRow,
    dateColumns,
    sheetName: sheet.getName(),
  };
}

// ── Role Filtering Helper ───────────────────────────────────

/**
 * isUnassignedRole_
 * Returns true if the assigned name indicates no one is actually assigned.
 * Matches: empty, null, whitespace-only, "TBD" (case-insensitive).
 * @param {string} name - The assigned name value.
 * @return {boolean}
 */
function isUnassignedRole_(name) {
  if (!name) return true;
  var trimmed = name.toString().trim();
  if (!trimmed) return true;
  if (trimmed.toUpperCase() === "TBD") return true;
  return false;
}

/**
 * isLegendRow_
 * Returns true if the role label is a color-coded legend/placeholder row
 * (e.g. "Green =", "Berry =", "Red Text =", "Yellow =").
 * These rows explain the color coding and should not appear as roles.
 * @param {string} roleLabel - The role label from column A.
 * @return {boolean}
 */
function isLegendRow_(roleLabel) {
  if (!roleLabel) return false;
  var lower = roleLabel.toString().trim().toLowerCase();
  // Match patterns like "green =", "berry =", "red text =", "yellow ="
  if (/^(green|berry|red\s*text|yellow)\s*=/.test(lower)) return true;
  return false;
}

// ── API Handlers ────────────────────────────────────────────

/**
 * api_getSchedule_
 * Returns upcoming meetings with all role assignments and statuses.
 * Filters out roles where no one is assigned (TBD, empty, blank)
 * and color-coded legend rows.
 * @return {Object} { meetings: [...] }
 */
function api_getSchedule_() {
  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, rolesHeaderRow, dateColumns } = parsed;
  const now = new Date();
  now.setHours(0, 0, 0, 0);

  const meetings = [];

  dateColumns.forEach(dc => {
    if (dc.dateObj < now) return; // skip past meetings

    const meeting = {
      date: dc.date,
      theme: dc.theme,
      location: dc.location || WEBAPP_CONFIG_.DEFAULT_MEETING_LOCATION,
      format: dc.format || "",
      roles: [],
    };

    let speechCounter = 1, evaluatorCounter = 1;
    for (let r = rolesHeaderRow + 1; r < data.length; r++) {
      let roleRaw = (data[r][0] || "").toString().trim();
      if (!roleRaw) continue;

      // Skip color-coded legend rows (e.g. "Green =", "Berry =")
      if (isLegendRow_(roleRaw)) continue;

      const roleLower = roleRaw.toLowerCase();
      let roleLabel = roleRaw;
      if (roleLower.startsWith("speech"))    roleLabel = "Speech " + speechCounter++;
      else if (roleLower.startsWith("evaluator")) roleLabel = "Evaluator " + evaluatorCounter++;

      const assignedName = (data[r][dc.colIndex] || "").toString().trim();

      // Skip roles where no one is assigned
      if (isUnassignedRole_(assignedName)) continue;

      const bgColor = (backgrounds[r][dc.colIndex] || "").toLowerCase();

      let status = "unconfirmed";
      if (isRoughlyGreen(bgColor))  status = "confirmed";
      else if (isRoughlyRed(bgColor))    status = "unable";
      else if (isRoughlyYellow(bgColor)) status = "emailed";

      meeting.roles.push({
        role: roleLabel,
        name: assignedName,
        status,
      });
    }

    meetings.push(meeting);
  });

  return { meetings };
}

/**
 * api_getMeetingDetails_
 * Returns full details for a specific meeting date.
 * @param {string} dateStr - Date string like "4/17/2026".
 * @return {Object} Meeting details with roles, theme, location, word of the day.
 */
function api_getMeetingDetails_(dateStr) {
  const schedule = api_getSchedule_();
  if (schedule.error) return schedule;
  const meeting = schedule.meetings.find(m => m.date === dateStr);
  if (!meeting) return { error: "No meeting found for " + dateStr };
  return { meeting };
}

/**
 * api_getMyRoles_
 * Returns the current user's upcoming role assignments.
 *
 * Identity comes from Session.getActiveUser().getEmail(), which resolves
 * to the visitor's own email under our "Execute as: User accessing the
 * web app" deployment. If we cannot identify the visitor (no email) or
 * the email is not in the member directory, returns
 * `{ roles: [], unidentified: true, message: ... }` so the client can
 * show a clean "we couldn't identify you" message. We deliberately do
 * NOT return a name picker — that path let visitors impersonate anyone.
 *
 * @return {Object} { roles: [{ date, role, status }], ... }
 */
function api_getMyRoles_() {
  // DEBUG: remove once identity confirmed working
  Logger.log('api_getMyRoles_: email=' + getCurrentUserEmail_());
  const email = getCurrentUserEmail_();

  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;

  // No identifiable email → return a single "unidentified" signal. The
  // client shows a clean message explaining how to fix it. No name
  // picker (that allowed impersonation).
  if (!email) {
    return {
      roles: [],
      unidentified: true,
      message: "We couldn't identify you. Make sure you're signed into " +
               "the Google account on file with the club.",
    };
  }

  // Find all names associated with this email. Both sides are normalized
  // through normalizeEmail_ so capital letters in the sheet entry never
  // silently break the match.
  const myNames = new Set();
  Object.entries(members).forEach(([name, memberEmail]) => {
    if (normalizeEmail_(memberEmail) === email) myNames.add(name);
  });

  if (myNames.size === 0) {
    return {
      roles: [],
      unidentified: true,
      email: email,
      message: "We couldn't identify you. Make sure you're signed into " +
               "the Google account on file with the club. " +
               "(Signed in as " + email + ", which isn't in the Sierra Speakers directory.)",
    };
  }

  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const myRoles = [];

  dateColumns.forEach(dc => {
    if (dc.dateObj < now) return;

    let speechCounter = 1, evaluatorCounter = 1;
    for (let r = rolesHeaderRow + 1; r < data.length; r++) {
      let roleRaw = (data[r][0] || "").toString().trim();
      if (!roleRaw) continue;

      // Skip color-coded legend rows
      if (isLegendRow_(roleRaw)) continue;

      const roleLower = roleRaw.toLowerCase();
      let roleLabel = roleRaw;
      if (roleLower.startsWith("speech"))    roleLabel = "Speech " + speechCounter++;
      else if (roleLower.startsWith("evaluator")) roleLabel = "Evaluator " + evaluatorCounter++;

      const assignedName = (data[r][dc.colIndex] || "").toString().trim();
      if (!myNames.has(assignedName)) continue;

      const bgColor = (backgrounds[r][dc.colIndex] || "").toLowerCase();
      let status = "unconfirmed";
      if (isRoughlyGreen(bgColor))  status = "confirmed";
      else if (isRoughlyRed(bgColor))    status = "unable";
      else if (isRoughlyYellow(bgColor)) status = "emailed";

      myRoles.push({
        date: dc.date,
        theme: dc.theme,
        role: roleLabel,
        status,
        rowIndex: r,
        colIndex: dc.colIndex,
      });
    }
  });

  return { roles: myRoles };
}

/**
 * api_confirmRole_
 * Marks the current user's role cell green (confirmed) in the sheet.
 * @param {string} dateStr - Meeting date string.
 * @param {string} roleLabel - Role label like "Speech 1", "Table Topics Master".
 * @return {Object} { success: true } or { error: string }
 */
function api_confirmRole_(dateStr, roleLabel) {
  const myRoles = api_getMyRoles_();
  if (myRoles.error) return myRoles;
  if (myRoles.unidentified) return { error: myRoles.message };

  const match = myRoles.roles.find(r => r.date === dateStr && r.role === roleLabel);
  if (!match) return { error: "Role not found for you on this date." };

  const ss = SpreadsheetApp.openById(WEBAPP_CONFIG_.SPREADSHEET_ID);
  const sheet = getLatestSchedSheet_(ss);
  if (!sheet) return { error: "Sheet not found." };

  // Set the cell background to green
  sheet.getRange(match.rowIndex + 1, match.colIndex + 1).setBackground("#93c47d");
  return { success: true, role: roleLabel, date: dateStr };
}

/**
 * api_triggerConfirmations_
 * (Admin only) Kicks off the role confirmation email flow for a given date.
 * This calls the existing startRoleConfirmations function concept but adapted
 * to work programmatically without UI prompts.
 * @param {string} dateStr - Meeting date.
 * @return {Object} Status message.
 */
function api_triggerConfirmations_(dateStr) {
  if (!isAdmin_()) return { error: "Admin access required." };

  // For the mobile app, we prepare the data and return a summary.
  // The actual email sending still needs to happen from the sheet UI
  // because it requires interactive review of each email draft.
  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(d => d.date === dateStr);
  if (!dc) return { error: "Date not found: " + dateStr };

  let speechCounter = 1, evaluatorCounter = 1;
  const summary = { date: dateStr, theme: dc.theme, roles: [] };

  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    let roleRaw = (data[r][0] || "").toString().trim();
    if (!roleRaw) continue;
    const roleLower = roleRaw.toLowerCase();
    let roleLabel = roleRaw;
    if (roleLower.startsWith("speech"))    roleLabel = "Speech " + speechCounter++;
    else if (roleLower.startsWith("evaluator")) roleLabel = "Evaluator " + evaluatorCounter++;

    const name = (data[r][dc.colIndex] || "").toString().trim();
    const bgColor = (backgrounds[r][dc.colIndex] || "").toLowerCase();
    let status = "unconfirmed";
    if (isRoughlyGreen(bgColor))  status = "confirmed";
    else if (isRoughlyRed(bgColor))    status = "unable";
    else if (isRoughlyYellow(bgColor)) status = "emailed";

    const email = members[name] || "";
    summary.roles.push({ role: roleLabel, name: name || "TBD", status, email });
  }

  return {
    message: "Confirmation summary prepared. Open the Google Sheet to send emails interactively.",
    summary,
    sheetUrl: `https://docs.google.com/spreadsheets/d/${WEBAPP_CONFIG_.SPREADSHEET_ID}/edit`,
  };
}

/**
 * api_triggerAgenda_
 * (Admin only) Returns info needed to generate the agenda.
 * Actual generation still runs from the sheet menu because it requires
 * interactive dialogs (WOTD, speech selection, etc.)
 * @param {string} dateStr - Meeting date.
 * @return {Object} Status message with sheet URL.
 */
function api_triggerAgenda_(dateStr) {
  if (!isAdmin_()) return { error: "Admin access required." };
  return {
    message: "To generate the agenda, open the Google Sheet and use Toastmasters > Generate Meeting Agenda.",
    sheetUrl: `https://docs.google.com/spreadsheets/d/${WEBAPP_CONFIG_.SPREADSHEET_ID}/edit`,
    date: dateStr,
  };
}

// ── Admin Backend: Role Confirmations ───────────────────────

/**
 * getConfirmationData_
 * Private helper that builds the confirmation data structure for a given date.
 * Returns roles array with all details needed for the confirmation UI.
 * @param {string} dateStr - Meeting date.
 * @return {Object} { roles: [...], theme, toastmasterName, members, wod, meetingFormat, hasGrammarian, hasSpeeches }
 */
function getConfirmationData_(dateStr) {
  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(d => d.date === dateStr);
  if (!dc) return { error: "Date not found: " + dateStr };

  const ss = SpreadsheetApp.openById(WEBAPP_CONFIG_.SPREADSHEET_ID);
  const sheet = getLatestSchedSheet_(ss);

  const wod = lookupWodCache_(dateStr);
  const meetingFormat = resolveMeetingFormatForColumn_(data, dc.colIndex);

  const roles = [];
  let speechCounter = 1, evaluatorCounter = 1;
  let toastmasterName = "";
  let hasGrammarian = false, hasSpeeches = false;

  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    let roleRaw = (data[r][0] || "").toString().trim();
    if (!roleRaw) continue;

    const roleLower = roleRaw.toLowerCase();
    let roleLabel = roleRaw;
    let roleType = "other";

    if (roleLower.startsWith("speech")) {
      roleLabel = "Speech " + speechCounter++;
      roleType = "speech";
      hasSpeeches = true;
    } else if (roleLower.startsWith("evaluator")) {
      roleLabel = "Evaluator " + evaluatorCounter++;
      roleType = "evaluator";
    } else if (roleLower.startsWith("toastmaster")) {
      roleType = "toastmaster";
    } else if (roleLower.startsWith("grammarian")) {
      roleType = "grammarian";
      hasGrammarian = true;
    }

    const assignedName = (data[r][dc.colIndex] || "").toString().trim();
    const bgColor = (backgrounds[r][dc.colIndex] || "").toLowerCase();
    const email = members[assignedName] || "";
    let status = "unconfirmed";
    if (isRoughlyGreen(bgColor)) status = "confirmed";
    else if (isRoughlyRed(bgColor)) status = "unable";
    else if (isRoughlyYellow(bgColor)) status = "emailed";

    if (roleType === "toastmaster") toastmasterName = assignedName;

    roles.push({
      role: roleLabel,
      roleType: roleType,
      name: assignedName || "TBD",
      email: email,
      status: status,
      fuzzyMatched: false,
      note: "",
      rowIndex: r,
      colIndex: dc.colIndex,
      currentBg: bgColor,
    });
  }

  const membersList = Object.keys(members);

  return {
    roles: roles,
    theme: dc.theme,
    toastmasterName: toastmasterName,
    members: membersList,
    cachedWod: wod ? wod.word : "",
    meetingFormat: meetingFormat || "undecided",
    hasGrammarian: hasGrammarian,
    hasSpeeches: hasSpeeches,
  };
}

/**
 * sendRoleConfirmations_
 * Private helper that creates the confirmation emails for a given date.
 * @param {Object} params - { dateStr, theme, wotd, meetingFormat, meetingAddress, senderType, senderName, introQuestion }
 * @return {Object} { success, draftsCreated, draftsUrl }
 */
function sendRoleConfirmations_(params) {
  const { dateStr, theme, wotd, meetingFormat, meetingAddress, senderType, senderName, introQuestion } = params;

  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(d => d.date === dateStr);
  if (!dc) return { error: "Date not found: " + dateStr };

  const ss = SpreadsheetApp.openById(WEBAPP_CONFIG_.SPREADSHEET_ID);
  const sheet = getLatestSchedSheet_(ss);

  const wod_cached = lookupWodCache_(dateStr);
  const wordOfTheDay = wotd || (wod_cached ? wod_cached.word : "");
  const address = meetingAddress || WEBAPP_CONFIG_.DEFAULT_MEETING_LOCATION;

  // Build location/attendance lines using same logic as Code.gs proceedToEmails
  let locationLine, attendanceNote;
  if (meetingFormat === "hybrid") {
    locationLine   = "";
    attendanceNote = "We'll have a hybrid meeting with our in-person location at " + address +
      ". Please update the attendance sheet if you haven't already. This will let us know whether you'll be joining in person or virtually.";
  } else if (meetingFormat === "virtual") {
    locationLine   = "We'll be meeting virtually on Zoom.";
    attendanceNote = "";
  } else if (meetingFormat === "in_person") {
    locationLine   = "We'll be meeting in person at " + address + ".";
    attendanceNote = "";
  } else {
    locationLine   = "Heads up: the meeting format (in person, virtual, or hybrid) is still being decided. We will confirm before the meeting.";
    attendanceNote = "";
  }

  // Determine sender name
  const toastmasterEntry = getConfirmationData_(dateStr);
  const tmName = toastmasterEntry.toastmasterName || "the Toastmaster";
  const finalSenderName = (senderType === "tm") ? tmName : (senderName || tmName);

  // Store intro question for generateAgenda to retrieve later
  if (introQuestion) {
    PropertiesService.getScriptProperties().setProperty("LAST_INTRO_QUESTION", introQuestion);
  }

  // Build confirmations list (matching Code.gs entry format)
  const confirmations = [];
  let speechCounter = 1, evaluatorCounter = 1;

  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    let roleRaw = (data[r][0] || "").toString().trim();
    if (!roleRaw) continue;

    const roleLower = roleRaw.toLowerCase();
    let roleLabel = roleRaw;
    let roleType = "general";
    if (roleLower.startsWith("speech"))    { roleLabel = "Speech " + speechCounter++;    roleType = "speech"; }
    else if (roleLower.startsWith("evaluator")) { roleLabel = "Evaluator " + evaluatorCounter++; roleType = "evaluator"; }
    else if (roleLower === "toastmaster")   roleType = "toastmaster";
    else if (roleLower.includes("grammarian")) roleType = "grammarian";
    else if (roleLower.includes("table topics master")) roleType = "tabletopics";

    const assignedName = (data[r][dc.colIndex] || "").toString().trim();
    const bgColor = (backgrounds[r][dc.colIndex] || "").toLowerCase();

    if (!assignedName || assignedName.toUpperCase() === "TBD") continue;

    const email = members[assignedName] || "";
    if (!email) continue;

    confirmations.push({
      role: roleLabel,
      roleType: roleType,
      name: assignedName,
      email: email,
      rowIndex: r,
      colIndex: dc.colIndex,
      currentBg: bgColor,
    });
  }

  // Group by email (matching Code.gs proceedToEmails logic)
  const groupMap = {};
  const eligibleEntries = confirmations.filter(function(e) {
    return e.roleType !== "toastmaster" && e.email;
  });
  eligibleEntries.forEach(function(entry) {
    if (!groupMap[entry.email]) groupMap[entry.email] = [];
    groupMap[entry.email].push(entry);
  });

  const emailGroups = Object.values(groupMap).map(function(entries) {
    const roles = entries.map(function(e) { return e.role; });
    var subject;
    if (roles.length === 1) {
      subject = "Sierra Speakers Toastmasters: " + roles[0] + " Confirmation for " + dateStr;
    } else if (roles.length === 2) {
      subject = "Sierra Speakers Toastmasters: " + roles[0] + " & " + roles[1] + " Confirmation for " + dateStr;
    } else {
      subject = "Sierra Speakers Toastmasters: Multiple Roles Confirmation for " + dateStr;
    }

    var body;
    if (entries.length === 1) {
      body = buildEmailBody(entries[0], theme, wordOfTheDay, {}, confirmations,
        finalSenderName, dateStr, introQuestion || "", locationLine, attendanceNote);
    } else {
      body = buildCombinedEmailBody_(entries, theme, wordOfTheDay, confirmations,
        finalSenderName, dateStr, introQuestion || "", locationLine, attendanceNote);
    }

    return {
      email: entries[0].email,
      name: entries[0].name,
      entries: entries,
      subject: subject,
      body: body,
    };
  });

  // Create drafts and color cells
  const draftsUrl = "https://mail.google.com/mail/u/0/#drafts";
  let draftsCreated = 0;

  emailGroups.forEach(function(group) {
    const allRoles = group.entries.map(function(e) { return e.role; });
    const htmlBody = buildFancyHtml_(group.body, allRoles);
    createGmailDraft_(group.email, group.subject, group.body, htmlBody);
    draftsCreated++;

    // Color non-colored cells yellow
    group.entries.forEach(function(entry) {
      if (!isAlreadyColored(entry.currentBg)) {
        sheet.getRange(entry.rowIndex + 1, entry.colIndex + 1).setBackground("#ffff00");
      }
    });
  });

  return {
    success: true,
    draftsCreated: draftsCreated,
    draftsUrl: draftsUrl,
  };
}

/**
 * getConfirmationData
 * PUBLIC wrapper. Returns all data needed for the confirmation wizard.
 * @param {string} dateStr - Meeting date.
 * @return {Object} Confirmation data or error.
 */
function getConfirmationData(dateStr) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return getConfirmationData_(dateStr);
  } catch (err) {
    return { error: err.toString() };
  }
}

/**
 * sendRoleConfirmations
 * PUBLIC wrapper. Creates and sends role confirmation emails.
 * @param {Object} params - { dateStr, theme, wotd, meetingFormat, meetingAddress, senderType, senderName, introQuestion }
 * @return {Object} { success, draftsCreated, draftsUrl } or { error }
 */
function sendRoleConfirmations(params) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return sendRoleConfirmations_(params);
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Admin Backend: Agenda Generation ────────────────────────

/**
 * getAgendaData_
 * Private helper that builds the agenda data structure for a given date.
 * @param {string} dateStr - Meeting date.
 * @return {Object} { speeches, evaluators, allRoles, theme, wod, meetingFormat }
 */
function getAgendaData_(dateStr) {
  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(d => d.date === dateStr);
  if (!dc) return { error: "Date not found: " + dateStr };

  const wod = lookupWodCache_(dateStr);
  const meetingFormat = resolveMeetingFormatForColumn_(data, dc.colIndex);

  const speeches = [];
  const evaluators = [];
  const allRoles = {};

  let speechCounter = 1, evaluatorCounter = 1;

  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    let roleRaw = (data[r][0] || "").toString().trim();
    if (!roleRaw) continue;

    const roleLower = roleRaw.toLowerCase();
    let roleLabel = roleRaw;
    let key = "";

    if (roleLower.startsWith("speech")) {
      roleLabel = "Speech " + speechCounter;
      key = "speech_" + speechCounter;
      speechCounter++;
      const assignedName = (data[r][dc.colIndex] || "").toString().trim();
      const bgColor = (backgrounds[r][dc.colIndex] || "").toLowerCase();
      speeches.push({
        key: key,
        name: assignedName || "TBD",
        isRed: isRoughlyRed(bgColor),
        isTbd: !assignedName,
      });
    } else if (roleLower.startsWith("evaluator")) {
      roleLabel = "Evaluator " + evaluatorCounter;
      key = "evaluator_" + evaluatorCounter;
      evaluatorCounter++;
      const assignedName = (data[r][dc.colIndex] || "").toString().trim();
      evaluators.push({
        key: key,
        name: assignedName || "TBD",
      });
    }

    const assignedName = (data[r][dc.colIndex] || "").toString().trim();
    allRoles[key] = assignedName || "TBD";
  }

  return {
    speeches: speeches,
    evaluators: evaluators,
    allRoles: allRoles,
    theme: dc.theme,
    cachedWod: wod ? wod.word : "",
    meetingFormat: meetingFormat,
  };
}

/**
 * lookupWodDefinitions_
 * Private helper that looks up WOD definitions via Merriam-Webster and Gemini APIs.
 * @param {string} word - The word to look up.
 * @param {string} theme - The meeting theme.
 * @return {Object} { definitions: [{pos, def, ex, pronunciation, source}] }
 */
function lookupWodDefinitions_(word, theme) {
  const definitions = [];

  // Try Merriam-Webster API
  const MW_API_KEY = PropertiesService.getScriptProperties().getProperty("MW_API_KEY");
  if (MW_API_KEY) {
    try {
      const mwUrl = "https://www.dictionaryapi.com/api/v3/references/collegiate/json/" +
                    encodeURIComponent(word) + "?key=" + MW_API_KEY;
      const resp = UrlFetchApp.fetch(mwUrl, { muteHttpExceptions: true });
      const parsed = JSON.parse(resp.getContentText());
      if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === "object" && parsed[0].meta) {
        const entry = parsed[0];
        let mwPronunciation = "";
        try { mwPronunciation = entry.hwi.prs[0].mw || ""; } catch(pe) {}
        const pos = entry.fl || "";
        (entry.shortdef || []).forEach(function(d) {
          if (d && definitions.length < 5) {
            definitions.push({ pos: pos, def: d, ex: "", pronunciation: mwPronunciation, source: "mw" });
          }
        });
        // Try to extract example sentences
        try {
          var exIdx = 0;
          for (var di = 0; di < entry.def.length && exIdx < definitions.length; di++) {
            for (var si = 0; si < entry.def[di].sseq.length && exIdx < definitions.length; si++) {
              for (var ei = 0; ei < entry.def[di].sseq[si].length && exIdx < definitions.length; ei++) {
                var sense = entry.def[di].sseq[si][ei];
                if (sense[0] === "sense" && sense[1] && sense[1].dt) {
                  for (var dti = 0; dti < sense[1].dt.length; dti++) {
                    if (sense[1].dt[dti][0] === "vis" && sense[1].dt[dti][1] && sense[1].dt[dti][1].length > 0) {
                      var ex = (sense[1].dt[dti][1][0].t || "").replace(/\{[^}]+\}/g, "").replace(/\s+/g, " ").trim();
                      if (ex) { definitions[exIdx].ex = ex; exIdx++; }
                    }
                  }
                }
              }
            }
          }
        } catch(exErr) {}
      }
    } catch (e) {
      // Fall through to Gemini
    }
  }

  // Try Gemini API
  const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (GEMINI_API_KEY) {
    const geminiPrompt =
      "You are helping prepare a Toastmasters meeting agenda. " +
      "The Word of the Day is \"" + word + "\". " +
      (theme ? "The meeting theme is \"" + theme + "\". " : "") +
      "Please respond with ONLY a valid JSON object (no markdown, no explanation) in this exact format: " +
      "{\"definition\": \"a concise, clear definition\", " +
      "\"partOfSpeech\": \"the part of speech, e.g. noun, verb, adjective\", " +
      "\"pronunciation\": \"phonetic pronunciation using simple syllable notation like \\\"kon-SISE\\\"\", " +
      "\"example\": \"a vivid example sentence that naturally relates to the meeting theme\"}";

    for (var attempt = 1; attempt <= 3; attempt++) {
      try {
        var aiModel = getAiModel_();
        var geminiResp = UrlFetchApp.fetch(
          "https://generativelanguage.googleapis.com/v1beta/models/" + aiModel.model + ":generateContent?key=" + GEMINI_API_KEY,
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
          var geminiJson = JSON.parse(geminiResp.getContentText());
          var rawText = (geminiJson.candidates && geminiJson.candidates[0] &&
                        geminiJson.candidates[0].content && geminiJson.candidates[0].content.parts &&
                        geminiJson.candidates[0].content.parts[0])
                        ? geminiJson.candidates[0].content.parts[0].text : "";
          var cleaned = rawText.replace(/```json|```/gi, "").trim();
          try {
            var gParsed = JSON.parse(cleaned);
            if (gParsed.definition) {
              recordAiPing_(aiModel.label);
              definitions.push({
                pos: gParsed.partOfSpeech || "",
                def: gParsed.definition,
                ex: gParsed.example || "",
                pronunciation: gParsed.pronunciation || "",
                source: aiModel.label
              });
              break;
            }
          } catch (parseErr) { /* retry */ }
        }
        if (attempt < 3) Utilities.sleep(2000);
      } catch (geminiErr) {
        if (attempt < 3) Utilities.sleep(2000);
      }
    }
  }

  return { definitions: definitions };
}

/**
 * getAgendaData
 * PUBLIC wrapper. Returns agenda data for a given date.
 * @param {string} dateStr - Meeting date.
 * @return {Object} Agenda data or error.
 */
function getAgendaData(dateStr) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return getAgendaData_(dateStr);
  } catch (err) {
    return { error: err.toString() };
  }
}

/**
 * lookupWodDefinitions
 * PUBLIC wrapper. Looks up word definitions.
 * @param {string} word - The word.
 * @param {string} theme - The meeting theme.
 * @return {Object} { definitions: [...] } or { error }
 */
function lookupWodDefinitions(word, theme) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return lookupWodDefinitions_(word, theme);
  } catch (err) {
    return { error: err.toString() };
  }
}

/**
 * generateAgendaFromApp_
 * Private helper that generates and saves the agenda document.
 * @param {Object} params - { dateStr, wordOfTheDay, wotdPronunciation, wotdPartOfSpeech, wotdDefinition, wotdExample, wotdSource, keptSpeechKeys, keptEvaluatorKeys, agendaMode }
 * @return {Object} { success, agendaUrl } or { error }
 */
function generateAgendaFromApp_(params) {
  const { dateStr, wordOfTheDay, wotdPronunciation, wotdPartOfSpeech,
          wotdDefinition, wotdExample, wotdSource, keptSpeechKeys,
          keptEvaluatorKeys, agendaMode, scanInbox } = params;

  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(function(d) { return d.date === dateStr; });
  if (!dc) return { error: "Date not found: " + dateStr };

  const colIndex = dc.colIndex;
  const meetingTheme = dc.theme || "";

  // Build meeting date object for formatting
  const meetingDateObj = dc.dateObj;
  const formattedLongDate = Utilities.formatDate(meetingDateObj, Session.getScriptTimeZone(), "MMMM d, yyyy");

  // Build roles map and speech/evaluator key arrays (matching Code.gs format)
  const roles = {};
  let speechCounter = 1, evaluatorCounter = 1;
  const evaluatorKeys = [];
  const speechFlags = {};

  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    const roleRaw = (data[r][0] || "").toString().trim();
    const assignedRaw = (data[r][colIndex] || "").toString().trim();
    if (!roleRaw) continue;
    const roleLower = roleRaw.toLowerCase();
    let key = roleRaw;
    if (roleLower.startsWith("speech")) {
      key = "Speech " + speechCounter++;
      const bgColor = (backgrounds[r] && backgrounds[r][colIndex]) ? backgrounds[r][colIndex].toLowerCase() : "";
      const isRed = isRoughlyRed(bgColor);
      const isTbd = !assignedRaw || assignedRaw.toUpperCase() === "TBD";
      speechFlags[key] = { name: assignedRaw, isRed: isRed, isTbd: isTbd, rowIndex: r };
    }
    else if (roleLower.startsWith("evaluator")) {
      key = "Evaluator " + evaluatorCounter++;
      evaluatorKeys.push(key);
    }
    roles[key] = (!assignedRaw || assignedRaw.toUpperCase() === "TBD") ? "" : assignedRaw;
  }

  // Map webapp keys (speech_1) to Code.gs keys (Speech 1)
  const mappedSpeechKeys = (keptSpeechKeys || []).map(function(k) {
    return "Speech " + k.replace("speech_", "");
  });
  const mappedEvalKeys = (keptEvaluatorKeys || []).map(function(k) {
    return "Evaluator " + k.replace("evaluator_", "");
  });

  // Build nameToEmail for inbox scan
  const nameToEmail = {};
  Object.entries(members).forEach(function(pair) {
    nameToEmail[pair[0]] = pair[1];
  });

  // Inbox scan for speech details (if scanInbox is true, or legacy agendaMode === "scan")
  let speechDetails = {};
  const storedIntroQ = PropertiesService.getScriptProperties().getProperty("LAST_INTRO_QUESTION") || "";

  if (scanInbox === true || agendaMode === "scan") {
    const speakerEntries = mappedSpeechKeys.map(function(key) {
      const name = roles[key] || "";
      const email = nameToEmail[name] || "";
      return { speechKey: key, name: name, email: email };
    }).filter(function(e) { return e.email; });

    if (speakerEntries.length > 0) {
      try {
        const scanResult = scanSpeakerEmails_(speakerEntries, storedIntroQ);
        if (scanResult && scanResult.details) speechDetails = scanResult.details;
      } catch(scanErr) {
        // Continue without scan results
      }
    }
  }

  // Calculate times (matching Code.gs logic)
  const totalAvailableSpeeches = evaluatorKeys.length;
  const speechStartMins = 6 * 60 + 20;
  const activeSpeechCount = mappedSpeechKeys.length;
  const evalMins = speechStartMins + totalAvailableSpeeches * 10 + 16;
  const ttMins   = speechStartMins + activeSpeechCount * 10;

  // Build the docx via Code.gs builder
  const docxBase64 = buildAgendaDocx({
    date: formattedLongDate,
    theme: meetingTheme,
    roles: roles,
    speechKeys: mappedSpeechKeys,
    evaluatorKeys: mappedEvalKeys,
    wordOfTheDay: wordOfTheDay || "",
    wotdPronunciation: wotdPronunciation || "",
    wotdPartOfSpeech: wotdPartOfSpeech || "",
    wotdDefinition: wotdDefinition || "",
    wotdExample: wotdExample || "",
    speechStartMins: speechStartMins,
    ttMins: ttMins,
    evalMins: evalMins,
    fmtTime: fmtTime_,
    speechDetails: speechDetails,
    impromptuKeys: [],
  });

  // Save to Google Drive
  const filename = "Toastmasters Agenda " + dateStr;
  const blob = Utilities.newBlob(
    Utilities.base64Decode(docxBase64),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    filename + ".docx"
  );
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
  const agendaUrl = file.getUrl();

  // Persist for hype email auto-population
  PropertiesService.getScriptProperties().setProperties({
    LAST_AGENDA_URL:    agendaUrl,
    LAST_AGENDA_DATE:   dateStr,
    LAST_AGENDA_WOTD:   wordOfTheDay   || "",
    LAST_AGENDA_THEME:  meetingTheme   || "",
    LAST_AGENDA_WOTD_DEF: wotdDefinition || "",
    LAST_AGENDA_WOTD_EX:  wotdExample    || "",
  });

  // Save WOD to cache
  if (wordOfTheDay && wotdDefinition) {
    saveWodToCache_(dateStr, wordOfTheDay, wotdDefinition, wotdPronunciation || "",
      wotdPartOfSpeech || "", wotdExample || "", wotdSource || "",
      meetingTheme || "", wotdSource || "");
  }

  return {
    success: true,
    agendaUrl: agendaUrl,
  };
}

/**
 * generateAgendaFromApp
 * PUBLIC wrapper. Generates the agenda document.
 * @param {Object} params - Agenda generation parameters.
 * @return {Object} { success, agendaUrl } or { error }
 */
function generateAgendaFromApp(params) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return generateAgendaFromApp_(params);
  } catch (err) {
    return { error: err.toString() };
  }
}

/**
 * scanSpeakerInbox_
 * Private helper that scans Gmail for speaker emails for a given date.
 * Builds speaker entries from the schedule and calls scanSpeakerEmails_
 * from Code.gs.
 * @param {string} dateStr - Meeting date.
 * @param {Array} speechKeys - Array of speech keys like ["speech_1", "speech_2"].
 * @return {Object} { details: { [speechKey]: {title, pathway, ...} }, foundKeys: [] }
 */
function scanSpeakerInbox_(dateStr, speechKeys) {
  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, backgrounds, members, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(function(d) { return d.date === dateStr; });
  if (!dc) return { error: "Date not found: " + dateStr };

  // Build nameToEmail
  const nameToEmail = {};
  Object.entries(members).forEach(function(pair) {
    nameToEmail[pair[0]] = pair[1];
  });

  // Build roles map
  const roles = {};
  let speechCounter = 1;
  for (var r = rolesHeaderRow + 1; r < data.length; r++) {
    var roleRaw = (data[r][0] || "").toString().trim();
    var assignedRaw = (data[r][dc.colIndex] || "").toString().trim();
    if (!roleRaw) continue;
    var roleLower = roleRaw.toLowerCase();
    var key = roleRaw;
    if (roleLower.startsWith("speech")) {
      key = "Speech " + speechCounter++;
    }
    roles[key] = (!assignedRaw || assignedRaw.toUpperCase() === "TBD") ? "" : assignedRaw;
  }

  // Map webapp keys (speech_1) to Code.gs keys (Speech 1)
  var mappedKeys = (speechKeys || []).map(function(k) {
    return "Speech " + k.replace("speech_", "");
  });

  var speakerEntries = mappedKeys.map(function(key) {
    var name = roles[key] || "";
    var email = nameToEmail[name] || "";
    return { speechKey: key, name: name, email: email };
  }).filter(function(e) { return e.email; });

  if (speakerEntries.length === 0) {
    return { details: {}, foundKeys: [] };
  }

  var storedIntroQ = PropertiesService.getScriptProperties().getProperty("LAST_INTRO_QUESTION") || "";

  try {
    var scanResult = scanSpeakerEmails_(speakerEntries, storedIntroQ);
    var foundKeys = [];
    if (scanResult.emailFoundKeys) {
      scanResult.emailFoundKeys.forEach(function(k) { foundKeys.push(k); });
    }
    return {
      details: scanResult.details || {},
      foundKeys: foundKeys,
    };
  } catch (scanErr) {
    return { error: "Inbox scan failed: " + scanErr.toString() };
  }
}

/**
 * scanSpeakerInbox
 * PUBLIC wrapper. Scans Gmail for speech details from speakers.
 * @param {string} dateStr - Meeting date.
 * @param {Array} speechKeys - Array of speech keys like ["speech_1", "speech_2"].
 * @return {Object} { details, foundKeys } or { error }
 */
function scanSpeakerInbox(dateStr, speechKeys) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return scanSpeakerInbox_(dateStr, speechKeys);
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Admin Backend: Club Email ───────────────────────────────

/**
 * getClubEmailData_
 * Private helper that builds pre-populated data for the club email draft.
 * @param {string} dateStr - Meeting date.
 * @return {Object} { theme, speakers, meetingFormat, wod, agendaUrl, wotdDef, allUpcomingDates }
 */
function getClubEmailData_(dateStr) {
  const parsed = parseScheduleData_();
  if (parsed.error) return parsed;

  const { data, rolesHeaderRow, dateColumns } = parsed;
  const dc = dateColumns.find(d => d.date === dateStr);
  if (!dc) return { error: "Date not found: " + dateStr };

  const wod = lookupWodCache_(dateStr);
  const meetingFormat = resolveMeetingFormatForColumn_(data, dc.colIndex);

  // Build speakers for selected date
  const speakers = [];
  for (let r = rolesHeaderRow + 1; r < data.length; r++) {
    const roleRaw = (data[r][0] || "").toString().trim();
    if (!roleRaw) continue;
    if (/^Speech\s*\d+$/i.test(roleRaw)) {
      const name = (data[r][dc.colIndex] || "").toString().trim();
      if (name && name.toUpperCase() !== "TBD") {
        speakers.push({ role: roleRaw, name: name });
      }
    }
  }

  // Build per-date data for all upcoming dates (so date change updates fields)
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  const perDateData = {};
  const allDates = [];
  dateColumns.forEach(function(ud) {
    if (ud.dateObj < now) return;
    allDates.push(ud.date);
    const sp = [];
    for (let r = rolesHeaderRow + 1; r < data.length; r++) {
      const role = (data[r][0] || "").toString().trim();
      const name = (data[r][ud.colIndex] || "").toString().trim();
      if (/^Speech\s*\d+$/i.test(role) && name && name.toUpperCase() !== "TBD") {
        sp.push({ role: role, name: name });
      }
    }
    const fmt = resolveMeetingFormatForColumn_(data, ud.colIndex) || "hybrid";
    perDateData[ud.date] = { theme: ud.theme || "", speakers: sp, meetingFormat: fmt };
  });

  const props = PropertiesService.getScriptProperties();
  const agendaUrl = props.getProperty("LAST_AGENDA_URL") || "";

  return {
    selectedDate: dateStr,
    theme: dc.theme || "",
    speakers: speakers,
    meetingFormat: meetingFormat || "hybrid",
    wotd: wod ? wod.word : "",
    wotdDef: wod ? wod.definition : "",
    wotdEx: wod ? wod.example : "",
    agendaUrl: agendaUrl,
    allDates: allDates,
    perDateData: perDateData,
  };
}

/**
 * draftClubEmail_
 * Private helper that creates the club hype email draft.
 * @param {Object} params - { dateStr, theme, wotd, wotdDef, wotdEx, agendaUrl, guestEmails, guestMode, meetingFormat, speakers, longDate }
 * @return {Object} { success, draftsUrl }
 */
function draftClubEmail_(params) {
  const { dateStr, theme, wotd, wotdDef, wotdEx, agendaUrl, guestEmails, guestMode, meetingFormat, speakers, longDate } = params;

  const opts = {
    dateStr: dateStr,
    theme: theme,
    wotd: wotd,
    wotdDef: wotdDef,
    wotdEx: wotdEx,
    agendaUrl: agendaUrl,
    guestEmails: guestEmails || [],
    guestMode: guestMode || false,
    meetingFormat: meetingFormat,
    speakers: speakers,
    longDate: longDate,
  };

  createClubHypeEmailDraft_(opts);

  return {
    success: true,
    draftsUrl: "https://mail.google.com/mail/u/0/#drafts",
  };
}

/**
 * getClubEmailData
 * PUBLIC wrapper. Returns pre-populated club email data.
 * @param {string} dateStr - Meeting date.
 * @return {Object} Club email data or error.
 */
function getClubEmailData(dateStr) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return getClubEmailData_(dateStr);
  } catch (err) {
    return { error: err.toString() };
  }
}

/**
 * draftClubEmail
 * PUBLIC wrapper. Creates the club email draft.
 * @param {Object} params - Club email parameters.
 * @return {Object} { success, draftsUrl } or { error }
 */
function draftClubEmail(params) {
  try {
    if (!isAdmin_()) return { error: "Admin access required." };
    return draftClubEmail_(params);
  } catch (err) {
    return { error: err.toString() };
  }
}

// ── Public Wrappers for google.script.run ──────────────────
// Functions ending with _ are PRIVATE in Google Apps Script and
// cannot be called from the client via google.script.run.
// These public wrappers expose the API to the client-side JS.
//
// Identity is resolved server-side via Session.getActiveUser() under
// our "Execute as: User accessing the web app" deployment. There is
// intentionally no name-based fallback / member-list picker — that
// would let any visitor impersonate anyone in the directory.

function getScheduleData()                        { return api_getSchedule_(); }
function getMyRoles()                             { return api_getMyRoles_(); }
function confirmRole(dateStr, roleLabel)          { return api_confirmRole_(dateStr, roleLabel); }
function triggerConfirmations(dateStr)            { return api_triggerConfirmations_(dateStr); }

/**
 * Training expiry notifier
 * Reads the sheet named Training, keeps only the newest record per person and course,
 * then emails a weekly summary of expired, thirty day, and sixty day items.
 */

function getConfig_() {
  const props = PropertiesService.getScriptProperties();
  return {
    SHEET_ID: props.getProperty("SHEET_ID") || "1jIAtvQ87J9Fs9SGB4XYA7xzRVzTjpzWi8dP3k1WCNPc",
    SHEET_NAME: props.getProperty("SHEET_NAME") || "Training",
    RECIPIENT_EMAIL: props.getProperty("RECIPIENT_EMAIL") || "sIndarjit@meadowb.com",
    CC_EMAIL: props.getProperty("CC_EMAIL") || "amohammed@meadowb.com, apagniello@meadowb.com",
    SUBJECT_LINE: props.getProperty("SUBJECT_LINE") || "Weekly Training Expiry Summary",
    DATE_FORMAT: props.getProperty("DATE_FORMAT") || "MMM dd, yyyy"
  };
}

/**
 * One time helper to set script properties from code if you prefer that route.
 * Edit values, run once, then remove or comment out.
 */
function setConfigOnce_() {
  const props = PropertiesService.getScriptProperties();
  props.setProperties({
    SHEET_ID: "1jIAtvQ87J9Fs9SGB4XYA7xzRVzTjpzWi8dP3k1WCNPc",
    SHEET_NAME: "Training",
    RECIPIENT_EMAIL: "sIndarjit@meadowb.com",
    CC_EMAIL: "amohammed@meadowb.com, apagniello@meadowb.com",
    SUBJECT_LINE: "Weekly Training Expiry Summary",
    DATE_FORMAT: "MMM dd, yyyy"
  }, true);
}

/**
 * Create a time based trigger to run every Monday at seven in the morning.
 */
function createTimeDrivenTrigger() {
  ScriptApp.newTrigger("sendWeeklyExpiryEmail")
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(7)
    .create();
}

/**
 * Main entry point
 */
function sendWeeklyExpiryEmail() {
  const cfg = getConfig_();

  const ss = SpreadsheetApp.openById(cfg.SHEET_ID);
  const sheet = ss.getSheetByName(cfg.SHEET_NAME);
  if (!sheet) {
    Logger.log("Training sheet not found");
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) {
    Logger.log("No data rows beyond header");
    return;
  }

  // Field picks from zero based array values
  // trainingName = row[1]
  // firstName = row[2]
  // lastName = row[3]
  // expiryDate = row[7]
  // department = row[8]

  const normCourse = s =>
    String(s || "")
      .split(/\n/)[0]
      .split("(")[0]
      .trim()
      .toLowerCase();

  const normPerson = (first, last) =>
    `${String(first || "").trim()} ${String(last || "").trim()}`
      .replace(/\s+/g, " ")
      .toLowerCase();

  const now = new Date();
  const nowTime = now.getTime();
  const dayMs = 24 * 60 * 60 * 1000;
  const plus30 = new Date(nowTime + 30 * dayMs);
  const plus60 = new Date(nowTime + 60 * dayMs);

  // Keep only the newest entry per person and course
  const latest = new Map();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const trainingRaw = row[1];
    const firstName = row[2];
    const lastName = row[3];
    const expiryDate = row[7];
    const department = row[8];

    if (!(expiryDate instanceof Date)) continue;

    const key = `${normPerson(firstName, lastName)}||${normCourse(trainingRaw)}`;
    const expiryTime = expiryDate.getTime();
    const existing = latest.get(key);
    if (!existing || expiryTime > existing.expiryTime) {
      latest.set(key, {
        department,
        firstName,
        lastName,
        trainingName: trainingRaw,
        expiryDate,
        expiryTime
      });
    }
  }

  const expiredItems = [];
  const expiring30Items = [];
  const expiring60Items = [];

  latest.forEach(item => {
    if (item.expiryTime < nowTime) {
      expiredItems.push(item);
    } else if (item.expiryTime <= plus30.getTime()) {
      expiring30Items.push(item);
    } else if (item.expiryTime <= plus60.getTime()) {
      expiring60Items.push(item);
    }
  });

  const expiredList = formatList_(expiredItems, cfg.DATE_FORMAT);
  const expiring30List = formatList_(expiring30Items, cfg.DATE_FORMAT);
  const expiring60List = formatList_(expiring60Items, cfg.DATE_FORMAT);

  if (!expiredList && !expiring30List && !expiring60List) {
    Logger.log("Nothing to report");
    return;
  }

  let emailBody = "Good morning,\n\n";
  if (expiredList) {
    emailBody += "<strong>The following trainings are expired:</strong>\n" + expiredList + "\n\n";
  }
  if (expiring30List) {
    emailBody += "<strong>The following trainings will expire within 30 days:</strong>\n" + expiring30List + "\n\n";
  }
  if (expiring60List) {
    emailBody += "<strong>The following will expire within 60 days:</strong>\n" + expiring60List + "\n\n";
  }

  MailApp.sendEmail({
    to: cfg.RECIPIENT_EMAIL,
    cc: cfg.CC_EMAIL,
    subject: cfg.SUBJECT_LINE,
    htmlBody: emailBody.replace(/\n/g, "<br>")
  });

  Logger.log("Weekly expiry email sent");
}

/**
 * Sort items by department and build the list text with color coding.
 */
function formatList_(items, dateFormat) {
  if (!items || !items.length) return "";

  const departmentColors = {
    "Production": "#FF0000",
    "Estimating": "#FFA500",
    "Human Resources": "#8B00FF",
    "Service": "#008000",
    "Sales": "#0000FF"
  };

  items.sort((a, b) => (a.department || "").localeCompare(b.department || ""));

  let listText = "";
  items.forEach((item, index) => {
    const { department, firstName, lastName, trainingName, expiryDate } = item;
    const formattedDate = Utilities.formatDate(
      expiryDate,
      Session.getScriptTimeZone(),
      dateFormat || "MMM dd, yyyy"
    );
    const color = departmentColors[department] || "#000000";
    listText += `${index + 1}. <span style="color: ${color};"><b>${department}</b></span> ` +
                `- <b>${firstName} ${lastName}</b>; ${trainingName}; Expiry Date: ${formattedDate}\n`;
  });

  return listText.trim();
}

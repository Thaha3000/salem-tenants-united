const SHEETS = {
  "Community Updates": [
    "ID", "Title", "Details", "Date", "Image URL", "Published"
  ],
  "Notices": [
    "ID", "Title", "Details", "Date", "Published"
  ],
  "Events": [
    "ID", "Title", "Details", "Date", "Signup Open", "Published"
  ],
  "Photos": [
    "ID", "Image URL", "Caption", "Published"
  ],
  "Comments": [
    "ID", "Submitted At", "Name", "Comment", "Approved"
  ],
  "Email Subscribers": [
    "ID", "Submitted At", "Email"
  ],
  "Event Signups": [
    "ID", "Submitted At", "Event ID", "Event Title", "Name", "Email"
  ]
};

function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty("SPREADSHEET_ID");

  const ss = id
    ? SpreadsheetApp.openById(id)
    : SpreadsheetApp.create(
        "Salem Tenants United Essex County - Website Database"
      );

  if (!id) {
    props.setProperty("SPREADSHEET_ID", ss.getId());
  }

  Object.entries(SHEETS).forEach(([name, headers]) => {
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    sheet.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground("#426B51")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold");

    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  });

  const updates = ss.getSheetByName("Community Updates");

  if (updates.getLastRow() === 1) {
    updates.appendRow([
      Utilities.getUuid(),
      "Welcome to Salem Tenants United",
      "A community space for Salem and Essex County tenants.",
      "Community Announcement",
      "",
      "Yes"
    ]);
  }

  Logger.log("GOOGLE SHEET: " + ss.getUrl());
}

function getSpreadsheet() {
  const id = PropertiesService.getScriptProperties()
    .getProperty("SPREADSHEET_ID");

  if (!id) {
    throw new Error("Run setup first.");
  }

  return SpreadsheetApp.openById(id);
}

function getRows(sheetName) {
  const sheet = getSpreadsheet().getSheetByName(sheetName);
  const values = sheet.getDataRange().getDisplayValues();
  const headers = values.shift();

  return values
    .filter(row => row.some(value => value.trim()))
    .map(row => Object.fromEntries(
      headers.map((header, index) => [
        header, row[index] || ""
      ])
    ));
}

function isYes(value) {
  return ["yes", "true", "1", "approved"]
    .includes(String(value).trim().toLowerCase());
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Salem Tenants United Essex County")
    .addMetaTag(
      "viewport",
      "width=device-width, initial-scale=1"
    );
}

function getPublicData() {
  return {
    updates: getRows("Community Updates")
      .filter(row => isYes(row.Published)).reverse(),

    notices: getRows("Notices")
      .filter(row => isYes(row.Published)).reverse(),

    events: getRows("Events")
      .filter(row => isYes(row.Published)).reverse(),

    photos: getRows("Photos")
      .filter(row => isYes(row.Published)).reverse(),

    comments: getRows("Comments")
      .filter(row => isYes(row.Approved)).reverse()
  };
}

function clean(value, maxLength) {
  return String(value || "")
    .trim()
    .slice(0, maxLength)
    .replace(/[\u0000-\u001f]/g, " ");
}

function rateLimit(type) {
  const cache = CacheService.getScriptCache();
  const key = type + "_" +
    (Session.getTemporaryActiveUserKey() || "anonymous");

  if (cache.get(key)) {
    throw new Error("Please wait before submitting again.");
  }

  cache.put(key, "1", 20);
}

function appendData(sheetName, values) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    getSpreadsheet()
      .getSheetByName(sheetName)
      .appendRow(values);
  } finally {
    lock.releaseLock();
  }
}

function submitComment(form) {
  rateLimit("comment");

  const name = clean(form.name, 80);
  const comment = clean(form.comment, 1200);

  if (name.length < 2 || comment.length < 3) {
    throw new Error("Please enter your name and comment.");
  }

  appendData("Comments", [
    Utilities.getUuid(),
    new Date(),
    name,
    comment,
    "No"
  ]);

  return "Thank you! Your comment was submitted for review.";
}

function subscribe(form) {
  rateLimit("email");

  const email = clean(form.email, 180);

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error("Please enter a valid email address.");
  }

  appendData("Email Subscribers", [
    Utilities.getUuid(),
    new Date(),
    email
  ]);

  return "Thank you! Your email has been saved.";
}

function eventSignup(form) {
  rateLimit("signup");

  const eventId = clean(form.eventId, 100);
  const name = clean(form.name, 80);
  const email = clean(form.email, 180);

  if (
    !eventId ||
    name.length < 2 ||
    !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)
  ) {
    throw new Error("Please complete the signup form.");
  }

  const event = getRows("Events").find(row =>
    row.ID === eventId &&
    isYes(row.Published) &&
    isYes(row["Signup Open"])
  );

  if (!event) {
    throw new Error("Signups are closed for this event.");
  }

  appendData("Event Signups", [
    Utilities.getUuid(),
    new Date(),
    eventId,
    event.Title,
    name,
    email
  ]);

  return "You have signed up successfully!";
}

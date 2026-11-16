// ========= INITIALIZATION & UI =========
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("NPS Tools")
    .addItem("Check Venmo payments", "checkVenmoPayments")
    .addItem("Check Zelle payments", "checkZellePayments")
    .addItem("Check PayPal payments", "checkPayPalPayments")
    .addSeparator()
    .addItem("Run ALL payment checks", "runAllPaymentChecks")
    .addSeparator()
    .addItem("🧹 Remove Search Bars", "removeSearchBars")
    .addToUi();
}

// ========= REMOVE SEARCH BARS UTILITY =========
function removeSearchBars() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ["Venmo", "PayPal", "Zelle"];
  
  sheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    
    // Check if search bar exists (row 1 has "🔍 SEARCH:")
    if (sheet.getRange("A1").getValue() === "🔍 SEARCH:") {
      // Delete the first 3 rows (search interface)
      sheet.deleteRows(1, 3);
      
      // Reset header formatting
      sheet.getRange(1, 1, 1, 7).setBackground("#ffffff").setFontWeight("bold");
    }
  });
  
  SpreadsheetApp.getUi().alert("✅ Search bars removed from all payment sheets!");
}

// ========= ADMIN / MAINTENANCE UTILITIES =========
function clearAllPaymentSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const PAYMENT_SHEETS = [
    "Venmo",
    "Zelle",
    "PayPal"
  ];

  PAYMENT_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // Find where data actually starts (skip search bars if present)
    const dataStartRow = sheet.getRange("A1").getValue() === "🔍 SEARCH:" ? 5 : 2;
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    // Only clear data rows (keep headers)
    if (lastRow >= dataStartRow) {
      sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).clearContent();
    }
  });
  
  SpreadsheetApp.getUi().alert("✅ All payment data cleared!");
}

// ========= PAYMENT PROCESSING LOGIC =========
function runAllPaymentChecks() {
  importPayPalEmails();
  importVenmoEmails();
  importZelleEmails();
}

// ========== VENMO FUNCTION ==========
function importVenmoEmails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create "Venmo" sheet
  var sheet = spreadsheet.getSheetByName("Venmo");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Venmo");
    // Add headers
    sheet.getRange(1, 1, 1, 7).setValues([["Date", "Type (payment or charge)", "Amount", "Person", "Note", "Status", "Cinnamon Buns to Make"]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }
  
  // Determine where data starts (handle search bars if present)
  var hasSearchBar = sheet.getRange("A1").getValue() === "🔍 SEARCH:";
  var dataStartRow = hasSearchBar ? 5 : 2;
  var headerRow = hasSearchBar ? 4 : 1;
  
  // Only search for "paid you" emails (money you received)
  var query =
  'from:venmo@venmo.com ' +
  '(' +
    'subject:"paid you" OR ' +
    'subject:"paid to your Venmo account" OR ' +
    'subject:"paid your" OR ' +
    'subject:"paid $" ' +
  ') ' +
   '-subject:"You paid" ' +
  '-subject:("charge" OR "confirmation" OR "verify")';

  var threads = GmailApp.search(query, 0, 100);

  // Get existing data to check for duplicates
  var lastRow = sheet.getLastRow();
  var existingData = [];
  var existingKeys = new Set();
  
  if (lastRow >= dataStartRow) {
    existingData = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 7).getValues();
    // Create unique keys from existing data (date + person + amount)
    for (var i = 0; i < existingData.length; i++) {
      var existingDate = existingData[i][0];
      var existingDateStr = (existingDate instanceof Date)
        ? Utilities.formatDate(existingDate, 'America/New_York', 'M/d/yyyy')
        : String(existingDate).trim();

      var existingPerson = String(existingData[i][3] || "").trim();
      var existingAmount = parseFloat(existingData[i][2]) || 0;

      var key = existingDateStr + "|" + existingPerson + "|" + existingAmount;
      existingKeys.add(key);
    }
  }
  
  var newRows = [];
  var skippedCount = 0;
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var htmlBody = message.getBody();
      var subject = message.getSubject();
      var date = message.getDate();
      
      // Extract amount from HTML
      var amountMatch = htmlBody.match(/\$(\d+\.\d{2})/);
      var amount = amountMatch ? parseFloat(amountMatch[1]) : 0;
      
      // Extract person (who paid you) from subject
      var person = "";
      var personMatch =
        subject.match(/^(.+?)\s+paid you/i) ||
        subject.match(/^(.+?)\s+paid\s+\$[\d,]+(?:\.\d{2})?\s+to your Venmo account/i) ||
        subject.match(/^(.+?)\s+paid your\s+\$[\d,]+(?:\.\d{2})?\s+request/i) || 
        subject.match(/^(.+?)\s+paid\s+\$[\d,]+(?:\.\d{2})?/i);

      var person = personMatch ? personMatch[1].trim() : "";

      // Extract note from HTML
      var note = "";
      var noteMatch = htmlBody.match(/class="transaction-note[^>]*>([^<]+)</i);
      if (noteMatch) {
        note = noteMatch[1].trim();
      }
      
      if (!note) {
        var noteMatch2 = htmlBody.match(/transaction-note.*?>(.+?)<\/p>/is);
        if (noteMatch2) {
          note = noteMatch2[1].replace(/<[^>]*>/g, '').trim();
        }
      }
      
      // Create unique key for this transaction
      var dateStr = Utilities.formatDate(date, 'America/New_York', 'M/d/yyyy');
      var uniqueKey = dateStr + "|" + person + "|" + amount;
      
      // Skip if duplicate
      if (existingKeys.has(uniqueKey)) {
        skippedCount++;
        continue;
      }
      
      // Calculate buns using helper function
      var bunCount = calculateBuns(amount);
      
      newRows.push([
        dateStr,
        "Payment Received",
        amount,
        person,
        note,
        "Completed",
        bunCount
      ]);
      
      Logger.log("Added: " + person + " - $" + amount);
    }
  }
  
  Logger.log("New transactions: " + newRows.length);
  Logger.log("Skipped duplicates: " + skippedCount);
  
  // Add to sheet
  if (newRows.length > 0) {
    // Sort by date (newest first)
    newRows.sort(function(a, b) {
      return new Date(b[0]) - new Date(a[0]);
    });
    
    sheet.insertRowsAfter(headerRow, newRows.length);
    sheet.getRange(dataStartRow, 1, newRows.length, 7).setValues(newRows);
    
    // Calculate total buns
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(dataStartRow, 7, lastRow - dataStartRow + 1, 1).getValues();
    var totalBuns = 0;
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        totalBuns += parseInt(data[i][0]);
      }
    }
    
    sheet.getRange("H" + headerRow).setValue("🍥 Total Buns");
    sheet.getRange("H" + dataStartRow).setValue(totalBuns);
    
    Browser.msgBox('✅ VENMO: Added ' + newRows.length + ' new transactions.\n' + 
                   'Skipped ' + skippedCount + ' duplicates.\n' +
                   'Total buns needed: ' + totalBuns);
  } else {
    Browser.msgBox('No new Venmo payments found.\n' + 
                   'Skipped ' + skippedCount + ' existing transactions.');
  }
}

// ========== PAYPAL FUNCTION (writes to separate sheet) ==========
function importPayPalEmails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create "PayPal" sheet
  var sheet = spreadsheet.getSheetByName("PayPal");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("PayPal");
    // Add headers
    sheet.getRange(1, 1, 1, 7).setValues([["Date", "Type (payment or charge)", "Amount", "Person", "Note", "Status", "Cinnamon Buns to Make"]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");
  }
  
  // Determine where data starts
  var hasSearchBar = sheet.getRange("A1").getValue() === "🔍 SEARCH:";
  var dataStartRow = hasSearchBar ? 5 : 2;
  var headerRow = hasSearchBar ? 4 : 1;
  
  // Only search for legitimate PayPal emails
  var threads = GmailApp.search('from:(@paypal.com) ' +
    '(subject:"You received" OR subject:"sent you" OR subject:"Payment received") ' +
    '-subject:"You sent" -subject:"You paid"', 0, 100);
  
  Logger.log("Found " + threads.length + " PayPal email threads");
  
  var lastRow = sheet.getLastRow();
  var existingData = [];
  var existingKeys = new Set();
  
  if (lastRow >= dataStartRow) {
    existingData = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 7).getValues();
    for (var i = 0; i < existingData.length; i++) {
      var existingDate = existingData[i][0];
      var existingDateStr = (existingDate instanceof Date)
        ? Utilities.formatDate(existingDate, 'America/New_York', 'M/d/yyyy')
        : String(existingDate).trim();
      
      var existingPerson = String(existingData[i][3] || "").trim();
      var existingAmount = parseFloat(existingData[i][2]) || 0;
      
      var key = existingDateStr + "|" + existingPerson + "|" + existingAmount;
      existingKeys.add(key);
    }
  }
  
  var newRows = [];
  var skippedCount = 0;
  
  for (var i = 0; i < threads.length; i++) {
    var messages = threads[i].getMessages();
    
    for (var j = 0; j < messages.length; j++) {
      var message = messages[j];
      var htmlBody = message.getBody();
      var plainBody = message.getPlainBody();
      var subject = message.getSubject();
      var date = message.getDate();
      
      // Extract amount (PayPal can have $X.XX or $X,XXX.XX format)
      var amountMatch = 
        htmlBody.match(/\$(\d{1,3}(?:,\d{3})*\.\d{2})/i) ||
        plainBody.match(/\$(\d{1,3}(?:,\d{3})*\.\d{2})/i);
      var amount = 0;
      if (amountMatch) {
        amount = parseFloat(amountMatch[1].replace(/,/g, ''));
      }
      
      // Extract person who paid you
      var person = "";
      var personMatch = 
        subject.match(/(.+?)\s+sent you/i) ||
        subject.match(/payment from\s+(.+?)$/i) ||
        subject.match(/You received.+?from\s+(.+?)$/i) ||
        htmlBody.match(/From:\s*<[^>]*>([^<]+)</i) ||
        plainBody.match(/From:\s*(.+?)[\n\r]/i);
      person = personMatch ? personMatch[1].trim() : "";
      
      // Extract note/message
      var note = "";
      var noteMatch = 
        htmlBody.match(/Note:\s*<[^>]*>([^<]+)</i) ||
        htmlBody.match(/Note:\s*(.+?)</i) ||
        plainBody.match(/Note:\s*(.+?)[\n\r]/i) ||
        htmlBody.match(/Message:\s*<[^>]*>([^<]+)</i) ||
        htmlBody.match(/Message:\s*(.+?)</i) ||
        plainBody.match(/Message:\s*(.+?)[\n\r]/i);
      if (noteMatch) {
        note = noteMatch[1].replace(/<[^>]*>/g, '').trim();
      }
      
      var dateStr = Utilities.formatDate(date, 'America/New_York', 'M/d/yyyy');
      var uniqueKey = dateStr + "|" + person + "|" + amount;
      
      if (existingKeys.has(uniqueKey)) {
        skippedCount++;
        continue;
      }
      
      var bunCount = calculateBuns(amount);
      
      newRows.push([dateStr, "Payment Received", amount, person, note, "Completed", bunCount]);
      Logger.log("Added: " + person + " - $" + amount);
    }
  }
  
  Logger.log("New transactions: " + newRows.length);
  Logger.log("Skipped duplicates: " + skippedCount);
  
  if (newRows.length > 0) {
    newRows.sort(function(a, b) {
      return new Date(b[0]) - new Date(a[0]);
    });
    
    sheet.insertRowsAfter(headerRow, newRows.length);
    sheet.getRange(dataStartRow, 1, newRows.length, 7).setValues(newRows);
    
    var lastRow = sheet.getLastRow();
    var data = sheet.getRange(dataStartRow, 7, lastRow - dataStartRow + 1, 1).getValues();
    var totalBuns = 0;
    
    for (var i = 0; i < data.length; i++) {
      if (data[i][0]) {
        totalBuns += parseInt(data[i][0]);
      }
    }
    
    sheet.getRange("H" + headerRow).setValue("🍥 Total Buns");
    sheet.getRange("H" + dataStartRow).setValue(totalBuns);
    
    Browser.msgBox('✅ PAYPAL: Added ' + newRows.length + ' new transactions to PayPal sheet.\n' + 
                   'Skipped ' + skippedCount + ' duplicates.\n' +
                   'Total buns needed: ' + totalBuns);
  } else {
    Browser.msgBox('No new PayPal payments found.\n' + 
                   'Skipped ' + skippedCount + ' existing transactions.');
  }
}

// ========== ZELLE FUNCTION (writes to separate sheet) ==========
function importZelleEmails() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get or create "Zelle" sheet
  var sheet = spreadsheet.getSheetByName("Zelle");
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Zelle");
    sheet.getRange(1, 1, 1, 7).setValues([["Date", "Type (payment or charge)", "Amount", "Person", "Note", "Status", "Cinnamon Buns to Make"]]);
    sheet.getRange(1, 1, 1, 7).setFontWeight("bold");

    // Simple config row (optional)
    sheet.getRange("A9").setValue("Config (optional):");
    sheet.getRange("A10").setValue("Gmail query override");
    sheet.getRange("A11").setValue("Exclude outgoing? (TRUE/FALSE)");
    sheet.getRange("B11").setValue("TRUE");
  }

  // Determine where data starts
  var hasSearchBar = sheet.getRange("A1").getValue() === "🔍 SEARCH:";
  var dataStartRow = hasSearchBar ? 5 : 2;
  var headerRow = hasSearchBar ? 4 : 1;

  // Allow user override query from sheet (B10)
  var overrideQuery = String(sheet.getRange("B10").getValue() || "").trim();
  var excludeOutgoing = String(sheet.getRange("B11").getValue() || "TRUE").toUpperCase() === "TRUE";

  // Default: bank-agnostic Zelle RECEIPT search (received money)
  var query =
    overrideQuery ||
    '(' +
      '(zelle OR subject:"Zelle" OR "with Zelle" OR "with Zelle®") ' +
      'AND (subject:"sent you" OR subject:"you received" OR subject:"received $" OR subject:"is now available" OR subject:"money is available" OR subject:"payment is available")' +
    ') ' +
    '-subject:("enrollment" OR "enroll" OR "verify" OR "verification" OR "security" OR "password" OR "code" OR "fraud" OR "alert" OR "scheduled" OR "canceled" OR "cancelled")';

  // Optional: exclude outgoing payments
  if (excludeOutgoing) {
    query += ' -subject:("you sent" OR "you paid" OR "payment sent" OR "you made")';
  }

  var threads = GmailApp.search(query, 0, 100);
  Logger.log("Found " + threads.length + " Zelle email threads with query: " + query);

  // Dedup keys from existing sheet data
  var lastRow = sheet.getLastRow();
  var existingKeys = new Set();

  if (lastRow >= dataStartRow) {
    var existingData = sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, 7).getValues();
    for (var i = 0; i < existingData.length; i++) {
      var existingDate = existingData[i][0];
      var existingDateStr = (existingDate instanceof Date)
        ? Utilities.formatDate(existingDate, 'America/New_York', 'M/d/yyyy')
        : String(existingDate).trim();

      var existingPerson = String(existingData[i][3] || "").trim();
      var existingAmount = parseFloat(existingData[i][2]) || 0;

      existingKeys.add(existingDateStr + "|" + existingPerson + "|" + existingAmount);
    }
  }

  var newRows = [];
  var skippedCount = 0;
  var parseFailCount = 0;

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();

    for (var m = 0; m < messages.length; m++) {
      var message = messages[m];
      var subject = message.getSubject() || "";
      var date = message.getDate();
      var htmlBody = message.getBody() || "";
      var plainBody = message.getPlainBody() || "";

      var note = "";
      
      // A) Common "Memo/Message/Note" labels in BANK emails (HTML)
      var noteMatch =
        htmlBody.match(/(?:Memo|Message|Note)\s*:\s*<\/?[^>]*>\s*([^<\r\n]{1,200})/i) ||
        htmlBody.match(/(?:Memo|Message|Note)[^<]*<\/?[^>]*>\s*([^<\r\n]{1,200})/i);

      if (noteMatch) {
        note = noteMatch[1].replace(/<[^>]*>/g, "").trim();
      }

      // B) Plain text labeled lines
      if (!note) {
        var noteMatch2 =
          plainBody.match(/^(?:Memo|Message|Note)\s*:\s*(.+)$/mi) ||
          plainBody.match(/^(?:Memo|Message|Note)\s+(.+)$/mi);
        if (noteMatch2) {
          note = String(noteMatch2[1]).trim();
        }
      }

      // C) "sent you $X for <something>" subject style
      if (!note) {
        var noteMatch3 = subject.match(/\s+for\s+(.+?)\s*$/i);
        if (noteMatch3) {
          note = String(noteMatch3[1]).replace(/[)\]]+$/, "").trim();
        }
      }

      // Final cleanup
      if (note) {
        note = note.replace(/\s{2,}/g, " ").trim();
      }

      // Extra safety: skip obvious outgoing
      if (/^You\s+(sent|paid)\b/i.test(subject)) {
        continue;
      }

      // Amount: prefer subject, then plain body, then html
      var amount = extractAmount(subject) || extractAmount(plainBody) || extractAmount(htmlBody) || 0;

      if (!amount || amount <= 0) {
        parseFailCount++;
        Logger.log("[ZELLE PARSE FAIL] No amount. Subject: " + subject);
        continue;
      }

      // Person: best-effort generic patterns
      var person = extractZelleCounterparty(subject, plainBody, htmlBody);
      if (!person) person = "Unknown";

      var dateStr = Utilities.formatDate(date, 'America/New_York', 'M/d/yyyy');
      var uniqueKey = dateStr + "|" + person + "|" + amount;

      if (existingKeys.has(uniqueKey)) {
        skippedCount++;
        continue;
      }

      var bunCount = calculateBuns(amount);

      var typeLabel = "Zelle Received";
      newRows.push([dateStr, typeLabel, amount, person, note, "Completed", bunCount]);

      Logger.log("ZELLE Added: " + person + " - $" + amount + " (" + subject + ")");
    }
  }

  if (newRows.length > 0) {
    newRows.sort(function(a, b) { return new Date(b[0]) - new Date(a[0]); });
    sheet.insertRowsAfter(headerRow, newRows.length);
    sheet.getRange(dataStartRow, 1, newRows.length, 7).setValues(newRows);

    var lastRow2 = sheet.getLastRow();
    var data = sheet.getRange(dataStartRow, 7, lastRow2 - dataStartRow + 1, 1).getValues();
    var totalBuns = 0;
    for (var k = 0; k < data.length; k++) {
      if (data[k][0]) totalBuns += parseInt(data[k][0], 10);
    }
    sheet.getRange("H" + headerRow).setValue("🍥 Total Buns");
    sheet.getRange("H" + dataStartRow).setValue(totalBuns);
  }

  Browser.msgBox(
    "✅ ZELLE: Added " + newRows.length +
    " new transactions.\nSkipped " + skippedCount +
    " duplicates.\nParse failures (skipped): " + parseFailCount +
    "\n\nTip: If parse failures are high, set a Gmail query override in Zelle!B10."
  );
}

// Extract $ amounts like "$38.00" or "$1,234.56"
function extractAmount(text) {
  if (!text) return 0;
  var m = String(text).match(/\$([\d,]+(?:\.\d{2})?)/);
  return m ? parseFloat(m[1].replace(/,/g, "")) : 0;
}

// Generic best-effort counterparty extraction across banks
function extractZelleCounterparty(subject, plainBody, htmlBody) {
  var s = subject || "";
  var p = plainBody || "";
  var h = htmlBody || "";

  // OLD FORMAT: "NAME sent you $56.00 with Zelle®"
  var m = s.match(/^(.+?)\s+sent you\s+\$[\d,]+(?:\.\d{2})?\s+with\s+Zelle/i);
  if (m) return cleanupName(m[1]);

  // Common patterns
  var m =
    s.match(/^\s*[\(\[]?\s*(.+?)\s+sent you\s+\$[\d,]+(?:\.\d{2})?/i) ||
    s.match(/received\s+\$[\d,]+(?:\.\d{2})?\s+from\s+(.+)/i) ||
    p.match(/received\s+\$[\d,]+(?:\.\d{2})?\s+from\s+(.+)/i) ||
    s.match(/from\s+(.+?)\s+via\s+zelle/i) ||
    p.match(/from\s+(.+?)\s+via\s+zelle/i);

  if (m) return cleanupName(m[1]);

  // Fallback: look for "From: NAME" in body
  m = p.match(/From:\s*(.+)/i);
  if (m) return cleanupName(m[1]);

  return "";
}

function cleanupName(name) {
  if (!name) return "";
  return String(name)
    .replace(/<[^>]*>/g, "")
    .replace(/[\r\n]+/g, " ")
    .replace(/\s{2,}/g, " ")
    .trim();
}

// ========== HELPER FUNCTIONS ==========
function calculateBuns(amount) {
  var bunCount = 0;
  if (amount == 2) bunCount = 1;
  else if (amount == 5) bunCount = 3;
  else if (amount == 10) bunCount = 5;
  else if (amount == 4) bunCount = 2;
  else if (amount == 7) bunCount = 4;
  else if (amount == 12) bunCount = 6;
  else if (amount == 15) bunCount = 8;
  else if (amount == 20) bunCount = 10;
  else {
    var remaining = amount;
    while (remaining >= 10) {
      bunCount += 5;
      remaining -= 10;
    }
    while (remaining >= 5) {
      bunCount += 3;
      remaining -= 5;
    }
    while (remaining >= 2) {
      bunCount += 1;
      remaining -= 2;
    }
    if (remaining >= 1) {
      bunCount += Math.floor(remaining / 2);
    }
  }
  return bunCount;
}

function clearAllData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
    Browser.msgBox('✅ Cleared all data! Run any import function to reimport.');
  } else {
    Browser.msgBox('Sheet is already empty.');
  }
}
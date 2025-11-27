function updateFreshLeadOnly() {
  try {
    const CLEAN_SHEET = "Your Sheet Name";
    const SUBJECT_LINE = "Email subject line";
    const threads = GmailApp.search(`subject:"${SUBJECT_LINE}"`, 0, 1);
    if (threads.length === 0) throw new Error("No email found with this subject");
    const lastMessage = threads[0].getMessages().pop();
    const attachments = lastMessage.getAttachments();
    if (!attachments.length) throw new Error("CSV attachment missing");
    
    const csvContent = attachments[0].getDataAsString("UTF-8");
    let csvData = Utilities.parseCsv(csvContent);

    csvData = csvData.map(row =>
      row.map(cell =>
        (typeof cell === "string" && cell.startsWith('"') && cell.endsWith('"'))
          ? cell.slice(1, -1)
          : cell
      )
    );

    const headers = csvData[0];
    const rows = csvData.slice(1);

    // CONDITION MATCH 1 – Filter based on city, status & date
    const validConditions1 = ["Online Without Date", "Online with Date", "AS-N", "AS-Y"];
    const validCities = ["Chennai", "Bangalore", "Hyderabad", "Kolkata", "Mumbai", "Pune", "Thane", "Navi Mumbai","Ahmedabad","Gandhinagar"];

    const now = new Date();
    const yesterday6pm = new Date(now);
    yesterday6pm.setDate(now.getDate() - 1);
    yesterday6pm.setHours(18, 0, 0, 0);

    const filtered = rows.filter(r => {
      const buyerId = r[1];
      const city = r[9];
      const status = r[18];
      const rowDate = tryParseDate(r[11]);
      return (
        buyerId &&
        city &&
        status &&
        rowDate &&
        validConditions1.includes(String(status).trim()) &&
        validCities.includes(String(city).trim()) &&
        rowDate >= yesterday6pm &&
        rowDate <= now
      );
    });

    const seen = new Set();
    const uniqueFiltered = filtered.filter(r => {
      const id = String(r[1]).trim();
      if (seen.has(id)) return false;
      seen.add(id);
      return true;
    });

    const today6pm = new Date(now);
    today6pm.setHours(18, 0, 0, 0);

    const finalData = uniqueFiltered.map(r => {
      const rowDate = tryParseDate(r[11]);
      const flag = rowDate && rowDate >= today6pm ? "Lead Created - Today after 6PM" : "";
      return r.concat([flag]);
    });

    const newHeaders = headers.concat(["Lead Flag"]);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CLEAN_SHEET);
    if (!sheet) sheet = ss.insertSheet(CLEAN_SHEET);

    sheet.clear();
    sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);

    if (finalData.length > 0) {
      sheet.getRange(2, 1, finalData.length, newHeaders.length).setValues(finalData);
    }

  } catch (err) {
    Logger.log("Error: " + err.message);
  }
}

function tryParseDate(val) {
  if (val instanceof Date) return val;
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d;
}

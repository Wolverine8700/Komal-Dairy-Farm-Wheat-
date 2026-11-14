// Google Sheets -> Extensions -> Apps Script માં આ કોડ પેસ્ટ કરીને Deploy -> Web App (Anyone) કરો
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var raw = e.postData && e.postData.contents ? e.postData.contents : '{}';
    var data = JSON.parse(raw);

    function writeSheet(name, headers, rows) {
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clear();
      sh.appendRow(headers);
      if (rows && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    writeSheet('Cows', ['Tag No', 'Breed', 'AI Date', 'Inseminator', 'Status', 'Dry-Off Date', 'Delivery Date', 'Note'],
      (data.cows || []).map(function(c) {
        return [c.tag || '', c.breed || '', c.aiDate || '', c.doctor || '', c.status || '', c.dryDate || '', c.dueDate || '', c.note || ''];
      }));

    writeSheet('MilkLog', ['Date', 'Morning (L)', 'Evening (L)', 'Total (L)', 'Fat %', 'Rate', 'Amount (Rs)', 'Note'],
      (data.milk || []).map(function(m) {
        var tot = Number(m.morning || 0) + Number(m.evening || 0);
        return [m.date || '', Number(m.morning || 0), Number(m.evening || 0), tot, Number(m.fat || 0), Number(m.rate || 0), tot * Number(m.rate || 0), m.note || ''];
      }));

    writeSheet('SickLog', ['Date', 'Tag No', 'Disease', 'Medicine', 'Withdrawal Days', 'Status'],
      (data.sick || []).map(function(s) {
        return [s.date || '', s.tag || '', s.disease || '', s.medicine || '', Number(s.withdrawal || 0), s.status || ''];
      }));

    writeSheet('Expenses', ['Date', 'Category', 'Note', 'Amount (Rs)'],
      (data.expenses || []).map(function(ex) {
        return [ex.date || '', ex.cat || '', ex.note || '', Number(ex.amount || 0)];
      }));

    writeSheet('WheatCustomers', ['id', 'name', 'village', 'mobile', 'business', 'vehicle', 'vehicleNo', 'createdAt'],
      (data.customers || []).map(function(c) {
        return [c.id || '', c.name || '', c.village || '', c.mobile || '', c.business || '', c.vehicle || '', c.vehicleNo || '', c.createdAt || ''];
      }));

    writeSheet('WheatBilling', ['Date', 'Farmer Name', 'Village', 'Mobile', 'Weight (Kg)', 'Rate/Kg', 'Total (Rs)', 'Payment'],
      (data.wheatBills || []).map(function(wb) {
        return [wb.date || '', wb.name || '', wb.village || '', wb.mobile || '', Number(wb.kg || 0), Number(wb.rate || 0), Number(wb.total || 0), wb.payStatus || 'રોકડ'];
      }));

    writeSheet('WheatThinPurchase', ['Date', 'Farmer Name', 'Mobile', 'Thin Wheat (Kg)', 'Rate/Kg', 'Purchase (Rs)', 'Labour Due (Rs)', 'Net Balance (Rs)'],
      (data.wheatPurchases || []).map(function(wp) {
        return [wp.date || '', wp.name || '', wp.mobile || '', Number(wp.kg || 0), Number(wp.rate || 0), Number(wp.purchaseAmt || 0), Number(wp.labourAmt || 0), Number(wp.netAmt || 0)];
      }));

    var stateSh = ss.getSheetByName('_CombinedStateJSON') || ss.insertSheet('_CombinedStateJSON');
    stateSh.clear();
    var chunkSize = 45000;
    var chunks = [];
    for (var i = 0; i < raw.length; i += chunkSize) chunks.push([raw.substring(i, i + chunkSize)]);
    if (chunks.length) stateSh.getRange(1, 1, chunks.length, 1).setValues(chunks);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName('_CombinedStateJSON');
    if (!sh || sh.getLastRow() === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'empty' })).setMimeType(ContentService.MimeType.JSON);
    }
    var vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    var full = vals.map(function(r) { return r[0] || ''; }).join('') || JSON.stringify({ status: 'empty' });
    return ContentService.createTextOutput(full).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Komal Dairy Farm & Wheat — Secure Cloud Sync (Google Apps Script)
 * 1) PIN Verification (Owner & Staff)
 * 2) Safe Merge by id + updatedAt (never wipes concurrent edits)
 */
const PROP_KEY_OWNER_PIN = 'WHEAT_OWNER_PIN_HASH';
const PROP_KEY_STAFF_PIN = 'WHEAT_STAFF_PIN_HASH';

function isAuthorized(pinHash) {
  var props = PropertiesService.getScriptProperties();
  var owner = props.getProperty(PROP_KEY_OWNER_PIN) || '';
  var staff = props.getProperty(PROP_KEY_STAFF_PIN) || '';
  if (!owner && pinHash) {
    props.setProperty(PROP_KEY_OWNER_PIN, String(pinHash));
    return true;
  }
  if (!pinHash) return false;
  return String(pinHash) === owner || (staff && String(pinHash) === staff);
}

function mergeArraysById(existingArr, incomingArr) {
  var map = {};
  (existingArr || []).forEach(function(item) {
    if (item && item.id != null) map[String(item.id)] = item;
  });
  (incomingArr || []).forEach(function(item) {
    if (!item || item.id == null) return;
    var k = String(item.id);
    var prev = map[k];
    if (!prev || Number(item.updatedAt || 0) >= Number(prev.updatedAt || 0)) {
      map[k] = item;
    }
  });
  return Object.keys(map).map(function(k) { return map[k]; });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var raw = e.postData && e.postData.contents ? e.postData.contents : '{}';
    var incoming = JSON.parse(raw);
    if (!isAuthorized(incoming.pinHash)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', error: 'Unauthorized PIN' })).setMimeType(ContentService.MimeType.JSON);
    }

    var stateSh = ss.getSheetByName('_CombinedStateJSON') || ss.insertSheet('_CombinedStateJSON');
    var existing = { cows: [], milk: [], sick: [], expenses: [], customers: [], wheatBills: [], wheatPurchases: [] };
    if (stateSh.getLastRow() > 0) {
      var vals = stateSh.getRange(1, 1, stateSh.getLastRow(), 1).getValues();
      var prevRaw = vals.map(function(r) { return r[0] || ''; }).join('');
      if (prevRaw) { try { existing = JSON.parse(prevRaw); } catch (err) {} }
    }

    var data = {
      cows: mergeArraysById(existing.cows, incoming.cows),
      milk: mergeArraysById(existing.milk, incoming.milk),
      sick: mergeArraysById(existing.sick, incoming.sick),
      expenses: mergeArraysById(existing.expenses, incoming.expenses),
      customers: mergeArraysById(existing.customers, incoming.customers),
      wheatBills: mergeArraysById(existing.wheatBills, incoming.wheatBills),
      wheatPurchases: mergeArraysById(existing.wheatPurchases, incoming.wheatPurchases)
    };

    function writeSheet(name, headers, rows) {
      var sh = ss.getSheetByName(name) || ss.insertSheet(name);
      sh.clearContents();
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (rows && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    writeSheet('Cows', ['ID', 'Tag No', 'Breed', 'AI Date', 'Inseminator', 'Status', 'Dry-Off Date', 'Delivery Date', 'Note', 'Deleted', 'UpdatedAt'],
      (data.cows || []).map(function(c) {
        return [c.id || '', c.tag || '', c.breed || '', c.aiDate || '', c.doctor || '', c.status || '', c.dryDate || '', c.dueDate || '', c.note || '', !!c.deleted, c.updatedAt || 0];
      }));
    writeSheet('MilkLog', ['ID', 'Date', 'Morning (L)', 'Evening (L)', 'Total (L)', 'Fat %', 'Rate', 'Amount (Rs)', 'Note', 'Deleted', 'UpdatedAt'],
      (data.milk || []).map(function(m) {
        var tot = Number(m.morning || 0) + Number(m.evening || 0);
        return [m.id || '', m.date || '', Number(m.morning || 0), Number(m.evening || 0), tot, Number(m.fat || 0), Number(m.rate || 0), tot * Number(m.rate || 0), m.note || '', !!m.deleted, m.updatedAt || 0];
      }));
    writeSheet('SickLog', ['ID', 'Date', 'Tag No', 'Disease', 'Medicine', 'Withdrawal Days', 'Status', 'Deleted', 'UpdatedAt'],
      (data.sick || []).map(function(s) {
        return [s.id || '', s.date || '', s.tag || '', s.disease || '', s.medicine || '', Number(s.withdrawal || 0), s.status || '', !!s.deleted, s.updatedAt || 0];
      }));
    writeSheet('Expenses', ['ID', 'Date', 'Category', 'Note', 'Amount (Rs)', 'Deleted', 'UpdatedAt'],
      (data.expenses || []).map(function(ex) {
        return [ex.id || '', ex.date || '', ex.cat || '', ex.note || '', Number(ex.amount || 0), !!ex.deleted, ex.updatedAt || 0];
      }));
    writeSheet('WheatCustomers', ['id', 'name', 'village', 'mobile', 'business', 'vehicle', 'vehicleNo', 'createdAt', 'Deleted', 'UpdatedAt'],
      (data.customers || []).map(function(c) {
        return [c.id || '', c.name || '', c.village || '', c.mobile || '', c.business || '', c.vehicle || '', c.vehicleNo || '', c.createdAt || '', !!c.deleted, c.updatedAt || 0];
      }));
    writeSheet('WheatBilling', ['ID', 'Date', 'Farmer Name', 'Village', 'Mobile', 'Weight (Kg)', 'Rate/Kg', 'Total (Rs)', 'Payment', 'Deleted', 'UpdatedAt'],
      (data.wheatBills || []).map(function(wb) {
        return [wb.id || '', wb.date || '', wb.name || '', wb.village || '', wb.mobile || '', Number(wb.kg || 0), Number(wb.rate || 0), Number(wb.total || 0), wb.payStatus || 'રોકડ', !!wb.deleted, wb.updatedAt || 0];
      }));
    writeSheet('WheatThinPurchase', ['ID', 'Date', 'Farmer Name', 'Mobile', 'Thin Wheat (Kg)', 'Rate/Kg', 'Purchase (Rs)', 'Labour Due (Rs)', 'Net Balance (Rs)', 'Deleted', 'UpdatedAt'],
      (data.wheatPurchases || []).map(function(wp) {
        return [wp.id || '', wp.date || '', wp.name || '', wp.mobile || '', Number(wp.kg || 0), Number(wp.rate || 0), Number(wp.purchaseAmt || 0), Number(wp.labourAmt || 0), Number(wp.netAmt || 0), !!wp.deleted, wp.updatedAt || 0];
      }));

    var mergedJson = JSON.stringify(data);
    stateSh.clearContents();
    var chunkSize = 45000;
    var chunks = [];
    for (var i = 0; i < mergedJson.length; i += chunkSize) chunks.push([mergedJson.substring(i, i + chunkSize)]);
    if (chunks.length) stateSh.getRange(1, 1, chunks.length, 1).setValues(chunks);

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: String(err) })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var pin = (e && e.parameter && e.parameter.pin) ? e.parameter.pin : '';
    if (!isAuthorized(pin)) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', error: 'Unauthorized PIN' })).setMimeType(ContentService.MimeType.JSON);
    }
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

// ============================================================
//  StylesHub Pro — Code.gs
//  Google Apps Script Backend
//  SME Systems Studio — Patience Oduori
// ============================================================

var SS_ID = '1hHS0oaOTrnGPgul6sXekZBMm4zWfS2IQQI5FzBjKm0k';

function ss() { return SpreadsheetApp.openById(SS_ID); }
function sh(name) { return ss().getSheetByName(name); }
function rows(name) { return sh(name).getDataRange().getValues(); }
function tz() { return Session.getScriptTimeZone(); }
function fdate(d) { try { return Utilities.formatDate(new Date(d), tz(), 'yyyy-MM-dd'); } catch(e) { return ''; } }
function uid(p) { return (p||'X') + '_' + new Date().getTime() + '_' + Math.floor(Math.random()*9999); }

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('StylesHub Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ─── LOGIN ────────────────────────────────────────────────────────────────────
function loginUser(u, p) {
  var users = {
    admin: { p:'admin123', name:'Admin', role:'admin' },
    cashier: { p:'cashier123', name:'Cashier', role:'cashier' }
  };
  if (users[u] && users[u].p === p) return { success:true, user:{ name:users[u].name, role:users[u].role } };
  return { success:false, error:'Invalid username or password' };
}

// ─── SETTINGS ────────────────────────────────────────────────────────────────
function getSettings() {
  try {
    var data = rows('Settings');
    var s = {};
    for (var i = 1; i < data.length; i++) { if (data[i][0]) s[String(data[i][0])] = data[i][1]; }
    return { success:true, data:s };
  } catch(e) { return { success:true, data:{ salonName:'StylesHub Pro', currency:'KES' } }; }
}

function saveSettings(data) {
  try {
    var sheet = sh('Settings');
    var existing = sheet.getDataRange().getValues();
    var map = {};
    for (var i = 1; i < existing.length; i++) map[existing[i][0]] = i + 1;
    for (var k in data) {
      if (map[k]) sheet.getRange(map[k], 2).setValue(data[k]);
      else sheet.appendRow([k, data[k]]);
    }
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── DASHBOARD ───────────────────────────────────────────────────────────────
function getDashboard() {
  try {
    var salesData = rows('Sales Log');
    var apptData = rows('Appointments');
    var clientData = rows('Clients');
    var stylistData = rows('Stylists');

    var totalRev = 0, totalTxns = 0;
    var catRev = {}, dateRev = {};
    var saleIdSet = {};
    var recent = [];

    for (var i = 1; i < salesData.length; i++) {
      var r = salesData[i];
      if (!r[0]) continue;
      var total = parseFloat(r[6]) || 0;
      var d = fdate(r[1]);
      totalRev += total;
      totalTxns++;
      dateRev[d] = (dateRev[d] || 0) + total;
      saleIdSet[String(r[0])] = true;
    }

    // Build catRev from Sale Items
    try {
      var itemsData = rows('Sale Items');
      for (var j = 1; j < itemsData.length; j++) {
        if (!itemsData[j][0]) continue;
        if (!saleIdSet[String(itemsData[j][1])]) continue;
        var itemName = String(itemsData[j][3] || 'Other');
        var lineTotal = parseFloat(itemsData[j][6]) || 0;
        catRev[itemName] = (catRev[itemName] || 0) + lineTotal;
      }
    } catch(e2) {}

    // Trend - last 7 dates with sales, sorted oldest to newest
    var sortedDates = Object.keys(dateRev).sort();
    var trend = {};
    var last7 = sortedDates.slice(-7);
    for (var i = 0; i < last7.length; i++) trend[last7[i]] = dateRev[last7[i]];

    // Recent 8 sales
    for (var i = salesData.length - 1; i >= 1 && recent.length < 8; i--) {
      var r = salesData[i];
      if (!r[0]) continue;
      recent.push({ date:fdate(r[1]), receiptNo:String(r[2]), client:String(r[3]||'Walk-in'), stylist:String(r[5]||'—'), total:parseFloat(r[6])||0, payment:String(r[7]||'') });
    }

    // Appointments
    var today = fdate(new Date());
    var todayAppts = 0, pendingAppts = 0, confirmedAppts = 0;
    for (var i = 1; i < apptData.length; i++) {
      if (!apptData[i][0]) continue;
      var d = fdate(apptData[i][3]);
      var st = String(apptData[i][10] || '');
      if (d === today) todayAppts++;
      if (st === 'Pending') pendingAppts++;
      if (st === 'Confirmed') confirmedAppts++;
    }

    var totalClients = 0;
    for (var i = 1; i < clientData.length; i++) { if (clientData[i][0]) totalClients++; }
    var totalStylists = 0;
    for (var i = 1; i < stylistData.length; i++) { if (stylistData[i][0]) totalStylists++; }

    return {
      success: true,
      totalRev: totalRev, totalTxns: totalTxns,
      todayAppts: todayAppts, pendingAppts: pendingAppts, confirmedAppts: confirmedAppts,
      totalClients: totalClients, totalStylists: totalStylists,
      trend: trend, catRev: catRev, recent: recent
    };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── SERVICES ────────────────────────────────────────────────────────────────
function getServices() {
  try {
    var data = rows('Services');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({
        id:String(data[i][0]), name:String(data[i][1]), type:String(data[i][2]),
        category:String(data[i][3]), price:parseFloat(data[i][4])||0,
        cost:parseFloat(data[i][5])||0, duration:parseInt(data[i][6])||0,
        imageUrl:String(data[i][7]||''), emoji:String(data[i][8]||'✨'),
        description:String(data[i][9]||''), status:String(data[i][10]||'Active')
      });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function addService(d) {
  try {
    var id = uid('SVC');
    sh('Services').appendRow([id, d.name, d.type, d.category, parseFloat(d.price)||0, parseFloat(d.cost)||0, parseInt(d.duration)||0, d.imageUrl||'', d.emoji||'✨', d.description||'', d.status||'Active', new Date()]);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateService(d) {
  try {
    var sheet = sh('Services');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(d.id)) {
        sheet.getRange(i+1,1,1,12).setValues([[d.id, d.name, d.type, d.category, parseFloat(d.price)||0, parseFloat(d.cost)||0, parseInt(d.duration)||0, d.imageUrl||'', d.emoji||'✨', d.description||'', d.status||'Active', data[i][11]]]);
        return { success:true };
      }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteService(id) {
  try {
    var sheet = sh('Services');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── APPOINTMENTS ─────────────────────────────────────────────────────────────
function getAppointments(date) {
  try {
    var data = rows('Appointments');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var aDate = fdate(data[i][3]);
      if (date && aDate !== date) continue;
      out.push({
        id:String(data[i][0]), client:String(data[i][1]), phone:String(data[i][2]||''),
        date:aDate, time:String(data[i][4]||''), stylistId:String(data[i][5]||''),
        stylist:String(data[i][6]||''), serviceId:String(data[i][7]||''),
        service:String(data[i][8]||''), duration:parseInt(data[i][9])||60,
        status:String(data[i][10]||'Pending'), notes:String(data[i][11]||'')
      });
    }
    // Sort by time
    out.sort(function(a,b){ return a.time.localeCompare(b.time); });
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function getAllAppointments() {
  try {
    var data = rows('Appointments');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({
        id:String(data[i][0]), client:String(data[i][1]), phone:String(data[i][2]||''),
        date:fdate(data[i][3]), time:String(data[i][4]||''), stylistId:String(data[i][5]||''),
        stylist:String(data[i][6]||''), service:String(data[i][8]||''),
        duration:parseInt(data[i][9])||60, status:String(data[i][10]||'Pending')
      });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function saveAppointment(appt) {
  try {
    var sheet = sh('Appointments');
    var isEdit = appt.id && appt.id !== '';

    // Double booking check
    var allData = rows('Appointments');
    var newStart = timeToMins(appt.time);
    var newEnd = newStart + (parseInt(appt.duration) || 60);

    for (var i = 1; i < allData.length; i++) {
      if (!allData[i][0]) continue;
      if (isEdit && String(allData[i][0]) === String(appt.id)) continue;
      if (String(allData[i][5]) !== String(appt.stylistId)) continue;
      if (fdate(allData[i][3]) !== appt.date) continue;
      if (String(allData[i][10]) === 'Cancelled') continue;

      var aStart = timeToMins(String(allData[i][4]));
      var aEnd = aStart + (parseInt(allData[i][9]) || 60);
      if (newStart < aEnd && newEnd > aStart) {
        return { success:false, conflict:true, error:String(allData[i][6]) + ' already has "' + String(allData[i][8]) + '" booked at ' + String(allData[i][4]) + '. Please choose a different time or stylist.' };
      }
    }

    if (isEdit) {
      for (var i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(appt.id)) {
          sheet.getRange(i+1,1,1,13).setValues([[appt.id, appt.client, appt.phone||'', appt.date, appt.time, appt.stylistId, appt.stylist, appt.serviceId||'', appt.service, parseInt(appt.duration)||60, appt.status||'Pending', appt.notes||'', allData[i][12]]]);
          return { success:true };
        }
      }
    } else {
      var id = uid('APPT');
      sheet.appendRow([id, appt.client, appt.phone||'', appt.date, appt.time, appt.stylistId, appt.stylist, appt.serviceId||'', appt.service, parseInt(appt.duration)||60, 'Pending', appt.notes||'', new Date()]);
      // Add client if not exists
      addClientIfNew(appt.client, appt.phone||'', appt.stylist);
      return { success:true, id:id };
    }
    return { success:false, error:'Appointment not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateAppointmentStatus(id, status) {
  try {
    var sheet = sh('Appointments');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) {
        sheet.getRange(i+1, 11).setValue(status);
        return { success:true };
      }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteAppointment(id) {
  try {
    var sheet = sh('Appointments');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

function timeToMins(t) {
  if (!t) return 0;
  var parts = String(t).split(':');
  return parseInt(parts[0]) * 60 + parseInt(parts[1] || 0);
}

// ─── SALES ────────────────────────────────────────────────────────────────────
function createSale(sale) {
  try {
    var lock = LockService.getScriptLock();
    lock.tryLock(10000);

    var salesSheet = sh('Sales Log');
    var itemsSheet = sh('Sale Items');

    var receipt = 'SH-' + (1000 + salesSheet.getLastRow());
    var today = fdate(new Date());
    var saleId = uid('SALE');

    salesSheet.appendRow([saleId, today, receipt, sale.clientName||'Walk-in', sale.clientPhone||'', sale.stylistName||'', parseFloat(sale.total)||0, sale.paymentMethod||'Cash', parseFloat(sale.amountPaid)||0, sale.cashier||'Cashier', new Date()]);

    for (var i = 0; i < sale.items.length; i++) {
      var item = sale.items[i];
      itemsSheet.appendRow([uid('ITM'), saleId, receipt, item.name, item.qty, item.price, item.price * item.qty, item.category||'']);
    }

    // Update client stats
    updateClientStats(sale.clientName, parseFloat(sale.total)||0, today);

    // Update stylist stats
    if (sale.stylistId) updateStylistStats(sale.stylistId, parseFloat(sale.total)||0);

    lock.releaseLock();
    return { success:true, receipt:receipt, saleId:saleId };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateClientStats(name, amount, date) {
  try {
    if (!name || name === 'Walk-in') return;
    var sheet = sh('Clients');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === name.toLowerCase()) {
        sheet.getRange(i+1, 7).setValue((parseFloat(data[i][6])||0) + amount);
        sheet.getRange(i+1, 8).setValue((parseInt(data[i][7])||0) + 1);
        sheet.getRange(i+1, 9).setValue(date);
        return;
      }
    }
  } catch(e) {}
}

function updateStylistStats(stylistId, amount) {
  try {
    var sheet = sh('Stylists');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(stylistId)) {
        sheet.getRange(i+1, 7).setValue((parseFloat(data[i][6])||0) + amount);
        return;
      }
    }
  } catch(e) {}
}

function getSales(limit) {
  try {
    var data = rows('Sales Log');
    var out = [];
    var max = limit || 200;
    for (var i = data.length-1; i >= 1 && out.length < max; i--) {
      if (!data[i][0]) continue;
      out.push({ id:String(data[i][0]), date:fdate(data[i][1]), receiptNo:String(data[i][2]), client:String(data[i][3]||'Walk-in'), stylist:String(data[i][5]||'—'), total:parseFloat(data[i][6])||0, payment:String(data[i][7]||''), cashier:String(data[i][9]||'') });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function getSaleItems(saleId) {
  try {
    var data = rows('Sale Items');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]) === String(saleId)) out.push({ name:String(data[i][3]), qty:data[i][4], price:data[i][5], total:data[i][6] });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function getSalesReport(start, end) {
  try {
    var data = rows('Sales Log');
    var out = [], total = 0;
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var d = fdate(data[i][1]);
      if (d >= start && d <= end) {
        out.push({ date:d, receiptNo:String(data[i][2]), client:String(data[i][3]||'Walk-in'), stylist:String(data[i][5]||'—'), total:parseFloat(data[i][6])||0, payment:String(data[i][7]||'') });
        total += parseFloat(data[i][6])||0;
      }
    }
    return { success:true, data:out, total:total };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── STYLISTS ─────────────────────────────────────────────────────────────────
function getStylists() {
  try {
    var data = rows('Stylists');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({ id:String(data[i][0]), name:String(data[i][1]), role:String(data[i][2]||''), phone:String(data[i][3]||''), email:String(data[i][4]||''), specialties:String(data[i][5]||''), totalSales:parseFloat(data[i][6])||0, totalClients:parseInt(data[i][7])||0, dateAdded:fdate(data[i][8]) });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function addStylist(d) {
  try {
    var id = uid('STY');
    sh('Stylists').appendRow([id, d.name, d.role||'', d.phone||'', d.email||'', d.specialties||'', 0, 0, new Date()]);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function updateStylist(d) {
  try {
    var sheet = sh('Stylists');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(d.id)) {
        sheet.getRange(i+1,1,1,9).setValues([[d.id, d.name, d.role||'', d.phone||'', d.email||'', d.specialties||'', data[i][6], data[i][7], data[i][8]]]);
        return { success:true };
      }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteStylist(id) {
  try {
    var sheet = sh('Stylists');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── CLIENTS ──────────────────────────────────────────────────────────────────
function getClients() {
  try {
    var data = rows('Clients');
    var out = [];
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      out.push({ id:String(data[i][0]), name:String(data[i][1]), phone:String(data[i][2]||''), email:String(data[i][3]||''), hairType:String(data[i][4]||''), preferredStylist:String(data[i][5]||''), totalSpent:parseFloat(data[i][6])||0, visits:parseInt(data[i][7])||0, lastVisit:fdate(data[i][8]), birthday:fdate(data[i][9]), notes:String(data[i][10]||''), dateAdded:fdate(data[i][11]) });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function addClient(d) {
  try {
    var id = uid('CLT');
    sh('Clients').appendRow([id, d.name, d.phone||'', d.email||'', d.hairType||'', d.preferredStylist||'', 0, 0, '', d.birthday||'', d.notes||'', new Date()]);
    return { success:true, id:id };
  } catch(e) { return { success:false, error:e.message }; }
}

function addClientIfNew(name, phone, preferredStylist) {
  try {
    if (!name || name === 'Walk-in') return;
    var data = rows('Clients');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === name.toLowerCase()) return;
    }
    sh('Clients').appendRow([uid('CLT'), name, phone||'', '', '', preferredStylist||'', 0, 0, '', '', '', new Date()]);
  } catch(e) {}
}

function deleteClient(id) {
  try {
    var sheet = sh('Clients');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── EXPENSES ─────────────────────────────────────────────────────────────────
function getExpenses() {
  try {
    var data = rows('Expenses');
    var out = [];
    for (var i = data.length-1; i >= 1; i--) {
      if (!data[i][0]) continue;
      out.push({ id:String(data[i][0]), date:fdate(data[i][1]), category:String(data[i][2]||''), description:String(data[i][3]||''), amount:parseFloat(data[i][4])||0, notes:String(data[i][5]||''), recordedBy:String(data[i][6]||'') });
    }
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

function addExpense(d) {
  try {
    sh('Expenses').appendRow([uid('EXP'), fdate(new Date()), d.category, d.description, parseFloat(d.amount)||0, d.notes||'', d.recordedBy||'']);
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

function deleteExpense(id) {
  try {
    var sheet = sh('Expenses');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(id)) { sheet.deleteRow(i+1); return { success:true }; }
    }
    return { success:false, error:'Not found' };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── REPORTS ──────────────────────────────────────────────────────────────────
function getProfitLoss(start, end) {
  try {
    var salesData = rows('Sales Log');
    var itemsData = rows('Sale Items');
    var expData = rows('Expenses');
    var svcData = rows('Services');

    var costMap = {};
    for (var i = 1; i < svcData.length; i++) { if (svcData[i][1]) costMap[String(svcData[i][1])] = parseFloat(svcData[i][5])||0; }

    var rev = 0, saleIds = {};
    for (var i = 1; i < salesData.length; i++) {
      if (!salesData[i][0]) continue;
      var d = fdate(salesData[i][1]);
      if (d >= start && d <= end) { rev += parseFloat(salesData[i][6])||0; saleIds[salesData[i][0]] = true; }
    }

    var cogs = 0;
    for (var i = 1; i < itemsData.length; i++) {
      if (saleIds[itemsData[i][1]]) cogs += (costMap[String(itemsData[i][3])]||0) * (parseInt(itemsData[i][4])||0);
    }

    var exps = 0;
    for (var i = 1; i < expData.length; i++) {
      if (!expData[i][0]) continue;
      var d = fdate(expData[i][1]);
      if (d >= start && d <= end) exps += parseFloat(expData[i][4])||0;
    }

    var gross = rev - cogs, net = gross - exps;
    return { success:true, data:{ revenue:rev, cogs:cogs, grossProfit:gross, expenses:exps, netProfit:net, grossMargin:rev>0?Math.round(gross/rev*100):0, netMargin:rev>0?Math.round(net/rev*100):0 } };
  } catch(e) { return { success:false, error:e.message }; }
}

function getStylistReport(start, end) {
  try {
    var data = rows('Sales Log');
    var stylistRev = {}, stylistTxns = {};
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var d = fdate(data[i][1]);
      if (d >= start && d <= end) {
        var name = String(data[i][5]||'Unknown');
        stylistRev[name] = (stylistRev[name]||0) + (parseFloat(data[i][6])||0);
        stylistTxns[name] = (stylistTxns[name]||0) + 1;
      }
    }
    var out = Object.keys(stylistRev).map(function(name){ return { name:name, revenue:stylistRev[name], txns:stylistTxns[name] }; });
    out.sort(function(a,b){ return b.revenue - a.revenue; });
    return { success:true, data:out };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── SEED DATA ────────────────────────────────────────────────────────────────
function seedSampleData() {
  try {
    // Stylists
    var stylistSheet = sh('Stylists');
    stylistSheet.getRange(2,1,10,9).clearContent();
    stylistSheet.appendRow(['STY_001','Amina Wanjiku','Senior Hair Stylist','+254 712 111 222','amina@styleshub.co.ke','Hair, Colour, Treatment',145600,48,new Date()]);
    stylistSheet.appendRow(['STY_002','Faith Muthoni','Nail Technician','+254 723 222 333','faith@styleshub.co.ke','Nails, Lashes',98400,35,new Date()]);
    stylistSheet.appendRow(['STY_003','Grace Otieno','Beauty Therapist','+254 734 333 444','grace@styleshub.co.ke','Skin, Makeup, Lashes',112000,42,new Date()]);

    // Services
    var svcSheet = sh('Services');
    svcSheet.getRange(2,1,20,12).clearContent();
    var svcs = [
      ['SVC_001','Blowout & Style','Service','Hair',1500,200,60,'https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?w=300&q=80','💇‍♀️','Full blowout with styling','Active'],
      ['SVC_002','Full Set Acrylic Nails','Service','Nails',2500,400,90,'https://images.unsplash.com/photo-1604654894610-df63bc536371?w=300&q=80','💅','Full acrylic nail set','Active'],
      ['SVC_003','Hair Relaxer','Service','Hair',2000,500,120,'https://images.unsplash.com/photo-1580618672591-eb180b1a973f?w=300&q=80','✨','Full hair relaxer treatment','Active'],
      ['SVC_004','Box Braids','Service','Braids',5000,800,300,'https://images.unsplash.com/photo-1611080626919-7cf5a9dbab12?w=300&q=80','🌿','Medium box braids','Active'],
      ['SVC_005','Facial Treatment','Service','Skin',2200,600,60,'https://images.unsplash.com/photo-1512290923902-8a9f81dc236c?w=300&q=80','🌸','Deep cleansing facial','Active'],
      ['SVC_006','Lash Extensions','Service','Lashes',3500,700,120,'https://images.unsplash.com/photo-1516975080664-ed2fc6a32937?w=300&q=80','👁️','Classic lash set','Active'],
      ['SVC_007','Gel Manicure','Service','Nails',1800,300,60,'https://images.unsplash.com/photo-1604654894610-df63bc536371?w=300&q=80','💅','Gel manicure with design','Active'],
      ['SVC_008','Hair Colour','Service','Hair',4500,1200,180,'https://images.unsplash.com/photo-1522337360788-8b13dee7a37e?w=300&q=80','🎨','Full hair colour','Active'],
      ['SVC_009','OGX Shampoo 385ml','Product','Product',1200,700,0,'https://images.unsplash.com/photo-1522338242992-e1a54906a8da?w=300&q=80','🧴','OGX strengthening shampoo','Active'],
      ['SVC_010','Hair Mask Treatment','Service','Treatment',1800,400,45,'https://images.unsplash.com/photo-1571781926291-c477ebfd024b?w=300&q=80','🌿','Deep conditioning mask','Active'],
      ['SVC_011','Makeup — Full Face','Service','Makeup',3500,500,90,'https://images.unsplash.com/photo-1487412947147-5cebf100ffc2?w=300&q=80','💄','Full glam makeup','Active'],
      ['SVC_012','Cantu Leave-In Cream','Product','Product',950,550,0,'https://images.unsplash.com/photo-1522338242992-e1a54906a8da?w=300&q=80','🧴','Cantu shea butter leave-in','Active'],
    ];
    for (var i = 0; i < svcs.length; i++) {
      svcSheet.appendRow(svcs[i].concat([new Date()]));
    }

    // Clients
    var clientSheet = sh('Clients');
    clientSheet.getRange(2,1,10,12).clearContent();
    clientSheet.appendRow(['CLT_001','Mary Kamau','+254 745 444 555','mary@gmail.com','Natural','Amina Wanjiku',28500,12,'2024-05-20','','Sensitive scalp. Loves natural styles.',new Date()]);
    clientSheet.appendRow(['CLT_002','Aisha Mohamed','+254 756 555 666','','Relaxed','Amina Wanjiku',45000,18,'2024-05-22','','Regular client every 2 weeks.',new Date()]);
    clientSheet.appendRow(['CLT_003','Cynthia Odhiambo','+254 767 666 777','cynthia@gmail.com','Natural','Faith Muthoni',18000,8,'2024-05-18','','Loves gel nails with art.',new Date()]);
    clientSheet.appendRow(['CLT_004','Njeri Githae','+254 778 777 888','','Loc\'d','Grace Otieno',32000,15,'2024-05-25','','Monthly facial + lash appointments.',new Date()]);
    clientSheet.appendRow(['CLT_005','Wanjiru Kariuki','+254 789 888 999','wanjiru@gmail.com','Colour Treated','Amina Wanjiku',67000,24,'2024-05-28','','Colour client. Always books Amina.',new Date()]);

    // Expenses
    var expSheet = sh('Expenses');
    expSheet.getRange(2,1,10,7).clearContent();
    expSheet.appendRow(['EXP_001','2024-05-01','Rent','May salon rent',35000,'','Admin']);
    expSheet.appendRow(['EXP_002','2024-05-05','Products Purchase','OGX and Cantu restock',18000,'','Admin']);
    expSheet.appendRow(['EXP_003','2024-05-10','Salaries','Staff salaries — 3 stylists',85000,'','Admin']);
    expSheet.appendRow(['EXP_004','2024-05-15','Utilities','Electricity and water',8500,'','Admin']);
    expSheet.appendRow(['EXP_005','2024-05-20','Supplies','Nail supplies and chemicals',12000,'','Admin']);

    // Sales
    var salesSheet = sh('Sales Log');
    var itemsSheet = sh('Sale Items');
    salesSheet.getRange(2,1,30,11).clearContent();
    itemsSheet.getRange(2,1,60,8).clearContent();
    var salesSeed = [
      ['SALE_001','2024-05-01','SH-1001','Mary Kamau','+254 745 444 555','Amina Wanjiku',1500,'M-Pesa',1500,'Admin'],
      ['SALE_002','2024-05-02','SH-1002','Aisha Mohamed','+254 756 555 666','Amina Wanjiku',2000,'M-Pesa',2000,'Admin'],
      ['SALE_003','2024-05-03','Cynthia Odhiambo','+254 767 666 777','Faith Muthoni',2500,'Cash',2500,'Cashier'],
      ['SALE_004','2024-05-04','SH-1004','Walk-in','','Faith Muthoni',1800,'Cash',1800,'Cashier'],
      ['SALE_005','2024-05-05','SH-1005','Njeri Githae','+254 778 777 888','Grace Otieno',2200,'M-Pesa',2200,'Admin'],
      ['SALE_006','2024-05-07','SH-1006','Wanjiru Kariuki','+254 789 888 999','Amina Wanjiku',4500,'M-Pesa',4500,'Admin'],
      ['SALE_007','2024-05-08','SH-1007','Mary Kamau','+254 745 444 555','Amina Wanjiku',1200,'M-Pesa',1200,'Admin'],
      ['SALE_008','2024-05-10','SH-1008','Walk-in','','Grace Otieno',3500,'Cash',3500,'Cashier'],
      ['SALE_009','2024-05-12','SH-1009','Aisha Mohamed','+254 756 555 666','Amina Wanjiku',5000,'M-Pesa',5000,'Admin'],
      ['SALE_010','2024-05-14','SH-1010','Cynthia Odhiambo','+254 767 666 777','Faith Muthoni',3500,'M-Pesa',3500,'Admin'],
      ['SALE_011','2024-05-15','SH-1011','Njeri Githae','+254 778 777 888','Grace Otieno',1800,'Cash',1800,'Cashier'],
      ['SALE_012','2024-05-17','SH-1012','Wanjiru Kariuki','+254 789 888 999','Amina Wanjiku',2200,'M-Pesa',2200,'Admin'],
      ['SALE_013','2024-05-18','SH-1013','Walk-in','','Faith Muthoni',950,'Cash',950,'Cashier'],
      ['SALE_014','2024-05-20','SH-1014','Mary Kamau','+254 745 444 555','Amina Wanjiku',2000,'M-Pesa',2000,'Admin'],
      ['SALE_015','2024-05-22','SH-1015','Aisha Mohamed','+254 756 555 666','Amina Wanjiku',1800,'M-Pesa',1800,'Admin'],
    ];
    for (var i = 0; i < salesSeed.length; i++) {
      var s = salesSeed[i];
      if (s.length === 10) salesSheet.appendRow([s[0],s[1],s[2],s[3],s[4],s[5],s[6],s[7],s[8],s[9],new Date()]);
    }

    // Appointments
    var apptSheet = sh('Appointments');
    apptSheet.getRange(2,1,20,13).clearContent();
    var today = fdate(new Date());
    var tomorrow = new Date(); tomorrow.setDate(tomorrow.getDate()+1); var tmr = fdate(tomorrow);
    apptSheet.appendRow(['APPT_001','Mary Kamau','+254 745 444 555',today,'09:00','STY_001','Amina Wanjiku','SVC_001','Blowout & Style',60,'Confirmed','',new Date()]);
    apptSheet.appendRow(['APPT_002','Aisha Mohamed','+254 756 555 666',today,'10:30','STY_001','Amina Wanjiku','SVC_003','Hair Relaxer',120,'Confirmed','Uses OGX shampoo only',new Date()]);
    apptSheet.appendRow(['APPT_003','Cynthia Odhiambo','+254 767 666 777',today,'11:00','STY_002','Faith Muthoni','SVC_002','Full Set Acrylic Nails',90,'Pending','',new Date()]);
    apptSheet.appendRow(['APPT_004','Njeri Githae','+254 778 777 888',today,'14:00','STY_003','Grace Otieno','SVC_005','Facial Treatment',60,'Confirmed','Monthly facial',new Date()]);
    apptSheet.appendRow(['APPT_005','Wanjiru Kariuki','+254 789 888 999',tmr,'10:00','STY_001','Amina Wanjiku','SVC_008','Hair Colour',180,'Confirmed','Full colour retouch',new Date()]);
    apptSheet.appendRow(['APPT_006','Walk-in','',tmr,'13:00','STY_002','Faith Muthoni','SVC_007','Gel Manicure',60,'Pending','',new Date()]);

    Logger.log('✅ StylesHub seed data complete!');
    return { success:true };
  } catch(e) { return { success:false, error:e.message }; }
}

// ─── BOOKING PORTAL FUNCTIONS ────────────────────────────────────────────────

// Serve the booking portal
function doGet(e) {
  var page = e && e.parameter && e.parameter.page;
  if (page === 'booking') {
    return HtmlService.createHtmlOutputFromFile('booking')
      .setTitle('Book an Appointment')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('StylesHub Pro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Load all data needed for the portal
function getPortalData() {
  try {
    var svcResult = getServices();
    var stylistResult = getStylists();
    var apptResult = getAllAppointments();
    var settingsResult = getSettings();
    return {
      success: true,
      services: svcResult.success ? svcResult.data : [],
      stylists: stylistResult.success ? stylistResult.data : [],
      appointments: apptResult.success ? apptResult.data : [],
      settings: settingsResult.success ? settingsResult.data : {}
    };
  } catch(e) { return { success:false, error:e.message }; }
}

// Client submits a booking from the portal
function clientBookAppointment(appt) {
  try {
    var result = saveAppointment(appt);
    return result;
  } catch(e) { return { success:false, error:e.message }; }
}
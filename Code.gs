var SHEET_NAME = "Penyewa";

function getConfig(key) {
  var value = PropertiesService.getScriptProperties().getProperty(key);
  if (!value) {
    throw new Error('Sila set Script Property "' + key + '" di Project Settings Apps Script.');
  }
  return value;
}

function getSpreadsheet() {
  return SpreadsheetApp.openById(getConfig("SPREADSHEET_ID"));
}

function getSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(["Timestamp", "Nama Penuh", "No IC", "Tempat Tinggal", "No Waris"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function servePage(file, title) {
  var out = HtmlService.createTemplateFromFile(file).evaluate();
  out.setTitle(title);
  return out;
}

function doPost(e) {
  try {
    var raw = (e.postData && e.postData.contents) ? e.postData.contents : "";
    var body = JSON.parse(raw);

    var nama = String(body.nama || "").trim();
    var ic = String(body.ic || "").trim();
    var alamat = String(body.alamat || "").trim();
    var waris = String(body.waris || "").trim();

    if (!nama || !ic || !alamat || !waris) {
      return jsonResponse({ ok: false, error: "Semua medan wajib diisi." });
    }
    if (!/^\d{12}$/.test(ic)) {
      return jsonResponse({ ok: false, error: "No IC mesti 12 digit (cth: 950101141234)." });
    }
    if (!/^\d[\d\s-]{7,12}$/.test(waris)) {
      return jsonResponse({ ok: false, error: "Nombor waris tidak sah." });
    }

    var sheet = getSheet();
    sheet.appendRow([new Date(), nama, ic, alamat, waris]);
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setNumberFormat("@").setValue(ic);
    sheet.getRange(lastRow, 5).setNumberFormat("@").setValue(waris);
    return jsonResponse({ ok: true, message: "Borang diterima. Selamat memancing!" });
  } catch (err) {
    return jsonResponse({ ok: false, error: "Ralat berlaku. Sila cuba lagi." });
  }
}

function doGet(e) {
  var page = e.parameter.page || "";
  var token = e.parameter.token || "";
  var callback = e.parameter.callback || "";
  var mode = e.parameter.mode || "";
  var tarikh = e.parameter.tarikh || "";

  if (mode === "senarai") {
    var sheetS = getSheet();
    var vals = sheetS.getDataRange().getValues();
    var sen = [];
    for (var j = 1; j < vals.length; j++) {
      var w = vals[j];
      if (!w[0]) continue;
      var tim = new Date(w[0]);
      if (tarikh) {
        var hri = Utilities.formatDate(tim, "Asia/Kuala_Lumpur", "yyyy-MM-dd");
        if (hri !== tarikh) continue;
      }
      sen.push({ timestamp: w[0], nama: w[1], waris: w[4] });
    }
    sen.sort(function (a, b) {
      return new Date(b.timestamp) - new Date(a.timestamp);
    });
    var payloadS = { ok: true, rows: sen };
    if (callback) {
      var cbS = String(callback).replace(/[^A-Za-z0-9_$.]/g, "");
      if (!cbS) cbS = "callback";
      return ContentService.createTextOutput(cbS + "(" + JSON.stringify(payloadS) + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonResponse(payloadS);
  }

  if (token) {
    var payload;
    if (token !== getConfig("ADMIN_TOKEN")) {
      payload = { ok: false, error: "Token tidak sah." };
    } else {
      var sheet = getSheet();
      var values = sheet.getDataRange().getValues();
      var rows = [];
      for (var i = 1; i < values.length; i++) {
        var v = values[i];
        if (!v[0]) continue;
        rows.push({
          timestamp: v[0],
          nama: v[1],
          ic: v[2],
          alamat: v[3],
          waris: v[4]
        });
      }
      rows.sort(function (a, b) {
        return new Date(b.timestamp) - new Date(a.timestamp);
      });
      payload = { ok: true, rows: rows };
    }
    if (callback) {
      var cb = String(callback).replace(/[^A-Za-z0-9_$.]/g, "");
      if (!cb) cb = "callback";
      return ContentService.createTextOutput(cb + "(" + JSON.stringify(payload) + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonResponse(payload);
  }

  if (page === "form") {
    return servePage("index", "Borang Pra-Naik Bot");
  }

  if (page === "pengurusan") {
    return servePage("admin", "Panel Pentadbir");
  }

  if (page === "senarai") {
    return servePage("senarai", "Senarai Peserta");
  }

  return servePage("home", "Pra-Naik Bot");
}

function clearEmptyRows() {
  var sheet = getSheet();
  for (var i = sheet.getLastRow(); i >= 2; i--) {
    if (!sheet.getRange(i, 2).getValue()) {
      sheet.deleteRow(i);
    }
  }
}
const SPREADSHEET_URL = 'https://docs.google.com/spreadsheets/d/1DMC1H-Pz-29HYRUU0SMioMKu1sCW69lOGRZkp212sP4/edit?gid=0#gid=0';
const WEB_APP_URL     = 'https://script.google.com/macros/s/AKfycbwV0ccaRmGAEITzL-p6kmWWRwKuXdUSf8Wr-Ax55v0zTn6yZZ_bUmYeQxzft1zyQVl-zw/exec';

/**
 * Opens the spreadsheet defined by SPREADSHEET_URL.
 * @return {SpreadsheetApp.Spreadsheet}
 */
function getSS() {
  try {
    return SpreadsheetApp.openByUrl(SPREADSHEET_URL);
  } catch (err) {
    throw new Error('Unable to open spreadsheet. Check SPREADSHEET_URL. ' + err);
  }
}

/**
 * Retrieves the Links sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getLinksSheet() {
  var ss = getSS();
  var sh = ss.getSheetByName('Links');
  if (!sh) throw new Error('Sheet "Links" not found. Run setupAndFill() first.');
  return sh;
}

/**
 * Retrieves the Clicks sheet.
 * @return {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getClicksSheet() {
  var ss = getSS();
  var sh = ss.getSheetByName('Clicks');
  if (!sh) throw new Error('Sheet "Clicks" not found. Run setupAndFill() first.');
  return sh;
}

/**
 * Builds a header map for a sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Object<string, number>}
 */
function getHeaderMap_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) {
    if (h) map[h] = i + 1;
  });
  return map;
}

/**
 * Handles /exec?id=<ID> requests, logs the click, and redirects.
 * @param {Object} e Event parameter
 * @return {HtmlService.HtmlOutput}
 */
function doGet(e) {
  var id = e && e.parameter && e.parameter.id;
  if (!id) {
    return HtmlService.createHtmlOutput('Missing id parameter')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var linksSheet;
  try {
    linksSheet = getLinksSheet();
  } catch (err) {
    return HtmlService.createHtmlOutput(err.message)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  var headerMap = getHeaderMap_(linksSheet);
  var data = linksSheet.getDataRange().getValues();
  var title = '';
  var youtubeUrl = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][headerMap['ID'] - 1] === id) {
      title = data[i][headerMap['Title'] - 1] || '';
      youtubeUrl = data[i][headerMap['YouTubeURL'] - 1] || '';
      break;
    }
  }
  if (!youtubeUrl) {
    return HtmlService.createHtmlOutput('ID not found or YouTubeURL missing.')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  try {
    var clicksSheet = getClicksSheet();
    clicksSheet.appendRow([
      new Date(),
      id,
      title,
      youtubeUrl,
      (e.parameter.ua || ''),
      (e.parameter.ref || ''),
      JSON.stringify(e.parameter || {})
    ]);
  } catch (err) {
    return HtmlService.createHtmlOutput(err.message)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  var sanitized = youtubeUrl.replace(/"/g, '&quot;');
  var html = '<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Redirecting...</title>' +
             '<script>function go(){var t=' + JSON.stringify(youtubeUrl) + ';' +
             'try{window.open(t,"_blank","noopener,noreferrer");}catch(e){}' +
             'try{window.top.location.href=t;}catch(e){try{window.parent.location.href=t;}catch(e){window.location.assign(t);}}' +
             '}</script></head>' +
             '<body onload="go()">' +
             '<noscript><meta http-equiv="refresh" content="0;url=' + sanitized + '"></noscript>' +
             '<p>Redirecting... <a href="' + sanitized + '" target="_top">Click here</a></p>' +
             '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Creates required sheets and headers, populates formulas.
 */
function setupAndFill() {
  var ss = getSS();
  var linksSheet = ss.getSheetByName('Links');
  if (!linksSheet) linksSheet = ss.insertSheet('Links');
  var clicksSheet = ss.getSheetByName('Clicks');
  if (!clicksSheet) clicksSheet = ss.insertSheet('Clicks');

  linksSheet.getRange(1, 1, 1, 6).setValues([["ID", "Title", "YouTubeURL", "RedirectURL", "QR", "ClickCount"]]);
  clicksSheet.getRange(1, 1, 1, 7).setValues([["Timestamp", "ID", "Title", "YouTubeURL", "UserAgent", "Referrer", "QueryString"]]);

  var headerMap = getHeaderMap_(linksSheet);
  var lastRow = linksSheet.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var id = linksSheet.getRange(r, headerMap['ID']).getValue();
    var youtubeUrl = linksSheet.getRange(r, headerMap['YouTubeURL']).getValue();
    if (!id || !youtubeUrl) continue;

    var redirectUrl = WEB_APP_URL + '?id=' + encodeURIComponent(id);
    linksSheet.getRange(r, headerMap['RedirectURL']).setValue(redirectUrl);

    var redirectA1 = 'Links!D' + r;
    var idA1 = 'Links!A' + r;
    linksSheet.getRange(r, headerMap['QR']).setFormula('=IMAGE("https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" & ENCODEURL(' + redirectA1 + '))');
    linksSheet.getRange(r, headerMap['ClickCount']).setFormula('=COUNTIF(Clicks!B:B,' + idA1 + ')');
  }
}

/**
 * Proves connectivity to the spreadsheet.
 */
function selfCheck() {
  var sh = getLinksSheet();
  sh.getRange('H1').setValue('Connected at ' + new Date());
}

/**
 * Exports individual QR PNGs to Drive and writes download links.
 */
function exportQRCodesToDrive_Standalone() {
  var ss = getSS();
  var linksSheet = getLinksSheet();
  var headerMap = getHeaderMap_(linksSheet);
  var folderName = 'QR-Codes - ' + ss.getName();
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var pngCol = headerMap['PNGLink'];
  if (!pngCol) {
    pngCol = linksSheet.getLastColumn() + 1;
    linksSheet.getRange(1, pngCol).setValue('PNGLink');
    headerMap['PNGLink'] = pngCol;
  }

  var lastRow = linksSheet.getLastRow();
  for (var r = 2; r <= lastRow; r++) {
    var id = linksSheet.getRange(r, headerMap['ID']).getValue();
    var redirectUrl = linksSheet.getRange(r, headerMap['RedirectURL']).getValue();
    if (!id || !redirectUrl) continue;
    var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=1024x1024&data=' + encodeURIComponent(redirectUrl);
    var blob = UrlFetchApp.fetch(qrUrl).getBlob().setName(id + '.png');
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var link = 'https://drive.google.com/uc?export=download&id=' + file.getId();
    linksSheet.getRange(r, headerMap['PNGLink']).setValue(link);
  }

  ss.toast('QR PNGs folder: ' + folder.getUrl());
}

/**
 * Builds a ZIP of all QR PNGs and shares it.
 */
function exportQRCodesZip_Standalone() {
  var ss = getSS();
  var linksSheet = getLinksSheet();
  var headerMap = getHeaderMap_(linksSheet);
  var lastRow = linksSheet.getLastRow();
  var blobs = [];
  for (var r = 2; r <= lastRow; r++) {
    var id = linksSheet.getRange(r, headerMap['ID']).getValue();
    var redirectUrl = linksSheet.getRange(r, headerMap['RedirectURL']).getValue();
    if (!id || !redirectUrl) continue;
    var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=1024x1024&data=' + encodeURIComponent(redirectUrl);
    var blob = UrlFetchApp.fetch(qrUrl).getBlob().setName(id + '.png');
    blobs.push(blob);
  }
  if (!blobs.length) {
    ss.toast('No QR codes to export.');
    return;
  }
  var ts = Utilities.formatDate(new Date(), 'GMT', 'yyyyMMdd-HHmmss');
  var zipBlob = Utilities.zip(blobs, 'QR-Codes-' + ts + '.zip');
  var file = DriveApp.createFile(zipBlob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  ss.toast('QR ZIP file: ' + file.getUrl());
}

/*
How to use
1) Save file, run selfCheck() — check Links!H1 updates.
2) Run setupAndFill() — verify RedirectURL/QR/ClickCount filled.
3) Deploy Web App (Execute as: Me, Who has access: Anyone) or Redeploy after changes.
4) Open a RedirectURL/scan QR — YouTube should open in a new tab; Clicks and ClickCount update.
5) For downloads: run exportQRCodesToDrive_Standalone() for per-QR PNG links, or exportQRCodesZip_Standalone() for a single ZIP.
*/

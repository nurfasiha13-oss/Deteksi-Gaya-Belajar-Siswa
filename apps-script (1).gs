// Google Apps Script - penerima data Deteksi Gaya Belajar
const PASS = "gurubk234";   // samakan dengan password guru BK
const SH = "Hasil";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SH);
  if (!s) {
    s = ss.insertSheet(SH);
    s.appendRow(["Waktu","Nama","Kelas","Visual","Auditori","Kinestetik","Gaya Dominan"]);
  }
  return s;
}
function doPost(e) {
  const d = JSON.parse(e.postData.contents);
  const L = LockService.getScriptLock();
  L.waitLock(10000);
  sheet_().appendRow([d.waktu, d.nama, d.kelas, d.V, d.A, d.K, d.dominan]);
  L.releaseLock();
  return ContentService.createTextOutput("ok");
}
function doGet(e) {
  if (!e.parameter.pw || e.parameter.pw !== PASS) return out_({error: "unauthorized"});
  const v = sheet_().getDataRange().getValues().slice(1);
  return out_(v.map(r => ({waktu: r[0], nama: r[1], kelas: r[2], V: r[3], A: r[4], K: r[5], dominan: r[6]})));
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

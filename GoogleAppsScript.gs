/**
 * 組員出勤與操作紀錄 — Google Apps Script 後端
 * 用法：
 *  1. 開一份「新的」Google 試算表（建議跟監看板的分開，因為是考核資料）
 *  2. 擴充功能 → Apps Script，把這份程式碼整份貼上取代
 *  3. 把下面的 PASSWORD 改成你自己的密碼
 *  4. 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *     執行身分：我　／　誰可以存取：所有人 → 部署，複製網址
 *  5. 打開網頁，貼上網址與密碼即可
 */
const PASSWORD = '請改成你的密碼';
const SHEET_NAME = '組員紀錄';
const HEADERS = ['日期', '組員', '班別', '出勤狀況', '機台', '操作狀況', '備註', '記錄時間'];
const BACKEND_VERSION = 'crew-2026-11-11-v1';

function doGet(e){ return handle_((e && e.parameter) || {}); }
function doPost(e){
  let body = {};
  try { body = JSON.parse(e.postData.contents); } catch(err){ return out_({ok:false, error:'資料格式錯誤'}); }
  return handle_(body);
}

function handle_(p){
  if(String(p.pw || '') !== PASSWORD) return out_({ok:false, error:'密碼不正確'});
  try{
    const sh = sheet_();
    switch(p.action){
      case 'load': return out_({ok:true, version:BACKEND_VERSION, rows:loadRows_(sh)});
      case 'add': {
        const d = clean_(p.data);
        sh.appendRow(HEADERS.map(h => h === '記錄時間' ? now_() : d[h]));
        fixText_(sh, sh.getLastRow());
        return out_({ok:true, rows:loadRows_(sh)});
      }
      case 'update': {
        const row = Number(p.row);
        if(!(row >= 2 && row <= sh.getLastRow())) return out_({ok:false, error:'找不到這筆資料'});
        const d = clean_(p.data);
        const old = sh.getRange(row, 8).getValue();
        sh.getRange(row, 1, 1, 8).setValues([HEADERS.map(h => h === '記錄時間' ? (old || now_()) : d[h])]);
        fixText_(sh, row);
        return out_({ok:true, rows:loadRows_(sh)});
      }
      case 'delete': {
        const row = Number(p.row);
        if(!(row >= 2 && row <= sh.getLastRow())) return out_({ok:false, error:'找不到這筆資料'});
        sh.deleteRow(row);
        return out_({ok:true, rows:loadRows_(sh)});
      }
      default: return out_({ok:false, error:'未知的動作'});
    }
  }catch(err){ return out_({ok:false, error:String(err && err.message || err)}); }
}

function clean_(d){
  d = d || {};
  const o = {};
  HEADERS.forEach(h => { o[h] = String(d[h] == null ? '' : d[h]).trim(); });
  if(!/^\d{4}-\d{2}-\d{2}$/.test(o['日期'])) throw new Error('日期格式不正確');
  if(!o['組員']) throw new Error('請選擇組員');
  return o;
}

function sheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if(!sh){
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

// 日期、時間一律存成文字，避免試算表自動轉格式
function fixText_(sh, row){ sh.getRange(row, 1, 1, 8).setNumberFormat('@'); }
function now_(){ return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm'); }

function loadRows_(sh){
  const last = sh.getLastRow();
  if(last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues();
  return vals.map((r, i) => {
    const o = {_row: i + 2};
    HEADERS.forEach((h, k) => { o[h] = r[k]; });
    return o;
  });
}

function out_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

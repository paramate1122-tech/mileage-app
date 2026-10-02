// Mileage app — shared project list API (Google Apps Script)
// วางโค้ดนี้ใน Google Sheet: Extensions → Apps Script แล้ว Deploy → New deployment → Web app
//   Execute as: Me   ·   Who has access: Anyone
// แล้วเอา URL ที่ลงท้ายด้วย /exec ไปใส่ใน SHEET_API_URL ของ index.html
//
// แถวที่ 1 = หัวตาราง: Project number | Project name | Project Manager | Nickname
// (หาคอลัมน์จากชื่อหัวตาราง ไม่เจอใช้ A–D ตามลำดับ)

const SHEET_NAME = ''; // ว่าง = แผ่นแรก

function sheet_(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return (SHEET_NAME && ss.getSheetByName(SHEET_NAME)) || ss.getSheets()[0];
}
function norm_(v){ return String(v == null ? '' : v).toUpperCase().trim(); }
function key_(v){ return String(v == null ? '' : v).toLowerCase().replace(/[^a-z0-9฀-๿]/g, ''); }

function columns_(sh){
  if(sh.getLastRow() === 0){
    sh.appendRow(['Project number', 'Project name', 'Project Manager', 'Nickname']);
  }
  const width = Math.max(4, sh.getLastColumn());
  const h = sh.getRange(1, 1, 1, width).getValues()[0].map(key_);
  const find = keys => h.findIndex(v => keys.indexOf(v) >= 0);
  const col = {
    number: find(['projectnumber','projectno','projno','projectcode','number','no','เลขโปรเจกต์','เลขโปรเจค']),
    name: find(['projectname','name','ชื่อโปรเจกต์','ชื่อโปรเจค']),
    manager: find(['projectmanager','manager','pm']),
    nickname: find(['nickname','ชื่อเล่น']),
  };
  if(col.number < 0) return { number:0, name:1, manager:2, nickname:3, width:width };
  col.width = width;
  return col;
}
function rows_(sh, col){
  const n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, col.width).getValues() : [];
}
function list_(){
  const sh = sheet_(), col = columns_(sh);
  const cell = (r, k) => col[k] >= 0 ? String(r[col[k]] == null ? '' : r[col[k]]).trim() : '';
  return rows_(sh, col).map(r => ({
    number: norm_(cell(r, 'number')), name: cell(r, 'name'), manager: cell(r, 'manager'), nickname: cell(r, 'nickname'),
  })).filter(p => p.number);
}
function findRow_(sh, col, number){
  const rows = rows_(sh, col);
  for(let i = 0; i < rows.length; i++) if(norm_(rows[i][col.number]) === number) return i + 2;
  return -1;
}
function upsert_(p){
  const sh = sheet_(), col = columns_(sh);
  const number = norm_(p.number);
  if(!number) throw new Error('Missing project number');
  let row = p.oldNumber ? findRow_(sh, col, norm_(p.oldNumber)) : -1;
  if(row < 0) row = findRow_(sh, col, number);
  const values = row > 0 ? sh.getRange(row, 1, 1, col.width).getValues()[0] : new Array(col.width).fill('');
  values[col.number] = number;
  if(col.name >= 0) values[col.name] = p.name || '';
  if(col.manager >= 0) values[col.manager] = p.manager || '';
  if(col.nickname >= 0) values[col.nickname] = p.nickname || '';
  if(row > 0) sh.getRange(row, 1, 1, col.width).setValues([values]);
  else sh.appendRow(values);
  // ถ้าเปลี่ยนเลขแล้วมีแถวเลขใหม่ซ้ำอยู่อีกแถว ให้ลบแถวซ้ำ
  if(p.oldNumber && norm_(p.oldNumber) !== number){
    const rows = rows_(sh, col);
    let seen = false;
    for(let i = rows.length - 1; i >= 0; i--){
      if(norm_(rows[i][col.number]) !== number) continue;
      if(seen) sh.deleteRow(i + 2); else seen = true;
    }
  }
}
function remove_(number){
  const sh = sheet_(), col = columns_(sh);
  const rows = rows_(sh, col);
  for(let i = rows.length - 1; i >= 0; i--) if(norm_(rows[i][col.number]) === norm_(number)) sh.deleteRow(i + 2);
}
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function doGet(){
  try{ return json_({ ok:true, projects:list_() }); }
  catch(err){ return json_({ ok:false, error:String(err.message || err) }); }
}
function doPost(e){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(15000);
    const req = JSON.parse(e.postData.contents);
    if(req.action === 'upsert') upsert_(req);
    else if(req.action === 'delete') remove_(req.number);
    else throw new Error('Unknown action');
    return json_({ ok:true });
  }catch(err){
    return json_({ ok:false, error:String(err.message || err) });
  }finally{
    try{ lock.releaseLock(); }catch(_){}
  }
}

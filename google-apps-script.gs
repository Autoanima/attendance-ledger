/**
 * 跨群選修點名｜Google Apps Script
 *
 * 用途：讓 index.html 把點名結果寫進指定的 Google 試算表。
 * Web App 以「執行身分：我」部署後，網頁把試算表網址與點名資料 POST 過來。
 *
 * 試算表格式：
 * A 欄：班級
 * B 欄：座號
 * C 欄：姓名
 * D 欄開始：每次點名一個新欄，第一列寫「日期｜第X節」
 * 例如：D1 = 2026/09/06｜第3節
 *       D2:Dn = 出席／未出席
 */

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  if (p.action === 'test') {
    const callback = String(p.callback || '').replace(/[^a-zA-Z0-9_$.]/g, '');
    let result;
    try {
      const sheetUrl = String(p.sheetUrl || '').trim();
      if (!sheetUrl) throw new Error('請先填入 Google 試算表網址');
      const ss = SpreadsheetApp.openByUrl(sheetUrl);
      const sheet = ss.getSheets()[0];
      result = {ok:true, message:'連線成功', sheetName:sheet.getName()};
    } catch (err) {
      result = {ok:false, message:String(err && err.message ? err.message : err)};
    }
    const text = JSON.stringify(result);
    if (callback) return ContentService.createTextOutput(callback + '(' + text + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return jsonOutput(result);
  }
  return ContentService
    .createTextOutput('跨群選修點名 Google 紀錄服務正常運作')
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput({ok:false, message:'沒有收到資料'});
    }

    const data = JSON.parse(e.postData.contents);
    const sheetUrl = String(data.sheetUrl || '').trim();
    if (!sheetUrl) return jsonOutput({ok:false, message:'缺少 Google 試算表網址'});

    const ss = SpreadsheetApp.openByUrl(sheetUrl);
    const sheet = ss.getSheets()[0];

    const rows = Array.isArray(data.rows) ? data.rows : [];
    if (!rows.length) return jsonOutput({ok:false, message:'沒有學生資料'});

    const date = String(data.date || '').trim() || Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd');
    const period = String(data.period || '').trim() || '非上課時段';
    const periodLabel = period === '非上課時段' ? period : '第' + period + '節';
    const header = date + '｜' + periodLabel;

    // 確保前三欄標題存在。
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, 3).setValues([['班級', '座號', '姓名']]);
    } else {
      const first = sheet.getRange(1, 1, 1, 3).getValues()[0];
      if (!first[0] && !first[1] && !first[2]) {
        sheet.getRange(1, 1, 1, 3).setValues([['班級', '座號', '姓名']]);
      }
    }

    // 每次點名都新增一欄，不覆蓋既有紀錄。
    const newCol = Math.max(sheet.getLastColumn() + 1, 4);
    sheet.getRange(1, newCol).setValue(header);

    // 讀取既有學生資料，依「班級 + 座號」找學生所在列。
    const lastRow = Math.max(sheet.getLastRow(), 2);
    const existing = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, 3).getValues() : [];
    const rowMap = new Map();
    existing.forEach((r, i) => {
      const key = makeKey(r[0], r[1]);
      if (key) rowMap.set(key, i + 2);
    });

    // 若新名單出現原本沒有的學生，就加到最下面。
    const newStudents = [];
    rows.forEach(item => {
      const key = makeKey(item.className, item.seat);
      if (key && !rowMap.has(key)) {
        newStudents.push([
          String(item.className || ''),
          String(item.seat || ''),
          String(item.name || '')
        ]);
      }
    });

    if (newStudents.length) {
      const startRow = Math.max(sheet.getLastRow() + 1, 2);
      sheet.getRange(startRow, 1, newStudents.length, 3).setValues(newStudents);
      newStudents.forEach((r, i) => {
        const key = makeKey(r[0], r[1]);
        rowMap.set(key, startRow + i);
      });
    }

    // 寫入本次點名狀態。
    rows.forEach(item => {
      const key = makeKey(item.className, item.seat);
      if (!key) return;
      const row = rowMap.get(key);
      if (!row) return;
      sheet.getRange(row, newCol).setValue(String(item.status || ''));
    });

    // 簡單格式化，讓期末查看比較清楚。
    sheet.getRange(1, 1, 1, newCol).setFontWeight('bold');
    sheet.getRange(1, newCol).setWrap(true);
    sheet.setFrozenRows(1);
    sheet.setFrozenColumns(3);
    sheet.autoResizeColumns(1, 3);

    return jsonOutput({ok:true, message:'已新增一欄點名紀錄', column:newCol, header:header});
  } catch (err) {
    return jsonOutput({ok:false, message:String(err && err.message ? err.message : err)});
  } finally {
    lock.releaseLock();
  }
}

function makeKey(cls, seat) {
  const c = String(cls || '').trim();
  const s = String(seat || '').trim();
  return c && s ? c + '|' + s : '';
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

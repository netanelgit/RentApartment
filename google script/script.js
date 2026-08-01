function doGet() {
  // ========= CONFIG =========
  const SPREADSHEET_ID = '18LLbthvm3P3s-C3Yzl9oaaJrqCab21M-gy9vyBMQwsA';
  const SHEET_NAME = 'חישוב תשלומים';
  const HEADERS_WANTED = [
    'שנה',
    'חודש',
    'מים צריכה(קוב)',
    'מים תשלום',
    'ביוב תשלום',
    'חשמל צריכה(קוט"ש)',
    'חשמל תשלום',
    'חיוב קבוע - מונה חשמל',
    'ארנונה',
    'שכירות',
    'סה"כ מח"א',
    'סה"כ',
  ];
  // ==========================

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) throw new Error('Spreadsheet not found by ID');

    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error('לא נמצא גיליון בשם: ' + SHEET_NAME);

    const lastRow = sh.getLastRow();
    const lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) throw new Error('הגיליון ריק');

    const headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0] || [];
    const data = lastRow > 1
      ? sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues()
      : [];

    const idxByHeader = Object.fromEntries(headers.map((h, i) => [String(h).trim(), i]));
    const missing = HEADERS_WANTED.filter(h => !(h in idxByHeader));
    if (missing.length) {
      throw new Error('כותרות חסרות בגיליון: ' + missing.join(', '));
    }

    // בונה טבלה מסוננת
    const filtered = [HEADERS_WANTED];
    for (const row of data) {
      const monthVal = row[idxByHeader['חודש']] ?? '';
      if (!String(monthVal).trim()) continue;
      filtered.push(HEADERS_WANTED.map(h => row[idxByHeader[h]] ?? ''));
    }

    // תעריפים
    let tariffs = [['תעריפים'], ['(TariffRange לא מוגדר)']];
    const tariffRange = ss.getRangeByName('TariffRange');
    if (tariffRange) {
      tariffs = tariffRange.getDisplayValues();
    }

    // טעינת חתימות (מאפס חודש -> חתימה)
    const signatures = getSignatures(ss, filtered);

    const tpl = HtmlService.createTemplateFromFile('index');
    tpl.payments = filtered;
    tpl.tariffs = tariffs;
    tpl.signatures = signatures;

    return tpl.evaluate()
      .setTitle('סיכום תשלומים')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  } catch (err) {
    const html = HtmlService.createHtmlOutput(
      '<pre style="direction:rtl;font-family:system-ui,monospace;white-space:pre-wrap">' +
      Utilities.formatString('שגיאה: %s', err && err.message ? err.message : err) +
      '</pre>'
    );
    return html;
  }
}

// פונקציה לטעינת חתימות קיימות
function getSignatures(ss, payments) {
  const sigSheet = ss.getSheetByName('Signatures');
  const result = {};
  
  if (!sigSheet) return result;
  
  const sigData = sigSheet.getDataRange().getValues();
  // שורה 0 היא כותרות, מתחיל מ-1
  for (let i = 1; i < sigData.length; i++) {
    const [month, signatureData, timestamp] = sigData[i];
    if (month && signatureData) {
      result[month] = {
        data: signatureData,
        date: timestamp ? Utilities.formatDate(new Date(timestamp), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') : ''
      };
    }
  }
  
  return result;
}

// פונקציה לשמירת חתימה לחודש ספציפי
function saveSignature(month, signatureData) {
  try {
    const ss = SpreadsheetApp.openById('18LLbthvm3P3s-C3Yzl9oaaJrqCab21M-gy9vyBMQwsA');
    let sigSheet = ss.getSheetByName('Signatures');
    
    // יצירת גיליון אם לא קיים
    if (!sigSheet) {
      sigSheet = ss.insertSheet('Signatures');
      sigSheet.getRange('A1:C1').setValues([['חודש', 'חתימה', 'תאריך']]);
      sigSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#3498db').setFontColor('#ffffff');
      sigSheet.setColumnWidth(1, 120);
      sigSheet.setColumnWidth(2, 400);
      sigSheet.setColumnWidth(3, 150);
    }
    
    // בדיקה אם החודש כבר קיים
    const data = sigSheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === month) {
        rowIndex = i + 1; // שורות מתחילות מ-1 בממשק
        break;
      }
    }
    
    const timestamp = new Date();
    
    if (rowIndex > 0) {
      // עדכון חתימה קיימת
      sigSheet.getRange(rowIndex, 2, 1, 2).setValues([[signatureData, timestamp]]);
    } else {
      // הוספת חתימה חדשה
      sigSheet.appendRow([month, signatureData, timestamp]);
    }
    
    return {
      success: true,
      date: Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
    };
    } catch (err) {
      const html = HtmlService.createHtmlOutput(
      '<pre style="direction:rtl;font-family:system-ui,monospace;white-space:pre-wrap">' +
      'שגיאה: ' + (err && err.message ? err.message : err) + '\n\n' +
      'Stack: ' + (err && err.stack ? err.stack : '(אין stack)') +
      '</pre>'
    );
      return html;
  }
}
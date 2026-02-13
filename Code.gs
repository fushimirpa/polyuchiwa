const CONFIG_SHEET_NAME = 'Config';
const ORDERS_SHEET_NAME = 'Orders';
const UPLOAD_FOLDER_ID = '1Hbh4q5EoyOR73KfbjWlD9p5hJiSGrGqH'; 

function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  try {
    const initialData = getDataForClient(50);
    template.jsonPayload = JSON.stringify({ success: true, data: initialData });
  } catch (e) {
    template.jsonPayload = JSON.stringify({ success: false, error: e.toString() });
  }
  return template.evaluate()
    .setTitle('うちわ管理システム')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDataForClient(limit = 0, keyword = '', status = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let orderSheet = ss.getSheetByName(ORDERS_SHEET_NAME);
  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!orderSheet) {
    orderSheet = ss.insertSheet(ORDERS_SHEET_NAME);
    orderSheet.appendRow(['ID', '作成日時', '更新日時', '基本:ステータス', '基本:顧客名']); 
  }

  const configValues = configSheet ? configSheet.getDataRange().getValues() : [];
  const options = {};
  if (configValues.length > 0) {
    const configHeader = configValues.shift();
    configHeader.forEach((colName, index) => {
      options[colName] = configValues.map(row => row[index]).filter(val => val !== "");
    });
  }

  const allData = orderSheet.getDataRange().getValues();
  const headers = allData.shift(); 

  let orders = [];
  if (allData.length > 0) {
    orders = allData.map(row => {
      let obj = {};
      headers.forEach((key, index) => {
        if (row[index] instanceof Date) {
          obj[key] = Utilities.formatDate(row[index], Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          obj[key] = row[index];
        }
      });
      return obj;
    });
    orders.reverse();
  }

  if (keyword || status) {
    const lowerKey = keyword.toLowerCase();
    orders = orders.filter(order => {
      const searchTarget = Object.values(order).join(' ').toLowerCase();
      const statusKey = Object.keys(order).find(k => k.includes('ステータス')) || '';
      return (!keyword || searchTarget.includes(lowerKey)) &&
             (!status || order[statusKey] === status);
    });
  }
  if (limit > 0 && orders.length > limit) orders = orders.slice(0, limit);
  return { options: options, orders: orders, headers: headers };
}

function searchOrders(keyword, status) {
  return getDataForClient(0, keyword, status);
}

function saveOrder(formData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: '混雑しています' };
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ORDERS_SHEET_NAME);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const folder = UPLOAD_FOLDER_ID ? DriveApp.getFolderById(UPLOAD_FOLDER_ID) : DriveApp.getRootFolder();
    
    Object.keys(formData).forEach(key => {
      if (key.includes('画像') && formData[key] && formData[key].data) {
         const blob = Utilities.newBlob(
           Utilities.base64Decode(formData[key].data.split(',')[1]),
           formData[key].mimeType,
           formData[key].name
         );
         const file = folder.createFile(blob);
         file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         formData[key] = file.getUrl();
      }
    });

    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm:ss');
    let targetRow;
    if (!formData['ID']) {
      formData['ID'] = Utilities.getUuid();
      formData['作成日時'] = now;
      targetRow = sheet.getLastRow() + 1;
    } else {
      const ids = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow() - 1), 1).getValues().flat();
      const index = ids.indexOf(formData['ID']);
      targetRow = (index === -1) ? sheet.getLastRow() + 1 : index + 2;
    }
    formData['更新日時'] = now;

    const values = headers.map(h => (formData[h] !== undefined ? formData[h] : ''));
    sheet.getRange(targetRow, 1, 1, values.length).setValues([values]);
    return { success: true, message: '保存しました', id: formData['ID'] };
  } catch (e) {
    return { success: false, message: 'エラー: ' + e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteOrder(id) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: '混雑しています' };
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ORDERS_SHEET_NAME);
    const ids = sheet.getRange(2, 1, Math.max(1, sheet.getLastRow()-1), 1).getValues().flat();
    const index = ids.indexOf(id);
    if (index !== -1) { sheet.deleteRow(index + 2); return { success: true, message: '削除しました' }; }
    return { success: false, message: 'データが見つかりません' };
  } catch (e) { return { success: false, message: 'エラー: ' + e.toString() };
  } finally { lock.releaseLock(); }
}
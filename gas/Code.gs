const PRODUCT_LIST = [
  {id:'p01',name:'金牌紅燒牛肉麵',price:420},
  {id:'p02',name:'金牌紅燒半筋半肉牛肉麵',price:420},
  {id:'p03',name:'金牌麻辣半筋半肉牛肉麵',price:420},
  {id:'p04',name:'金牌清燉牛肉麵',price:420},
  {id:'p05',name:'金牌清燉半筋半肉牛肉麵',price:420},
  {id:'p06',name:'金牌蕃茄牛肉麵',price:420},
  {id:'p07',name:'金牌蕃茄半筋半肉牛肉麵',price:420},
  {id:'p08',name:'金牌紅燒排骨麵',price:420},
  {id:'p09',name:'金牌番茄排骨麵',price:420},
  {id:'p10',name:'經典胡椒豬肚雞麵',price:420},
  {id:'p11',name:'藥燉排骨麵線',price:420},
  {id:'p12',name:'酸白菜豬肉寬粉',price:420},
  {id:'p13',name:'豚骨拉麵',price:470},
  {id:'p14',name:'辛味噌豚骨拉麵',price:470},
  {id:'p15',name:'麻辣鴨血寬粉',price:320},
  {id:'p16',name:'椒燥味噌乾拌麵',price:190},
  {id:'p17',name:'和牛肉燥乾拌麵',price:190},
  {id:'p18',name:'珍珠鮑干貝乾拌麵',price:120},
  {id:'p19',name:'香菇雞湯',price:155},
  {id:'p20',name:'剝皮辣椒雞湯',price:155},
  {id:'p21',name:'紅豆紫米粥',price:69},
  {id:'p22',name:'鮑片雞蓉玉米粥',price:130},
  {id:'p23',name:'吻仔魚海帶蘿蔔粥',price:130},
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getOrCreateSpreadsheet();
    const sheet = ss.getSheets()[0];
    const params = e.parameter;
    const name = params.name || '';
    const summary = params.summary || '';
    const total = parseInt(params.total || '0', 10);
    const note = params.note || '';
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');

    // Build row
    const row = [timestamp, name];
    const itemLines = [];
    PRODUCT_LIST.forEach(p => {
      const q = parseInt(params[p.id] || '0', 10);
      row.push(q > 0 ? q : '');
      if (q > 0) itemLines.push(`${p.name} ${q}盒 單價${p.price}元`);
    });

    // 訂購摘要
    row.push(summary);
    // 總金額
    row.push(total);
    // 備註及回覆
    const replyMsg = `${name}你好，你所訂購的產品如下：${itemLines.join('、')}、總金額：${total}元${note ? '\n備註：' + note : ''}`;
    row.push(replyMsg);

    const lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1, 1, row.length).setValues([row]);

    // Format: center qty columns
    const qtyRange = sheet.getRange(lastRow, 3, 1, PRODUCT_LIST.length);
    qtyRange.setHorizontalAlignment('center');

    // Wrap summary & reply columns
    sheet.getRange(lastRow, 3 + PRODUCT_LIST.length, 1, 1).setWrap(true);
    sheet.getRange(lastRow, 3 + PRODUCT_LIST.length + 2, 1, 1).setWrap(true);

    // Alternating row color
    const bgColor = lastRow % 2 === 0 ? '#e8f0f5' : '#ffffff';
    sheet.getRange(lastRow, 1, 1, row.length).setBackground(bgColor);

    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({status:'ok'}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    lock.releaseLock();
    return ContentService.createTextOutput(JSON.stringify({status:'error',message:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');

  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) { /* recreate */ }
  }

  const ss = SpreadsheetApp.create('藍海饌雙11團購訂單');
  ssId = ss.getId();
  props.setProperty('SPREADSHEET_ID', ssId);
  setupHeaders(ss.getSheets()[0]);
  return ss;
}

function setupHeaders(sheet) {
  const headers = ['時間戳記', '姓名'];
  PRODUCT_LIST.forEach(p => headers.push(`${p.name}\n$${p.price}`));
  headers.push('訂購摘要', '總金額（元）', '備註及回覆');

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a3a5c')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);

  sheet.setRowHeight(1, 48);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);

  // Column widths
  sheet.setColumnWidth(1, 155); // 時間
  sheet.setColumnWidth(2, 75);  // 姓名
  for (let i = 0; i < PRODUCT_LIST.length; i++) {
    sheet.setColumnWidth(3 + i, 78);
  }
  const afterProducts = 3 + PRODUCT_LIST.length;
  sheet.setColumnWidth(afterProducts, 220);     // 訂購摘要
  sheet.setColumnWidth(afterProducts + 1, 90);  // 總金額
  sheet.setColumnWidth(afterProducts + 2, 160); // 備註及回覆
}

function getSpreadsheetUrl() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    Logger.log('試算表網址：https://docs.google.com/spreadsheets/d/' + ssId);
  } else {
    Logger.log('尚未建立試算表，請先送出第一筆訂單。');
  }
}

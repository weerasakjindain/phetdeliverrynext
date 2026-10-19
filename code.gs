// ==========================================
// 1. การตั้งค่าระบบและการเรนเดอร์หน้าเว็บ
// ==========================================
function doGet(e) {
  // เรนเดอร์ไฟล์ Index.html
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PHET Delivery+')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 2. ฟังก์ชันตรวจสอบและสร้างฐานข้อมูลอัตโนมัติ
// ==========================================
function initializeSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. สร้าง/ตรวจสอบแท็บ Menus
  let menuSheet = ss.getSheetByName('Menus');
  if (!menuSheet) {
    menuSheet = ss.insertSheet('Menus');
    menuSheet.appendRow(['Menu_ID', 'Title', 'URL', 'Icon_Name', 'Clicks', 'Status']);
    menuSheet.getRange("A1:F1").setFontWeight("bold").setBackground("#d2e3fc");
    
    // ข้อมูลเริ่มต้น 6 เมนู
    const initialMenus = [
      [1, 'PHET CDD AI "ผู้ช่วยพัฒนากรอัจฉริยะ"', 'https://cddata.cdd.go.th/cddai/chat.html', 'fa-robot', 0, 'Active'],
      [2, 'PHET Inspection "ตรวจราชการ"', 'https://script.google.com/macros/s/AKfycbwL-ZhIgdjRo90JUEDPm0FXd_KjwvXIkvFLDzqFgtdPXgkNcn6MH51LkTM1fZoX95a1/exec', 'fa-magnifying-glass', 0, 'Active'],
      [3, 'PHET Knowledge "คลังความรู้"', 'https://script.google.com/macros/s/AKfycbwLqwZi15zOFiWEHbQfoLXY-gSLUBwgkUq-gaf2S8yQrBIWNphb8DcFFLLe1POBq7i9-g/exec', 'fa-book-open', 0, 'Active'],
      [4, 'PHET OTOP Market "ตลาดชุมชน"', 'https://script.google.com/macros/s/AKfycbz8mPfkWSDu5SK3sfDSzQ7Q3SCYnR9-l9PWzC_GWDoXzfnZ6H9_iPoH2eQl-Ycv9PuN7Q/exec', 'fa-store', 0, 'Active'],
      [5, 'PHET Personal "สำหรับเจ้าหน้าที่"', 'https://script.google.com/macros/s/AKfycbyk_2ZmlersqQ0-ye8lVbtntd9wbpMUYmFOiSKfKdXNb8Qoi4PLCYqwyLfM2pp4essS/exec', 'fa-users', 0, 'Active'],
      [6, 'PHET Service "รวมบริการจาก ศพช."', 'https://script.google.com/macros/s/AKfycbzpPOJxpSPlaD8blZPGKIEcNDeDTpwCO5ar315lHMujUDQsZD2P3_WLmjHwxXgFwkIGMg/exec', 'fa-box-archive', 0, 'Active']
    ];
    menuSheet.getRange(2, 1, initialMenus.length, 6).setValues(initialMenus);
  }

  // 2. สร้าง/ตรวจสอบแท็บ Visits
  let visitSheet = ss.getSheetByName('Visits');
  if (!visitSheet) {
    visitSheet = ss.insertSheet('Visits');
    visitSheet.appendRow(['Date', 'Visit_Count']);
    visitSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#d2e3fc");
  }

  // 3. สร้าง/ตรวจสอบแท็บ Settings
  let settingSheet = ss.getSheetByName('Settings');
  if (!settingSheet) {
    settingSheet = ss.insertSheet('Settings');
    settingSheet.appendRow(['Setting_Name', 'Value']);
    settingSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#d2e3fc");
    
    const initialSettings = [
      ['Admin_Password', 'admin1234'],
      ['System_Name', 'PHET Delivery+'],
      ['Maintenance_Mode', 'FALSE']
    ];
    settingSheet.getRange(2, 1, initialSettings.length, 2).setValues(initialSettings);
  }
}

// ==========================================
// 3. API สำหรับผู้ใช้งานทั่วไป (Frontend)
// ==========================================

// ดึงข้อมูลสำหรับหน้าแรก (เมนู และสถิติวันนี้)
function getAppData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Menus')) {
    initializeSystem(); // สร้างฐานข้อมูลอัตโนมัติถ้ายังไม่มี
  }

  const menuSheet = ss.getSheetByName('Menus');
  const menuData = menuSheet.getDataRange().getValues();
  const menus = [];
  
  // แปลงข้อมูลเมนู
  for (let i = 1; i < menuData.length; i++) {
    if (menuData[i][5] === 'Active') {
      menus.push({
        id: menuData[i][0],
        title: menuData[i][1],
        url: menuData[i][2],
        icon: menuData[i][3],
        clicks: menuData[i][4]
      });
    }
  }

  // บันทึกและดึงยอดเข้าชมวันนี้
  const todayVisits = recordDailyVisit();

  return {
    menus: menus,
    todayVisits: todayVisits
  };
}

// นับยอดเข้าชมรายวัน
function recordDailyVisit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const visitSheet = ss.getSheetByName('Visits');
  
  // สร้างวันที่ปัจจุบันรูปแบบ DD/MM/YYYY
  const today = new Date();
  const dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  const visitData = visitSheet.getDataRange().getValues();
  let found = false;
  let currentCount = 0;
  
  for (let i = 1; i < visitData.length; i++) {
    let rowDate = visitData[i][0];
    if (rowDate instanceof Date) {
      rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    } else {
      rowDate = String(rowDate);
    }

    if (rowDate === dateString) {
      currentCount = parseInt(visitData[i][1]) || 0;
      currentCount += 1;
      visitSheet.getRange(i + 1, 2).setValue(currentCount); // อัปเดตยอด
      found = true;
      break;
    }
  }
  
  if (!found) {
    currentCount = 1;
    visitSheet.appendRow([dateString, currentCount]); // เพิ่มวันใหม่
  }
  
  return currentCount;
}

// บันทึกยอดคลิกเมนู
function recordMenuClick(menuId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Menus');
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == menuId) {
      let clicks = parseInt(data[i][4]) || 0;
      sheet.getRange(i + 1, 5).setValue(clicks + 1);
      return true;
    }
  }
  return false;
}

// ==========================================
// 4. API สำหรับผู้ดูแลระบบ (Admin)
// ==========================================

function verifyAdmin(password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'Admin_Password') {
      return data[i][1].toString() === password;
    }
  }
  return false;
}

function getAdminDashboard(password) {
  if (!verifyAdmin(password)) throw new Error("Unauthorized");
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get Menus
  const menuData = ss.getSheetByName('Menus').getDataRange().getValues();
  const menus = [];
  for (let i = 1; i < menuData.length; i++) {
    menus.push({
      id: menuData[i][0],
      title: menuData[i][1],
      url: menuData[i][2],
      icon: menuData[i][3],
      clicks: menuData[i][4],
      status: menuData[i][5]
    });
  }
  
  // Get Visits
  const visitData = ss.getSheetByName('Visits').getDataRange().getValues();
  const visits = [];
  for (let i = 1; i < visitData.length; i++) {
    let rowDate = visitData[i][0];
    if (rowDate instanceof Date) rowDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
    visits.push({
      row: i + 1,
      date: rowDate,
      count: visitData[i][1]
    });
  }
  
  return {
    menus: menus,
    visits: visits.reverse(), // ล่าสุดขึ้นก่อน
    sheetUrl: ss.getUrl()
  };
}

// บันทึกการเพิ่ม/แก้ไข เมนู (CRUD)
function saveMenuAction(password, action, menuObj) {
  if (!verifyAdmin(password)) throw new Error("Unauthorized");
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Menus');
  
  if (action === 'ADD') {
    const newId = new Date().getTime(); // Gen ID แบบง่าย
    sheet.appendRow([newId, menuObj.title, menuObj.url, menuObj.icon, parseInt(menuObj.clicks) || 0, menuObj.status]);
  } 
  else if (action === 'EDIT') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == menuObj.id) {
        sheet.getRange(i + 1, 2).setValue(menuObj.title);
        sheet.getRange(i + 1, 3).setValue(menuObj.url);
        sheet.getRange(i + 1, 4).setValue(menuObj.icon);
        sheet.getRange(i + 1, 5).setValue(parseInt(menuObj.clicks));
        sheet.getRange(i + 1, 6).setValue(menuObj.status);
        break;
      }
    }
  }
  else if (action === 'DELETE') {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == menuObj.id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
  }
  return true;
}

// อัปเดตสถิติรายวันจาก Dashboard
function updateVisitStat(password, row, newCount) {
  if (!verifyAdmin(password)) throw new Error("Unauthorized");
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visits');
  sheet.getRange(row, 2).setValue(parseInt(newCount));
  return true;
}

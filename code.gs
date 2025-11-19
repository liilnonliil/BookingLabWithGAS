// Code.gs - Google Apps Script Backend

// ชื่อ Spreadsheet sheets
const SHEET_NAMES = {
  BOOKINGS: 'Bookings',
  ROOMS: 'Rooms',
  USERS: 'Users',
  CONFIG: 'Config',
  BLOCKED_DATES: 'BlockedDates'
};

// เปิดหน้า Web App
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบจองห้องและเครื่องคอม')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ดึงข้อมูลผู้ใช้ที่ login
function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  return {
    email: email,
    name: email.split('@')[0]
  };
}
// ดึงข้อมูล Config
function getConfig() {
  // ตรวจสอบสิทธิ์ Admin
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    sheet.getRange('A1:B1').setValues([['Key', 'Value']]);
    
    // ค่า default
    const defaultConfig = [
      ['bookingStartDate', ''],
      ['bookingEndDate', ''],
      ['emailEnabled', 'true'],
      ['adminEmail', Session.getActiveUser().getEmail()],
      ['adminUsers', Session.getActiveUser().getEmail()]
    ];
    sheet.getRange(2, 1, defaultConfig.length, 2).setValues(defaultConfig);
  }
  
  const data = sheet.getDataRange().getValues();
  const config = {};
  
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  
  return JSON.stringify(config);
}

// บันทึก Config
function saveConfig(configData) {
  // ตรวจสอบสิทธิ์ Admin
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    sheet.getRange('A1:B1').setValues([['Key', 'Value']]);
  }
  
  const data = sheet.getDataRange().getValues();
  
  // Update existing or append new
  for (let key in configData) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(configData[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, configData[key]]);
    }
  }
  
  return { success: true };
}

// ดึงวันที่ปิดจอง
function getBlockedDates() {
  // ตรวจสอบสิทธิ์ Admin
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BLOCKED_DATES);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.BLOCKED_DATES);
    sheet.getRange('A1:E1').setValues([['Date', 'Room', 'Computer', 'Reason', 'ReasonEN']]);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const blockedDates = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) {
      blockedDates.push({
        date: Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        room: String(data[i][1] || ''),
        computer: String(data[i][2] || ''),
        reason: String(data[i][3] || ''),
        reasonEN: String(data[i][4] || '')
      });
    }
  }
  
  return blockedDates;
}

// เพิ่มวันปิดจอง
function addBlockedDate(dateStr, room, computer, reason, reasonEN) {
  // ตรวจสอบสิทธิ์ Admin
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BLOCKED_DATES);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.BLOCKED_DATES);
    sheet.getRange('A1:E1').setValues([['Date', 'Room', 'Computer', 'Reason', 'ReasonEN']]);
  }
  
  sheet.appendRow([new Date(dateStr), room, computer, reason, reasonEN]);
  return { success: true };
}

// ลบวันปิดจอง
function removeBlockedDate(dateStr, room, computer) {
  // ตรวจสอบสิทธิ์ Admin
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BLOCKED_DATES);
  
  if (!sheet) return { success: false };
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    const rowDate = Utilities.formatDate(new Date(data[i][0]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (rowDate === dateStr && String(data[i][1]) === room && String(data[i][2]) === computer) {
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  
  return { success: false };
}

// ✅ ฟังก์ชันตรวจสอบสิทธิ์ Admin (รองรับหลาย Admin)
function isAdmin() {
  const userEmail = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!configSheet) {
    // ถ้ายังไม่มี config sheet ให้ user แรก (เจ้าของ sheet) เป็น admin
    return true;
  }
  
  const data = configSheet.getDataRange().getValues();
  
  // หา adminUsers
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'adminUsers') {
      const adminList = String(data[i][1]).split(',').map(e => e.trim());
      return adminList.includes(userEmail);
    }
  }
  
  // ถ้าไม่มี adminUsers ให้ user แรกเป็น admin
  return true;
}

// ✅ ดึงรายชื่อ Admin ทั้งหมด
function getAdminUsers() {
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!configSheet) {
    return [];
  }
  
  const data = configSheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'adminUsers') {
      const adminList = String(data[i][1]).split(',').map(e => e.trim());
      return adminList.filter(e => e); // ลบค่าว่าง
    }
  }
  
  return [];
}

// ✅ ฟังก์ชันตรวจสอบสิทธิ์ (ส่งกลับให้ frontend)
function checkAdminAccess() {
  return {
    isAdmin: isAdmin(),
    email: Session.getActiveUser().getEmail()
  };
}


// ✅ เพิ่ม Admin ใหม่
function addAdminUser(email) {
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!configSheet) {
    configSheet = ss.insertSheet(SHEET_NAMES.CONFIG);
    configSheet.getRange('A1:B1').setValues([['Key', 'Value']]);
  }
  
  const data = configSheet.getDataRange().getValues();
  
  // หา adminUsers
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'adminUsers') {
      const adminList = String(data[i][1]).split(',').map(e => e.trim());
      
      // ตรวจสอบซ้ำ
      if (adminList.includes(email)) {
        return { success: false, error: 'User is already admin' };
      }
      
      adminList.push(email);
      configSheet.getRange(i + 1, 2).setValue(adminList.join(', '));
      return { success: true, message: 'Admin added successfully' };
    }
  }
  
  // ถ้าไม่มี adminUsers ให้สร้างใหม่
  configSheet.appendRow(['adminUsers', email]);
  return { success: true, message: 'Admin added successfully' };
}

// ✅ ลบ Admin
function removeAdminUser(email) {
  if (!isAdmin()) {
    throw new Error('Access Denied: Admin only');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = ss.getSheetByName(SHEET_NAMES.CONFIG);
  
  if (!configSheet) {
    return { success: false, error: 'Config sheet not found' };
  }
  
  const data = configSheet.getDataRange().getValues();
  
  // หา adminUsers
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === 'adminUsers') {
      const adminList = String(data[i][1]).split(',').map(e => e.trim());
      const filteredList = adminList.filter(e => e !== email);
      
      if (filteredList.length === adminList.length) {
        return { success: false, error: 'User is not admin' };
      }
      
      // ตรวจสอบว่าต้องมี admin อย่างน้อย 1 คน
      if (filteredList.length === 0) {
        return { success: false, error: 'Must have at least one admin' };
      }
      
      configSheet.getRange(i + 1, 2).setValue(filteredList.join(', '));
      return { success: true, message: 'Admin removed successfully' };
    }
  }
  
  return { success: false, error: 'Admin users not found' };
}

// ดึงข้อมูลห้องและเครื่องทั้งหมด
function getRooms() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.ROOMS);
  
  // ถ้ายังไม่มี sheet ให้สร้างพร้อมข้อมูลตัวอย่าง
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.ROOMS);
    sheet.getRange('A1:E1').setValues([['RoomID', 'RoomName', 'ComputerID', 'OS', 'Specs']]);
    
    // ข้อมูลตัวอย่าง
    const sampleData = [
      ['R101', 'ห้อง 101', 'C1', 'Mac', 'Editor,Color,3D'],
      ['R101', 'ห้อง 101', 'C2', 'Mac', 'Editor,Visual effect'],
      ['R101', 'ห้อง 101', 'C3', 'Windows', '3D,Sound'],
      ['R102', 'ห้อง 102', 'C1', 'Windows', 'Editor,Graphic'],
      ['R102', 'ห้อง 102', 'C2', 'Windows', 'Color,3D'],
      ['R103', 'ห้อง 103', 'C1', 'Mac', 'Editor,Color,Visual effect,Sound']
    ];
    sheet.getRange(2, 1, sampleData.length, 5).setValues(sampleData);
  }
  
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rooms = {};
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const roomId = row[0];
    const roomName = row[1];
    const computerId = row[2];
    const os = row[3];
    const specs = row[4] ? row[4].split(',') : [];
    
    if (!rooms[roomId]) {
      rooms[roomId] = {
        id: roomId,
        name: roomName,
        computers: []
      };
    }
    
    rooms[roomId].computers.push({
      id: computerId,
      os: os,
      specs: specs
    });
  }
  
  return Object.values(rooms);
}

// ดึงข้อมูล Profile ผู้ใช้
function getUserProfile() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS);
    sheet.getRange('A1:F1').setValues([['Email', 'FirstName', 'LastName', 'Phone', 'StudentID', 'Language']]);
    return null;
  }
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      return {
        firstName: data[i][1],
        lastName: data[i][2],
        phone: data[i][3],
        studentId: data[i][4],
        language: data[i][5] || 'th'
      };
    }
  }
  
  return null;
}

// บันทึก Profile ผู้ใช้
function saveUserProfile(profile) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.USERS);
    sheet.getRange('A1:F1').setValues([['Email', 'FirstName', 'LastName', 'Phone', 'StudentID', 'Language']]);
  }
  
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  
  // ค้นหาว่ามีข้อมูลอยู่แล้วหรือไม่
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === email) {
      rowIndex = i + 1;
      break;
    }
  }
  
  const rowData = [email, profile.firstName, profile.lastName, profile.phone, profile.studentId, profile.language || 'th'];
  
  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, 6).setValues([rowData]);
  } else {
    sheet.appendRow(rowData);
  }
  
  return { success: true };
}

// ดึงข้อมูลการจองทั้งหมดของผู้ใช้
function getUserBookings() {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BOOKINGS);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.BOOKINGS);
    sheet.getRange('A1:O1').setValues([
      ['BookingID', 'Email', 'FirstName', 'LastName', 'Phone', 'StudentID', 
       'Date', 'Time', 'Room', 'Computer', 'OS', 'Usage', 'Subject', 'Reason', 'CreatedAt']
    ]);
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  
  for (let i = 1; i < data.length; i++) {
    // ข้ามแถวว่าง
    if (!data[i][0]) continue;
    
    if (data[i][1] === email) {
      // แปลงวันที่ให้เป็น string
      let bookingDate = data[i][6];
      if (bookingDate instanceof Date) {
        bookingDate = Utilities.formatDate(bookingDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      
      let createdDate = data[i][14];
      if (createdDate instanceof Date) {
        createdDate = createdDate.toISOString();
      }
      
      bookings.push({
        id: String(data[i][0]),
        email: String(data[i][1]),
        firstName: String(data[i][2]),
        lastName: String(data[i][3]),
        phone: String(data[i][4]),
        studentId: String(data[i][5]),
        date: bookingDate,
        time: String(data[i][7]),
        room: String(data[i][8]),
        computer: String(data[i][9]),
        os: String(data[i][10]),
        usage: data[i][11] ? String(data[i][11]).split(',') : [],
        subject: String(data[i][12]),
        reason: String(data[i][13]),
        createdAt: createdDate,
        status: 'confirmed'
      });
    }
  }
  
  Logger.log('Found ' + bookings.length + ' bookings for user: ' + email);
  return bookings;
}

// ดึงข้อมูลการจองทั้งหมด (สำหรับตารางการจอง)
function getAllBookings(selectedDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BOOKINGS);
  
  if (!sheet) {
    return [];
  }
  
  const data = sheet.getDataRange().getValues();
  const bookings = [];
  
  for (let i = 1; i < data.length; i++) {
    // ข้ามแถวว่าง
    if (!data[i][0]) continue;
    
    const bookingDate = Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    
    if (!selectedDate || bookingDate === selectedDate) {
      bookings.push({
        room: String(data[i][8]),
        computer: String(data[i][9]),
        time: String(data[i][7]),
        date: bookingDate,
        os: String(data[i][10]),
        email: String(data[i][1]),
        firstName: String(data[i][2]),
        lastName: String(data[i][3])
      });
    }
  }
  
  Logger.log('Found ' + bookings.length + ' bookings for date: ' + selectedDate);
  return bookings;
}

// ส่ง Email แจ้งการจอง
function sendBookingEmail(bookingData) {
  const config = getConfig();
  
  if (config.emailEnabled !== 'true') {
    Logger.log('Email disabled');
    return;
  }
  
  const userProfile = getUserProfile();
  const language = userProfile ? userProfile.language : 'th';
  
  let subject, body;
  
  if (language === 'en') {
    subject = '✅ Booking Confirmation - ' + bookingData.bookingId;
    body = `
      <h2 style="color: #4F46E5;">Booking Confirmed!</h2>
      <p>Dear ${bookingData.firstName} ${bookingData.lastName},</p>
      <p>Your booking has been confirmed successfully. Here are the details:</p>
      
      <div style="background: #F3F4F6; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <p><strong>Booking ID:</strong> ${bookingData.bookingId}</p>
        <p><strong>Date:</strong> ${bookingData.date}</p>
        <p><strong>Time:</strong> ${bookingData.time}</p>
        <p><strong>Room:</strong> ${bookingData.room}</p>
        <p><strong>Computer:</strong> ${bookingData.computer}</p>
        <p><strong>OS:</strong> ${bookingData.os}</p>
        <p><strong>Usage:</strong> ${bookingData.usage.join(', ')}</p>
        <p><strong>Subject:</strong> ${bookingData.subject}</p>
        <p><strong>Reason:</strong> ${bookingData.reason}</p>
      </div>
      
      <p style="color: #6B7280;">Please arrive on time. If you need to cancel, please contact the administrator.</p>
      <p>Thank you!</p>
    `;
  } else {
    subject = '✅ ยืนยันการจอง - ' + bookingData.bookingId;
    body = `
      <h2 style="color: #4F46E5;">การจองสำเร็จ!</h2>
      <p>เรียน คุณ${bookingData.firstName} ${bookingData.lastName}</p>
      <p>การจองของคุณได้รับการยืนยันเรียบร้อยแล้ว รายละเอียดดังนี้:</p>
      
      <div style="background: #F3F4F6; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <p><strong>รหัสการจอง:</strong> ${bookingData.bookingId}</p>
        <p><strong>วันที่:</strong> ${bookingData.date}</p>
        <p><strong>เวลา:</strong> ${bookingData.time}</p>
        <p><strong>ห้อง:</strong> ${bookingData.room}</p>
        <p><strong>เครื่อง:</strong> ${bookingData.computer}</p>
        <p><strong>ระบบปฏิบัติการ:</strong> ${bookingData.os}</p>
        <p><strong>การใช้งาน:</strong> ${bookingData.usage.join(', ')}</p>
        <p><strong>วิชา:</strong> ${bookingData.subject}</p>
        <p><strong>เหตุผล:</strong> ${bookingData.reason}</p>
      </div>
      
      <p style="color: #6B7280;">กรุณามาตรงเวลา หากต้องการยกเลิกกรุณาติดต่อผู้ดูแลระบบ</p>
      <p>ขอบคุณครับ/ค่ะ</p>
    `;
  }
  
  try {
    MailApp.sendEmail({
      to: bookingData.email,
      subject: subject,
      htmlBody: body
    });
    
    // ส่ง CC ให้ admin ถ้ามี
    if (config.adminEmail) {
      MailApp.sendEmail({
        to: config.adminEmail,
        subject: '[Admin Copy] ' + subject,
        htmlBody: body
      });
    }
    
    Logger.log('Email sent to: ' + bookingData.email);
  } catch (e) {
    Logger.log('Error sending email: ' + e.message);
  }
}

// สร้างการจองใหม่
function createBooking(bookingData) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BOOKINGS);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.BOOKINGS);
    sheet.getRange('A1:O1').setValues([
      ['BookingID', 'Email', 'FirstName', 'LastName', 'Phone', 'StudentID', 
       'Date', 'Time', 'Room', 'Computer', 'OS', 'Usage', 'Subject', 'Reason', 'CreatedAt']
    ]);
  }
  
  // ตรวจสอบว่ามีการจองซ้ำหรือไม่
  const data = sheet.getDataRange().getValues();
  const newDate = Utilities.formatDate(new Date(bookingData.date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  
  for (let i = 1; i < data.length; i++) {
    // ข้ามแถวว่าง
    if (!data[i][0]) continue;
    
    const existingDate = Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const existingTime = String(data[i][7]);
    const existingRoom = String(data[i][8]);
    const existingComputer = String(data[i][9]);
    
    Logger.log('Checking booking: Date=' + existingDate + ', Time=' + existingTime + ', Room=' + existingRoom + ', Computer=' + existingComputer);
    Logger.log('New booking: Date=' + newDate + ', Time=' + bookingData.time + ', Room=' + bookingData.room + ', Computer=' + bookingData.computer);
    
    // ตรวจสอบการจองซ้ำ
    if (existingDate === newDate && 
        existingRoom === bookingData.room && 
        existingComputer === bookingData.computer) {
      
      // ถ้าจองทั้งวัน หรือ เวลาซ้ำกัน
      if (existingTime === 'ทั้งวัน' || bookingData.time === 'ทั้งวัน') {
        return { 
          success: false, 
          error: 'มีผู้จองเครื่องนี้ไปแล้ว (ทั้งวัน)' 
        };
      }
      
      // ถ้าเวลาซ้ำกัน
      if (existingTime === bookingData.time) {
        return { 
          success: false, 
          error: 'มีผู้จองเครื่องนี้ในช่วงเวลา' + bookingData.time + 'แล้ว' 
        };
      }
    }
  }
  
  // สร้าง BookingID
  const bookingId = 'B' + new Date().getTime();
  const createdAt = new Date();
  
  const rowData = [
    bookingId,
    email,
    bookingData.firstName,
    bookingData.lastName,
    bookingData.phone,
    bookingData.studentId,
    new Date(bookingData.date),
    bookingData.time,
    bookingData.room,
    bookingData.computer,
    bookingData.os,
    bookingData.usage.join(','),
    bookingData.subject,
    bookingData.reason,
    createdAt
  ];
  
  sheet.appendRow(rowData);
  
  // ส่ง Email
  bookingData.email = email;
  bookingData.bookingId = bookingId;
  sendBookingEmail(bookingData);
  
  return { 
    success: true, 
    bookingId: bookingId,
    message: 'จองสำเร็จ!' 
  };
}

// ยกเลิกการจอง
function cancelBooking(bookingId) {
  const email = Session.getActiveUser().getEmail();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAMES.BOOKINGS);
  
  if (!sheet) return { success: false, error: 'Sheet not found' };
  
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === bookingId && String(data[i][1]) === email) {
      // ส่งอีเมลแจ้งยกเลิก
      sendCancellationEmail({
        bookingId: String(data[i][0]),
        email: String(data[i][1]),
        firstName: String(data[i][2]),
        lastName: String(data[i][3]),
        date: Utilities.formatDate(new Date(data[i][6]), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        time: String(data[i][7]),
        room: String(data[i][8]),
        computer: String(data[i][9])
      });
      
      sheet.deleteRow(i + 1);
      return { success: true, message: 'ยกเลิกการจองสำเร็จ' };
    }
  }
  
  return { success: false, error: 'Booking not found' };
}

// ส่ง Email แจ้งการยกเลิก
function sendCancellationEmail(bookingData) {
  const config = getConfig();
  
  if (config.emailEnabled !== 'true') {
    Logger.log('Email disabled');
    return;
  }
  
  const userProfile = getUserProfile();
  const language = userProfile ? userProfile.language : 'th';
  
  let subject, body;
  
  if (language === 'en') {
    subject = '❌ Booking Cancelled - ' + bookingData.bookingId;
    body = `
      <h2 style="color: #DC2626;">Booking Cancelled</h2>
      <p>Dear ${bookingData.firstName} ${bookingData.lastName},</p>
      <p>Your booking has been cancelled.</p>
      
      <div style="background: #FEF2F2; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #DC2626;">
        <p><strong>Booking ID:</strong> ${bookingData.bookingId}</p>
        <p><strong>Date:</strong> ${bookingData.date}</p>
        <p><strong>Time:</strong> ${bookingData.time}</p>
        <p><strong>Room:</strong> ${bookingData.room}</p>
        <p><strong>Computer:</strong> ${bookingData.computer}</p>
      </div>
      
      <p>If this was a mistake, please make a new booking.</p>
    `;
  } else {
    subject = '❌ ยกเลิกการจอง - ' + bookingData.bookingId;
    body = `
      <h2 style="color: #DC2626;">การจองถูกยกเลิกแล้ว</h2>
      <p>เรียน คุณ${bookingData.firstName} ${bookingData.lastName}</p>
      <p>การจองของคุณได้ถูกยกเลิกเรียบร้อยแล้ว</p>
      
      <div style="background: #FEF2F2; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #DC2626;">
        <p><strong>รหัสการจอง:</strong> ${bookingData.bookingId}</p>
        <p><strong>วันที่:</strong> ${bookingData.date}</p>
        <p><strong>เวลา:</strong> ${bookingData.time}</p>
        <p><strong>ห้อง:</strong> ${bookingData.room}</p>
        <p><strong>เครื่อง:</strong> ${bookingData.computer}</p>
      </div>
      
      <p>หากต้องการใช้บริการ กรุณาทำการจองใหม่อีกครั้ง</p>
    `;
  }
  
  try {
    MailApp.sendEmail({
      to: bookingData.email,
      subject: subject,
      htmlBody: body
    });
    
    Logger.log('Cancellation email sent to: ' + bookingData.email);
  } catch (e) {
    Logger.log('Error sending cancellation email: ' + e.message);
  }
}

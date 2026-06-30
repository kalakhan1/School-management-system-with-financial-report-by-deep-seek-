// ==================== GLOBAL CONFIGURATION ====================
const SPREADSHEET_ID = 'sheet id get from sheet url';
const CACHE_PREFIX = 'SCHOOL_ERP_';
const SESSION_EXPIRY = 30;

// ==================== WEB APP ====================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('School Management ERP')
    .setFaviconUrl('https://img.icons8.com/color/48/000000/school.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, shrink-to-fit=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== CACHE SERVICE ====================
function getC(key) {
  try {
    const c = CacheService.getScriptCache();
    return c.get(CACHE_PREFIX + key);
  } catch(e) { return null; }
}

function setC(key, val, sec = 21600) {
  try {
    CacheService.getScriptCache().put(CACHE_PREFIX + key, val, sec);
  } catch(e) {}
}

function delC(key) {
  try {
    const keys = key ? [key] : ['SETTINGS','STUDENTS','TEACHERS','CLASSES','FEES','SALARY','ARCHIVE_STU','ARCHIVE_TCH'];
    const c = CacheService.getScriptCache();
    keys.forEach(k => c.remove(CACHE_PREFIX + k));
  } catch(e) {}
}

// ==================== ERROR LOGGING ====================
function logErr(src, desc, err) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(5000);
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('ErrorLogs');
    if (!sheet) {
      sheet = doc.insertSheet('ErrorLogs');
      sheet.getRange(1,1,1,4).setValues([['Timestamp','Source','Description','Error']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([new Date(), src, desc, err ? err.toString() : '']);
    lock.releaseLock();
  } catch(e) { console.error(src, desc, err); }
}

// ==================== AUTHENTICATION ====================
function adminLogin(pwd) {
  try {
    const s = getSettings();
    if (!s || !s.adminPassword) {
      createDefaultSettings();
      return { success: false, message: 'System initialized. Default: admin123' };
    }
    if (pwd === s.adminPassword) {
      const token = Utilities.getUuid();
      setC('SESSION_' + token, JSON.stringify({ role: 'admin', time: new Date().toISOString() }), 3600);
      return { success: true, token: token, settings: s };
    }
    return { success: false, message: 'Invalid password' };
  } catch(e) { return { success: false, message: 'Login error' }; }
}

function validateSession(token) {
  if (!token) return false;
  const d = getC('SESSION_' + token);
  if (!d) return false;
  try {
    const session = JSON.parse(d);
    const exp = new Date(session.time);
    exp.setMinutes(exp.getMinutes() + SESSION_EXPIRY);
    if (new Date() > exp) { delC('SESSION_' + token); return false; }
    session.time = new Date().toISOString();
    setC('SESSION_' + token, JSON.stringify(session), 3600);
    return true;
  } catch(e) { return false; }
}

function adminLogout(token) {
  if (token) delC('SESSION_' + token);
  return { success: true };
}

// ==================== SETTINGS ====================
function getSettings() {
  const cached = getC('SETTINGS');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Settings');
    if (!sheet) return createDefaultSettings();
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return createDefaultSettings();
    const settings = {};
    for (let i = 1; i < data.length; i++) {
      if (data[i][0]) settings[data[i][0]] = data[i][1] || '';
    }
    setC('SETTINGS', JSON.stringify(settings));
    return settings;
  } catch(e) { return {}; }
}

function createDefaultSettings() {
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Settings');
    if (sheet) sheet.clear();
    else sheet = doc.insertSheet('Settings');
    
    const defaults = [
      ['schoolName', 'The Excellence School'],
      ['adminPassword', 'admin123'],
      ['logoUrl', 'https://img.icons8.com/color/96/000000/school.png'],
      ['address', '123 Education Lane, Model Town, Lahore'],
      ['phone', '+92-300-1234567'],
      ['email', 'info@excellenceschool.edu.pk'],
      ['website', 'www.excellenceschool.edu.pk'],
      ['principalName', 'Dr. Muhammad Ahmed'],
      ['motto', 'Knowledge is Power'],
      ['primaryColor', '#1a237e'],
      ['secondaryColor', '#0d47a1'],
      ['accentColor', '#ff6f00'],
      ['theme', 'light'],
      ['language', 'en'],
      ['session', '2024-2025']
    ];
    sheet.getRange(1,1,1,2).setValues([['Key','Value']]);
    sheet.getRange(2,1,defaults.length,2).setValues(defaults);
    sheet.setFrozenRows(1);
    delC('SETTINGS');
    const obj = Object.fromEntries(defaults);
    setC('SETTINGS', JSON.stringify(obj));
    return obj;
  } catch(e) { logErr('createDefaultSettings', 'Init error', e); return {}; }
}

function updateSettings(obj) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Settings');
    if (!sheet) { sheet = doc.insertSheet('Settings'); sheet.getRange(1,1,1,2).setValues([['Key','Value']]); }
    const existing = {};
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { if (data[i][0]) existing[data[i][0]] = data[i][1]; }
    const merged = { ...existing, ...obj };
    const entries = Object.entries(merged);
    sheet.clear();
    sheet.getRange(1,1,1,2).setValues([['Key','Value']]);
    if (entries.length > 0) sheet.getRange(2,1,entries.length,2).setValues(entries);
    sheet.setFrozenRows(1);
    delC('SETTINGS');
    return { success: true };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

// ==================== CLASSES & SECTIONS ====================
function getClasses() {
  return [
    'Play Group', 'Nursery', 'Prep', 'Class 1', 'Class 2', 'Class 3', 
    'Class 4', 'Class 5', 'Class 6', 'Class 7', 'Class 8', 
    'Class 9', 'Class 10', 'Class 11', 'Class 12'
  ];
}

function getSections() {
  return ['A', 'B', 'C', 'D'];
}

// ==================== STUDENT CRUD ====================
function addStudent(data, imgBase64) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Students');
    if (!sheet) { sheet = doc.insertSheet('Students'); createStudentHeaders(sheet); }
    
    const sid = 'STU-' + Utilities.getUuid().substring(0,8).toUpperCase();
    const grNo = 'GR-' + new Date().getFullYear() + '-' + String(sheet.getLastRow()).padStart(4,'0');
    const img = imgBase64 ? saveImage(sid, imgBase64) : (data.imageUrl || '');
    
    sheet.appendRow([
      sid, grNo, data.fullName||'', data.fatherName||'', data.motherName||'',
      data.gender||'', data.dob||'', data.age||'', data.cnic||'', data.phone||'',
      data.whatsapp||'', data.email||'', data.address||'', data.city||'',
      data.country||'Pakistan', data.religion||'', data.bloodGroup||'',
      data.class||'', data.section||'', data.rollNo||'', data.house||'',
      new Date(), data.fee||0, data.discount||0, data.status||'Active',
      data.notes||'', img, data.fatherPhone||'', data.motherPhone||'',
      data.emergencyContact||''
    ]);
    delC('STUDENTS');
    return { success: true, studentId: sid, grNo: grNo };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateStudent(sid, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Students');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === sid) {
        Object.keys(updates).forEach(k => {
          const idx = headers.indexOf(k);
          if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]);
        });
        delC('STUDENTS');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteStudent(sid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Students');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    for (let i = data.length-1; i >= 1; i--) {
      if (data[i][0] === sid) {
        archiveRecord('ArchiveStudents', data[i]);
        sheet.deleteRow(i+1);
        delC('STUDENTS');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getStudents() {
  const cached = getC('STUDENTS');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Students');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    setC('STUDENTS', JSON.stringify(result));
    return result;
  } catch(e) { return []; }
}

function getStudentById(sid) {
  const students = getStudents();
  return students.find(s => s.StudentID === sid) || null;
}

// ==================== TEACHER CRUD ====================
function addTeacher(data, imgBase64) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Teachers');
    if (!sheet) { sheet = doc.insertSheet('Teachers'); createTeacherHeaders(sheet); }
    
    const tid = 'TCH-' + Utilities.getUuid().substring(0,8).toUpperCase();
    const empCode = 'EMP-' + String(sheet.getLastRow()).padStart(4,'0');
    const img = imgBase64 ? saveImage(tid, imgBase64) : (data.imageUrl || '');
    
    sheet.appendRow([
      tid, empCode, data.name||'', data.fatherName||'', data.gender||'',
      data.dob||'', data.cnic||'', data.phone||'', data.whatsapp||'',
      data.email||'', data.address||'', data.city||'', data.qualification||'',
      data.experience||'', data.subjects||'', data.classTeacher||'',
      new Date(), data.salaryType||'Monthly', data.perLecturePay||0,
      data.monthlyPackage||0, data.status||'Active', img, data.notes||''
    ]);
    delC('TEACHERS');
    return { success: true, teacherId: tid, empCode: empCode };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateTeacher(tid, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Teachers');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === tid) {
        Object.keys(updates).forEach(k => {
          const idx = headers.indexOf(k);
          if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]);
        });
        delC('TEACHERS');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteTeacher(tid) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Teachers');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    for (let i = data.length-1; i >= 1; i--) {
      if (data[i][0] === tid) {
        archiveRecord('ArchiveTeachers', data[i]);
        sheet.deleteRow(i+1);
        delC('TEACHERS');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getTeachers() {
  const cached = getC('TEACHERS');
  if (cached) { try { return JSON.parse(cached); } catch(e) {} }
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Teachers');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    const result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    setC('TEACHERS', JSON.stringify(result));
    return result;
  } catch(e) { return []; }
}

// ==================== ARCHIVE ====================
function archiveRecord(sheetName, rowData) {
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName(sheetName);
    if (!sheet) {
      sheet = doc.insertSheet(sheetName);
      if (sheetName === 'ArchiveStudents') createStudentHeaders(sheet);
      else createTeacherHeaders(sheet);
    }
    sheet.appendRow(rowData);
  } catch(e) { logErr('archiveRecord', 'Archive error', e); }
}

function getArchivedStudents() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArchiveStudents');
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1);
  } catch(e) { return []; }
}

function getArchivedTeachers() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ArchiveTeachers');
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1);
  } catch(e) { return []; }
}

// ==================== FEE MANAGEMENT ====================
function addFeePayment(studentId, amount, month, receiptNo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('FeeHistory');
    if (!sheet) {
      sheet = doc.insertSheet('FeeHistory');
      sheet.getRange(1,1,1,5).setValues([['StudentID','Date','Amount','Month','ReceiptNo']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([studentId, new Date(), amount, month, receiptNo || Utilities.getUuid().substring(0,8)]);
    delC('FEES');
    return { success: true };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getFeeHistory(studentId) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('FeeHistory');
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1).filter(r => r[0] === studentId);
  } catch(e) { return []; }
}

function getAllFeeHistory() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('FeeHistory');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    return data.slice(1).map(r => { const o={}; headers.forEach((h,i)=>o[h]=r[i]||''); return o; });
  } catch(e) { return []; }
}

// ==================== SALARY MANAGEMENT ====================
function addSalaryPayment(teacherId, month, amount, lectures, deductions) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('SalaryHistory');
    if (!sheet) {
      sheet = doc.insertSheet('SalaryHistory');
      sheet.getRange(1,1,1,7).setValues([['TeacherID','Month','Amount','Lectures','Deductions','NetPay','Date']]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([teacherId, month, amount, lectures||0, deductions||0, amount-(deductions||0), new Date()]);
    delC('SALARY');
    return { success: true };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getSalaryHistory(teacherId) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SalaryHistory');
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1).filter(r => r[0] === teacherId);
  } catch(e) { return []; }
}

function getAllSalaryHistory() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('SalaryHistory');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    return data.slice(1).map(r => { const o={}; headers.forEach((h,i)=>o[h]=r[i]||''); return o; });
  } catch(e) { return []; }
}

// ==================== IMAGE HANDLING ====================
function saveImage(id, base64Data) {
  try {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), 'image/png', id + '.png');
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const parents = DriveApp.getFileById(SPREADSHEET_ID).getParents();
    if (!parents.hasNext()) return '';
    const parentFolder = parents.next();
    let folder;
    const folders = parentFolder.getFoldersByName('School_Images');
    if (folders.hasNext()) folder = folders.next();
    else folder = parentFolder.createFolder('School_Images');
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/uc?export=view&id=' + file.getId();
  } catch(e) { return ''; }
}

// ==================== MAINTENANCE ====================
function resetProject() {
  try {
    delC();
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = doc.getSheetByName('Settings');
    if (sheet) sheet.clear();
    createDefaultSettings();
    return { success: true, message: 'School ERP reset. Password: admin123' };
  } catch(e) { return { success: false }; }
}

function clearSystemCache() {
  delC();
  return { success: true, message: 'Cache cleared' };
}

function getErrorLogs() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('ErrorLogs');
    if (!sheet) return [];
    return sheet.getDataRange().getValues().slice(1);
  } catch(e) { return []; }
}

// ==================== DUMMY DATA ====================
function generateDummyData() {
  try {
    const students = [
      { fullName:'Ahmed Khan', fatherName:'Muhammad Khan', motherName:'Fatima Khan', gender:'Male', dob:'2015-03-15', age:'9', cnic:'35202-1234567-1', phone:'0300-1111111', whatsapp:'0300-1111111', email:'parent@test.com', address:'123 Main St', city:'Lahore', country:'Pakistan', religion:'Islam', bloodGroup:'B+', class:'Class 4', section:'A', rollNo:'01', house:'Blue', fee:5000, discount:500, status:'Active', fatherPhone:'0300-9999999', motherPhone:'0300-8888888', emergencyContact:'0300-7777777' },
      { fullName:'Fatima Ali', fatherName:'Ali Raza', motherName:'Sara Ali', gender:'Female', dob:'2012-07-22', age:'12', cnic:'35202-9876543-2', phone:'0300-2222222', whatsapp:'0300-2222222', email:'parent2@test.com', address:'456 Model Town', city:'Karachi', country:'Pakistan', religion:'Islam', bloodGroup:'A+', class:'Class 7', section:'B', rollNo:'15', house:'Green', fee:6000, discount:0, status:'Active', fatherPhone:'0300-6666666', motherPhone:'0300-5555555', emergencyContact:'0300-4444444' },
      { fullName:'Usman Raza', fatherName:'Raza Ahmed', motherName:'Noor Raza', gender:'Male', dob:'2018-11-10', age:'6', cnic:'35202-5555555-3', phone:'0300-3333333', whatsapp:'0300-3333333', email:'parent3@test.com', address:'789 College Rd', city:'Islamabad', country:'Pakistan', religion:'Islam', bloodGroup:'O+', class:'Class 1', section:'C', rollNo:'08', house:'Red', fee:4000, discount:200, status:'Active', fatherPhone:'0300-3333333', motherPhone:'0300-2222222', emergencyContact:'0300-1111111' }
    ];
    
    const teachers = [
      { name:'Prof. Abdul Qadir', fatherName:'Late Ahmed', gender:'Male', dob:'1980-05-15', cnic:'35202-1111111-1', phone:'0300-6666666', whatsapp:'0300-6666666', email:'qadir@school.edu.pk', address:'789 College Rd', city:'Lahore', qualification:'M.Sc Mathematics', experience:'15 years', subjects:'Mathematics', classTeacher:'Class 10', salaryType:'Monthly', perLecturePay:0, monthlyPackage:75000, status:'Active' },
      { name:'Ms. Sara Ahmed', fatherName:'Ahmed Ali', gender:'Female', dob:'1985-08-20', cnic:'35202-2222222-2', phone:'0300-7777777', whatsapp:'0300-7777777', email:'sara@school.edu.pk', address:'123 Garden Town', city:'Lahore', qualification:'M.A English', experience:'10 years', subjects:'English', classTeacher:'Class 7', salaryType:'Monthly', perLecturePay:0, monthlyPackage:60000, status:'Active' }
    ];
    
    students.forEach(s => addStudent(s));
    teachers.forEach(t => addTeacher(t));
    
    return { success: true, message: 'Dummy data generated!' };
  } catch(e) { return { success: false }; }
}

// ==================== SHEET HEADERS ====================
function createStudentHeaders(sheet) {
  sheet.getRange(1,1,1,30).setValues([['StudentID','GRNo','FullName','FatherName','MotherName','Gender','DOB','Age','CNIC','Phone','WhatsApp','Email','Address','City','Country','Religion','BloodGroup','Class','Section','RollNo','House','AdmissionDate','Fee','Discount','Status','Notes','ImageURL','FatherPhone','MotherPhone','EmergencyContact']]);
  sheet.setFrozenRows(1);
}

function createTeacherHeaders(sheet) {
  sheet.getRange(1,1,1,22).setValues([['TeacherID','EmpCode','Name','FatherName','Gender','DOB','CNIC','Phone','WhatsApp','Email','Address','City','Qualification','Experience','Subjects','ClassTeacher','JoiningDate','SalaryType','PerLecturePay','MonthlyPackage','Status','ImageURL','Notes']]);
  sheet.setFrozenRows(1);
}

function onOpen() {
  try { createDefaultSettings(); } catch(e) {}
}

// ==================== EXPENSE MANAGEMENT ====================
function addExpense(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const doc = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = doc.getSheetByName('Expenses');
    if (!sheet) {
      sheet = doc.insertSheet('Expenses');
      sheet.getRange(1,1,1,8).setValues([['ExpenseID','Date','Category','Description','Amount','PaymentMode','BillNo','Notes']]);
      sheet.setFrozenRows(1);
    }
    const expId = 'EXP-' + Utilities.getUuid().substring(0,8).toUpperCase();
    sheet.appendRow([
      expId,
      data.date || new Date(),
      data.category || 'General',
      data.description || '',
      parseFloat(data.amount) || 0,
      data.paymentMode || 'Cash',
      data.billNo || '',
      data.notes || ''
    ]);
    delC('EXPENSES');
    return { success: true, expenseId: expId };
  } catch(e) { return { success: false, message: e.toString() }; }
  finally { lock.releaseLock(); }
}

function updateExpense(expId, data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Expenses');
    if (!sheet) return { success: false };
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][0] === expId) {
        const updates = {
          'Date': data.date || allData[i][1],
          'Category': data.category || allData[i][2],
          'Description': data.description || allData[i][3],
          'Amount': parseFloat(data.amount) || allData[i][4],
          'PaymentMode': data.paymentMode || allData[i][5],
          'BillNo': data.billNo || allData[i][6],
          'Notes': data.notes || allData[i][7]
        };
        Object.keys(updates).forEach(k => {
          const idx = headers.indexOf(k);
          if (idx !== -1) sheet.getRange(i+1, idx+1).setValue(updates[k]);
        });
        delC('EXPENSES');
        return { success: true };
      }
    }
    return { success: false, message: 'Expense not found' };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function deleteExpense(expId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Expenses');
    if (!sheet) return { success: false };
    const data = sheet.getDataRange().getValues();
    for (let i = data.length-1; i >= 1; i--) {
      if (data[i][0] === expId) {
        sheet.deleteRow(i+1);
        delC('EXPENSES');
        return { success: true };
      }
    }
    return { success: false };
  } catch(e) { return { success: false }; }
  finally { lock.releaseLock(); }
}

function getExpenses(month, year) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Expenses');
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];
    const headers = data[0];
    let result = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = data[i][idx] || '');
      result.push(obj);
    }
    
    // Filter by month/year if provided
    if (month || year) {
      result = result.filter(r => {
        const d = new Date(r.Date);
        const m = d.getMonth() + 1;
        const y = d.getFullYear();
        if (month && year) return m === parseInt(month) && y === parseInt(year);
        if (month) return m === parseInt(month);
        if (year) return y === parseInt(year);
        return true;
      });
    }
    
    return result;
  } catch(e) { return []; }
}

function getExpenseById(expId) {
  const expenses = getExpenses();
  return expenses.find(e => e.ExpenseID === expId) || null;
}

// ==================== EXPENSE CATEGORIES ====================
function getExpenseCategories() {
  return [
    'Salary', 'Rent', 'Utilities', 'Stationery', 'Books',
    'Maintenance', 'Transport', 'Food/Canteen', 'Events',
    'Marketing', 'Software', 'Internet', 'Phone Bill',
    'Office Supplies', 'Cleaning', 'Security', 'General', 'Other'
  ];
}

// ==================== FINANCIAL REPORT ====================
function getFinancialReport(month, year) {
  try {
    // Get all income (fee payments)
    const feeHistory = getAllFeeHistory();
    const expenses = getExpenses(month, year);
    
    // Filter fee history by month/year
    let filteredFees = feeHistory;
    if (month || year) {
      filteredFees = feeHistory.filter(f => {
        const d = new Date(f.Date);
        const m = d.getMonth() + 1;
        const y = d.getFullYear();
        if (month && year) return m === parseInt(month) && y === parseInt(year);
        if (month) return m === parseInt(month);
        if (year) return y === parseInt(year);
        return true;
      });
    }
    
    const totalIncome = filteredFees.reduce((sum, f) => sum + parseFloat(f.Amount || f[2] || 0), 0);
    const totalExpenses = expenses.reduce((sum, e) => sum + parseFloat(e.Amount || 0), 0);
    const profit = totalIncome - totalExpenses;
    
    // Category-wise expense breakdown
    const expensesByCategory = {};
    expenses.forEach(e => {
      const cat = e.Category || 'General';
      expensesByCategory[cat] = (expensesByCategory[cat] || 0) + parseFloat(e.Amount || 0);
    });
    
    // Monthly breakdown for current year
    const monthlyBreakdown = [];
    const currentYear = year || new Date().getFullYear();
    for (let m = 1; m <= 12; m++) {
      const monthFees = feeHistory.filter(f => {
        const d = new Date(f.Date);
        return d.getMonth() + 1 === m && d.getFullYear() === currentYear;
      });
      const monthExpenses = expenses.filter(e => {
        const d = new Date(e.Date);
        return d.getMonth() + 1 === m && d.getFullYear() === currentYear;
      });
      monthlyBreakdown.push({
        month: m,
        monthName: new Date(currentYear, m-1).toLocaleString('default', { month: 'long' }),
        income: monthFees.reduce((s,f) => s + parseFloat(f.Amount||f[2]||0), 0),
        expenses: monthExpenses.reduce((s,e) => s + parseFloat(e.Amount||0), 0)
      });
    }
    
    return {
      success: true,
      totalIncome,
      totalExpenses,
      profit,
      expensesByCategory,
      monthlyBreakdown,
      expenseCount: expenses.length,
      feeCount: filteredFees.length,
      period: { month: month || 'All', year: year || 'All' }
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

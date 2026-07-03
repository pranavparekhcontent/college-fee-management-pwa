/**
 * College Fee Management System — Google Apps Script Central API (v1.0)
 * Handles CRUD for students, transactions, settings, reports, and backup/restore.
 * Proxies access to individual college sheets via sheetId.
 */

var _ssCache = {};
function _getSpreadsheet(sheetId) {
  if (!sheetId) {
    throw new Error("Missing sheetId parameter");
  }
  if (!_ssCache[sheetId]) {
    _ssCache[sheetId] = SpreadsheetApp.openById(sheetId);
  }
  return _ssCache[sheetId];
}

// Helper to ensure all required sheets exist
function _ensureSheets(ss) {
  var sheets = {
    'Students': [
      'id', 'serial_no', 'name', 'gender', 'student_mobile', 'parent_mobile', 
      'student_email', 'parent_email', 'permanent_address', 'local_address', 
      'year_of_admission', 'category', 'current_class', 'current_academic_year', 
      'fee_applicable', 'fee_paid', 'fee_pending', 'scholarship_expected', 
      'scholarship_received', 'scholarship_pending', 'remark', 'admission_cancelled'
    ],
    'Transactions': [
      'id', 'student_id', 'student_name', 'transaction_id', 'date', 'amount', 
      'bank_name', 'source', 'status', 'verified', 'raw_remark', 'type', 'other_reason'
    ],
    'Settings': [
      'key', 'value'
    ]
  };

  for (var name in sheets) {
    var ws = ss.getSheetByName(name);
    if (!ws) {
      ws = ss.insertSheet(name);
      ws.appendRow(sheets[name]);
      // Format headers nicely
      ws.getRange(1, 1, 1, sheets[name].length)
        .setFontWeight('bold')
        .setBackground('#F1F5F9')
        .setHorizontalAlignment('center');
      ws.setFrozenRows(1);
    } else {
      var existingHeaders = ws.getRange(1, 1, 1, ws.getLastColumn()).getValues()[0];
      var expectedHeaders = sheets[name];
      for (var k = 0; k < expectedHeaders.length; k++) {
        var header = expectedHeaders[k];
        if (existingHeaders.indexOf(header) === -1) {
          ws.getRange(1, existingHeaders.length + 1).setValue(header)
            .setFontWeight('bold')
            .setBackground('#F1F5F9')
            .setHorizontalAlignment('center');
          existingHeaders.push(header);
        }
      }
    }
  }
}

function doGet(e) {
  try {
    var action = e.parameter.action;
    var sheetId = e.parameter.sheetId;
    var result;

    if (!sheetId) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Missing sheetId' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = _getSpreadsheet(sheetId);
    _ensureSheets(ss);

    switch (action) {
      case 'getAllData':
        result = getAllData(ss);
        break;
      case 'getStudents':
        result = getStudents(ss);
        break;
      case 'getTransactions':
        result = getTransactions(ss);
        break;
      case 'getDashboard':
        result = getDashboard(ss);
        break;
      case 'exportBackup':
        result = exportBackup(ss);
        break;
      case 'getReport':
        result = getReport(ss, e.parameter.reportType);
        break;
      default:
        result = { success: false, error: 'Unknown action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var action = data.action || (e.parameter && e.parameter.action);
    var sheetId = data.sheetId || (e.parameter && e.parameter.sheetId);
    var result;

    if (!sheetId) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: 'Missing sheetId' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var ss = _getSpreadsheet(sheetId);
    _ensureSheets(ss);

    switch (action) {
      case 'addStudent':
        result = addStudent(ss, data);
        break;
      case 'updateStudent':
        result = updateStudent(ss, data);
        break;
      case 'recordPayment':
        result = recordPayment(ss, data);
        break;
      case 'uploadStudents':
      case 'bulkSync':
        result = bulkSync(ss, data.rows || data.records || []);
        break;
      case 'sendNotification':
        result = sendNotification(ss, data);
        break;
      case 'verifyTransaction':
        result = verifyTransaction(ss, data.transactionId || data.id, data.verified);
        break;
      case 'saveSettings':
        result = saveSettings(ss, data.settings || {});
        break;
      case 'setSecurityPin':
        result = setSecurityPin(ss, data);
        break;
      case 'verifyPin':
        result = verifyPin(ss, data);
        break;
      case 'restoreBackup':
        result = restoreBackup(ss, data.students || [], data.transactions || []);
        break;
      default:
        result = { success: false, error: 'Unknown POST action: ' + action };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ─── READ OPERATIONS ───

function getAllData(ss) {
  var students = getStudents(ss).students || [];
  var transactions = getTransactions(ss).transactions || [];
  var settings = getSettings(ss).settings || {};
  
  return {
    success: true,
    students: students,
    transactions: transactions,
    settings: settings
  };
}

function getStudents(ss) {
  var sheet = ss.getSheetByName('Students');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var students = [];

  for (var i = 1; i < data.length; i++) {
    var student = {};
    for (var j = 0; j < headers.length; j++) {
      student[headers[j]] = data[i][j];
    }
    students.push(student);
  }

  return { success: true, students: students };
}

function getTransactions(ss) {
  var sheet = ss.getSheetByName('Transactions');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var transactions = [];

  for (var i = 1; i < data.length; i++) {
    var txn = {};
    for (var j = 0; j < headers.length; j++) {
      txn[headers[j]] = data[i][j];
    }
    transactions.push(txn);
  }

  return { success: true, transactions: transactions };
}

function getSettings(ss) {
  var sheet = ss.getSheetByName('Settings');
  var data = sheet.getDataRange().getValues();
  var settings = {};

  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    var value = data[i][1];
    if (key) {
      settings[key] = value;
    }
  }

  return { success: true, settings: settings };
}

function getDashboard(ss) {
  var students = getStudents(ss).students || [];
  var transactions = getTransactions(ss).transactions || [];

  var activeStudents = 0;
  var cancelledAdmissions = 0;
  var totalFeeApplicable = 0;
  var totalFeePaid = 0;
  var totalFeePending = 0;
  var totalScholarshipExpected = 0;
  var totalScholarshipReceived = 0;
  var totalScholarshipPending = 0;

  students.forEach(function(s) {
    var cancelled = Number(s.admission_cancelled) === 1;
    if (cancelled) {
      cancelledAdmissions++;
    } else {
      activeStudents++;
      totalFeeApplicable += Number(s.fee_applicable || 0);
      totalFeePaid += Number(s.fee_paid || 0);
      totalFeePending += Number(s.fee_pending || 0);
      totalScholarshipExpected += Number(s.scholarship_expected || 0);
      totalScholarshipReceived += Number(s.scholarship_received || 0);
      totalScholarshipPending += Number(s.scholarship_pending || 0);
    }
  });

  return {
    success: true,
    stats: {
      activeStudents: activeStudents,
      cancelledAdmissions: cancelledAdmissions,
      totalFeeApplicable: totalFeeApplicable,
      totalFeePaid: totalFeePaid,
      totalFeePending: totalFeePending,
      totalScholarshipExpected: totalScholarshipExpected,
      totalScholarshipReceived: totalScholarshipReceived,
      totalScholarshipPending: totalScholarshipPending,
      collectionRate: totalFeeApplicable > 0 ? (totalFeePaid / totalFeeApplicable) * 100 : 0
    }
  };
}

function exportBackup(ss) {
  var students = getStudents(ss).students || [];
  var transactions = getTransactions(ss).transactions || [];
  var settings = getSettings(ss).settings || {};

  return {
    success: true,
    timestamp: new Date().toISOString(),
    students: students,
    transactions: transactions,
    settings: settings
  };
}

function getReport(ss, reportType) {
  var students = getStudents(ss).students || [];
  var classMap = {};

  students.forEach(function(s) {
    if (Number(s.admission_cancelled) === 1) return;
    var cls = s.current_class || 'Unclassed';
    if (!classMap[cls]) {
      classMap[cls] = {
        class: cls,
        students: 0,
        male: 0,
        female: 0,
        fee_applicable: 0,
        fee_paid: 0,
        fee_pending: 0,
        scholarship_expected: 0,
        scholarship_received: 0,
        scholarship_pending: 0
      };
    }

    var c = classMap[cls];
    c.students++;
    if (String(s.gender).toLowerCase() === 'male') c.male++;
    else if (String(s.gender).toLowerCase() === 'female') c.female++;
    
    c.fee_applicable += Number(s.fee_applicable || 0);
    c.fee_paid += Number(s.fee_paid || 0);
    c.fee_pending += Number(s.fee_pending || 0);
    c.scholarship_expected += Number(s.scholarship_expected || 0);
    c.scholarship_received += Number(s.scholarship_received || 0);
    c.scholarship_pending += Number(s.scholarship_pending || 0);
  });

  var reports = [];
  for (var k in classMap) {
    var c = classMap[k];
    c.collection_rate = c.fee_applicable > 0 ? (c.fee_paid / c.fee_applicable) * 100 : 0;
    reports.push(c);
  }

  return { success: true, reports: reports };
}

// ─── WRITE OPERATIONS ───

function addStudent(ss, data) {
  var sheet = ss.getSheetByName('Students');
  var sData = sheet.getDataRange().getValues();
  var headers = sData[0];
  
  // Find new ID
  var maxId = 0;
  for (var i = 1; i < sData.length; i++) {
    var cid = Number(sData[i][0]);
    if (cid > maxId) maxId = cid;
  }
  var newId = maxId + 1;

  // Auto-calculated fields
  var applicable = Number(data.fee_applicable || 0);
  var paid = Number(data.fee_paid || 0);
  var expectedSch = Number(data.scholarship_expected || 0);
  var receivedSch = Number(data.scholarship_received || 0);

  data.id = newId;
  data.serial_no = newId; // Same as ID
  data.fee_pending = applicable - paid;
  data.scholarship_pending = expectedSch - receivedSch;
  data.admission_cancelled = data.admission_cancelled || 0;

  var row = [];
  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    var val = data[header];
    row.push(val !== undefined ? val : '');
  }
  sheet.appendRow(row);
  return { success: true, student: data };
}

function updateStudent(ss, data) {
  var sheet = ss.getSheetByName('Students');
  var sData = sheet.getDataRange().getValues();
  var headers = sData[0];
  var idColIdx = 0; // 'id' column is the first one
  
  var targetRowIdx = -1;
  for (var i = 1; i < sData.length; i++) {
    if (Number(sData[i][idColIdx]) === Number(data.id)) {
      targetRowIdx = i + 1; // 1-indexed row number in sheet
      break;
    }
  }
  
  if (targetRowIdx === -1) {
    return { success: false, error: 'Student ID not found' };
  }
  
  // Merge and calculate
  var student = {};
  for (var j = 0; j < headers.length; j++) {
    student[headers[j]] = sData[targetRowIdx - 1][j];
  }
  
  for (var key in data) {
    student[key] = data[key];
  }
  
  student.fee_pending = Number(student.fee_applicable || 0) - Number(student.fee_paid || 0);
  student.scholarship_pending = Number(student.scholarship_expected || 0) - Number(student.scholarship_received || 0);
  
  // Write back row
  var row = [];
  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    var val = student[header];
    row.push(val !== undefined ? val : '');
  }
  
  sheet.getRange(targetRowIdx, 1, 1, headers.length).setValues([row]);
  return { success: true, student: student };
}

function recordPayment(ss, data) {
  var sheet = ss.getSheetByName('Transactions');
  var sData = sheet.getDataRange().getValues();
  var headers = sData[0];
  
  var maxId = 0;
  for (var i = 1; i < sData.length; i++) {
    var cid = Number(sData[i][0]);
    if (cid > maxId) maxId = cid;
  }
  var newId = maxId + 1;
  
  data.id = newId;
  data.date = data.date || new Date().toISOString().substring(0, 10);
  data.status = data.status || 'Unverified';
  data.verified = Number(data.verified) === 1 ? 1 : 0;
  
  var row = [];
  for (var j = 0; j < headers.length; j++) {
    var header = headers[j];
    var val = data[header];
    row.push(val !== undefined ? val : '');
  }
  sheet.appendRow(row);
  
  // Update student paid amount
  if (data.student_id) {
    var studentsSheet = ss.getSheetByName('Students');
    var studentsData = studentsSheet.getDataRange().getValues();
    var sHeaders = studentsData[0];
    var idColIdx = 0;
    
    var sRowIdx = -1;
    for (var i = 1; i < studentsData.length; i++) {
      if (Number(studentsData[i][idColIdx]) === Number(data.student_id)) {
        sRowIdx = i + 1;
        break;
      }
    }
    
    if (sRowIdx !== -1) {
      var student = {};
      for (var j = 0; j < sHeaders.length; j++) {
        student[sHeaders[j]] = studentsData[sRowIdx - 1][j];
      }
      
      var amount = Number(data.amount || 0);
      if (data.type === 'Scholarship') {
        student.scholarship_received = Number(student.scholarship_received || 0) + amount;
        student.scholarship_pending = Number(student.scholarship_expected || 0) - student.scholarship_received;
      } else if (data.type !== 'Other') {
        student.fee_paid = Number(student.fee_paid || 0) + amount;
        student.fee_pending = Number(student.fee_applicable || 0) - Number(student.scholarship_expected || 0) - student.fee_paid;
      }
      
      var sRow = [];
      for (var j = 0; j < sHeaders.length; j++) {
        sRow.push(student[sHeaders[j]] !== undefined ? student[sHeaders[j]] : '');
      }
      studentsSheet.getRange(sRowIdx, 1, 1, sHeaders.length).setValues([sRow]);
    }
  }
  
  return { success: true, transaction: data };
}

function bulkSync(ss, records) {
  var count = 0;
  records.forEach(function(r) {
    if (!r.action) return;
    if (r.action === 'addStudent') {
      addStudent(ss, r);
      count++;
    } else if (r.action === 'updateStudent') {
      updateStudent(ss, r);
      count++;
    } else if (r.action === 'recordPayment') {
      recordPayment(ss, r);
      count++;
    }
  });
  return { success: true, imported: count };
}

function sendNotification(ss, data) {
  try {
    var adminEmail = data.recipientEmail || '';
    if (!adminEmail) {
      var settings = getSettings(ss).settings || {};
      adminEmail = settings.gmail_address || '';
    }
    
    if (!adminEmail) {
      return { success: false, error: 'No admin email configured.' };
    }
    
    var subject = data.subject || 'FeeNext Alert';
    var message = data.message || 'Notification message.';
    
    var attachments = [];
    if (data.backupTrigger) {
      var file = DriveApp.getFileById(ss.getId());
      var pdfBlob = file.getAs('application/pdf');
      pdfBlob.setName(ss.getName() + ' - Backup.pdf');
      attachments.push(pdfBlob);
    }
    
    MailApp.sendEmail({
      to: adminEmail,
      subject: subject,
      body: message,
      attachments: attachments
    });
    
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function verifyTransaction(ss, txnId, verified) {
  var sheet = ss.getSheetByName('Transactions');
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var idColIdx = headers.indexOf('id');
  var txnIdColIdx = headers.indexOf('transaction_id');
  var verifiedColIdx = headers.indexOf('verified');
  var statusColIdx = headers.indexOf('status');
  
  var targetRowIdx = -1;
  for (var i = 1; i < data.length; i++) {
    if (Number(data[i][idColIdx]) === Number(txnId) || String(data[i][txnIdColIdx]) === String(txnId)) {
      targetRowIdx = i + 1;
      break;
    }
  }
  
  if (targetRowIdx === -1) {
    return { success: false, error: 'Transaction not found' };
  }
  
  var valVerified = Number(verified) === 1 ? 1 : 0;
  var valStatus = valVerified ? 'Verified' : 'Unverified';
  
  sheet.getRange(targetRowIdx, verifiedColIdx + 1).setValue(valVerified);
  sheet.getRange(targetRowIdx, statusColIdx + 1).setValue(valStatus);
  return { success: true };
}

function saveSettings(ss, settings) {
  var sheet = ss.getSheetByName('Settings');
  var data = sheet.getDataRange().getValues();
  
  for (var key in settings) {
    var val = String(settings[key]);
    var foundIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        foundIdx = i + 1;
        break;
      }
    }
    if (foundIdx !== -1) {
      sheet.getRange(foundIdx, 2).setValue(val);
    } else {
      sheet.appendRow([key, val]);
    }
  }
  return { success: true };
}

function setSecurityPin(ss, data) {
  var pin = data.pin;
  if (!pin) return { success: false, error: 'PIN empty' };
  saveSettings(ss, { 'security_pin': pin });
  return { success: true };
}

function verifyPin(ss, data) {
  var pin = data.pin;
  var settings = getSettings(ss).settings || {};
  var correctPin = settings.security_pin || '0000';
  if (String(pin) === String(correctPin)) {
    return { success: true };
  }
  return { success: false, error: 'Incorrect PIN' };
}

function restoreBackup(ss, students, transactions) {
  if (students && students.length > 0) {
    var sSheet = ss.getSheetByName('Students');
    sSheet.clearContents();
    var sHeaders = [
      'id', 'serial_no', 'name', 'gender', 'student_mobile', 'parent_mobile', 
      'student_email', 'parent_email', 'permanent_address', 'local_address', 
      'year_of_admission', 'category', 'current_class', 'current_academic_year', 
      'fee_applicable', 'fee_paid', 'fee_pending', 'scholarship_expected', 
      'scholarship_received', 'scholarship_pending', 'remark', 'admission_cancelled'
    ];
    sSheet.appendRow(sHeaders);
    students.forEach(function(s) {
      var row = [];
      sHeaders.forEach(function(h) {
        row.push(s[h] !== undefined ? s[h] : '');
      });
      sSheet.appendRow(row);
    });
  }
  
  if (transactions && transactions.length > 0) {
    var tSheet = ss.getSheetByName('Transactions');
    tSheet.clearContents();
    var tHeaders = [
      'id', 'student_id', 'student_name', 'transaction_id', 'date', 'amount', 
      'bank_name', 'source', 'status', 'verified', 'raw_remark', 'type', 'other_reason'
    ];
    tSheet.appendRow(tHeaders);
    transactions.forEach(function(t) {
      var row = [];
      tHeaders.forEach(function(h) {
        row.push(t[h] !== undefined ? t[h] : '');
      });
      tSheet.appendRow(row);
    });
  }
  
  return { success: true };
}
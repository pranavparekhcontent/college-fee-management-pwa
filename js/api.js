/**
 * Fee Manager PWA — Centralized API Module
 * Handles all communication with Google Apps Script + offline queue
 * Pattern: Same as Smart Attendance Central API
 */

const API = (() => {
      const PENDING_KEY = 'fee_pending_sync';
      const CACHE_PREFIX = 'fee_cache_';
      const MAX_RETRIES = 3;
      const RETRY_DELAY_MS = 2000;

      function _getBaseUrl() {
            return (window.appStartContext && window.appStartContext.serverUrl) || '';
      }

      // ─── Internal HTTP helpers ───

      async function _get(action, params = {}) {
            if (window.appStartContext && window.appStartContext.sheetId) {
                  params.sheetId = window.appStartContext.sheetId;
            }
            let url = _getBaseUrl() + '?action=' + encodeURIComponent(action) + '&ts=' + Date.now();
            for (const k in params) {
                  url += '&' + encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
            }

            let lastErr;
            for (let i = 0; i < MAX_RETRIES; i++) {
                  try {
                        const res = await fetch(url, { method: 'GET', redirect: 'follow', cache: 'no-store' });
                        if (!res.ok) throw new Error('HTTP ' + res.status);
                        let data = await res.json();
                        if (data && data.success && data.data && typeof data.data === 'object') {
                              data = { ...data, ...data.data };
                        }
                        return data;
                  } catch (err) {
                        lastErr = err;
                        if (i < MAX_RETRIES - 1) await _sleep(RETRY_DELAY_MS * (i + 1));
                  }
            }
            throw lastErr;
      }

      async function _post(action, body) {
            let url = _getBaseUrl() + '?action=' + encodeURIComponent(action);
            if (window.appStartContext && window.appStartContext.sheetId) {
                  if (typeof body === 'object' && body !== null) {
                        body.sheetId = window.appStartContext.sheetId;
                  }
            }

            let lastErr;
            for (let i = 0; i < MAX_RETRIES; i++) {
                  try {
                        const res = await fetch(url, {
                              method: 'POST',
                              body: JSON.stringify(body),
                              headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                              redirect: 'follow'
                        });
                        if (!res.ok) throw new Error('HTTP ' + res.status);
                        let data = await res.json();
                        if (data && data.success && data.data && typeof data.data === 'object') {
                              data = { ...data, ...data.data };
                        }
                        return data;
                  } catch (err) {
                        lastErr = err;
                        if (i < MAX_RETRIES - 1) await _sleep(RETRY_DELAY_MS * (i + 1));
                  }
            }
            throw lastErr;
      }

      function _sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

      // ─── Cache ───

      function _setCache(key, data) {
            try {
                  localStorage.setItem(CACHE_PREFIX + key, JSON.stringify({ ts: Date.now(), data }));
            } catch (e) { /* quota exceeded */ }
      }

      function _getCache(key) {
            try {
                  const raw = localStorage.getItem(CACHE_PREFIX + key);
                  if (!raw) return null;
                  return JSON.parse(raw).data;
            } catch { return null; }
      }

      // ─── Pending queue ───

      function _getPending() {
            try {
                  const raw = localStorage.getItem(PENDING_KEY);
                  return raw ? JSON.parse(raw) : [];
            } catch { return []; }
      }

      function _setPending(arr) {
            localStorage.setItem(PENDING_KEY, JSON.stringify(arr));
      }

      // Auto-sync on reconnect
      if (typeof window !== 'undefined') {
            window.addEventListener('online', () => {
                  const pending = _getPending();
                  if (pending.length > 0) {
                        _post('bulkSync', { records: pending }).then(r => {
                              if (r && r.success) {
                                    localStorage.removeItem(PENDING_KEY);
                                    if (window.Toast) Toast.show('✅ Synced ' + pending.length + ' offline records', 'success');
                              }
                        });
                  }
            });
      }

      // ─── Offline cache updates ───

      function _updateCacheWithOfflineRecord(action, record) {
            const cached = _getCache('all_data');
            if (!cached || !cached.success) return { success: true, offline: true };

            if (action === 'addStudent') {
                  const maxId = cached.students.reduce((max, s) => Math.max(max, Number(s.id || 0)), 0);
                  const newStudent = {
                        ...record,
                        id: maxId + 1,
                        serial_no: maxId + 1,
                        fee_pending: Number(record.fee_applicable || 0) - Number(record.fee_paid || 0),
                        scholarship_pending: Number(record.scholarship_expected || 0) - Number(record.scholarship_received || 0)
                  };
                  cached.students.push(newStudent);
                  _setCache('all_data', cached);
                  return { success: true, student: newStudent, offline: true };
            } else if (action === 'updateStudent') {
                  const idx = cached.students.findIndex(s => Number(s.id) === Number(record.id));
                  if (idx !== -1) {
                        const student = { ...cached.students[idx], ...record };
                        student.fee_pending = Number(student.fee_applicable || 0) - Number(student.fee_paid || 0);
                        student.scholarship_pending = Number(student.scholarship_expected || 0) - Number(student.scholarship_received || 0);
                        cached.students[idx] = student;
                        _setCache('all_data', cached);
                        return { success: true, student: student, offline: true };
                  }
            } else if (action === 'recordPayment') {
                  const maxTxnId = cached.transactions.reduce((max, t) => Math.max(max, Number(t.id || 0)), 0);
                  const newTxn = {
                        ...record,
                        id: maxTxnId + 1,
                        date: record.date || new Date().toISOString().substring(0, 10),
                        status: record.status || 'Unverified',
                        verified: Number(record.verified) === 1 ? 1 : 0
                  };
                  cached.transactions.push(newTxn);

                  if (record.student_id) {
                        const sIdx = cached.students.findIndex(s => Number(s.id) === Number(record.student_id));
                        if (sIdx !== -1) {
                              const student = cached.students[sIdx];
                              const amount = Number(record.amount || 0);
                              if (record.type === 'Scholarship') {
                                    student.scholarship_received = Number(student.scholarship_received || 0) + amount;
                                    student.scholarship_pending = Number(student.scholarship_expected || 0) - student.scholarship_received;
                              } else if (record.type !== 'Other') {
                                    student.fee_paid = Number(student.fee_paid || 0) + amount;
                                    student.fee_pending = Number(student.fee_applicable || 0) - Number(student.scholarship_expected || 0) - student.fee_paid;
                              }
                        }
                  }

                  _setCache('all_data', cached);
                  return { success: true, transaction: newTxn, offline: true };
            }
            return { success: true, offline: true };
      }

      async function _postOrQueue(action, body) {
            if (typeof navigator !== 'undefined' && !navigator.onLine) {
                  const pending = _getPending();
                  pending.push({ ...body, action: action });
                  _setPending(pending);
                  return _updateCacheWithOfflineRecord(action, body);
            }
            try {
                  return await _post(action, body);
            } catch (err) {
                  const pending = _getPending();
                  pending.push({ ...body, action: action });
                  _setPending(pending);
                  return _updateCacheWithOfflineRecord(action, body);
            }
      }

      // ─── Public API functions ───

      async function getAllData() {
            try {
                  const data = await _get('getAllData');
                  if (data && data.success) {
                        _setCache('all_data', data);
                  }
                  return data;
            } catch (err) {
                  const cached = _getCache('all_data');
                  if (cached) return cached;
                  throw err;
            }
      }

      async function getAllDataFromUrl(url) {
            let cleanUrl = url;
            if (window.appStartContext && window.appStartContext.sheetId) {
                  cleanUrl += (cleanUrl.includes('?') ? '&' : '?') + 'sheetId=' + encodeURIComponent(window.appStartContext.sheetId);
            }
            cleanUrl += (cleanUrl.includes('?') ? '&' : '?') + 'action=getAllData&ts=' + Date.now();
            const res = await fetch(cleanUrl);
            if (!res.ok) throw new Error('HTTP ' + res.status);
            return res.json();
      }

      async function getStudents() {
            return _get('getStudents');
      }

      async function addStudent(student) {
            return _postOrQueue('addStudent', student);
      }

      async function updateStudent(student) {
            return _postOrQueue('updateStudent', student);
      }

      async function getFeeStructure() {
            return _get('getFeeStructure');
      }

      async function recordPayment(payment) {
            return _postOrQueue('recordPayment', payment);
      }

      async function getTransactions() {
            return _get('getTransactions');
      }

      async function getDashboard() {
            return _get('getDashboard');
      }

      async function uploadStudents(rows) {
            return _post('uploadStudents', { rows });
      }

      async function sendNotification(payload) {
            return _post('sendNotification', payload);
      }

      async function exportBackup() {
            return _get('exportBackup');
      }

      async function getReport(reportType) {
            return _get('getReport', { reportType });
      }

      async function verifyTransaction(transactionId, verified) {
            return _post('verifyTransaction', { transactionId, verified });
      }

      async function saveSettings(settings) {
            return _post('saveSettings', { settings });
      }

      async function restoreBackup(students, transactions) {
            return _post('restoreBackup', { students, transactions });
      }

      return {
            getAllData,
            getAllDataFromUrl,
            getStudents,
            addStudent,
            updateStudent,
            getFeeStructure,
            recordPayment,
            getTransactions,
            getDashboard,
            uploadStudents,
            sendNotification,
            exportBackup,
            getReport,
            verifyTransaction,
            saveSettings,
            restoreBackup
      };
})();

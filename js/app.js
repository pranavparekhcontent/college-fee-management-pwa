/**
 * Fee Manager PWA — Core App Logic (js/app.js)
 * Coordinates navigation, dashboard visualization, data table rendering,
 * client-side Excel import/export (SheetJS), and client-side PDF parsing (PDF.js).
 */

const Toast = {
  show(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;

    let icon = 'ℹ️';
    if (type === 'success') icon = '✅';
    if (type === 'error') icon = '❌';
    if (type === 'warning') icon = '⚠️';

    toast.innerHTML = `<span style="margin-right: 8px;">${icon}</span><span>${message}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(-20px)';
      toast.style.transition = 'all 0.4s ease';
      setTimeout(() => toast.remove(), 400);
    }, 3500);
  }
};

/**
 * Safely converts any value (number, null, undefined, boolean, string)
 * to a lowercase string. Prevents runtime crashes from calling
 * .toLowerCase() on non-string types.
 */
function safeLower(value) {
  return String(value ?? '').toLowerCase();
}

/**
 * Extracts meaningful search keywords from a natural language prompt
 * by stripping common question/filler words.
 * e.g., "who is doing MBBS" → ["mbbs"]
 * e.g., "neha patil record" → ["neha", "patil"]
 * e.g., "show SC students in B.Pharm 1st Year" → ["sc", "b.pharm", "1st", "year"]
 */
const NOISE_WORDS = new Set([
  'who', 'what', 'where', 'when', 'how', 'why', 'which',
  'is', 'are', 'was', 'were', 'am', 'be', 'been', 'being',
  'do', 'does', 'did', 'doing', 'done',
  'has', 'have', 'had', 'having',
  'will', 'would', 'shall', 'should', 'may', 'might', 'can', 'could',
  'the', 'a', 'an', 'this', 'that', 'these', 'those',
  'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by', 'from', 'about',
  'and', 'or', 'but', 'not', 'no', 'nor',
  'i', 'me', 'my', 'we', 'our', 'you', 'your',
  'it', 'its', 'they', 'them', 'their',
  'show', 'list', 'find', 'get', 'give', 'tell', 'display',
  'fetch', 'search', 'look', 'lookup', 'check', 'see',
  'all', 'any', 'every', 'each', 'some',
  'student', 'students', 'record', 'records', 'data', 'details', 'info', 'information',
  'please', 'thanks', 'thank', 'hello', 'hi', 'hey',
  'fee', 'fees', 'payment', 'payments', 'pending', 'paid', 'due', 'dues',
  'scholarship', 'scholarships',
  'class', 'course', 'enrolled', 'studying', 'admission',
  'category', 'caste', 'type', 'group', 'department', 'dept',
  'year', 'sem', 'semester',
  'male', 'female', 'gender',
  'total', 'count', 'number', 'many', 'much',
  'name', 'called', 'named',
  'there', 'here', 'only', 'just', 'also', 'us', 'sir', 'madam', 'mam',
  // Domain-specific action/context words
  'transaction', 'transactions', 'history', 'receipt', 'receipts',
  'amount', 'amounts', 'balance', 'status', 'report', 'reports',
  'summary', 'overview', 'profile', 'account', 'slip', 'challan',
  'collection', 'collections', 'refund', 'refunds', 'concession',
  'result', 'results', 'detail', 'exam', 'marks',
  'utr', 'id', 'no', 'roll', 'serial', 'adm', 'rollno', 'admno',
  // Common Hindi particles (bilingual clerks)
  'ka', 'ki', 'ke', 'ko', 'se', 'hai', 'hain', 'kya', 'kaise',
  'wala', 'wale', 'wali', 'mein', 'par', 'aur', 'ya',
  'batao', 'bataye', 'dikhao', 'dikhaye', 'bolo', 'boliye',
  'kaun', 'kitne', 'kitna', 'kitni', 'kon', 'kiska', 'kiski',
  // Additional requested Hindi words
  'kaunsa', 'kaunse', 'konsa', 'mujhe', 'humko', 'unhe', 'pata', 'batana', 'karna'
]);

function extractSearchKeywords(prompt) {
  return safeLower(prompt)
    .replace(/[^a-z0-9.\s-]/g, ' ')  // keep dots and hyphens for "B.Pharm", "1st-year"
    .split(/\s+/)
    .filter(word => word.length > 0 && !NOISE_WORDS.has(word));
}

const App = (() => {
  // Global State
  const state = {
    students: [],
    transactions: [],
    settings: {},
    collegeName: '',
    managementName: '',

    outputSheetId: '',
    activeView: 'dashboard',
    dashboardMode: 'collection', // 'collection' or 'pending'
    filters: {
      class: '',
      year: ''
    },
    searchQuery: '',
    showCancelled: false,
    charts: {}
  };

  let __liveSyncTimer = null;
  const __startLiveSync = () => {
    if (__liveSyncTimer) return;
    const lockOverlay = document.getElementById('pin-lock-overlay');
    if (lockOverlay && lockOverlay.style.display === 'flex') return;
    __liveSyncTimer = setInterval(async () => {
      if (document.hidden || !navigator.onLine) return;
      try {
        const changed = await refreshData({ silent: true });
        if (changed) {
          populatePillFilters();
          populateExportOptions();
          switchView(state.activeView);
          Toast.show('Live Sync: Database updated.', 'info');
        }
      } catch (_) {}
    }, 10000);
  };

  const __stopLiveSync = () => {
    if (__liveSyncTimer) {
      clearInterval(__liveSyncTimer);
      __liveSyncTimer = null;
    }
  };

  // List of standard column checkboxes for Custom Export
  const EXPORT_COLUMNS = [
    { key: 'serial_no', label: 'Admission No.' },
    { key: 'name', label: 'Full Name' },
    { key: 'gender', label: 'Gender' },
    { key: 'student_mobile', label: 'Student Mobile' },
    { key: 'parent_mobile', label: 'Parent Mobile' },
    { key: 'student_email', label: 'Student Email' },
    { key: 'parent_email', label: 'Parent Email' },
    { key: 'local_address', label: 'Local Address' },
    { key: 'permanent_address', label: 'Permanent Address' },
    { key: 'year_of_admission', label: 'Admission Year' },
    { key: 'category', label: 'Category' },
    { key: 'current_class', label: 'Current Class' },
    { key: 'current_academic_year', label: 'Academic Year' },
    { key: 'fee_applicable', label: 'Fee Applicable' },
    { key: 'fee_paid', label: 'Fee Paid' },
    { key: 'fee_pending', label: 'Fee Pending' },
    { key: 'scholarship_expected', label: 'Scholarship Expected' },
    { key: 'scholarship_received', label: 'Scholarship Received' },
    { key: 'scholarship_pending', label: 'Scholarship Pending' },
    { key: 'remark', label: 'Remarks' }
  ];

  // ─── PIN SECURITY MODULE ───
  function getAppPin() {
    const p = String(state.settings.app_pin || '').trim();
    return p.length === 4 ? p : '1234'; // fallback default
  }

  // Per-action gate. Usage:  if (!(await requirePin())) return;
  function requirePin() {
    return new Promise(resolve => {
      const correct = getAppPin();
      // Reuse the confirm modal area as a quick PIN prompt
      const title = document.getElementById('confirm-title');
      const body  = document.getElementById('confirm-body');
      const backdrop = document.getElementById('confirm-backdrop');
      const okBtn = document.getElementById('confirm-ok');
      const cancelBtn = document.getElementById('confirm-cancel');

      title.textContent = 'Enter PIN';
      body.innerHTML = `
        <div style="display:flex;flex-direction:column;gap:8px;">
          <div>Enter your 4-digit PIN to authorize this action.</div>
          <input id="action-pin-input" type="password" inputmode="numeric" maxlength="4"
            style="text-align:center;font-size:1.4rem;letter-spacing:0.5rem;padding:8px;
            background:var(--bg-primary);border:1px solid var(--border);border-radius:8px;color:var(--text-primary);">
          <div id="action-pin-error" style="color:var(--danger);font-size:0.8rem;min-height:14px;"></div>
        </div>`;

      const cleanup = () => {
        backdrop.classList.remove('visible');
        okBtn.replaceWith(okBtn.cloneNode(true));
        cancelBtn.replaceWith(cancelBtn.cloneNode(true));
      };
      const newOk = document.getElementById('confirm-ok');
      const newCancel = document.getElementById('confirm-cancel');

      const tryConfirm = () => {
        const val = (document.getElementById('action-pin-input').value || '').trim();
        if (val === correct) { cleanup(); resolve(true); }
        else { document.getElementById('action-pin-error').textContent = 'Incorrect PIN.'; }
      };
      newOk.onclick = tryConfirm;
      newCancel.onclick = () => { cleanup(); resolve(false); };

      backdrop.classList.add('visible');
      setTimeout(() => {
        const inp = document.getElementById('action-pin-input');
        if (inp) { inp.focus(); inp.addEventListener('keydown', e => { if (e.key === 'Enter') tryConfirm(); }); }
      }, 100);
    });
  }

  // Full-screen blocking lock (used at app start AND for Lock button)
  function showLockScreen() {
    return new Promise(resolve => {
      const correct = getAppPin();
      const overlay = document.getElementById('pin-lock-overlay');
      const input = document.getElementById('pin-lock-input');
      const errEl = document.getElementById('pin-lock-error');
      const btn = document.getElementById('pin-lock-btn');
      if (!overlay) { resolve(true); return; }

      overlay.style.display = 'flex';
      input.value = '';
      errEl.textContent = '';
      setTimeout(() => input.focus(), 100);

      const tryUnlock = () => {
        if ((input.value || '').trim() === correct) {
          overlay.style.display = 'none';
          resolve(true);
          __startLiveSync();
        } else {
          errEl.textContent = 'Incorrect PIN. Try again.';
          input.value = '';
          input.focus();
        }
      };
      btn.onclick = tryUnlock;
      input.onkeydown = e => { if (e.key === 'Enter') tryUnlock(); };
    });
  }

  function lockApp() { 
    __stopLiveSync();
    showLockScreen(); 
  }

  // Silent automated backup triggers for Launch and Close/Unload
  async function triggerAutoBackup(triggerType) {
    console.log(`[Safety] Automated backup triggered on: ${triggerType}`);
    try {
      const baseUrl = (window.appStartContext && window.appStartContext.serverUrl) || '';
      if (!baseUrl || baseUrl === "DEMO_MODE") return;

      const now = new Date();
      const pad = (n) => String(n).padStart(2, '0');
      const dateStr = `${pad(now.getDate())}/${pad(now.getMonth() + 1)}/${now.getFullYear()}`;
      const timeStr = `${pad(now.getHours())}:${pad(now.getMinutes())}:${pad(now.getSeconds())}`;
      const subject = `FeeFlow Auto Backup [${triggerType}] - ${dateStr} ${timeStr}`;

      const payload = {
        recipientGroup: 'admin',
        method: 'email',
        message: `Safety Backup: Automatically generated on app ${triggerType.toLowerCase()}.`,
        backupTrigger: triggerType,
        subject: subject
      };

      if (triggerType === 'Close') {
        const action = 'sendNotification';
        if (window.appStartContext && window.appStartContext.sheetId) {
          payload.sheetId = window.appStartContext.sheetId;
        }
        payload.action = action;

        const url = baseUrl + '?action=' + encodeURIComponent(action);
        
        // Standard keepalive fetch prevents CORS preflight OPTIONS request issues
        fetch(url, {
          method: 'POST',
          body: JSON.stringify(payload),
          headers: { 'Content-Type': 'text/plain;charset=utf-8' },
          redirect: 'follow',
          keepalive: true
        }).catch(() => {});
      } else {
        // Normal launch notification API call
        await API.sendNotification(payload);
      }
    } catch (e) {
      console.warn('[Safety] Auto backup failed to send:', e);
    }
  }

  // ─── INITIALIZATION ───

  async function initFromEngine(context) {
    console.log("Initializing App with context:", context);
    state.collegeName = context.collegeName || 'College Fee Management';
    state.managementName = context.managementName || '';
    state.outputSheetId = context.outputSheetId || '';

    // Set header & sidebar details
    document.getElementById('sidebar-college').textContent = state.collegeName;

    document.getElementById('header-username').textContent = 'Admin (' + state.collegeName.substring(0, 8) + '...)';

    const initials = state.collegeName.split(' ').map(w => w[0]).join('').substring(0, 2).toUpperCase();
    document.getElementById('header-avatar').textContent = initials || 'AD';

    // Populate data from pre-fetched boot context
    const bootData = context.fetchedData?.allData;
    if (bootData && bootData.success) {
      state.students = bootData.students || [];
      state.transactions = bootData.transactions || [];
      state.settings = bootData.settings || {};
    } else {
      // Fallback: fetch manually if engine boot data was partial
      await refreshData();
    }

    // Require PIN at app start before showing anything
    await showLockScreen();

    // Clean up the AppStart overlay
    const appStartObj = window.AppStart || (typeof AppStart !== 'undefined' ? AppStart : null);
    if (appStartObj && typeof appStartObj.cleanup === 'function') {
      await appStartObj.cleanup();
    }

    setupEventListeners();
    initClock();
    initTheme();

    // Force refresh data on every launch to ensure latest data from Google Sheets BEFORE rendering pills
    await refreshData();

    // Populate pill filters and export options with fresh data
    populatePillFilters();
    populateExportOptions();

    // Initial render with fresh data
    switchView('dashboard');

    // Start Live Sync and bind tab visibility change events
    __startLiveSync();
    document.addEventListener('visibilitychange', () => {
      if (document.hidden) {
        __stopLiveSync();
      } else {
        __startLiveSync();
      }
    });

    // Silent safety backup on launch
    triggerAutoBackup('Launch');
  }

  async function refreshData(options = {}) {
    const silent = options.silent || false;
    try {
      const res = await API.getAllData();
      if (res.success) {
        const newStudents = res.students || [];
        const newTxns = res.transactions || [];
        const newSettings = res.settings || {};

        // Smart change detection to avoid UI redraw/flicker
        const oldHash = JSON.stringify({ s: state.students, t: state.transactions, st: state.settings });
        const newHash = JSON.stringify({ s: newStudents, t: newTxns, st: newSettings });
        const changed = oldHash !== newHash;

        if (changed || state.students.length === 0) {
          state.students = newStudents;
          state.transactions = newTxns;
          state.settings = newSettings;
          return true; // Changes detected
        }
        return false; // No changes
      } else {
        if (!silent) {
          Toast.show(res.error || 'Failed to fetch data.', 'error');
        }
      }
    } catch (e) {
      console.error(e);
      if (!silent) {
        Toast.show('Network error while refreshing database.', 'error');
      }
    }
    return false;
  }

  // ─── CLOCK & THEME ───

  function initClock() {
    const clockEl = document.getElementById('header-clock');
    const update = () => {
      const d = new Date();
      clockEl.textContent = d.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    };
    setInterval(update, 1000);
    update();
  }

  function initTheme() {
    const saved = localStorage.getItem('theme_preference') || 'dark';
    document.documentElement.setAttribute('data-theme', saved);
    updateThemeUI(saved);

    document.getElementById('theme-toggle').addEventListener('click', () => {
      const current = document.documentElement.getAttribute('data-theme');
      const next = current === 'light' ? 'dark' : 'light';
      document.documentElement.setAttribute('data-theme', next);
      localStorage.setItem('theme_preference', next);
      updateThemeUI(next);
      // Re-render dashboard to apply theme colors to ECharts
      renderDashboard();
      // Resize charts for layout adjust
      setTimeout(resizeCharts, 200);
    });
  }

  function updateThemeUI(theme) {
    const textEl = document.getElementById('theme-text');
    const iconEl = document.getElementById('theme-icon');
    if (textEl) {
      textEl.textContent = theme === 'light' ? 'Dark Mode' : 'Light Mode';
    }
    if (iconEl) {
      iconEl.className = theme === 'light' ? 'ph-bold ph-moon theme-toggle-icon' : 'ph-bold ph-sun theme-toggle-icon';
    }

    // Support the compact toggle state
    const toggleContainer = document.getElementById('theme-toggle');
    if (toggleContainer) {
      if (theme === 'light') {
        toggleContainer.classList.add('theme-light');
        toggleContainer.classList.remove('theme-dark');
      } else {
        toggleContainer.classList.add('theme-dark');
        toggleContainer.classList.remove('theme-light');
      }
    }
  }

  // ─── VIEW ROUTING & SIDEBAR ───

  function setupEventListeners() {
    // Sidebar nav clicks
    document.querySelectorAll('.nav-item').forEach(item => {
      item.addEventListener('click', (e) => {
        const view = e.currentTarget.getAttribute('data-view');
        switchView(view);

        // Mobile auto-collapse sidebar
        if (window.innerWidth <= 768) {
          document.getElementById('sidebar').classList.remove('visible');
        }
      });
    });

    // Lock App button
    const btnLock = document.getElementById('btn-lock-app');
    if (btnLock) btnLock.addEventListener('click', () => lockApp());

    // Sync Excel Data button
    const btnSync = document.getElementById('btn-sync-data');
    if (btnSync) {
      btnSync.addEventListener('click', async () => {
        btnSync.classList.add('syncing');
        btnSync.disabled = true;
        const span = btnSync.querySelector('span');
        span.textContent = 'Syncing...';
        Toast.show('Fetching latest data from server...', 'info');

        try {
          await refreshData();
          populatePillFilters();
          populateExportOptions();
          switchView(state.activeView);
          Toast.show('Sync completed successfully!', 'success');
        } catch (err) {
          console.error(err);
          Toast.show('Failed to sync: ' + (err.message || err), 'error');
        } finally {
          btnSync.classList.remove('syncing');
          btnSync.disabled = false;
          span.textContent = 'Sync';
        }
      });
    }

    // Hamburger button
    document.getElementById('hamburger-btn').addEventListener('click', () => {
      document.getElementById('sidebar').classList.toggle('visible');
    });

    // Real-time Student Search
    document.getElementById('student-search').addEventListener('input', (e) => {
      state.searchQuery = e.target.value;
      renderStudents();
    });

    // Show Cancelled Checkbox
    document.getElementById('show-cancelled-check').addEventListener('change', (e) => {
      state.showCancelled = e.target.checked;
      renderStudents();
    });

    // Dashboard toggle buttons
    const btnCollection = document.getElementById('btn-toggle-collection');
    const btnPending = document.getElementById('btn-toggle-pending');
    if (btnCollection && btnPending) {
      btnCollection.addEventListener('click', () => {
        if (state.dashboardMode === 'collection') return;
        state.dashboardMode = 'collection';
        btnCollection.classList.add('active');
        btnCollection.classList.add('toggle-success');
        btnPending.classList.remove('active');
        btnPending.classList.remove('toggle-warning');
        renderMainBarChart();
      });
      btnPending.addEventListener('click', () => {
        if (state.dashboardMode === 'pending') return;
        state.dashboardMode = 'pending';
        btnPending.classList.add('active');
        btnPending.classList.add('toggle-warning');
        btnCollection.classList.remove('active');
        btnCollection.classList.remove('toggle-success');
        renderMainBarChart();
      });
      // Apply initial color to the default active toggle
      btnCollection.classList.add('toggle-success');
    }

    // Modal close when clicking backdrop
    document.getElementById('modal-backdrop').addEventListener('click', (e) => {
      if (e.target.id === 'modal-backdrop') {
        closeModal();
      }
    });

    // New Admission Clear Button
    document.getElementById('btn-adm-clear').addEventListener('click', () => {
      document.getElementById('admission-form').reset();
    });

    // New Admission Form Submit
    document.getElementById('admission-form').addEventListener('submit', async (e) => {
      e.preventDefault();

      if (!(await requirePin())) return;

      const payload = {
        serial_no: parseInt(document.getElementById('adm-serial-no').value) || null,
        name: document.getElementById('adm-name').value.trim(),
        gender: document.getElementById('adm-gender').value,
        student_mobile: document.getElementById('adm-student-mobile').value.trim(),
        parent_mobile: document.getElementById('adm-parent-mobile').value.trim(),
        student_email: document.getElementById('adm-student-email').value.trim(),
        parent_email: document.getElementById('adm-parent-email').value.trim(),
        local_address: document.getElementById('adm-local-address').value.trim(),
        permanent_address: document.getElementById('adm-permanent-address').value.trim(),
        year_of_admission: parseInt(document.getElementById('adm-year').value) || new Date().getFullYear(),
        category: document.getElementById('adm-category').value,
        current_class: document.getElementById('adm-class').value,
        current_academic_year: document.getElementById('adm-academic-year').value.trim() || `${new Date().getFullYear()}-${new Date().getFullYear() + 1}`,
        fee_applicable: parseFloat(document.getElementById('adm-fee-applicable').value) || 0,
        fee_paid: parseFloat(document.getElementById('adm-fee-paid').value) || 0,
        scholarship_expected: parseFloat(document.getElementById('adm-sch-expected').value) || 0,
        scholarship_received: parseFloat(document.getElementById('adm-sch-received').value) || 0,
        remark: document.getElementById('adm-remark').value.trim(),
        admission_cancelled: 0
      };

      Toast.show('Registering student...', 'info');
      try {
        const res = await API.addStudent(payload);
        if (res.success) {
          Toast.show('Student registered successfully!', 'success');
          document.getElementById('admission-form').reset();
          await refreshData();
          populatePillFilters();
          switchView('students');
        } else {
          Toast.show(res.error || 'Failed to register student.', 'error');
        }
      } catch (err) {
        Toast.show('Network error while registering student.', 'error');
      }
    });

    // Overall filter mode is handled via Class and Year dynamic "All" pills

    // Excel import file handling
    const uploadZone = document.getElementById('upload-zone-students');
    ['dragenter', 'dragover'].forEach(name => {
      uploadZone.addEventListener(name, e => { e.preventDefault(); uploadZone.classList.add('hover'); });
    });
    ['dragleave', 'drop'].forEach(name => {
      uploadZone.addEventListener(name, e => { e.preventDefault(); uploadZone.classList.remove('hover'); });
    });
    uploadZone.addEventListener('drop', e => {
      const file = e.dataTransfer.files[0];
      if (file) handleExcelImport(file);
    });
    document.getElementById('student-file-input').addEventListener('change', e => {
      const file = e.target.files[0];
      if (file) handleExcelImport(file);
    });

    // Excel template download trigger
    document.getElementById('btn-download-template').addEventListener('click', e => {
      e.preventDefault();
      downloadBlankTemplate();
    });

    // Reports export trigger
    document.getElementById('btn-export-reports').addEventListener('click', () => {
      exportReportsExcel();
    });

    // Custom Export Submit
    document.getElementById('btn-custom-export-submit').addEventListener('click', () => {
      exportCustomExcel();
    });

    // Credentials Form save
    document.getElementById('credentials-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const gmail = document.getElementById('gmail-address').value.trim();
      const pass = document.getElementById('gmail-app-password').value.trim();
      const autoSend = document.getElementById('auto-send-emails').checked ? 'true' : 'false';

      Toast.show('Saving email credentials...', 'info');
      try {
        const res = await API.saveSettings({
          gmail_address: gmail,
          app_password: pass,
          auto_send_emails: autoSend
        });
        if (res.success) {
          Toast.show('Settings saved successfully!', 'success');
          state.settings.gmail_address = gmail;
          state.settings.app_password = pass;
          state.settings.auto_send_emails = autoSend;
        } else {
          Toast.show(res.error || 'Failed to save settings.', 'error');
        }
      } catch (err) {
        Toast.show('Error saving settings to GSheets.', 'error');
      }
    });

    // Communication Form submit (Announcements)
    document.getElementById('communication-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const recipient = document.getElementById('comm-recipient').value;
      const msg = document.getElementById('comm-message').value.trim();

      Toast.show('Preparing announcement email dispatch...', 'info');
      try {
        const res = await API.sendNotification({
          recipientGroup: recipient,
          method: 'email',
          message: msg
        });
        if (res.success) {
          Toast.show(`Announcements sent to ${recipient} via Email!`, 'success');
          document.getElementById('comm-message').value = '';
        } else {
          Toast.show(res.error || 'Dispatch failed.', 'error');
        }
      } catch (err) {
        console.error(err);
        Toast.show(err.message || 'Failed to connect to messaging server.', 'error');
      }
    });

    // Bulk Dues reminders click listener
    document.getElementById('btn-send-fee-dues-reminders').addEventListener('click', async () => {
      const activeUnpaid = state.students.filter(s => {
        return Number(s.admission_cancelled) !== 1 && parseFloat(s.fee_pending) > 0;
      });

      if (activeUnpaid.length === 0) {
        Toast.show('No active students with outstanding fee pending balance!', 'info');
        return;
      }

      const confirmSend = confirm(`Are you sure you want to send outstanding fee reminder emails to all ${activeUnpaid.length} unpaid students?`);
      if (!confirmSend) return;

      Toast.show(`Queueing automated fee dues email alerts for ${activeUnpaid.length} students...`, 'info');
      try {
        const res = await API.sendNotification({
          recipientGroup: 'both',
          method: 'email',
          message: 'Pending fee reminder trigger (Automated Bulk Run)'
        });
        if (res.success) {
          Toast.show(`Dues reminders email dispatch complete! Sent: ${activeUnpaid.length}`, 'success');
        } else {
          Toast.show(res.error || 'Dues reminders dispatch failed.', 'error');
        }
      } catch (err) {
        console.error(err);
        Toast.show('Failed to connect to messaging server.', 'error');
      }
    });

    // Manual backup download
    document.getElementById('btn-manual-backup').addEventListener('click', () => {
      downloadBackupExcel();
    });

    // Email backup request
    document.getElementById('btn-email-backup').addEventListener('click', async () => {
      Toast.show('Triggering spreadsheet backup email...', 'info');
      try {
        const res = await API.sendNotification({
          recipientGroup: 'admin',
          method: 'email',
          message: 'System manual backup request',
          backupTrigger: 'Manual'
        });
        if (res.success) {
          Toast.show('Excel spreadsheet backup sent to admin email!', 'success');
        } else {
          Toast.show(res.error || 'Failed to send backup.', 'error');
        }
      } catch (e) {
        Toast.show('Network error sending backup.', 'error');
      }
    });

    // Restore Backup Excel Upload
    document.getElementById('restore-file-input').addEventListener('change', e => {
      const file = e.target.files[0];
      if (file) handleRestoreBackup(file);
    });

    // Bank Statement PDF Verification Trigger
    document.getElementById('btn-upload-statement').addEventListener('click', () => {
      const input = document.createElement('input');
      input.type = 'file';
      input.accept = '.pdf';
      input.onchange = e => {
        const file = e.target.files[0];
        if (file) handlePdfStatementUpload(file);
      };
      input.click();
    });

    // Window Resize chart layout updates
    window.addEventListener('resize', resizeCharts);

    // --- AI Assistant UI Event Listeners ---
    // Settings panel collapse/expand
    const aiToggle = document.getElementById('ai-settings-toggle');
    const aiContent = document.getElementById('ai-settings-content');
    const aiCaret = document.getElementById('ai-settings-caret');
    if (aiToggle && aiContent && aiCaret) {
      aiToggle.addEventListener('click', () => {
        const isCollapsed = aiContent.style.maxHeight === '0px' || !aiContent.style.maxHeight || aiContent.style.maxHeight === '0';
        if (isCollapsed) {
          aiContent.style.maxHeight = '300px';
          aiCaret.style.transform = 'rotate(180deg)';
        } else {
          aiContent.style.maxHeight = '0px';
          aiCaret.style.transform = 'rotate(0deg)';
        }
      });
    }

    // Toggle mask/unmask API key
    const btnToggleAiKey = document.getElementById('btn-toggle-ai-key');
    const aiApiKeyInput = document.getElementById('ai-api-key');
    if (btnToggleAiKey && aiApiKeyInput) {
      btnToggleAiKey.addEventListener('click', () => {
        const isPassword = aiApiKeyInput.type === 'password';
        aiApiKeyInput.type = isPassword ? 'text' : 'password';
        const icon = btnToggleAiKey.querySelector('i');
        if (icon) {
          icon.className = isPassword ? 'ph-bold ph-eye-slash' : 'ph-bold ph-eye';
        }
      });
    }

    // Save Groq AI Settings
    const btnAiSave = document.getElementById('btn-ai-save');
    if (btnAiSave) {
      btnAiSave.addEventListener('click', async () => {
        const apiKey = document.getElementById('ai-api-key').value.trim();
        const model = 'mixtral-8x7b-32768';

        if (!(await requirePin())) return;

        Toast.show('Saving Groq API credentials...', 'info');
        try {
          const res = await API.saveSettings({
            openrouter_api_key: apiKey,
            openrouter_model: model
          });
          if (res.success) {
            state.settings.openrouter_api_key = apiKey;
            state.settings.openrouter_model = model;
            updateAiStatusBadge();
            Toast.show('Groq API key saved successfully!', 'success');
            // Refresh data to ensure settings are synced across the app
            await refreshData();
            // Update pill filters with fresh data
            populatePillFilters();
            populateExportOptions();
            // Collapse settings panel after successful save
            if (aiContent && aiCaret) {
              aiContent.style.maxHeight = '0px';
              aiCaret.style.transform = 'rotate(0deg)';
            }
          } else {
            Toast.show(res.error || 'Failed to save settings.', 'error');
          }
        } catch (err) {
          console.error(err);
          Toast.show('Error saving AI settings to GSheets.', 'error');
        }
      });
    }

    // Suggested query chips interaction
    document.querySelectorAll('.ai-chip').forEach(chip => {
      chip.addEventListener('click', () => {
        const prompt = chip.getAttribute('data-prompt');
        const input = document.getElementById('ai-chat-input');
        if (input && prompt) {
          input.value = prompt;
          input.focus();
          input.style.height = 'auto';
          input.style.height = input.scrollHeight + 'px';
        }
      });
    });

    // Chat submit form handler
    const aiForm = document.getElementById('ai-chat-form');
    if (aiForm) {
      aiForm.addEventListener('submit', async (e) => {
        e.preventDefault();
        const inputEl = document.getElementById('ai-chat-input');
        if (!inputEl) return;
        const prompt = inputEl.value.trim();
        if (!prompt) return;

        if (!state.settings.openrouter_api_key) {
          Toast.show('Please configure and save your OpenRouter API Key first.', 'warning');
          if (aiContent && aiCaret) {
            aiContent.style.maxHeight = '300px';
            aiCaret.style.transform = 'rotate(180deg)';
          }
          return;
        }

        // Add user message to screen
        appendChatBubble(prompt, 'user');

        // Clear input box
        inputEl.value = '';
        inputEl.style.height = '50px';

        // Add typing indicator
        const typingId = appendTypingIndicator();

        try {
          const responseText = await sendAiMessage(prompt);
          removeTypingIndicator(typingId);
          appendChatBubble(responseText, 'ai');
        } catch (err) {
          removeTypingIndicator(typingId);
          console.error(err);
          appendChatBubble(`⚠️ Error: ${err.message || 'Failed to get response from AI. Please check your key or internet.'}`, 'ai');
        }
      });
    }

    // Textarea enter key submit
    const chatInput = document.getElementById('ai-chat-input');
    if (chatInput) {
      chatInput.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.shiftKey) {
          e.preventDefault();
          const form = document.getElementById('ai-chat-form');
          if (form) {
            const event = new Event('submit', { cancelable: true });
            form.dispatchEvent(event);
          }
        }
      });
    }

    // Auto-backup on app close/unload
    window.addEventListener('beforeunload', () => {
      triggerAutoBackup('Close');
    });
  }

  // ─── INTENT DETECTION LAYER ───
  // Classifies user questions into predefined intents to avoid sending full DB to AI
  const INTENTS = {
    LIST_STUDENTS: 'list_students',
    FEE_INFO: 'fee_info',
    PENDING_DUES: 'pending_dues',
    SCHOLARSHIP_INFO: 'scholarship_info',
    GENERAL_CHAT: 'general_chat'
  };

  async function detectIntent(userMessage) {
    const apiKey = state.settings.openrouter_api_key;
    if (!apiKey) throw new Error('API Key missing');

    const endpoint = 'https://api.groq.com/openai/v1/chat/completions';

    const systemPrompt = `You are an intent classifier for a college fee management chatbot.
Classify the user's message into exactly ONE of these intents:
- list_students: User wants to see student lists (by course, caste, class, etc.)
- fee_info: User wants fee payment details for a specific student
- pending_dues: User wants to know pending fee amounts
- scholarship_info: User wants scholarship details
- general_chat: Any other question not related to the above

Respond with ONLY the intent name, nothing else.`;

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: 'llama-3.1-8b-instant',
        messages: [
          { role: 'system', content: systemPrompt },
          { role: 'user', content: userMessage }
        ],
        max_tokens: 100
      })
    });

    if (!response.ok) {
      const err = await response.text();
      throw new Error(`Intent detection failed: ${err}`);
    }

    const data = await response.json();
    let intent = data?.choices?.[0]?.message?.content?.trim()?.toLowerCase();

    // Validate intent
    if (!Object.values(INTENTS).includes(intent)) {
      intent = INTENTS.GENERAL_CHAT;
    }
    return intent;
  }

  // ─── LOCAL DATA HELPER FUNCTIONS ───
  // All data operations happen in JavaScript, not AI

  function filterStudents(criteria = {}) {
    let results = [...state.students];

    if (criteria.class) {
      results = results.filter(s =>
        safeLower(s.current_class).includes(safeLower(criteria.class))
      );
    }
    if (criteria.category) {
      results = results.filter(s =>
        safeLower(s.category).includes(safeLower(criteria.category))
      );
    }
    if (criteria.year) {
      results = results.filter(s =>
        safeLower(s.current_academic_year).includes(safeLower(criteria.year))
      );
    }
    if (criteria.adm_no) {
      results = results.filter(s =>
        safeLower(s.serial_no).includes(safeLower(criteria.adm_no))
      );
    }

    return results.map(s => ({
      adm_no: s.serial_no,
      name: s.name,
      class: s.current_class,
      category: s.category,
      paid: s.fee_paid,
      pending: s.fee_pending
    }));
  }

  function getFeeInfo(studentIdentifier) {
    if (!studentIdentifier) return null;

    const id = safeLower(studentIdentifier);

    // Try matching by serial_no / admission number first
    let student = state.students.find(s =>
      safeLower(s.serial_no).includes(id)
    );

    // If no ID match, try matching by student name (keyword-based)
    if (!student) {
      const keywords = extractSearchKeywords(studentIdentifier);
      if (keywords.length > 0) {
        student = state.students.find(s => {
          const name = safeLower(s.name);
          return keywords.every(kw => name.includes(kw));
        });
      }
    }

    if (!student) return null;

    const transactions = state.transactions
      .filter(t => String(t.serial_no) === String(student.serial_no))
      .map(t => ({
        amount: t.amount,
        date: t.date,
        method: t.method,
        utr: t.utr
      }));

    return {
      student: {
        adm_no: student.serial_no,
        name: student.name,
        class: student.current_class
      },
      total_paid: student.fee_paid,
      total_pending: student.fee_pending,
      transactions
    };
  }

  function getPendingDues(studentId = null) {
    let results = state.students.filter(s =>
      Number(s.fee_pending) > 0 && Number(s.admission_cancelled) !== 1
    );

    if (studentId) {
      results = results.filter(s => safeLower(s.serial_no).includes(safeLower(studentId)));
    }

    return results.map(s => ({
      adm_no: s.serial_no,
      name: s.name,
      class: s.current_class,
      pending: s.fee_pending
    })).sort((a, b) => b.pending - a.pending);
  }

  function getScholarshipInfo(studentId = null) {
    let results = state.students.filter(s =>
      Number(s.scholarship_pending) > 0 || Number(s.scholarship_expected) > 0
    );

    if (studentId) {
      results = results.filter(s => safeLower(s.serial_no).includes(safeLower(studentId)));
    }

    return results.map(s => ({
      adm_no: s.serial_no,
      name: s.name,
      expected: s.scholarship_expected,
      received: s.scholarship_received,
      pending: s.scholarship_pending
    }));
  }

  function listStudents(filterText = '') {
    const keywords = extractSearchKeywords(filterText);

    const results = state.students
      .filter(s => {
        if (keywords.length === 0) return true;

        // Build a combined searchable string for this student
        const searchable = [
          safeLower(s.name),
          safeLower(s.serial_no),
          safeLower(s.current_class),
          safeLower(s.category),
          safeLower(s.current_academic_year),
          safeLower(s.gender),
          safeLower(s.remark)
        ].join(' ');

        // ALL keywords must match somewhere in the student's fields
        return keywords.every(kw => searchable.includes(kw));
      })
      .filter(s => parseInt(s.admission_cancelled) !== 1) // exclude cancelled
      .slice(0, 20) // Limit to 20 for display
      .map(s => ({
        adm_no: s.serial_no,
        name: s.name,
        class: s.current_class,
        category: s.category
      }));

    return results;
  }

  function formatDataAsText(intent, data) {
    // Convert structured data to plain text for AI formatting
    if (intent === INTENTS.LIST_STUDENTS) {
      if (!data || data.length === 0) return 'No students found.';
      return data.map(s =>
        `${s.adm_no} - ${s.name} (${s.class}) - Category: ${s.category}`
      ).join('\n');
    }
    if (intent === INTENTS.FEE_INFO && data) {
      let text = `Student: ${data.student.name} (${data.student.adm_no})\n`;
      text += `Class: ${data.student.class}\n`;
      text += `Total Paid: ₹${data.total_paid}\n`;
      text += `Total Pending: ₹${data.total_pending}\n`;
      if (data.transactions.length > 0) {
        text += '\nPayment History:\n';
        data.transactions.forEach(t => {
          text += `  - ₹${t.amount} on ${t.date} (${t.method})${t.utr ? ` [UTR: ${t.utr}]` : ''}\n`;
        });
      }
      return text;
    }
    if (intent === INTENTS.PENDING_DUES && data) {
      if (data.length === 0) return 'No pending dues found.';
      let text = `Pending Fee Dues (${data.length} students):\n\n`;
      data.slice(0, 10).forEach(s => {
        text += `  ${s.adm_no} - ${s.name} - ₹${s.pending}\n`;
      });
      if (data.length > 10) text += `\n... and ${data.length - 10} more`;
      return text;
    }
    if (intent === INTENTS.SCHOLARSHIP_INFO && data) {
      if (data.length === 0) return 'No scholarship records found.';
      let text = `Scholarship Information:\n\n`;
      data.slice(0, 10).forEach(s => {
        text += `  ${s.adm_no} - ${s.name}\n`;
        text += `    Expected: ₹${s.expected} | Received: ₹${s.received} | Pending: ₹${s.pending}\n`;
      });
      return text;
    }
    return JSON.stringify(data);
  }

  // Trim chat history to last 6 messages
  function trimChatHistory(messages, maxMessages = 6) {
    if (messages.length <= maxMessages) return messages;
    return messages.slice(-maxMessages);
  }

  // ─── UPDATED AI MESSAGE SENDER ───
  // Now uses intent-based routing with minimal tokens

  async function sendAiMessage(userPrompt) {
    const apiKey = state.settings.openrouter_api_key;
    if (!apiKey) {
      throw new Error('API Key missing. Please configure Groq API Key in Settings.');
    }

    const endpoint = 'https://api.groq.com/openai/v1/chat/completions';
    const model = 'llama-3.1-8b-instant';

    // Step 1: Detect intent
    const intent = await detectIntent(userPrompt);

    // Step 2: Process data locally based on intent
    let dataText = '';
    let shouldFormat = false;

    switch (intent) {
      case INTENTS.LIST_STUDENTS:
        dataText = formatDataAsText(intent, listStudents(userPrompt));
        shouldFormat = true;
        break;
      case INTENTS.FEE_INFO: {
        // Try to extract student ID first (numeric pattern like 1234 or AB1234)
        const feeMatch = userPrompt.match(/(\d{4,10})|([A-Z]{2,4}\d{3,8})/i);
        // If no ID pattern found, pass the full prompt so getFeeInfo can do name-based lookup
        const feeStudentId = feeMatch ? feeMatch[0] : userPrompt;
        const feeData = getFeeInfo(feeStudentId);
        dataText = formatDataAsText(intent, feeData);
        shouldFormat = true;
        break;
      }
      case INTENTS.PENDING_DUES:
        dataText = formatDataAsText(intent, getPendingDues());
        shouldFormat = true;
        break;
      case INTENTS.SCHOLARSHIP_INFO:
        dataText = formatDataAsText(intent, getScholarshipInfo());
        shouldFormat = true;
        break;
      case INTENTS.GENERAL_CHAT:
      default:
        // For general chat, use a short context
        dataText = `College: ${state.collegeName}\nTotal Students: ${state.students.length}\nTotal Transactions: ${state.transactions.length}`;
        shouldFormat = false;
        break;
    }

    // Step 3: Build system prompt (short!)
    const systemPrompt = `You are the College Fee Assistant. Format the provided data clearly.
If data is empty, say 'No records found'.
Respond in Indian Rupees (₹) format.
Be concise and helpful.`;

    // Step 4: Build messages (trimmed history + current)
    const messages = [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: shouldFormat ? `Format this data: ${dataText}\n\nUser question: ${userPrompt}` : userPrompt }
    ];

    // Step 5: Make API call with max_tokens
    const maxTokens = intent === INTENTS.GENERAL_CHAT ? 500 : 300;

    const response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: model,
        messages: messages,
        max_tokens: maxTokens
      })
    });

    if (!response.ok) {
      const errText = await response.text();
      throw new Error(`API Error: ${errText}`);
    }

    const data = await response.json();
    const replyText = data?.choices?.[0]?.message?.content;
    if (!replyText) {
      throw new Error('Received empty response from API.');
    }

    return replyText;
  }

  function updateAiStatusBadge() {
    const badge = document.getElementById('ai-status-badge');
    if (!badge) return;
    if (state.settings.openrouter_api_key) {
      badge.textContent = 'API Key Configured';
      badge.style.background = 'rgba(34, 197, 94, 0.15)';
      badge.style.color = 'var(--success)';
    } else {
      badge.textContent = 'API Key Missing';
      badge.style.background = 'rgba(239, 68, 68, 0.15)';
      badge.style.color = 'var(--danger)';
    }
  }

  function getCompactDatabaseContext() {
    const compactStudents = state.students.map(s => ({
      adm_no: s.serial_no,
      name: s.name,
      gender: s.gender,
      class: s.current_class,
      year: s.current_academic_year,
      category: s.category,
      fees: s.fee_applicable,
      paid: s.fee_paid,
      pending: s.fee_pending,
      sch_expected: s.scholarship_expected,
      sch_received: s.scholarship_received,
      sch_pending: s.scholarship_pending,
      cancelled: Number(s.admission_cancelled) === 1
    }));

    const compactTxns = state.transactions.map(t => ({
      adm_no: t.serial_no,
      amount: t.amount,
      date: t.date,
      method: t.method,
      utr: t.utr,
      verified: Number(t.verified) === 1
    }));

    return JSON.stringify({
      college: state.collegeName,
      students: compactStudents,
      transactions: compactTxns
    });
  }

  function formatMarkdown(text) {
    if (!text) return '';

    let html = text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');

    const lines = html.split('\n');
    let inTable = false;
    let tableHtml = '';
    let processedLines = [];

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line.startsWith('|') && line.endsWith('|')) {
        if (!inTable) {
          inTable = true;
          tableHtml = '<table>';
        }

        const cells = line.split('|').slice(1, -1).map(c => c.trim());
        const isSeparator = cells.every(c => /^:-*-*:*$/.test(c) || /^-+$/.test(c));

        if (isSeparator) {
          continue;
        }

        tableHtml += '<tr>';
        cells.forEach(cell => {
          const nextLine = lines[i + 1]?.trim() || '';
          const isHeader = !tableHtml.includes('</tr>') || (nextLine.startsWith('|') && nextLine.includes('---'));
          const tag = isHeader ? 'th' : 'td';
          tableHtml += `<${tag}>${parseInlineMarkdown(cell)}</${tag}>`;
        });
        tableHtml += '</tr>';
      } else {
        if (inTable) {
          inTable = false;
          tableHtml += '</table>';
          processedLines.push(tableHtml);
          tableHtml = '';
        }
        processedLines.push(line);
      }
    }
    if (inTable) {
      tableHtml += '</table>';
      processedLines.push(tableHtml);
    }

    html = processedLines.join('\n');

    html = html.replace(/^\s*&gt;\s+(.*)$/gm, '<blockquote>$1</blockquote>');

    html = html.replace(/^\s*-\s+(.*)$/gm, '<li>$1</li>');
    html = html.replace(/^\s*\*\s+(.*)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>)/g, '<ul>$1</ul>');
    html = html.replace(/<\/ul>\s*<ul>/g, '');

    return html
      .split('\n\n')
      .map(p => {
        p = p.trim();
        if (!p) return '';
        if (p.startsWith('<table>') || p.startsWith('<ul>') || p.startsWith('<ol>') || p.startsWith('<blockquote>') || p.startsWith('<h3>')) {
          return p;
        }
        if (p.startsWith('###')) {
          return `<h3>${parseInlineMarkdown(p.substring(3).trim())}</h3>`;
        }
        if (p.startsWith('##')) {
          return `<h2>${parseInlineMarkdown(p.substring(2).trim())}</h2>`;
        }
        if (p.startsWith('#')) {
          return `<h1>${parseInlineMarkdown(p.substring(1).trim())}</h1>`;
        }
        return `<p>${parseInlineMarkdown(p)}</p>`;
      })
      .filter(Boolean)
      .join('\n');
  }

  function parseInlineMarkdown(text) {
    return text
      .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
      .replace(/\*(.*?)\*/g, '<em>$1</em>')
      .replace(/`(.*?)`/g, '<code>$1</code>')
      .replace(/\n/g, '<br>');
  }

  function appendChatBubble(text, sender) {
    const container = document.getElementById('ai-chat-messages');
    if (!container) return;

    const bubble = document.createElement('div');
    bubble.className = `chat-bubble ${sender}-bubble`;

    let headerHtml = '';
    if (sender === 'ai') {
      headerHtml = `
        <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 6px;">
          <div style="display: flex; align-items: center; gap: 6px;">
            <i class="ph-bold ph-sparkle" style="color: var(--accent); font-size: 1rem;"></i>
            <span style="font-weight: 600; font-size: 0.8rem; color: var(--text-secondary);">College Fee AI Assistant</span>
          </div>
        </div>
      `;
    } else {
      headerHtml = `
        <div style="display: flex; align-items: center; gap: 6px; margin-bottom: 6px;">
          <i class="ph-bold ph-user" style="color: var(--accent); font-size: 1rem;"></i>
          <span style="font-weight: 600; font-size: 0.8rem; color: var(--text-secondary);">You (College Clerk)</span>
        </div>
      `;
    }

    const contentHtml = `
      <div class="chat-message-text" style="font-size: 0.9rem; line-height: 1.5; color: var(--text-primary);">
        ${sender === 'ai' ? formatMarkdown(text) : parseInlineMarkdown(text)}
      </div>
    `;

    let actionsHtml = '';
    if (sender === 'ai') {
      actionsHtml = `
        <div class="chat-bubble-actions">
          <button class="btn-bubble-action btn-copy-reply" title="Copy reply to clipboard">
            <i class="ph-bold ph-copy"></i> Copy
          </button>
        </div>
      `;
    }

    bubble.innerHTML = headerHtml + contentHtml + actionsHtml;
    container.appendChild(bubble);

    if (sender === 'ai') {
      const copyBtn = bubble.querySelector('.btn-copy-reply');
      if (copyBtn) {
        copyBtn.addEventListener('click', () => {
          navigator.clipboard.writeText(text).then(() => {
            const originalText = copyBtn.innerHTML;
            copyBtn.innerHTML = '<i class="ph-bold ph-check" style="color: var(--success);"></i> Copied';
            setTimeout(() => {
              copyBtn.innerHTML = originalText;
            }, 2000);
          }).catch(err => {
            Toast.show('Failed to copy to clipboard', 'warning');
          });
        });
      }
    }

    container.scrollTop = container.scrollHeight;
  }

  function appendTypingIndicator() {
    const container = document.getElementById('ai-chat-messages');
    if (!container) return null;

    const id = 'typing-' + Date.now();
    const bubble = document.createElement('div');
    bubble.className = 'chat-bubble ai-bubble';
    bubble.id = id;

    bubble.innerHTML = `
      <div style="display: flex; align-items: center; gap: 6px; margin-bottom: 6px;">
        <i class="ph-bold ph-sparkle" style="color: var(--accent); font-size: 1rem;"></i>
        <span style="font-weight: 600; font-size: 0.8rem; color: var(--text-secondary);">College Fee AI Assistant</span>
      </div>
      <div class="typing-indicator">
        <span class="typing-dot"></span>
        <span class="typing-dot"></span>
        <span class="typing-dot"></span>
      </div>
    `;

    container.appendChild(bubble);
    container.scrollTop = container.scrollHeight;
    return id;
  }

  function removeTypingIndicator(id) {
    if (!id) return;
    const indicator = document.getElementById(id);
    if (indicator) indicator.remove();
  }

  async function switchView(viewName) {
    state.activeView = viewName;

    // Update sidebar active classes
    document.querySelectorAll('.nav-menu .nav-item').forEach(item => {
      if (item.getAttribute('data-view') === viewName) {
        item.classList.add('active');
      } else {
        item.classList.remove('active');
      }
    });

    // Update screen views
    document.querySelectorAll('.view-panel').forEach(panel => {
      if (panel.id === `view-${viewName}`) {
        panel.classList.add('active');
      } else {
        panel.classList.remove('active');
      }
    });

    // Update screen headers
    const titleMap = {
      dashboard: 'Dashboard Overview',
      students: 'Student Database Registry',
      admission: 'New Admission & Bulk Import',
      reports: 'Collection & Class Summaries',
      export: 'Custom Excel Report Export',
      communication: 'Broadcasting & Reminders',
      backup: 'System Backup & Restore',
      ai: 'College Fee AI Assistant'
    };
    document.getElementById('current-view-title').textContent = titleMap[viewName] || 'Fee Manager';

    // Execution hooks on views
    if (viewName === 'dashboard') {
      renderDashboard();
    } else if (viewName === 'students') {
      renderStudents();
    } else if (viewName === 'reports') {
      renderReports();
    } else if (viewName === 'communication') {
      // Pre-fill settings
      document.getElementById('gmail-address').value = state.settings.gmail_address || '';
      document.getElementById('gmail-app-password').value = state.settings.app_password || '';
      document.getElementById('auto-send-emails').checked = state.settings.auto_send_emails === 'true';
    } else if (viewName === 'ai') {
      // Refresh data to ensure AI has latest database
      Toast.show('Fetching latest data for AI...', 'info');
      await refreshData();
      Toast.show('Data updated!', 'success');
      // Pre-fill OpenRouter key
      document.getElementById('ai-api-key').value = state.settings.openrouter_api_key || '';
      updateAiStatusBadge();
    }
  }

  function populatePillFilters() {
    const classContainer = document.getElementById('class-filter-pills');
    if (classContainer) {
      // Only show classes that actually exist in the data
      const dbClasses = state.students.map(s => s.current_class).filter(Boolean);
      const studentClasses = [...new Set(dbClasses)].sort();
      classContainer.innerHTML = '';

      // Add "All Classes" pill
      const allClassBtn = document.createElement('button');
      allClassBtn.className = `filter-btn ${state.filters.class === '' ? 'active' : ''}`;
      allClassBtn.textContent = 'All Classes';
      allClassBtn.addEventListener('click', (e) => {
        classContainer.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
        e.currentTarget.classList.add('active');
        state.filters.class = '';
        renderDashboard();
      });
      classContainer.appendChild(allClassBtn);

      studentClasses.forEach(cls => {
        const btn = document.createElement('button');
        btn.className = `filter-btn ${safeLower(state.filters.class) === safeLower(cls) ? 'active' : ''}`;
        btn.textContent = cls;
        btn.addEventListener('click', (e) => {
          classContainer.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
          e.currentTarget.classList.add('active');
          state.filters.class = cls;
          renderDashboard();
        });
        classContainer.appendChild(btn);
      });
    }

    const yearContainer = document.getElementById('year-filter-pills');
    if (yearContainer) {
      // Get all unique academic years that actually have data
      const dbYears = state.students.map(s => s.current_academic_year).filter(Boolean);
      const yearsWithData = [...new Set(dbYears)];

      // Sort years descending (most recent first)
      yearsWithData.sort((a, b) => {
        const [aStart] = a.split('-').map(Number);
        const [bStart] = b.split('-').map(Number);
        return bStart - aStart;
      });

      yearContainer.innerHTML = '';

      // Add "All Years" pill
      const allYearBtn = document.createElement('button');
      allYearBtn.className = `filter-btn ${state.filters.year === '' ? 'active' : ''}`;
      allYearBtn.textContent = 'All Years';
      allYearBtn.addEventListener('click', (e) => {
        yearContainer.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
        e.currentTarget.classList.add('active');
        state.filters.year = '';
        renderDashboard();
      });
      yearContainer.appendChild(allYearBtn);

      // Create pills ONLY for years that have data - all will be active
      yearsWithData.forEach(yr => {
        const btn = document.createElement('button');
        // All pills are clickable since they all have data
        btn.className = `filter-btn ${safeLower(state.filters.year) === safeLower(yr) ? 'active' : ''}`;
        btn.textContent = yr;
        btn.addEventListener('click', (e) => {
          yearContainer.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
          e.currentTarget.classList.add('active');
          state.filters.year = yr;
          renderDashboard();
        });
        yearContainer.appendChild(btn);
      });
    }
  }

  function populateExportOptions() {
    // Custom Export columns
    const colsGrid = document.getElementById('export-columns-grid');
    colsGrid.innerHTML = '';
    EXPORT_COLUMNS.forEach(c => {
      const label = document.createElement('label');
      label.className = 'checkbox-label';
      label.innerHTML = `<input type="checkbox" class="export-col-check" value="${c.key}" checked> ${c.label}`;
      colsGrid.appendChild(label);
    });

    // Custom Export classes
    const classesGrid = document.getElementById('export-classes-grid');
    classesGrid.innerHTML = '';
    const dbClasses = state.students.map(s => s.current_class).filter(Boolean);
    const standardClasses = [
      "B.Pharm 1st Year",
      "B.Pharm 2nd Year",
      "B.Pharm 3rd Year",
      "B.Pharm Final Year",
      "M.Pharm 1st Year",
      "M.Pharm 2nd Year"
    ];
    const studentClasses = [...new Set([...standardClasses, ...dbClasses])];
    studentClasses.forEach(cls => {
      const label = document.createElement('label');
      label.className = 'checkbox-label';
      label.innerHTML = `<input type="checkbox" class="export-class-check" value="${cls}" checked> ${cls}`;
      classesGrid.appendChild(label);
    });
  }

  // ─── DASHBOARD VISUALIZATIONS (ECHARTS) ───

  function renderDashboard() {
    // Filter students by both Class and Year simultaneously
    let filtered = state.students;
    if (state.filters.class) {
      filtered = filtered.filter(s => safeLower(s.current_class) === safeLower(state.filters.class));
    }
    if (state.filters.year) {
      filtered = filtered.filter(s => safeLower(s.current_academic_year) === safeLower(state.filters.year));
    }

    // Calculations
    let active = 0;
    let cancelled = 0;
    let applicable = 0;
    let paid = 0;
    let pending = 0;
    let schExpected = 0;
    let schReceived = 0;
    let schPending = 0;
    let males = 0;
    let females = 0;
    let others = 0;
    const casteCounts = {};

    filtered.forEach(s => {
      if (parseInt(s.admission_cancelled) === 1) {
        cancelled++;
      } else {
        active++;
        applicable += parseFloat(s.fee_applicable) || 0;
        paid += parseFloat(s.fee_paid) || 0;
        pending += parseFloat(s.fee_pending) || 0;
        schExpected += parseFloat(s.scholarship_expected) || 0;
        schReceived += parseFloat(s.scholarship_received) || 0;
        schPending += parseFloat(s.scholarship_pending) || 0;

        const g = safeLower(s.gender);
        if (g === 'male') males++;
        else if (g === 'female') females++;
        else if (g === 'other') others++;

        const category = s.category || 'General';
        casteCounts[category] = (casteCounts[category] || 0) + 1;
      }
    });

    const setElText = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.textContent = val;
    };

    // Update UI text values safely
    setElText('m-active-students', active);
    setElText('m-cancelled-students', cancelled);
    setElText('m-collected-fees', '₹' + paid.toLocaleString('en-IN', { maximumFractionDigits: 0 }));
    setElText('m-pending-fees', '₹' + pending.toLocaleString('en-IN', { maximumFractionDigits: 0 }));
    setElText('m-sch-expected', '₹' + schExpected.toLocaleString('en-IN', { maximumFractionDigits: 0 }));
    setElText('m-sch-received', '₹' + schReceived.toLocaleString('en-IN', { maximumFractionDigits: 0 }));

    const collectionRate = applicable > 0 ? (paid / applicable) * 100 : 0;
    const totalCollected = paid + schReceived;
    const totalPending = pending + schPending;

    // Render visual types
    renderGaugeChart(collectionRate);
    renderBreakdownChart(paid, pending);
    renderGenderDonutChart(males, females, others);
    renderCasteDonutChart(casteCounts);
    renderMainBarChart();
    renderScholarshipChart(schReceived, schPending);
    renderTotalCollectionChart(totalCollected, totalPending);

    // Update sub labels under donut charts safely
    setElText('info-gauge-val', collectionRate.toFixed(1) + '%');
    setElText('info-breakdown-val', '₹' + paid.toLocaleString('en-IN', { maximumFractionDigits: 0 }));
    setElText('info-gender-val', others > 0 ? `${males} : ${females} : ${others}` : `${males} : ${females}`);
    setElText('info-scholarship-val', '₹' + schReceived.toLocaleString('en-IN', { maximumFractionDigits: 0 }));
    setElText('info-total-collection-val', '₹' + totalCollected.toLocaleString('en-IN', { maximumFractionDigits: 0 }));
  }

  function renderHighRiskStudents(students) {
    const listEl = document.getElementById('high-risk-students-list');
    if (!listEl) return;

    // Filter students: pending > 60% of applicable fee
    const highRisk = students.filter(s => {
      if (parseInt(s.admission_cancelled) === 1) return false;
      const applicable = parseFloat(s.fee_applicable) || 0;
      const pending = parseFloat(s.fee_pending) || 0;
      if (applicable <= 0) return false;
      return (pending / applicable) > 0.6;
    });

    // Sort by percentage pending (highest first), then absolute pending amount
    highRisk.sort((a, b) => {
      const aApp = parseFloat(a.fee_applicable) || 1;
      const bApp = parseFloat(b.fee_applicable) || 1;
      const aPer = (parseFloat(a.fee_pending) || 0) / aApp;
      const bPer = (parseFloat(b.fee_pending) || 0) / bApp;
      if (Math.abs(aPer - bPer) > 0.01) return bPer - aPer;
      return (parseFloat(b.fee_pending) || 0) - (parseFloat(a.fee_pending) || 0);
    });

    listEl.innerHTML = '';

    if (highRisk.length === 0) {
      listEl.innerHTML = `
        <div class="high-risk-empty">
          <i class="ph-fill ph-check-circle" style="font-size: 2rem; color: var(--success);"></i>
          <span>No high risk students found</span>
        </div>`;
      return;
    }

    // Take top 50 to prevent DOM overload
    const topRisk = highRisk.slice(0, 50);

    topRisk.forEach((s, index) => {
      const applicable = parseFloat(s.fee_applicable) || 0;
      const pending = parseFloat(s.fee_pending) || 0;
      const percentage = ((pending / applicable) * 100).toFixed(1);

      const li = document.createElement('li');
      li.className = 'high-risk-item';
      li.onclick = () => showStudentDetailModal(s.id);

      li.innerHTML = `
        <div class="high-risk-item-info">
          <span class="high-risk-name">${s.name || 'Unknown'}</span>
          <span class="high-risk-class">${s.current_class || 'N/A'}</span>
        </div>
        <div class="high-risk-item-info">
          <span class="high-risk-amount">₹${pending.toLocaleString('en-IN', { maximumFractionDigits: 0 })}</span>
          <span class="high-risk-percentage">${percentage}% Pending</span>
        </div>
      `;
      listEl.appendChild(li);
    });
  }

  function renderGaugeChart(rate) {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';
    const accentColor = isLight ? '#0284c7' : '#38bdf8';
    const valueColor = isLight ? '#0f172a' : '#f1f5f9';

    if (state.charts.gauge) {
      state.charts.gauge.dispose();
      state.charts.gauge = null;
    }

    const container = document.getElementById('chart-gauge');
    if (!container) return;

    const chart = echarts.init(container);
    state.charts.gauge = chart;

    chart.setOption({
      tooltip: { show: false },
      series: [
        {
          type: 'gauge',
          startAngle: 210,
          endAngle: -30,
          min: 0,
          max: 100,
          radius: '100%',
          center: ['50%', '50%'],
          pointer: { show: false },
          progress: {
            show: true,
            overlap: false,
            roundCap: true,
            clip: false,
            itemStyle: {
              color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
                { offset: 0, color: accentColor },
                { offset: 1, color: isLight ? '#7c3aed' : '#a78bfa' }
              ]),
              opacity: 0.8,
              borderColor: isLight ? 'rgba(255, 255, 255, 0.8)' : 'rgba(255, 255, 255, 0.3)',
              borderWidth: 1
            }
          },
          axisLine: {
            lineStyle: {
              width: 14,
              color: [[1, isLight ? 'rgba(0,0,0,0.06)' : 'rgba(255,255,255,0.06)']]
            }
          },
          splitLine: { show: false },
          axisTick: { show: false },
          axisLabel: { show: false },
          data: [{ value: parseFloat(rate.toFixed(1)) }],
          detail: { show: false }
        }
      ]
    });
  }

  function renderBreakdownChart(paid, pending) {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';

    if (state.charts.breakdown) {
      state.charts.breakdown.dispose();
      state.charts.breakdown = null;
    }

    const container = document.getElementById('chart-breakdown');
    if (!container) return;

    const chart = echarts.init(container);
    state.charts.breakdown = chart;

    chart.setOption({
      tooltip: { show: true, formatter: '{b}: {c} ({d}%)' },
      legend: { show: false },
      series: [
        {
          name: 'Payment Distribution',
          type: 'pie',
          radius: ['28%', '78%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: true,
          itemStyle: {
            borderRadius: 8,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            opacity: 0.70,
            shadowBlur: 10,
            shadowColor: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(0, 0, 0, 0.35)',
            shadowOffsetY: 4
          },
          label: {
            show: true,
            position: 'inside',
            formatter: '{b}\n{d}%',
            fontSize: 10,
            lineHeight: 12,
            fontWeight: '700',
            color: '#fff',
            rotate: 'radial'
          },
          labelLine: { show: false },
          data: [
            {
              value: paid,
              name: 'Collected',
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#10b981' },
                  { offset: 1, color: '#059669' }
                ])
              }
            },
            {
              value: pending,
              name: 'Pending',
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#fbbf24' },
                  { offset: 1, color: '#d97706' }
                ])
              }
            }
          ]
        }
      ]
    });
  }

  function renderGenderDonutChart(males, females, others = 0) {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';

    if (state.charts.genderDonut) {
      state.charts.genderDonut.dispose();
      state.charts.genderDonut = null;
    }

    const container = document.getElementById('chart-gender-donut');
    if (!container) return;

    const chart = echarts.init(container);
    state.charts.genderDonut = chart;

    const chartData = [
      {
        value: males,
        name: 'Male',
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: '#38bdf8' },
            { offset: 1, color: '#0284c7' }
          ])
        }
      },
      {
        value: females,
        name: 'Female',
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: '#f472b6' },
            { offset: 1, color: '#db2777' }
          ])
        }
      },
      {
        value: others,
        name: 'Other',
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: '#a78bfa' },
            { offset: 1, color: '#7c3aed' }
          ])
        }
      }
    ].filter(d => d.value > 0);

    chart.setOption({
      tooltip: { show: true, formatter: '{b}: {c} ({d}%)' },
      legend: { show: false },
      series: [
        {
          name: 'Gender Split',
          type: 'pie',
          radius: ['28%', '78%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: true,
          itemStyle: {
            borderRadius: 8,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            opacity: 0.70,
            shadowBlur: 10,
            shadowColor: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(0, 0, 0, 0.35)',
            shadowOffsetY: 4
          },
          label: {
            show: true,
            position: 'inside',
            formatter: '{b}\n{d}%',
            fontSize: 10,
            lineHeight: 12,
            fontWeight: '700',
            color: '#fff',
            rotate: 'radial'
          },
          labelLine: { show: false },
          data: chartData
        }
      ]
    });
  }

  function renderScholarshipChart(received, pending) {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';

    if (state.charts.scholarship) {
      state.charts.scholarship.dispose();
      state.charts.scholarship = null;
    }

    const container = document.getElementById('chart-scholarship-pie');
    if (!container) return;

    const chart = echarts.init(container);
    state.charts.scholarship = chart;

    chart.setOption({
      tooltip: { show: true, formatter: '{b}: {c} ({d}%)' },
      legend: { show: false },
      series: [
        {
          name: 'Scholarship Split',
          type: 'pie',
          radius: ['28%', '78%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: true,
          itemStyle: {
            borderRadius: 8,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            opacity: 0.70,
            shadowBlur: 10,
            shadowColor: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(0, 0, 0, 0.35)',
            shadowOffsetY: 4
          },
          label: {
            show: true,
            position: 'inside',
            formatter: '{b}\n{d}%',
            fontSize: 10,
            lineHeight: 12,
            fontWeight: '700',
            color: '#fff',
            rotate: 'radial'
          },
          labelLine: { show: false },
          data: [
            {
              value: received,
              name: 'Received',
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#f472b6' },
                  { offset: 1, color: '#db2777' }
                ])
              }
            },
            {
              value: pending,
              name: 'Pending',
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#a78bfa' },
                  { offset: 1, color: '#7c3aed' }
                ])
              }
            }
          ]
        }
      ]
    });
  }

  function renderTotalCollectionChart(collected, pending) {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';

    if (state.charts.totalCollection) {
      state.charts.totalCollection.dispose();
      state.charts.totalCollection = null;
    }

    const container = document.getElementById('chart-total-collection');
    if (!container) return;

    const chart = echarts.init(container);
    state.charts.totalCollection = chart;

    chart.setOption({
      tooltip: { show: true, formatter: '{b}: {c} ({d}%)' },
      legend: { show: false },
      series: [
        {
          name: 'Total Collection Split',
          type: 'pie',
          radius: ['28%', '78%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: true,
          itemStyle: {
            borderRadius: 8,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            opacity: 0.70,
            shadowBlur: 10,
            shadowColor: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(0, 0, 0, 0.35)',
            shadowOffsetY: 4
          },
          label: {
            show: true,
            position: 'inside',
            formatter: '{b}\n{d}%',
            fontSize: 10,
            lineHeight: 12,
            fontWeight: '700',
            color: '#fff',
            rotate: 'radial'
          },
          labelLine: { show: false },
          data: [
            {
              value: collected,
              name: 'Collected',
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#10b981' },
                  { offset: 1, color: '#059669' }
                ])
              }
            },
            {
              value: pending,
              name: 'Pending',
              itemStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                  { offset: 0, color: '#fbbf24' },
                  { offset: 1, color: '#d97706' }
                ])
              }
            }
          ]
        }
      ]
    });
  }

  function renderCasteDonutChart(casteCounts) {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';

    if (state.charts.casteDonut) {
      state.charts.casteDonut.dispose();
      state.charts.casteDonut = null;
    }

    const container = document.getElementById('chart-caste-donut');
    if (!container) return;

    const chart = echarts.init(container);
    state.charts.casteDonut = chart;

    const categoryColors = {
      'General': ['#3b82f6', '#1d4ed8'],
      'OBC': ['#10b981', '#047857'],
      'SC': ['#8b5cf6', '#6d28d9'],
      'ST': ['#ec4899', '#be185d'],
      'EWS': ['#f59e0b', '#b45309'],
      'NT': ['#06b6d4', '#0891b2'],
      'VJ-DT': ['#eab308', '#a16207']
    };

    const colorPalette = [
      ['#3b82f6', '#1d4ed8'],
      ['#10b981', '#047857'],
      ['#8b5cf6', '#6d28d9'],
      ['#ec4899', '#be185d'],
      ['#f59e0b', '#b45309'],
      ['#06b6d4', '#0891b2'],
      ['#eab308', '#a16207']
    ];

    const chartData = Object.keys(casteCounts).map((cat, idx) => {
      const colors = categoryColors[cat] || colorPalette[idx % colorPalette.length];
      return {
        value: casteCounts[cat],
        name: cat,
        itemStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: colors[0] },
            { offset: 1, color: colors[1] }
          ])
        }
      };
    });

    chart.setOption({
      tooltip: { show: true, formatter: '{b}: {c} ({d}%)' },
      legend: { show: false },
      series: [
        {
          name: 'Caste Split',
          type: 'pie',
          radius: ['28%', '78%'],
          center: ['50%', '50%'],
          avoidLabelOverlap: true,
          itemStyle: {
            borderRadius: 8,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            opacity: 0.70,
            shadowBlur: 10,
            shadowColor: isLight ? 'rgba(0, 0, 0, 0.08)' : 'rgba(0, 0, 0, 0.35)',
            shadowOffsetY: 4
          },
          label: {
            show: true,
            position: 'inside',
            formatter: '{b}\n{d}%',
            fontSize: 10,
            lineHeight: 12,
            fontWeight: '700',
            color: '#fff',
            rotate: 'radial'
          },
          labelLine: { show: false },
          data: chartData
        }
      ]
    });
  }

  function renderMainBarChart() {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';
    const textColor = isLight ? '#475569' : '#94a3b8';

    if (state.charts.mainBar) {
      state.charts.mainBar.dispose();
      state.charts.mainBar = null;
    }

    const container = document.getElementById('main-bar-chart');
    if (!container) return;

    const isRateMode = state.dashboardMode === 'collection';
    const isClassFilteredOnly = state.filters.class && !state.filters.year;
    const isMultiSeries = !isClassFilteredOnly && !state.filters.year;

    function formatChartLabel(amount, percentage) {
      let amtStr = '';
      if (amount >= 100000) {
        amtStr = '₹' + (amount / 100000).toFixed(2).replace(/\.?0+$/, '') + 'L';
      } else if (amount >= 1000) {
        amtStr = '₹' + (amount / 1000).toFixed(1).replace(/\.?0+$/, '') + 'k';
      } else {
        amtStr = '₹' + Math.round(amount);
      }
      return `${amtStr} (${Math.round(percentage)}%)`;
    }

    let xData = [];
    let titleText = '';
    const seriesList = [];
    let overallMaxVal = 0;

    // Dynamic palette: green for Collection, yellow/amber for Pending
    const colorPalette = isRateMode
      ? [
        { dark: '#065f46', light: '#34d399', cap: '#34d399' }, // Emerald
        { dark: '#14532d', light: '#4ade80', cap: '#4ade80' }, // Green
        { dark: '#115e59', light: '#2dd4bf', cap: '#2dd4bf' }, // Teal
        { dark: '#064e3b', light: '#6ee7b7', cap: '#6ee7b7' }, // Mint
        { dark: '#166534', light: '#86efac', cap: '#86efac' }, // Lime-Green
        { dark: '#0f766e', light: '#5eead4', cap: '#5eead4' }  // Cyan-Teal
      ]
      : [
        { dark: '#92400e', light: '#fbbf24', cap: '#fbbf24' }, // Amber
        { dark: '#78350f', light: '#f59e0b', cap: '#f59e0b' }, // Yellow-Gold
        { dark: '#854d0e', light: '#facc15', cap: '#facc15' }, // Yellow
        { dark: '#9a3412', light: '#fb923c', cap: '#fb923c' }, // Orange
        { dark: '#a16207', light: '#fde047', cap: '#fde047' }, // Lemon
        { dark: '#b45309', light: '#fdba74', cap: '#fdba74' }  // Peach-Amber
      ];

    if (isClassFilteredOnly) {
      // 1. Class-wise but show Year-Wise (category = Years)
      titleText = isRateMode
        ? `${state.filters.class} — Year-Wise Collection Rates`
        : `${state.filters.class} — Year-Wise Pending Fees`;

      const currentCalendarYear = new Date().getFullYear();
      const standardYears = [];
      for (let i = 0; i < 5; i++) {
        const start = currentCalendarYear - i;
        const end = start + 1;
        standardYears.push(`${start}-${end}`);
      }
      xData = standardYears; // newest to oldest (current year on the left)

      const rawVals = [];
      const rawAmounts = [];
      const rawPercentages = [];
      xData.forEach(yr => {
        const studentsInYr = state.students.filter(s =>
          safeLower(s.current_class) === safeLower(state.filters.class) &&
          safeLower(s.current_academic_year) === safeLower(yr) &&
          parseInt(s.admission_cancelled) !== 1
        );
        let app = 0, paid = 0, pendingSum = 0;
        studentsInYr.forEach(s => {
          app += parseFloat(s.fee_applicable) || 0;
          paid += parseFloat(s.fee_paid) || 0;
          pendingSum += parseFloat(s.fee_pending) || 0;
        });

        const collectionRate = app > 0 ? (paid / app) * 100 : 0;
        const pendingRate = app > 0 ? (pendingSum / app) * 100 : 0;

        if (isRateMode) {
          rawVals.push(collectionRate);
          rawAmounts.push(paid);
          rawPercentages.push(collectionRate);
        } else {
          rawVals.push(pendingSum);
          rawAmounts.push(pendingSum);
          rawPercentages.push(pendingRate);
        }
      });

      overallMaxVal = Math.max(...rawVals, 0);

      // Prepare data
      const data = rawVals.map((val, idx) => {
        const colors = colorPalette[idx % colorPalette.length];
        const v = Math.round(val);
        return {
          value: v,
          amount: rawAmounts[idx],
          percentage: rawPercentages[idx],
          itemStyle: {
            color: v > 0 ? new echarts.graphic.LinearGradient(0, 1, 0, 0, [
              { offset: 0, color: colors.dark },
              { offset: 1, color: colors.light }
            ]) : 'transparent',
            opacity: 0.72,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            borderRadius: [4, 4, 0, 0],
            shadowBlur: 8,
            shadowColor: isLight ? 'rgba(0,0,0,0.06)' : 'rgba(0,0,0,0.25)',
            shadowOffsetY: 2
          }
        };
      });

      seriesList.push({
        name: isRateMode ? 'Collection Rate' : 'Pending Fees',
        type: 'bar',
        xAxisIndex: 0,
        barWidth: '42%',
        label: {
          show: true,
          position: 'top',
          rotate: 0,
          align: 'center',
          verticalAlign: 'bottom',
          distance: 5,
          formatter: params => {
            if (params.data && typeof params.data.amount !== 'undefined') {
              return formatChartLabel(params.data.amount, params.data.percentage);
            }
            if (isRateMode) return params.value + '%';
            return params.value > 0 ? '₹' + Math.round(params.value / 1000) + 'k' : '₹0';
          },
          fontSize: 10,
          fontWeight: '700',
          color: isLight ? '#1e293b' : '#f1f5f9',
          textBorderColor: isLight ? '#ffffff' : '#1e293b',
          textBorderWidth: 2
        },
        data: data,
        z: 2
      });

    } else if (state.filters.year) {
      // 2. Class-Wise for a single selected year (category = Classes)
      if (state.filters.class) {
        titleText = isRateMode
          ? `${state.filters.class} — Collection Rate (${state.filters.year})`
          : `${state.filters.class} — Pending Fees (${state.filters.year})`;
        xData = [state.filters.class];
      } else {
        titleText = isRateMode
          ? `Class-Wise Collection Rates (${state.filters.year})`
          : `Class-Wise Pending Fees (${state.filters.year})`;
        const dbClasses = state.students.map(s => s.current_class).filter(Boolean);
        const standardClasses = [
          "B.Pharm 1st Year",
          "B.Pharm 2nd Year",
          "B.Pharm 3rd Year",
          "B.Pharm Final Year",
          "M.Pharm 1st Year",
          "M.Pharm 2nd Year"
        ];
        xData = [...new Set([...standardClasses, ...dbClasses])];
      }

      const rawVals = [];
      const rawAmounts = [];
      const rawPercentages = [];
      xData.forEach(cls => {
        const studentsInCls = state.students.filter(s =>
          safeLower(s.current_class) === safeLower(cls) &&
          safeLower(s.current_academic_year) === safeLower(state.filters.year) &&
          parseInt(s.admission_cancelled) !== 1
        );
        let app = 0, paid = 0, pendingSum = 0;
        studentsInCls.forEach(s => {
          app += parseFloat(s.fee_applicable) || 0;
          paid += parseFloat(s.fee_paid) || 0;
          pendingSum += parseFloat(s.fee_pending) || 0;
        });

        const collectionRate = app > 0 ? (paid / app) * 100 : 0;
        const pendingRate = app > 0 ? (pendingSum / app) * 100 : 0;

        if (isRateMode) {
          rawVals.push(collectionRate);
          rawAmounts.push(paid);
          rawPercentages.push(collectionRate);
        } else {
          rawVals.push(pendingSum);
          rawAmounts.push(pendingSum);
          rawPercentages.push(pendingRate);
        }
      });

      overallMaxVal = Math.max(...rawVals, 0);

      // Prepare data
      const data = rawVals.map((val, idx) => {
        const colors = colorPalette[idx % colorPalette.length];
        const v = Math.round(val);
        return {
          value: v,
          amount: rawAmounts[idx],
          percentage: rawPercentages[idx],
          itemStyle: {
            color: v > 0 ? new echarts.graphic.LinearGradient(0, 1, 0, 0, [
              { offset: 0, color: colors.dark },
              { offset: 1, color: colors.light }
            ]) : 'transparent',
            opacity: 0.72,
            borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
            borderWidth: 1.5,
            borderRadius: [4, 4, 0, 0],
            shadowBlur: 8,
            shadowColor: isLight ? 'rgba(0,0,0,0.06)' : 'rgba(0,0,0,0.25)',
            shadowOffsetY: 2
          }
        };
      });

      seriesList.push({
        name: isRateMode ? 'Collection Rate' : 'Pending Fees',
        type: 'bar',
        xAxisIndex: 0,
        barWidth: '42%',
        label: {
          show: true,
          position: 'top',
          rotate: 0,
          align: 'center',
          verticalAlign: 'bottom',
          distance: 5,
          formatter: params => {
            if (params.data && typeof params.data.amount !== 'undefined') {
              return formatChartLabel(params.data.amount, params.data.percentage);
            }
            if (isRateMode) return params.value + '%';
            return params.value > 0 ? '₹' + Math.round(params.value / 1000) + 'k' : '₹0';
          },
          fontSize: 12,
          fontWeight: '700',
          color: isLight ? '#1e293b' : '#f1f5f9',
          textBorderColor: isLight ? '#ffffff' : '#1e293b',
          textBorderWidth: 2
        },
        data: data,
        z: 2
      });

    } else {
      // 3. Year-Wise grouped by classes (category = Years, series = 6 classes)
      titleText = isRateMode
        ? `Year-Wise Collection Rates (Grouped)`
        : `Year-Wise Pending Fees (Grouped)`;

      const currentCalendarYear = new Date().getFullYear();
      const standardYears = [];
      for (let i = 0; i < 5; i++) {
        const start = currentCalendarYear - i;
        const end = start + 1;
        standardYears.push(`${start}-${end}`);
      }
      xData = standardYears; // newest to oldest (current year on the left)

      const classes = [
        "B.Pharm 1st Year",
        "B.Pharm 2nd Year",
        "B.Pharm 3rd Year",
        "B.Pharm Final Year",
        "M.Pharm 1st Year",
        "M.Pharm 2nd Year"
      ];

      // Compute raw values for all classes/years to find maxVal
      const allClassRates = [];
      const allClassAmounts = [];
      const allClassPercentages = [];

      classes.forEach(cls => {
        const rates = [];
        const amounts = [];
        const percentages = [];
        xData.forEach(yr => {
          const studentsInYrCls = state.students.filter(s =>
            safeLower(s.current_class) === safeLower(cls) &&
            safeLower(s.current_academic_year) === safeLower(yr) &&
            parseInt(s.admission_cancelled) !== 1
          );
          let app = 0, paid = 0, pendingSum = 0;
          studentsInYrCls.forEach(s => {
            app += parseFloat(s.fee_applicable) || 0;
            paid += parseFloat(s.fee_paid) || 0;
            pendingSum += parseFloat(s.fee_pending) || 0;
          });

          const collectionRate = app > 0 ? (paid / app) * 100 : 0;
          const pendingRate = app > 0 ? (pendingSum / app) * 100 : 0;

          if (isRateMode) {
            rates.push(collectionRate);
            amounts.push(paid);
            percentages.push(collectionRate);
          } else {
            rates.push(pendingSum);
            amounts.push(pendingSum);
            percentages.push(pendingRate);
          }
        });
        allClassRates.push(rates);
        allClassAmounts.push(amounts);
        allClassPercentages.push(percentages);
      });

      allClassRates.forEach(rates => {
        rates.forEach(v => {
          if (v > overallMaxVal) overallMaxVal = v;
        });
      });

      // For each class, build a separate bar series
      classes.forEach((cls, clsIdx) => {
        const colors = colorPalette[clsIdx % colorPalette.length];
        const rates = allClassRates[clsIdx];
        const amounts = allClassAmounts[clsIdx];
        const percentages = allClassPercentages[clsIdx];

        const data = rates.map((val, idx) => {
          const v = Math.round(val);
          return {
            value: v,
            amount: amounts[idx],
            percentage: percentages[idx],
            itemStyle: {
              color: v > 0 ? new echarts.graphic.LinearGradient(0, 1, 0, 0, [
                { offset: 0, color: colors.dark },
                { offset: 1, color: colors.light }
              ]) : 'transparent',
              opacity: 0.72,
              borderColor: isLight ? 'rgba(255, 255, 255, 0.85)' : 'rgba(255, 255, 255, 0.35)',
              borderWidth: 1.5,
              borderRadius: [3, 3, 0, 0],
              shadowBlur: 8,
              shadowColor: isLight ? 'rgba(0,0,0,0.06)' : 'rgba(0,0,0,0.25)',
              shadowOffsetY: 2
            }
          };
        });

        seriesList.push({
          name: cls,
          type: 'bar',
          xAxisIndex: 0,
          barWidth: '12%',
          barGap: '5%',
          label: {
            show: true,
            position: 'top',
            rotate: 0,
            align: 'center',
            verticalAlign: 'bottom',
            distance: 2,
            formatter: params => {
              if (params.data && typeof params.data.amount !== 'undefined') {
                let amtStr = '';
                const amt = params.data.amount;
                if (amt >= 100000) {
                  amtStr = '₹' + (amt / 100000).toFixed(1).replace(/\.?0+$/, '') + 'L';
                } else if (amt >= 1000) {
                  amtStr = '₹' + Math.round(amt / 1000) + 'k';
                } else {
                  amtStr = '₹' + Math.round(amt);
                }
                return `${amtStr}\n(${Math.round(params.data.percentage)}%)`;
              }
              if (isRateMode) return params.value + '%';
              return params.value > 0 ? '₹' + Math.round(params.value / 1000) + 'k' : '₹0';
            },
            fontSize: 8,
            lineHeight: 10,
            fontWeight: '700',
            color: isLight ? '#1e293b' : '#f1f5f9',
            textBorderColor: isLight ? '#ffffff' : '#1e293b',
            textBorderWidth: 2
          },
          data: data,
          z: 2
        });
      });
    }

    // Add background container cards behind the bars using the secondary xAxis
    const bgCardColor = isLight ? 'rgba(15, 23, 42, 0.04)' : 'rgba(255, 255, 255, 0.04)';
    const bgHeight = isRateMode ? 100 : (overallMaxVal > 0 ? overallMaxVal * 1.15 : 100000);
    seriesList.unshift({
      name: 'Background Card',
      type: 'bar',
      xAxisIndex: 1,
      barWidth: isMultiSeries ? '80%' : '45%',
      silent: true,
      itemStyle: {
        color: bgCardColor,
        borderRadius: [6, 6, 0, 0]
      },
      data: xData.map(() => bgHeight),
      z: 1
    });

    // Set title safely
    const titleEl = document.getElementById('bar-chart-title');
    if (titleEl) titleEl.textContent = titleText;

    const chart = echarts.init(container);
    state.charts.mainBar = chart;

    const isCategoryYear = isClassFilteredOnly || isMultiSeries;
    const isLabelHorizontal = isCategoryYear || xData.length <= 2;

    chart.setOption({
      legend: {
        show: false, // Always hide the legend to remove the row from screenshot
        textStyle: { color: textColor, fontWeight: '600', fontSize: 9.5 },
        top: '0%',
        left: 'center',
        itemWidth: 12,
        itemHeight: 12
      },
      tooltip: {
        show: true,
        trigger: 'axis',
        axisPointer: { type: 'shadow' },
        formatter: function (params) {
          if (Array.isArray(params)) {
            let tooltipStr = `<div style="font-weight:700;margin-bottom:4px;">${params[0].axisValue}</div>`;
            params.forEach(p => {
              if (p.seriesName === 'Background Card') return;
              const amt = p.data && typeof p.data.amount !== 'undefined' ? p.data.amount : p.value;
              const pct = p.data && typeof p.data.percentage !== 'undefined' ? p.data.percentage : p.value;
              const amtStr = amt.toLocaleString('en-IN', { maximumFractionDigits: 0 });
              const pctStr = pct.toFixed(1) + '%';

              tooltipStr += `<div style="display:flex;align-items:center;gap:6px;font-size:12px;">
                 <span style="display:inline-block;width:8px;height:8px;border-radius:50%;background-color:${p.color && p.color.colorStops ? p.color.colorStops[0].color : p.color};"></span>
                 ${p.seriesName}: <b>₹${amtStr} (${pctStr})</b>
               </div>`;
            });
            return tooltipStr;
          }
          if (params.seriesName === 'Background Card') return '';
          const amt = params.data && typeof params.data.amount !== 'undefined' ? params.data.amount : params.value;
          const pct = params.data && typeof params.data.percentage !== 'undefined' ? params.data.percentage : params.value;
          const amtStr = amt.toLocaleString('en-IN', { maximumFractionDigits: 0 });
          const pctStr = pct.toFixed(1) + '%';

          return `${params.seriesName} - ${params.name}: <br/><b>₹${amtStr} (${pctStr})</b>`;
        }
      },
      grid: {
        left: '2%',
        right: '2%',
        bottom: '5%',
        top: '8%',
        containLabel: true
      },
      xAxis: [
        {
          type: 'category',
          data: xData,
          axisLabel: {
            color: isLight ? '#0f172a' : '#ffffff',
            fontSize: 16,
            fontWeight: '700',
            interval: 0,
            rotate: isLabelHorizontal ? 0 : 45,
            inside: false,
            align: isLabelHorizontal ? 'center' : 'right',
            verticalAlign: isLabelHorizontal ? 'top' : 'middle',
            margin: 12,
            textBorderColor: isLight ? '#ffffff' : '#1e293b',
            textBorderWidth: isLabelHorizontal ? 0 : 3
          },
          axisLine: { lineStyle: { color: isLight ? '#cbd5e1' : '#334155' } },
          axisTick: { show: false }
        },
        {
          type: 'category',
          data: xData,
          show: false
        }
      ],
      yAxis: {
        type: 'value',
        min: 0,
        max: isRateMode ? 100 : undefined,
        axisLabel: {
          color: textColor,
          fontSize: 9.5,
          formatter: isRateMode ? '{value}%' : function (val) {
            return val >= 1000 ? '₹' + (val / 1000) + 'k' : '₹' + val;
          }
        },
        splitLine: { lineStyle: { color: isLight ? '#cbd5e1' : '#1e293b', type: 'dashed' } }
      },
      series: seriesList
    });
  }


  function resizeCharts() {
    for (const key in state.charts) {
      if (state.charts[key] && typeof state.charts[key].resize === 'function') {
        try {
          state.charts[key].resize();
        } catch (err) {
          console.warn("Chart resize failed:", key, err);
        }
      }
    }
  }

  // ─── STUDENT REGISTER DATABASE RENDER ───

  function renderStudents() {
    const tbody = document.getElementById('students-table-body');
    tbody.innerHTML = '';

    // Render class filters pills in db view
    const pContainer = document.getElementById('student-class-filters');
    pContainer.innerHTML = '';

    const filterAll = document.createElement('button');
    filterAll.className = `filter-btn ${state.filters.class === '' ? 'active' : ''}`;
    filterAll.textContent = 'All Classes';
    filterAll.onclick = () => {
      state.filters.class = '';
      renderStudents();
    };
    pContainer.appendChild(filterAll);

    const dbClasses = state.students.map(s => s.current_class).filter(Boolean);
    const standardClasses = [
      "B.Pharm 1st Year",
      "B.Pharm 2nd Year",
      "B.Pharm 3rd Year",
      "B.Pharm Final Year",
      "M.Pharm 1st Year",
      "M.Pharm 2nd Year"
    ];
    const studentClasses = [...new Set([...standardClasses, ...dbClasses])];
    studentClasses.forEach(cls => {
      const btn = document.createElement('button');
      btn.className = `filter-btn ${safeLower(state.filters.class) === safeLower(cls) ? 'active' : ''}`;
      btn.textContent = cls;
      btn.onclick = () => {
        state.filters.class = cls;
        renderStudents();
      };
      pContainer.appendChild(btn);
    });

    // Filter students
    const query = safeLower(state.searchQuery).trim();
    const showCancelled = state.showCancelled;
    const filterClass = state.filters.class;

    const filtered = state.students.filter(s => {
      // Cancelled filter
      const cancelled = parseInt(s.admission_cancelled) === 1;
      if (!showCancelled && cancelled) return false;

      // Class filter (case-insensitive)
      if (filterClass && safeLower(s.current_class) !== safeLower(filterClass)) return false;

      // Search query filter
      if (query) {
        const serialMatch = safeLower(s.serial_no).includes(query);
        const nameMatch = safeLower(s.name).includes(query);
        const remarkMatch = safeLower(s.remark).includes(query);
        const categoryMatch = safeLower(s.category).includes(query);
        return serialMatch || nameMatch || remarkMatch || categoryMatch;
      }

      return true;
    });

    if (filtered.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 24px; color: var(--text-muted);">No students found.</td></tr>`;
      return;
    }

    filtered.forEach(s => {
      const tr = document.createElement('tr');
      if (parseInt(s.admission_cancelled) === 1) {
        tr.className = 'cancelled';
      }
      tr.style.cursor = 'pointer';
      tr.addEventListener('click', () => showStudentDetailModal(s.id));

      const applicable = parseFloat(s.fee_applicable) || 0;
      const paid = parseFloat(s.fee_paid) || 0;
      const pending = parseFloat(s.fee_pending) || 0;

      // Status indicator
      let statusHtml = '';
      if (parseInt(s.admission_cancelled) === 1) {
        statusHtml = `<span class="badge badge-danger">Cancelled</span>`;
      } else if (pending <= 0) {
        statusHtml = `<span class="badge badge-success">✓ Paid</span>`;
      } else {
        statusHtml = `<span class="badge badge-warning">₹${pending.toLocaleString('en-IN')} due</span>`;
      }

      tr.innerHTML = `
        <td style="font-weight:600; color: var(--accent);">${s.name}</td>
        <td>${s.current_class}</td>
        <td><span class="badge badge-accent">${s.category || 'Gen'}</span></td>
        <td>₹${applicable.toLocaleString('en-IN')}</td>
        <td>${statusHtml}</td>
      `;
      tbody.appendChild(tr);
    });
  }

  // ─── MODAL CONTROLLERS ───

  function openModal(title, bodyHtml, footerHtml) {
    document.getElementById('modal-title').textContent = title;
    document.getElementById('modal-body').innerHTML = bodyHtml;
    document.getElementById('modal-footer').innerHTML = footerHtml || '';
    document.getElementById('modal-backdrop').classList.add('visible');
  }

  function closeModal() {
    document.getElementById('app-modal').style.maxWidth = '';
    document.getElementById('modal-backdrop').classList.remove('visible');
  }

  function appConfirm(title, message, onConfirm) {
    document.getElementById('confirm-title').textContent = title;
    document.getElementById('confirm-body').textContent = message;

    const backdrop = document.getElementById('confirm-backdrop');
    const okBtn = document.getElementById('confirm-ok');
    const cancelBtn = document.getElementById('confirm-cancel');

    const cleanup = () => {
      backdrop.classList.remove('visible');
      okBtn.replaceWith(okBtn.cloneNode(true));
      cancelBtn.replaceWith(cancelBtn.cloneNode(true));
    };

    const newOkBtn = document.getElementById('confirm-ok');
    const newCancelBtn = document.getElementById('confirm-cancel');

    newOkBtn.onclick = () => {
      cleanup();
      if (onConfirm) onConfirm();
    };

    newCancelBtn.onclick = () => {
      cleanup();
    };

    backdrop.classList.add('visible');
  }

  function showStudentDetailModal(studentId) {
    document.getElementById('app-modal').style.maxWidth = '1150px';
    const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
    const bodyHtml = renderDetailModalContent(studentId, false);

    const isCancelled = student && parseInt(student.admission_cancelled) === 1;
    const admissionBtn = isCancelled
      ? `<button class="btn btn-success btn-sm" onclick="App.toggleAdmission(${studentId}, 0); App.closeModal();">
           <i class="ph-bold ph-arrow-counter-clockwise"></i> Restore Admission
         </button>`
      : `<button class="btn btn-danger btn-sm" onclick="App.toggleAdmission(${studentId}, 1); App.closeModal();">
           <i class="ph-bold ph-prohibit"></i> Cancel Admission
         </button>`;

    const hasDues = student && parseFloat(student.fee_pending) > 0 && parseInt(student.admission_cancelled) !== 1;
    const dueReminderBtn = hasDues
      ? `<button class="btn btn-primary btn-sm" onclick="App.sendIndividualDueReminder(${studentId})">
           <i class="ph-bold ph-paper-plane"></i> Send Due Reminder
         </button>`
      : '';

    const footerHtml = `
      ${admissionBtn}
      ${dueReminderBtn}
      <div style="flex:1;"></div>
      <button class="btn btn-ghost" onclick="App.closeModal()">Close</button>
    `;
    openModal('Student Profile & Fee Ledger', bodyHtml, footerHtml);
  }

  async function sendIndividualDueReminder(studentId) {
    const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
    if (!student) return;

    const emails = Array.from(new Set(
      [student.student_email, student.parent_email]
        .map(e => (e || '').trim())
        .filter(e => e.length > 0)
    ));
    if (emails.length === 0) {
      alert('This student does not have any student or parent email address configured.');
      return;
    }

    const confirmSend = confirm(`Are you sure you want to send a due reminder email to ${student.name} (${emails.join(', ')})?`);
    if (!confirmSend) return;

    Toast.show('Sending due reminder email...', 'info');
    try {
      const res = await API.sendNotification({
        recipientEmail: emails.join(','),
        subject: `Fee Payment Due Reminder - ${state.collegeName}`,
        message: `Dear ${student.name},\n\nThis is a friendly reminder that you have an outstanding fee balance of INR ${student.fee_pending} for the academic year ${student.current_academic_year}. Please make the payment at your earliest convenience.\n\nBest regards,\nFee Collection Desk\n${state.collegeName}`
      });

      if (res.success) {
        Toast.show('Due reminder email sent successfully!', 'success');
      } else {
        Toast.show(res.error || 'Failed to send reminder email.', 'error');
      }
    } catch (err) {
      console.error(err);
      Toast.show('Network error sending reminder email.', 'error');
    }
  }

  function toggleDetailProfileEdit(studentId, isEdit) {
    const bodyHtml = renderDetailModalContent(studentId, isEdit);
    document.getElementById('modal-body').innerHTML = bodyHtml;
  }

  async function saveDetailProfile(studentId) {
    if (!(await requirePin())) return;
    const payload = {
      id: studentId,
      name: document.getElementById('det-edit-name').value.trim(),
      gender: document.getElementById('det-edit-gender').value,
      category: document.getElementById('det-edit-category').value,
      current_class: document.getElementById('det-edit-class').value,
      current_academic_year: document.getElementById('det-edit-ay').value.trim(),
      year_of_admission: parseInt(document.getElementById('det-edit-year').value) || new Date().getFullYear(),
      student_mobile: document.getElementById('det-edit-smob').value.trim(),
      parent_mobile: document.getElementById('det-edit-pmob').value.trim(),
      student_email: document.getElementById('det-edit-semail').value.trim(),
      parent_email: document.getElementById('det-edit-pemail').value.trim(),
      local_address: document.getElementById('det-edit-laddr').value.trim(),
      permanent_address: document.getElementById('det-edit-paddr').value.trim(),
      fee_applicable: parseFloat(document.getElementById('det-edit-fee').value) || 0,
      scholarship_expected: parseFloat(document.getElementById('det-edit-schexp').value) || 0,
      remark: document.getElementById('det-edit-remark').value.trim()
    };

    if (!payload.name) {
      alert('Please fill out the student name.');
      return;
    }

    Toast.show('Updating student profile...', 'info');
    try {
      const res = await API.updateStudent(payload);
      if (res.success) {
        Toast.show('Student profile updated!', 'success');
        await refreshData();
        renderStudents();
        toggleDetailProfileEdit(studentId, false);
      } else {
        Toast.show(res.error || 'Failed to update student profile.', 'error');
      }
    } catch (e) {
      Toast.show('Error updating student profile.', 'error');
    }
  }

  async function submitDetailPayment(studentId) {
    if (!(await requirePin())) return;
    const amount = parseFloat(document.getElementById('det-pay-amount').value) || 0;
    const utr = document.getElementById('det-pay-utr').value.trim();
    const date = document.getElementById('det-pay-date').value;
    const bank = document.getElementById('det-pay-bank').value.trim() || 'Manual Entry';
    const remark = document.getElementById('det-pay-remark').value.trim();
    const payTypeEl = document.getElementById('det-pay-type');
    const type = payTypeEl ? payTypeEl.value : 'Fee';
    let otherReason = '';
    if (type === 'Other') {
      otherReason = document.getElementById('det-pay-other-reason').value.trim();
      if (!otherReason) {
        alert('Please specify the reason for the other fee.');
        return;
      }
    }

    if (amount <= 0 || !utr) {
      alert('Please enter an amount and a valid Transaction ID/UTR No.');
      return;
    }

    Toast.show('Recording transaction...', 'info');
    try {
      const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
      const res = await API.recordPayment({
        student_id: studentId,
        student_name: student?.name || 'Unknown',
        transaction_id: utr,
        date: date,
        amount: amount,
        bank_name: bank,
        source: 'Manual Entry',
        status: 'Unverified',
        verified: 0,
        raw_remark: remark,
        type: type,
        other_reason: otherReason
      });

      if (res.success) {
        Toast.show('Manual payment recorded!', 'success');
        await refreshData();
        renderStudents();
        toggleDetailProfileEdit(studentId, false);
      } else {
        Toast.show(res.error || 'Recording payment failed.', 'error');
      }
    } catch (e) {
      Toast.show('Error connecting to Server API.', 'error');
    }
  }

  function setDetailPaymentType(type) {
    const input = document.getElementById('det-pay-type');
    if (!input) return;
    input.value = type;

    const feeBtn = document.getElementById('btn-paytype-fee');
    const schBtn = document.getElementById('btn-paytype-sch');
    const otherBtn = document.getElementById('btn-paytype-other');
    const otherReasonContainer = document.getElementById('det-pay-other-reason-container');

    if (feeBtn) feeBtn.classList.remove('active');
    if (schBtn) schBtn.classList.remove('active');
    if (otherBtn) otherBtn.classList.remove('active');

    if (type === 'Fee') {
      if (feeBtn) feeBtn.classList.add('active');
      const pendingFee = parseFloat(feeBtn ? feeBtn.dataset.pending : 0) || 0;
      document.getElementById('det-pay-amount').value = pendingFee > 0 ? pendingFee : '';
      if (otherReasonContainer) otherReasonContainer.style.display = 'none';
    } else if (type === 'Scholarship') {
      if (schBtn) schBtn.classList.add('active');
      const pendingSch = parseFloat(schBtn ? schBtn.dataset.pending : 0) || 0;
      document.getElementById('det-pay-amount').value = pendingSch > 0 ? pendingSch : '';
      if (otherReasonContainer) otherReasonContainer.style.display = 'none';
    } else if (type === 'Other') {
      if (otherBtn) otherBtn.classList.add('active');
      document.getElementById('det-pay-amount').value = '';
      if (otherReasonContainer) otherReasonContainer.style.display = 'block';
    }
  }

  async function toggleDetailTxnVerification(txnId, verifyState, studentId) {
    if (!(await requirePin())) return;
    const msg = verifyState
      ? 'Are you sure you want to mark this transaction as Verified?'
      : 'Are you sure you want to Unverify this transaction?';

    appConfirm(verifyState ? 'Verify Payment' : 'Unverify Payment', msg, async () => {
      Toast.show(verifyState ? 'Verifying payment...' : 'Unverifying payment...', 'info');
      try {
        const res = await API.verifyTransaction(txnId, verifyState);
        if (res.success) {
          Toast.show(verifyState ? 'Transaction verified!' : 'Transaction unverified!', 'success');
          await refreshData();
          renderStudents();
          toggleDetailProfileEdit(studentId, false);
        } else {
          Toast.show(res.error || 'Action failed.', 'error');
        }
      } catch (e) {
        Toast.show('Network error updating verification.', 'error');
      }
    });
  }

  function renderDetailModalContent(studentId, isEdit = false) {
    const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
    if (!student) return '';

    const txns = state.transactions.filter(t => parseInt(t.student_id) === parseInt(studentId));

    const applicable = parseFloat(student.fee_applicable) || 0;
    const paid = parseFloat(student.fee_paid) || 0;
    const pending = parseFloat(student.fee_pending) || 0;
    const schExpected = parseFloat(student.scholarship_expected) || 0;
    const schReceived = parseFloat(student.scholarship_received) || 0;
    const schPending = parseFloat(student.scholarship_pending) || 0;

    // LEFT COLUMN: PROFILE
    let leftHtml = '';
    if (!isEdit) {
      leftHtml = `
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--border); padding-bottom:8px;">
            <h4 style="margin:0; font-size:1.05rem; font-weight:700; color:var(--accent);">Student Profile</h4>
            <button class="btn btn-ghost btn-sm" onclick="App.toggleDetailProfileEdit(${studentId}, true)" style="padding: 4px 8px; font-size:0.85rem;">
              <i class="ph-bold ph-pencil-simple"></i> Edit Profile
            </button>
          </div>
          
          <div style="display:grid; grid-template-columns:repeat(4, 1fr); gap:12px; font-size:0.95rem;">
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Full Name</div>
              <div style="font-weight:600; color:var(--text-primary);">${student.name || 'N/A'}</div>
            </div>
            <div>
              <div style="color:var(--text-muted); font-size:0.85rem;">Gender</div>
              <div>${student.gender || 'N/A'}</div>
            </div>
            <div>
              <div style="color:var(--text-muted); font-size:0.85rem;">Category</div>
              <div><span class="badge badge-accent">${student.category || 'General'}</span></div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Class</div>
              <div style="font-weight:600;">${student.current_class || 'N/A'}</div>
            </div>
            <div>
              <div style="color:var(--text-muted); font-size:0.85rem;">Academic Year</div>
              <div>${student.current_academic_year || 'N/A'}</div>
            </div>
            <div>
              <div style="color:var(--text-muted); font-size:0.85rem;">Admission Year</div>
              <div>${student.year_of_admission || 'N/A'}</div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Student Mobile</div>
              <div>${student.student_mobile || 'N/A'}</div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Parent Mobile</div>
              <div>${student.parent_mobile || 'N/A'}</div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Student Email</div>
              <div>${student.student_email || 'N/A'}</div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Parent Email</div>
              <div>${student.parent_email || 'N/A'}</div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Local Address</div>
              <div style="font-size:0.9rem; opacity:0.9;">${student.local_address || 'N/A'}</div>
            </div>
            <div style="grid-column: span 2;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Permanent Address</div>
              <div style="font-size:0.9rem; opacity:0.9;">${student.permanent_address || 'N/A'}</div>
            </div>
            <div style="grid-column: span 4; border-top:1px dashed var(--border); padding-top:8px;">
              <div style="color:var(--text-muted); font-size:0.85rem;">Remarks</div>
              <div style="font-style:italic; font-size:0.9rem;">${student.remark || 'No remarks.'}</div>
            </div>
          </div>
        </div>
      `;
    } else {
      leftHtml = `
        <div style="display:flex; flex-direction:column; gap:16px;">
          <div style="display:flex; justify-content:space-between; align-items:center; border-bottom:1px solid var(--border); padding-bottom:8px;">
            <h4 style="margin:0; font-size:1.05rem; font-weight:700; color:var(--warning);">Edit Profile</h4>
            <div style="display:flex; gap:6px;">
              <button class="btn btn-ghost btn-sm" onclick="App.toggleDetailProfileEdit(${studentId}, false)" style="padding: 4px 8px; font-size:0.85rem;">Cancel</button>
              <button class="btn btn-primary btn-sm" onclick="App.saveDetailProfile(${studentId})" style="padding: 4px 8px; font-size:0.85rem;">Save</button>
            </div>
          </div>
          
          <div style="display:grid; grid-template-columns:repeat(2, 1fr); gap:10px; font-size:0.95rem;" id="detail-edit-form">
            <div style="grid-column: span 2;">
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Full Name *</label>
              <input type="text" class="form-input" id="det-edit-name" value="${student.name || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Gender</label>
              <select class="form-select" id="det-edit-gender" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
                <option value="Male" ${student.gender === 'Male' ? 'selected' : ''}>Male</option>
                <option value="Female" ${student.gender === 'Female' ? 'selected' : ''}>Female</option>
                <option value="Other" ${student.gender === 'Other' ? 'selected' : ''}>Other</option>
              </select>
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Category</label>
              <select class="form-select" id="det-edit-category" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
                <option value="General" ${student.category === 'General' ? 'selected' : ''}>General</option>
                <option value="OBC" ${student.category === 'OBC' ? 'selected' : ''}>OBC</option>
                <option value="SC" ${student.category === 'SC' ? 'selected' : ''}>SC</option>
                <option value="ST" ${student.category === 'ST' ? 'selected' : ''}>ST</option>
                <option value="NT" ${student.category === 'NT' ? 'selected' : ''}>NT</option>
                <option value="VJ-DT" ${student.category === 'VJ-DT' ? 'selected' : ''}>VJ-DT</option>
                <option value="EWS" ${student.category === 'EWS' ? 'selected' : ''}>EWS</option>
              </select>
            </div>
            <div style="grid-column: span 2;">
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Class</label>
              <select class="form-select" id="det-edit-class" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
                <option value="B.Pharm 1st Year" ${student.current_class === 'B.Pharm 1st Year' ? 'selected' : ''}>B.Pharm 1st Year</option>
                <option value="B.Pharm 2nd Year" ${student.current_class === 'B.Pharm 2nd Year' ? 'selected' : ''}>B.Pharm 2nd Year</option>
                <option value="B.Pharm 3rd Year" ${student.current_class === 'B.Pharm 3rd Year' ? 'selected' : ''}>B.Pharm 3rd Year</option>
                <option value="B.Pharm Final Year" ${student.current_class === 'B.Pharm Final Year' ? 'selected' : ''}>B.Pharm Final Year</option>
                <option value="M.Pharm 1st Year" ${student.current_class === 'M.Pharm 1st Year' ? 'selected' : ''}>M.Pharm 1st Year</option>
                <option value="M.Pharm 2nd Year" ${student.current_class === 'M.Pharm 2nd Year' ? 'selected' : ''}>M.Pharm 2nd Year</option>
              </select>
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Academic Year</label>
              <input type="text" class="form-input" id="det-edit-ay" value="${student.current_academic_year || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Admission Year</label>
              <input type="number" class="form-input" id="det-edit-year" value="${student.year_of_admission || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Student Mobile</label>
              <input type="tel" class="form-input" id="det-edit-smob" value="${student.student_mobile || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Parent Mobile</label>
              <input type="tel" class="form-input" id="det-edit-pmob" value="${student.parent_mobile || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Student Email</label>
              <input type="email" class="form-input" id="det-edit-semail" value="${student.student_email || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Parent Email</label>
              <input type="email" class="form-input" id="det-edit-pemail" value="${student.parent_email || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div style="grid-column: span 2;">
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Local Address</label>
              <input type="text" class="form-input" id="det-edit-laddr" value="${student.local_address || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div style="grid-column: span 2;">
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Permanent Address</label>
              <input type="text" class="form-input" id="det-edit-paddr" value="${student.permanent_address || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Fee Applicable (₹) *</label>
              <input type="number" class="form-input" id="det-edit-fee" value="${applicable}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div>
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Sch Expected (₹)</label>
              <input type="number" class="form-input" id="det-edit-schexp" value="${schExpected}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
            <div style="grid-column: span 2;">
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Remark</label>
              <input type="text" class="form-input" id="det-edit-remark" value="${student.remark || ''}" style="padding:6px 10px; font-size:0.95rem; background: var(--bg-primary);">
            </div>
          </div>
        </div>
      `;
    }

    // RIGHT COLUMN: FEE STATUS & TRANSACTION HISTORY
    const summaryCardHtml = `
      <div style="display:grid; grid-template-columns: 1fr 1fr; gap:12px;">
        <div style="background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-sm); padding:10px;">
          <div style="font-size:0.85rem; color:var(--text-muted); text-transform:uppercase; font-weight:600;">Fee Status</div>
          <div style="display:flex; flex-direction:column; gap:4px; margin-top:6px; font-size:0.95rem;">
            <div style="display:flex; justify-content:space-between; gap: 4px;"><span>App:</span><strong>₹${applicable.toLocaleString('en-IN')}</strong></div>
            <div style="display:flex; justify-content:space-between; gap: 4px; color:var(--success);"><span>Paid:</span><strong>₹${paid.toLocaleString('en-IN')}</strong></div>
            <div style="display:flex; justify-content:space-between; gap: 4px; color:${pending > 0 ? 'var(--warning)' : 'var(--success)'};"><span>Pending:</span><strong>₹${pending.toLocaleString('en-IN')}</strong></div>
          </div>
        </div>
        
        <div style="background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-sm); padding:10px;">
          <div style="font-size:0.85rem; color:var(--text-muted); text-transform:uppercase; font-weight:600;">Scholarship</div>
          <div style="display:flex; flex-direction:column; gap:4px; margin-top:6px; font-size:0.95rem;">
            <div style="display:flex; justify-content:space-between; gap: 4px; color:var(--purple);"><span>Exp:</span><strong>₹${schExpected.toLocaleString('en-IN')}</strong></div>
            <div style="display:flex; justify-content:space-between; gap: 4px; color:var(--success);"><span>Rec:</span><strong>₹${schReceived.toLocaleString('en-IN')}</strong></div>
            <div style="display:flex; justify-content:space-between; gap: 4px; color:var(--pink);"><span>Pending:</span><strong>₹${schPending.toLocaleString('en-IN')}</strong></div>
          </div>
        </div>
      </div>
    `;

    const inlinePaymentFormHtml = `
      <div style="background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-sm); padding:14px; display:flex; flex-direction:column; gap:10px;">
        <h5 style="margin:0; font-size:0.95rem; font-weight:700; color:var(--accent); text-transform:uppercase; letter-spacing:0.02em;">Record Manual Payment</h5>
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
          <div style="grid-column: span 2; display: flex; flex-direction: column; gap: 2px;">
            <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Payment Type</label>
            <div style="display:flex; background:var(--toggle-track-bg); padding:3px; border-radius:6px; border:1px solid var(--border);">
              <button id="btn-paytype-fee" class="btn-paytype active" style="flex:1;" data-pending="${pending}" onclick="App.setDetailPaymentType('Fee')">Fee</button>
              <button id="btn-paytype-sch" class="btn-paytype" style="flex:1;" data-pending="${schPending}" onclick="App.setDetailPaymentType('Scholarship')">Scholarship</button>
              <button id="btn-paytype-other" class="btn-paytype" style="flex:1;" onclick="App.setDetailPaymentType('Other')">Other</button>
            </div>
            <input type="hidden" id="det-pay-type" value="Fee">
            <div id="det-pay-other-reason-container" style="display:none; margin-top: 8px;">
              <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Reason for Other Fee *</label>
              <input type="text" class="form-input" id="det-pay-other-reason" placeholder="e.g. Fine, Contribution" style="padding:4px 8px; font-size:0.9rem; background: var(--bg-primary);">
            </div>
          </div>
          <div>
            <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Amount (₹) *</label>
            <input type="number" class="form-input" id="det-pay-amount" value="${pending > 0 ? pending : ''}" placeholder="Amount" style="padding:4px 8px; font-size:0.9rem; background: var(--bg-primary);">
          </div>
          <div>
            <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">UTR / Txn ID *</label>
            <input type="text" class="form-input" id="det-pay-utr" placeholder="UTR or Txn ID" style="padding:4px 8px; font-size:0.9rem; background: var(--bg-primary);">
          </div>
          <div>
            <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Payment Date</label>
            <input type="date" class="form-input" id="det-pay-date" value="${new Date().toISOString().split('T')[0]}" style="padding:4px 8px; font-size:0.9rem; background: var(--bg-primary);">
          </div>
          <div>
            <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Bank / Method</label>
            <input type="text" class="form-input" id="det-pay-bank" value="GPay" placeholder="Method" style="padding:4px 8px; font-size:0.9rem; background: var(--bg-primary);">
          </div>
          <div style="grid-column: span 2;">
            <label class="form-label" style="font-size:0.85rem; margin-bottom:2px;">Remark</label>
            <input type="text" class="form-input" id="det-pay-remark" placeholder="Remark details" style="padding:4px 8px; font-size:0.9rem; background: var(--bg-primary);">
          </div>
        </div>
        <button class="btn btn-primary btn-sm" onclick="App.submitDetailPayment(${studentId})" style="width:100%; margin-top:4px; padding:6px 12px; font-size:0.9rem;">
          <i class="ph-bold ph-plus-circle"></i> Record Payment
        </button>
      </div>
    `;

    let txnsHtml = '';
    if (txns.length === 0) {
      txnsHtml = `<div style="text-align:center; padding: 12px; color:var(--text-muted); font-size:0.9rem;">No payment records.</div>`;
    } else {
      txns.forEach(t => {
        const isVer = parseInt(t.verified) === 1;
        const badge = isVer
          ? `<span class="badge badge-success">Verified</span>`
          : `<span class="badge badge-warning">Pending</span>`;
        const actionBtn = isVer
          ? `<button class="btn btn-danger btn-sm" style="padding: 2px 6px; font-size:0.85rem;" onclick="App.toggleDetailTxnVerification(${t.id}, 0, ${studentId})">Unverify</button>`
          : `<button class="btn btn-success btn-sm" style="padding: 2px 6px; font-size:0.85rem;" onclick="App.toggleDetailTxnVerification(${t.id}, 1, ${studentId})">Verify</button>`;

        let typeBadge = '';
        if (t.type === 'Scholarship') {
          typeBadge = `<span class="badge badge-purple" style="font-size:0.85rem; padding: 1px 6px; margin-left: 6px;">Scholarship</span>`;
        } else if (t.type === 'Other') {
          typeBadge = `<span class="badge badge-accent" style="font-size:0.85rem; padding: 1px 6px; margin-left: 6px;">Other: ${t.other_reason || 'N/A'}</span>`;
        } else {
          typeBadge = `<span class="badge badge-success" style="font-size:0.85rem; padding: 1px 6px; margin-left: 6px;">Fee</span>`;
        }

        txnsHtml += `
          <div style="display:flex; justify-content:space-between; align-items:center; padding:8px 0; border-bottom:1px solid var(--border); font-size:0.9rem;">
            <div>
              <div style="font-weight:600; color:var(--text-primary); display:flex; align-items:center;">
                ₹${(parseFloat(t.amount) || 0).toLocaleString('en-IN')}
                ${typeBadge}
              </div>
              <div style="font-size:0.85rem; color:var(--text-muted);">UTR: <code>${t.transaction_id || 'N/A'}</code> · <span style="color:var(--text-primary); font-weight:600; display:inline-flex; align-items:center; gap:4px;"><i class="ph-bold ph-calendar-blank" style="color:var(--accent);"></i> ${t.date || 'N/A'}</span></div>
              <div style="font-size:0.85rem; color:var(--text-secondary);">${t.bank_name || 'N/A'} ${t.raw_remark ? '(' + t.raw_remark + ')' : ''}</div>
            </div>
            <div style="display:flex; flex-direction:column; align-items:flex-end; gap:4px;">
              ${badge}
              ${actionBtn}
            </div>
          </div>
        `;
      });
    }

    const txnHistoryHtml = `
      <div style="background:var(--surface); border:1px solid var(--border); border-radius:var(--radius-sm); padding:14px; flex:1; display:flex; flex-direction:column; min-height: 180px;">
        <h5 style="margin:0 0 8px; font-size:0.95rem; font-weight:700; color:var(--accent); text-transform:uppercase; letter-spacing:0.02em;">Payment History</h5>
        <div class="detail-txn-list" style="overflow-y:auto; max-height:220px; padding-right:4px; flex:1;">
          ${txnsHtml}
        </div>
      </div>
    `;

    return `
      <div class="detail-modal-container" style="display:flex; gap:24px; min-height:0; flex-wrap:nowrap;">
        <!-- Left Column -->
        <div class="detail-profile-col" style="flex: 1; min-width: 320px; border-right: 1px solid var(--border); padding-right: 24px;">
          ${leftHtml}
        </div>
        <!-- Right Column -->
        <div class="detail-payments-col" style="flex: 1.2; min-width: 340px; display:flex; flex-direction:column; gap:16px;">
          ${summaryCardHtml}
          ${inlinePaymentFormHtml}
          ${txnHistoryHtml}
        </div>
      </div>
    `;
  }

  function showPayModal(studentId) {
    const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
    if (!student) return;

    const body = `
      <div style="display:flex; flex-direction:column; gap:14px;">
        <div style="font-size:0.9rem; color:var(--text-secondary);">
          Recording payment for <strong>${student.name}</strong> (${student.current_class})
        </div>
        <div class="form-group">
          <label class="form-label">Amount to Pay (₹) *</label>
          <input type="number" class="form-input" id="pay-amount" value="${student.fee_pending}" required>
        </div>
        <div class="form-group">
          <label class="form-label">Transaction ID / UTR No. *</label>
          <input type="text" class="form-input" id="pay-utr" placeholder="UTR or Bank Reference No." required>
        </div>
        <div class="form-group">
          <label class="form-label">Date of Payment</label>
          <input type="date" class="form-input" id="pay-date" value="${new Date().toISOString().split('T')[0]}">
        </div>
        <div class="form-group">
          <label class="form-label">Bank Name / Payment Method</label>
          <input type="text" class="form-input" id="pay-bank" placeholder="e.g. HDFC Bank, GPay, Cash">
        </div>
        <div class="form-group">
          <label class="form-label">Remark</label>
          <input type="text" class="form-input" id="pay-remark" placeholder="e.g. Tuition fee part 1">
        </div>
      </div>
    `;

    const footer = `
      <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
      <button class="btn btn-primary" onclick="App.submitPayment(${studentId})">Record Payment</button>
    `;

    openModal('Record Manual Payment', body, footer);
  }

  async function submitPayment(studentId) {
    if (!(await requirePin())) return;
    const amount = parseFloat(document.getElementById('pay-amount').value) || 0;
    const utr = document.getElementById('pay-utr').value.trim();
    const date = document.getElementById('pay-date').value;
    const bank = document.getElementById('pay-bank').value.trim() || 'Manual Entry';
    const remark = document.getElementById('pay-remark').value.trim();

    if (amount <= 0 || !utr) {
      alert('Please fill all required fields correctly.');
      return;
    }

    closeModal();
    Toast.show('Recording transaction...', 'info');

    try {
      const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
      const res = await API.recordPayment({
        student_id: studentId,
        student_name: student?.name || 'Unknown',
        transaction_id: utr,
        date: date,
        amount: amount,
        bank_name: bank,
        source: 'Manual Entry',
        status: 'Unverified',
        verified: 0,
        raw_remark: remark
      });

      if (res.success) {
        Toast.show('Payment recorded successfully!', 'success');
        await refreshData();
        renderStudents();
      } else {
        Toast.show(res.error || 'Payment failed.', 'error');
      }
    } catch (e) {
      Toast.show('Error connecting to Server API.', 'error');
    }
  }

  function showHistoryModal(studentId) {
    const student = state.students.find(s => parseInt(s.id) === parseInt(studentId));
    if (!student) return;

    const txns = state.transactions.filter(t => parseInt(t.student_id) === parseInt(studentId));
    let rowsHtml = '';

    if (txns.length === 0) {
      rowsHtml = `<tr><td colspan="6" style="text-align:center; padding:12px; color:var(--text-muted)">No payment records.</td></tr>`;
    } else {
      txns.forEach(t => {
        const isVer = parseInt(t.verified) === 1;
        const btnText = isVer ? 'Unverify' : 'Verify';
        const btnClass = isVer ? 'btn-danger' : 'btn-success';
        const badge = isVer ? '<span class="badge badge-success">Verified</span>' : '<span class="badge badge-warning">Pending</span>';

        let typeBadgeHtml = '';
        if (t.type === 'Scholarship') {
          typeBadgeHtml = `<span class="badge badge-purple" style="font-size:0.75rem; margin-left:6px; padding:2px 6px;">Scholarship</span>`;
        } else if (t.type === 'Other') {
          typeBadgeHtml = `<span class="badge badge-accent" style="font-size:0.75rem; margin-left:6px; padding:2px 6px;">Other: ${t.other_reason || 'N/A'}</span>`;
        } else {
          typeBadgeHtml = `<span class="badge badge-success" style="font-size:0.75rem; margin-left:6px; padding:2px 6px;">Fee</span>`;
        }

        rowsHtml += `
          <tr>
            <td>${t.date || 'N/A'}</td>
            <td><strong>₹${(t.amount || 0).toLocaleString('en-IN')}</strong>${typeBadgeHtml}</td>
            <td><code>${t.transaction_id || 'N/A'}</code></td>
            <td>${t.bank_name || 'N/A'}</td>
            <td>${badge}</td>
            <td>
              <button class="btn ${btnClass} btn-sm" onclick="App.toggleTxnVerification(${t.id}, ${isVer ? 0 : 1}, ${studentId})">
                ${btnText}
              </button>
            </td>
          </tr>
        `;
      });
    }

    const body = `
      <div style="font-size:0.9rem; color:var(--text-secondary); margin-bottom:12px;">
        Payment history for <strong>${student.name}</strong> (${student.current_class})
      </div>
      <div class="table-wrapper" style="max-height: 350px;">
        <table class="data-table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Amount</th>
              <th>UTR/Txn ID</th>
              <th>Bank</th>
              <th>Status</th>
              <th>Action</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
      </div>
    `;

    const footer = `
      <button class="btn btn-ghost" onclick="App.closeModal()">Close</button>
    `;

    openModal('Transaction History', body, footer);
  }

  async function toggleTxnVerification(txnId, verifyState, studentId) {
    if (!(await requirePin())) return;
    closeModal();
    const msg = verifyState
      ? 'Are you sure you want to mark this transaction as Verified?'
      : 'Are you sure you want to Unverify this transaction?';

    appConfirm(verifyState ? 'Verify Payment' : 'Unverify Payment', msg, async () => {
      Toast.show(verifyState ? 'Verifying payment...' : 'Unverifying payment...', 'info');
      try {
        const res = await API.verifyTransaction(txnId, verifyState);
        if (res.success) {
          Toast.show(verifyState ? 'Transaction verified!' : 'Transaction unverified!', 'success');
          await refreshData();
          renderStudents();
          // Re-show history modal
          setTimeout(() => showHistoryModal(studentId), 200);
        } else {
          Toast.show(res.error || 'Action failed.', 'error');
        }
      } catch (e) {
        Toast.show('Network error updating verification.', 'error');
      }
    });
  }

  async function toggleAdmission(studentId, cancelState) {
    if (!(await requirePin())) return;
    const actionText = cancelState ? 'Cancel Admission' : 'Restore Admission';

    appConfirm('Confirm Action', `Are you sure you want to ${actionText}?`, async () => {
      Toast.show(cancelState ? 'Cancelling admission...' : 'Restoring student...', 'info');
      try {
        const res = await API.updateStudent({
          id: studentId,
          admission_cancelled: cancelState
        });
        if (res.success) {
          Toast.show(cancelState ? 'Admission cancelled.' : 'Student restored.', 'success');
          await refreshData();
          renderStudents();
        } else {
          Toast.show(res.error || 'Update failed.', 'error');
        }
      } catch (e) {
        Toast.show('Network error updating student profile.', 'error');
      }
    });
  }

  // ─── REPORTS GENERATOR & EXPORT ───

  function renderReports() {
    const tbody = document.getElementById('reports-table-body');
    const tfoot = document.getElementById('reports-table-foot');
    tbody.innerHTML = '';
    tfoot.innerHTML = '';

    // Aggregate class details
    const classData = {};
    let totalSt = 0, totalM = 0, totalF = 0;
    let totalApp = 0, totalPaid = 0, totalPend = 0;
    let totalSchExp = 0, totalSchRec = 0, totalSchPend = 0;

    state.students.forEach(s => {
      if (parseInt(s.admission_cancelled) === 1) return;
      const cls = s.current_class || 'Unclassed';
      if (!classData[cls]) {
        classData[cls] = {
          name: cls, students: 0, male: 0, female: 0,
          applicable: 0, paid: 0, pending: 0,
          schExp: 0, schRec: 0, schPend: 0
        };
      }

      const c = classData[cls];
      c.students++;
      totalSt++;

      if (safeLower(s.gender) === 'male') { c.male++; totalM++; }
      else if (safeLower(s.gender) === 'female') { c.female++; totalF++; }

      const appVal = parseFloat(s.fee_applicable) || 0;
      const paidVal = parseFloat(s.fee_paid) || 0;
      const pendVal = parseFloat(s.fee_pending) || 0;
      const schEVal = parseFloat(s.scholarship_expected) || 0;
      const schRVal = parseFloat(s.scholarship_received) || 0;
      const schPVal = parseFloat(s.scholarship_pending) || 0;

      c.applicable += appVal; totalApp += appVal;
      c.paid += paidVal; totalPaid += paidVal;
      c.pending += pendVal; totalPend += pendVal;
      c.schExp += schEVal; totalSchExp += schEVal;
      c.schRec += schRVal; totalSchRec += schRVal;
      c.schPend += schPVal; totalSchPend += schPVal;
    });

    const classes = Object.keys(classData);
    if (classes.length === 0) {
      tbody.innerHTML = `<tr><td colspan="10" style="text-align:center; padding: 24px; color: var(--text-muted);">No database entries available.</td></tr>`;
      return;
    }

    classes.forEach(k => {
      const c = classData[k];
      const rate = c.applicable > 0 ? (c.paid / c.applicable) * 100 : 0;
      const totalPendingSum = c.pending + c.schPend;
      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td style="font-weight:600;">${c.name}</td>
        <td style="text-align: center;">${c.students}</td>
        <td>₹${c.applicable.toLocaleString('en-IN')}</td>
        <td class="text-success">₹${c.paid.toLocaleString('en-IN')}</td>
        <td class="text-warning">₹${c.pending.toLocaleString('en-IN')}</td>
        <td>₹${c.schExp.toLocaleString('en-IN')}</td>
        <td class="text-purple">₹${c.schRec.toLocaleString('en-IN')}</td>
        <td class="text-pink">₹${c.schPend.toLocaleString('en-IN')}</td>
        <td class="text-danger" style="font-weight:700;">₹${totalPendingSum.toLocaleString('en-IN')}</td>
        <td style="font-weight:600;">${rate.toFixed(1)}%</td>
      `;
      tbody.appendChild(tr);
    });

    const footRate = totalApp > 0 ? (totalPaid / totalApp) * 100 : 0;
    const grandTotalPending = totalPend + totalSchPend;
    tfoot.innerHTML = `
      <tr>
        <td>TOTALS</td>
        <td style="text-align: center;">${totalSt}</td>
        <td>₹${totalApp.toLocaleString('en-IN')}</td>
        <td class="text-success">₹${totalPaid.toLocaleString('en-IN')}</td>
        <td class="text-warning">₹${totalPend.toLocaleString('en-IN')}</td>
        <td>₹${totalSchExp.toLocaleString('en-IN')}</td>
        <td class="text-purple">₹${totalSchRec.toLocaleString('en-IN')}</td>
        <td class="text-pink">₹${totalSchPend.toLocaleString('en-IN')}</td>
        <td class="text-danger" style="font-weight:700;">₹${grandTotalPending.toLocaleString('en-IN')}</td>
        <td style="font-weight:600;">${footRate.toFixed(1)}%</td>
      </tr>
    `;

    renderHighRiskStudents(state.students);
  }

  // ─── CLIENT-SIDE EXCEL EXPORTS (SHEETJS) ───

  function downloadBlankTemplate() {
    const headers = [
      'admission_no', 'name', 'gender', 'student_mobile', 'parent_mobile', 'student_email', 'parent_email',
      'local_address', 'permanent_address', 'year_of_admission', 'category', 'current_class',
      'current_academic_year', 'fee_applicable', 'fee_paid', 'scholarship_expected',
      'scholarship_received', 'remark'
    ];
    const ws = XLSX.utils.aoa_to_sheet([headers]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Import Template');
    XLSX.writeFile(wb, 'fee_import_template.xlsx');
  }

  function exportReportsExcel() {
    const wb = XLSX.utils.book_new();

    // Page 1: Overall Statistics / Class summaries
    const sumTable = [['Class Summary Report', '', '', '', '', '', '', '', '', ''], []];
    sumTable.push(['Class', 'Total Students', 'Male', 'Female', 'Fee Applicable', 'Fee Collected', 'Fee Pending', 'Scholarship Expected', 'Scholarship Received', 'Scholarship Pending', 'Collection Rate']);

    const classData = {};
    state.students.forEach(s => {
      if (parseInt(s.admission_cancelled) === 1) return;
      const cls = s.current_class || 'Unclassed';
      if (!classData[cls]) {
        classData[cls] = {
          name: cls, students: 0, male: 0, female: 0,
          applicable: 0, paid: 0, pending: 0,
          schExp: 0, schRec: 0, schPend: 0
        };
      }
      const c = classData[cls];
      c.students++;
      if (safeLower(s.gender) === 'male') c.male++;
      else if (safeLower(s.gender) === 'female') c.female++;
      c.applicable += parseFloat(s.fee_applicable) || 0;
      c.paid += parseFloat(s.fee_paid) || 0;
      c.pending += parseFloat(s.fee_pending) || 0;
      c.schExp += parseFloat(s.scholarship_expected) || 0;
      c.schRec += parseFloat(s.scholarship_received) || 0;
      c.schPend += parseFloat(s.scholarship_pending) || 0;
    });

    Object.keys(classData).forEach(k => {
      const c = classData[k];
      const rate = c.applicable > 0 ? (c.paid / c.applicable) * 100 : 0;
      sumTable.push([c.name, c.students, c.male, c.female, c.applicable, c.paid, c.pending, c.schExp, c.schRec, c.schPend, rate.toFixed(1) + '%']);
    });

    const wsSum = XLSX.utils.aoa_to_sheet(sumTable);
    XLSX.utils.book_append_sheet(wb, wsSum, 'Summary Report');

    // Page 2: All Student Database rows
    const dbHeaders = ['Admission No.', 'Name', 'Gender', 'Class', 'Category', 'Mobile', 'Email', 'Fee Applicable', 'Fee Paid', 'Fee Pending', 'Sch Expected', 'Sch Received', 'Sch Pending', 'Status'];
    const dbRows = state.students.map(s => [
      s.serial_no, s.name, s.gender, s.current_class, s.category, s.student_mobile, s.student_email,
      s.fee_applicable, s.fee_paid, s.fee_pending, s.scholarship_expected, s.scholarship_received, s.scholarship_pending,
      parseInt(s.admission_cancelled) === 1 ? 'Cancelled' : 'Active'
    ]);
    const wsDb = XLSX.utils.aoa_to_sheet([dbHeaders, ...dbRows]);
    XLSX.utils.book_append_sheet(wb, wsDb, 'All Students');

    XLSX.writeFile(wb, `${state.collegeName.replace(/[^a-zA-Z0-9]/g, '_')}_Fee_Report.xlsx`);
    Toast.show('Reports excel generated!', 'success');
  }

  function exportCustomExcel() {
    // 1. Get selected columns
    const selectedCols = Array.from(document.querySelectorAll('.export-col-check:checked')).map(el => el.value);
    // 2. Get selected classes
    const selectedClasses = Array.from(document.querySelectorAll('.export-class-check:checked')).map(el => el.value);

    if (selectedCols.length === 0 || selectedClasses.length === 0) {
      alert('Please select at least one column and one class to export.');
      return;
    }

    const filtered = state.students.filter(s => {
      if (parseInt(s.admission_cancelled) === 1) return false; // exclude soft deleted
      return selectedClasses.includes(s.current_class);
    });

    const headers = selectedCols.map(key => {
      const match = EXPORT_COLUMNS.find(c => c.key === key);
      return match ? match.label : key;
    });

    const rows = filtered.map(s => {
      return selectedCols.map(key => {
        return s[key] !== undefined ? s[key] : '';
      });
    });

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Custom Export');
    XLSX.writeFile(wb, `${state.collegeName.replace(/[^a-zA-Z0-9]/g, '_')}_Custom_Export.xlsx`);
    Toast.show('Custom report exported!', 'success');
  }

  function downloadBackupExcel() {
    const wb = XLSX.utils.book_new();

    // Students sheet schema keys (maps to GSheets columns)
    const studentDbKeys = [
      'id', 'serial_no', 'name', 'gender', 'student_mobile', 'parent_mobile',
      'student_email', 'parent_email', 'permanent_address', 'local_address',
      'year_of_admission', 'category', 'current_class', 'current_academic_year',
      'fee_applicable', 'fee_paid', 'fee_pending', 'scholarship_expected',
      'scholarship_received', 'scholarship_pending', 'remark', 'admission_cancelled'
    ];
    const sHeaders = studentDbKeys.map(h => h === 'serial_no' ? 'admission_no' : h);
    const sRows = state.students.map(s => studentDbKeys.map(k => s[k] !== undefined ? s[k] : ''));
    const wsS = XLSX.utils.aoa_to_sheet([sHeaders, ...sRows]);
    XLSX.utils.book_append_sheet(wb, wsS, 'Students');

    // Transactions sheet schema keys (maps to GSheets columns)
    const transactionDbKeys = [
      'id', 'student_id', 'student_name', 'transaction_id', 'date', 'amount',
      'bank_name', 'source', 'status', 'verified', 'raw_remark', 'type', 'other_reason'
    ];
    const tRows = state.transactions.map(t => transactionDbKeys.map(k => t[k] !== undefined ? t[k] : ''));
    const wsT = XLSX.utils.aoa_to_sheet([transactionDbKeys, ...tRows]);
    XLSX.utils.book_append_sheet(wb, wsT, 'Transactions');

    // Settings sheet
    const setHeaders = ['key', 'value'];
    const setRows = Object.entries(state.settings).map(([k, v]) => [k, v]);
    const wsSet = XLSX.utils.aoa_to_sheet([setHeaders, ...setRows]);
    XLSX.utils.book_append_sheet(wb, wsSet, 'Settings');

    XLSX.writeFile(wb, `${state.collegeName.replace(/[^a-zA-Z0-9]/g, '_')}_System_Backup.xlsx`);
    Toast.show('Backup excel file downloaded!', 'success');
  }

  // ─── CLIENT-SIDE EXCEL BULK IMPORTS ───

  function handleExcelImport(file) {
    const reader = new FileReader();

    const status = document.getElementById('upload-status');
    const fill = document.getElementById('upload-progress-bar');
    const txt = document.getElementById('upload-status-text');

    status.style.display = 'block';
    fill.style.width = '20%';
    txt.textContent = 'Reading excel file...';

    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Parse rows as raw JSON array
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        if (rawRows.length < 2) {
          throw new Error('Excel sheet must contain at least headers and one student row.');
        }

        fill.style.width = '50%';
        txt.textContent = 'Parsing column aliases...';

        // Column mapping helper aliases matching spec
        const aliases = {
          serial_no: ['serial_no', 'serial no', 'admission_no', 'admission no', 'adm no.', 'adm no', 'admission number', 'serial_no'],
          name: ['student_name', 'full_name', 'student name', 'full name', 'name'],
          gender: ['sex', 'gender'],
          student_mobile: ['student_phone', 'mobile', 'mobile_no', 'student mobile', 'student mobile no'],
          parent_mobile: ['parent_phone', 'father_mobile', 'parent mobile', 'parent mobile no'],
          student_email: ['email', 'student_email', 'student email'],
          parent_email: ['father_email', 'parent_email', 'parent email'],
          permanent_address: ['address', 'permanent_address', 'permanent address'],
          local_address: ['current_address', 'local_address', 'local address'],
          year_of_admission: ['admission_year', 'year', 'year_of_admission', 'admission year'],
          category: ['caste', 'category'],
          current_class: ['class', 'year_class', 'current_class', 'current class'],
          current_academic_year: ['academic_year', 'current_academic_year', 'academic year'],
          fee_applicable: ['total_fees', 'fees_applicable', 'total_fee', 'fee_applicable', 'fee applicable'],
          fee_paid: ['fees_paid', 'amount_paid', 'paid', 'fee_paid', 'fee paid'],
          scholarship_expected: ['expected_scholarship', 'scholarship_expected', 'scholarship expected'],
          scholarship_received: ['scholarship', 'scholorship', 'scholarship_received', 'scholarship received'],
          remark: ['remarks', 'note', 'remark']
        };

        const headers = rawRows[0].map(h => String(h || '').trim().toLowerCase());

        // Find columns index maps
        const colMap = {};
        for (const [key, searchKeys] of Object.entries(aliases)) {
          colMap[key] = headers.findIndex(h => searchKeys.includes(h));
        }

        // Process rows
        const parsedRows = [];
        for (let i = 1; i < rawRows.length; i++) {
          const rawRow = rawRows[i];
          if (rawRow.length === 0 || !rawRow.some(cell => cell !== undefined && cell !== '')) continue;

          const rowObject = {};
          for (const [key, idx] of Object.entries(colMap)) {
            if (idx !== -1 && rawRow[idx] !== undefined) {
              rowObject[key] = rawRow[idx];
            } else {
              rowObject[key] = ''; // default fallback
            }
          }

          // Auto-calculations if missing
          const app = parseFloat(rowObject.fee_applicable) || 0;
          const paid = parseFloat(rowObject.fee_paid) || 0;
          rowObject.fee_applicable = app;
          rowObject.fee_paid = paid;
          rowObject.fee_pending = app - paid;

          const expSch = parseFloat(rowObject.scholarship_expected) || 0;
          const recSch = parseFloat(rowObject.scholarship_received) || 0;
          rowObject.scholarship_expected = expSch;
          rowObject.scholarship_received = recSch;
          rowObject.scholarship_pending = expSch - recSch;

          rowObject.admission_cancelled = 0;

          parsedRows.push(rowObject);
        }

        fill.style.width = '75%';
        txt.textContent = `Uploading ${parsedRows.length} student records...`;

        const res = await API.uploadStudents(parsedRows);
        if (res.success) {
          fill.style.width = '100%';
          txt.textContent = 'Import Complete!';
          Toast.show(`Successfully imported ${parsedRows.length} student records!`, 'success');

          await refreshData();
          populatePillFilters();

          setTimeout(() => {
            status.style.display = 'none';
            switchView('students');
          }, 1500);
        } else {
          throw new Error(res.error || 'Bulk sync rejected.');
        }

      } catch (err) {
        console.error(err);
        fill.style.width = '0%';
        txt.textContent = 'Import Failed!';
        Toast.show(err.message || 'Error parsing imported file.', 'error');
      }
    };

    reader.readAsArrayBuffer(file);
  }

  // ─── DESTRUCTIVE DATABASE RESTORE ───

  function handleRestoreBackup(file) {
    const confirmRestore = confirm('🚨 DANGER: You are about to perform a full database restore. This will permanently overwrite all existing students and payments in Google Sheets. Proceed?');
    if (!confirmRestore) return;

    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sSheet = workbook.Sheets['Students'];
        const tSheet = workbook.Sheets['Transactions'];

        if (!sSheet) {
          throw new Error('Backup file must contain a "Students" worksheet.');
        }

        const rawStudents = XLSX.utils.sheet_to_json(sSheet);
        const studentDbKeys = [
          'id', 'serial_no', 'name', 'gender', 'student_mobile', 'parent_mobile',
          'student_email', 'parent_email', 'permanent_address', 'local_address',
          'year_of_admission', 'category', 'current_class', 'current_academic_year',
          'fee_applicable', 'fee_paid', 'fee_pending', 'scholarship_expected',
          'scholarship_received', 'scholarship_pending', 'remark', 'admission_cancelled'
        ];

        // Normalization helper for robust casing/spacing matching
        const normalize = str => String(str || '').trim().toLowerCase().replace(/[\s_-]+/g, '').replace(/\./g, '');

        const mappedStudents = rawStudents.map(s => {
          const newS = {};
          for (const rawKey of Object.keys(s)) {
            const normRawKey = normalize(rawKey);

            // Map variations of admission number to serial_no
            if (normRawKey === 'admissionno' || normRawKey === 'serialno' || normRawKey === 'admno' || normRawKey === 'admissionnumber') {
              newS.serial_no = s[rawKey];
              continue;
            }

            const matchKey = studentDbKeys.find(dbKey => normalize(dbKey) === normRawKey);
            if (matchKey) {
              newS[matchKey] = s[rawKey];
            } else {
              newS[rawKey] = s[rawKey];
            }
          }
          return newS;
        });

        const rawTransactions = tSheet ? XLSX.utils.sheet_to_json(tSheet) : [];
        const transactionDbKeys = [
          'id', 'student_id', 'student_name', 'transaction_id', 'date', 'amount',
          'bank_name', 'source', 'status', 'verified', 'raw_remark', 'type', 'other_reason'
        ];
        const mappedTransactions = rawTransactions.map(t => {
          const newT = {};
          for (const rawKey of Object.keys(t)) {
            const normRawKey = normalize(rawKey);
            const matchKey = transactionDbKeys.find(dbKey => normalize(dbKey) === normRawKey);
            if (matchKey) {
              newT[matchKey] = t[rawKey];
            } else {
              newT[rawKey] = t[rawKey];
            }
          }
          return newT;
        });

        Toast.show('Initiating cloud rebuild...', 'info');
        const res = await API.restoreBackup(mappedStudents, mappedTransactions);
        if (res.success) {
          Toast.show(`Restored database: ${rawStudents.length} students, ${rawTransactions.length} transactions.`, 'success');
          await refreshData();
          switchView('dashboard');
        } else {
          Toast.show(res.error || 'Restore failed.', 'error');
        }
      } catch (err) {
        console.error(err);
        Toast.show(err.message || 'Error processing backup file.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // ─── BANK STATEMENT CLIENT-SIDE PDF AUTO-VERIFY ───

  async function handlePdfStatementUpload(file) {
    Toast.show('Reading Bank Statement PDF...', 'info');
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const buffer = e.target.result;
        const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;
        let text = '';

        // Extract text from pages
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          text += content.items.map(item => item.str).join(' ') + '\n';
        }

        parseAndVerifyStatement(text);
      } catch (err) {
        console.error(err);
        Toast.show('Failed to read PDF text details.', 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function parseAndVerifyStatement(pdfText) {
    // Regex for UTR/Txn IDs (Common Indian Bank formats: UTIB..., BARB..., SBIN..., PUNB... or 12-digit numeric sequences)
    // Matches 8 to 22 characters of typical txn IDs
    const utrRegex = /\b([A-Z]{4}\d{6,14}|\b\d{12}\b|[A-Z]{3,4}[0-9A-Z]{8,15})\b/gi;

    // Regex for Amounts (matching commas and decimal, e.g. 15,000.00 or 15000)
    const amountRegex = /\b\d{1,3}(,\d{3})*(\.\d{2})?\b/g;

    const utrs = [...new Set(pdfText.match(utrRegex) || [])];

    // Match against unverified transactions
    const unverifiedTxns = state.transactions.filter(t => parseInt(t.verified) === 0);
    const matches = [];

    unverifiedTxns.forEach(t => {
      const matchUtr = utrs.find(u => safeLower(u) === safeLower(t.transaction_id).trim());
      if (matchUtr) {
        // Find if the amount exists nearby in the PDF (or just verification on the UTR match)
        const cleanedAmt = parseFloat(t.amount);
        matches.push({
          txnId: t.id,
          studentName: t.student_name,
          utr: t.transaction_id,
          amount: cleanedAmt,
          date: t.date
        });
      }
    });

    if (matches.length === 0) {
      alert('🔍 No matching transaction IDs (UTRs) found in statement for unverified items.');
      return;
    }

    // Modal matching results interface
    let listHtml = '';
    matches.forEach((m, idx) => {
      listHtml += `
        <div style="display:flex; justify-content:space-between; align-items:center; padding: 10px; border-bottom: 1px solid var(--border);">
          <label class="checkbox-label">
            <input type="checkbox" class="statement-match-check" value="${m.txnId}" checked>
            <div>
              <strong>${m.studentName}</strong>
              <div style="font-size:0.85rem; color:var(--text-muted);">UTR: <code>${m.utr}</code> · <span style="color:var(--text-primary); font-weight:600; display:inline-flex; align-items:center; gap:4px;"><i class="ph-bold ph-calendar-blank" style="color:var(--accent);"></i> ${m.date}</span></div>
            </div>
          </label>
          <span style="font-weight:700; color:var(--success);">₹${m.amount.toLocaleString('en-IN')}</span>
        </div>
      `;
    });

    const body = `
      <div style="font-size:0.9rem; color:var(--text-secondary); margin-bottom:14px;">
        The system scanned the statement and matched the following <strong>${matches.length}</strong> transactions:
      </div>
      <div style="max-height: 300px; overflow-y:auto; border: 1px solid var(--border); border-radius: var(--radius-sm); margin-bottom:16px;">
        ${listHtml}
      </div>
    `;

    const footer = `
      <button class="btn btn-ghost" onclick="App.closeModal()">Cancel</button>
      <button class="btn btn-primary" onclick="App.submitStatementVerification()">Verify Selected</button>
    `;

    openModal('Bank Statement Auto-Match Results', body, footer);
  }

  async function submitStatementVerification() {
    const selectedIds = Array.from(document.querySelectorAll('.statement-match-check:checked')).map(el => parseInt(el.value));

    closeModal();
    if (selectedIds.length === 0) return;

    Toast.show(`Verifying ${selectedIds.length} payments...`, 'info');
    try {
      let failed = 0;
      for (const id of selectedIds) {
        const res = await API.verifyTransaction(id, 1);
        if (!res.success) failed++;
      }

      if (failed > 0) {
        Toast.show(`Verified ${selectedIds.length - failed} transactions, ${failed} failed.`, 'warning');
      } else {
        Toast.show(`Successfully verified all ${selectedIds.length} transactions!`, 'success');
      }

      await refreshData();
      renderStudents();
    } catch (e) {
      Toast.show('Error during bulk statement verification.', 'error');
    }
  }

  return {
    initFromEngine,
    closeModal,
    appConfirm,
    lockApp,
    showPayModal,
    submitPayment,
    showHistoryModal,
    toggleTxnVerification,
    toggleAdmission,
    submitStatementVerification,
    showStudentDetailModal,
    sendIndividualDueReminder,
    toggleDetailProfileEdit,
    saveDetailProfile,
    submitDetailPayment,
    setDetailPaymentType,
    toggleDetailTxnVerification
  };
})();

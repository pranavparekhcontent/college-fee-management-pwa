<p align="center">
  <img src="icons/feeflow-logo.png" alt="FeeFlow Logo" width="200">
</p>

<h1 align="center">FeeFlow — College Fee Management System</h1>

<p align="center">
  <strong>A comprehensive, offline-first Progressive Web App for managing student fees, scholarships, and notifications — synced to Google Sheets.</strong>
</p>

<p align="center">
  <img src="https://img.shields.io/badge/PWA-Offline%20Ready-34d399?style=for-the-badge&logo=pwa&logoColor=white" alt="PWA">
  <img src="https://img.shields.io/badge/JavaScript-ES6+-fbbf24?style=for-the-badge&logo=javascript&logoColor=black" alt="JavaScript">
  <img src="https://img.shields.io/badge/Google%20Sheets-Synced-38bdf8?style=for-the-badge&logo=google-sheets&logoColor=white" alt="Google Sheets">
  <img src="https://img.shields.io/badge/ECharts-5.5-a78bfa?style=for-the-badge" alt="ECharts">
  <img src="https://img.shields.io/badge/License-Proprietary-f87171?style=for-the-badge" alt="License">
</p>

---

## ✨ Features

### 📊 Visual Dashboard
- Real-time metrics: Active Students, Cancelled, Fees Collected, Pending, Scholarships
- **5 Chart Types**: Gauge, Bar, Donut/Pie, Rose, Stacked Bar
- Class-wise and Year-wise filtering with animated counters
- Powered by ECharts 5.5

### 👥 Student Database
- Full CRUD operations for student records
- Real-time search across all columns
- Cancel/Restore admission (soft delete)
- View complete payment history per student
- Class-wise filtering with toggle controls

### 📝 Admission & Bulk Import
- Single student admission form with auto-calculated fields
- **Excel bulk import** with intelligent column name mapping
- Supports 20+ column aliases for flexible uploads
- Auto-calculates fee pending & scholarship pending

### 🏦 Auto Bank Statement Verification
- Upload bank statement PDFs
- Automatically matches transactions by UTR/Transaction ID + Amount
- One-click verification of student payments
- Built-in PDF viewer for manual cross-referencing

### 📈 Reports & Analytics
- Class-wise fee collection summaries
- Gender distribution per class
- Collection rate percentages
- **One-click Excel export** with separate sheets per class

### 📤 Custom Excel Export
- Checkbox-based column selection
- Class-wise filtering
- Download tailored Excel files with only the data you need

### 🤖 AI Brain Assistant
- Natural language queries about your fee data
- Supports **OpenRouter** and **Groq** APIs
- Powered by Llama 3.1 models
- System prompt auto-includes all student data and dashboard metrics

### 📧 Smart Notifications
- Send custom announcements via **Email** or **SMS**
- Recipients: All Students & Parents / Parents Only / Students Only
- **ADB SMS integration** for free SMS via connected Android phone
- **Automated monthly reminders** (1st of each month at 10 AM)

### 🛡️ Backup & Restore
- Auto email backup on app startup/shutdown
- Manual backup to Excel with Students + Transactions sheets
- One-click restore from backup file

### 🔐 Secure License System
- HWID-bound license key activation
- Base-64 encoded keys with HMAC-SHA256 verification
- Expiry date and college name binding
- Full-screen lock screen for unauthorized access

### 🌙 Dark & Light Themes
- Premium dark mode with glassmorphism
- Clean light mode for daytime use
- Smooth CSS variable transitions

### ☁️ Offline-First with Google Sheets Sync
- Complete offline data entry with queue system
- Service Worker caching for all assets
- Automatic sync to centralized Google Sheet when online

---

## 🏗️ Tech Stack

| Technology | Purpose |
|---|---|
| **HTML5 + CSS3** | Structure & styling with CSS custom properties |
| **Vanilla JavaScript (ES6+)** | Core application logic — zero frameworks |
| **PWA / Service Worker** | Offline support, caching, install prompt |
| **ECharts 5.5** | Interactive dashboard visualizations |
| **SheetJS (XLSX)** | Excel import/export capabilities |
| **PDF.js** | Bank statement PDF parsing |
| **Google Apps Script** | Cloud backend (CRUD, sync, notifications) |
| **Google Sheets** | Centralized database |
| **OpenRouter / Groq** | AI assistant integration |
| **Inter + Outfit** | Typography (Google Fonts) |
| **Phosphor Icons** | UI iconography |

---

## 📁 Project Structure

```
college-fee-management-pwa/
├── index.html                  # Landing page
├── app.html                    # Main PWA application
├── manifest.json               # PWA manifest
├── sw.js                       # Service worker
├── version.json                # App version tracking
├── .gitignore
├── css/
│   └── app.css                 # Application styles
├── js/
│   ├── app.js                  # Core app logic
│   └── api.js                  # API communication layer
├── appstart/
│   ├── appstart.js             # Boot sequence
│   ├── appstart.css            # Boot UI styles
│   ├── config.js               # App configuration
│   ├── license.js              # License verification
│   ├── keystore.js             # Local key storage
│   ├── schema.js               # Data schema definitions
│   └── translator.js           # Column name translation
├── icons/
│   ├── icon-192.png            # PWA icon (192x192)
│   ├── icon-512.png            # PWA icon (512x512)
│   ├── feeflow-logo.png        # Brand logo
│   └── ai-logo.png             # AI assistant icon
└── google_apps_script/
    └── Central_API.gs          # Google Apps Script backend
```

---

## 🚀 Getting Started

### 1. Deploy Google Apps Script
1. Open [Google Apps Script](https://script.google.com)
2. Create a new project and paste the contents of `google_apps_script/Central_API.gs`
3. Deploy as a **Web App** with "Anyone" access
4. Copy the deployment URL

### 2. Configure the App
1. Open `appstart/config.js`
2. Set your Google Apps Script deployment URL
3. Set your Google Sheet ID

### 3. Host the PWA
Host the files on any static server (GitHub Pages, Netlify, Vercel, etc.):
- The app runs entirely in the browser
- No backend server required — uses Google Apps Script as the API
- Service Worker enables full offline functionality

### 4. Activate License
1. Open the app — you'll see the license activation screen
2. Enter your activation key (contact the developer)
3. The key is HWID-bound and has an expiry date

---

## 📱 Install as PWA

FeeFlow can be installed as a native-like application:

1. Open the app in **Chrome** or **Edge**
2. Click the **"Install App"** button in the navbar (or browser's install prompt)
3. The app will appear on your desktop/home screen
4. Works fully offline after first load

---

## 🎨 Screenshots

> *Launch the app to experience the premium glassmorphism UI with animated dashboard, dual themes, and responsive design.*

---

## 👨‍💻 Author

**VibeMantra Studio**

---

## 📄 License

This project is proprietary software. Unauthorized distribution or modification is prohibited.

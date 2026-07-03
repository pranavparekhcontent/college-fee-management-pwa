// ============================================================
//  APPSTART CONFIG — FeeNext College Fee Management System
//  All other appstart/ files remain untouched between projects.
// ============================================================

const APP_CONFIG = {

  // ── App Identity ──────────────────────────────────────────
  APP_NAME: "FEENEXT",
  APP_VERSION: "1.0.52",

  // ── Layout ───────────────────────────────────────────────
  LAYOUT: "desktop-first",

  // ── Theme (Dark Blue — matches PROMPT.md spec) ──────────
  THEME: {
    primary: "#38bdf8",   // Cyan accent
    secondary: "#22d3ee",   // Cyan hover
    danger: "#f87171",   // Red for errors
    bg: "#0b1120",   // Dark background
    surface: "#111827",   // Card surfaces
    border: "#1e293b",   // Subtle borders
    text: "#f1f5f9",   // Light text
    muted: "#64748b",   // Muted labels
  },

  // ── License ───────────────────────────────────────────────
  LICENSE_STORAGE_KEY: "fee_manager_license",

  // ── Config Sheet ──────────────────────────────────────────
  CONFIG_SHEET_URL:
    "https://docs.google.com/spreadsheets/d/1u6Em-j9NjDBX_EdKtU_VI5x3Dyn8Xe7lo_1A2zIqDrg/gviz/tq?tqx=out:json",

  // ── Data Fetcher ──────────────────────────────────────────
  dataFetcher: async (serverUrl, sheetId = "") => {
    const cleanUrl = serverUrl.replace(/\/+$/, "").replace(/\?.*$/, "");
    let targetUrl = cleanUrl + '?action=getAllData&ts=' + Date.now();
    if (sheetId) {
      targetUrl += '&sheetId=' + encodeURIComponent(sheetId);
    }
    return {
      allData: fetch(targetUrl, { cache: 'no-store' })
        .then(r => {
          if (!r.ok) throw new Error('HTTP ' + r.status);
          return r.json();
        })
        .catch(err => {
          console.error("AppStart Data Fetcher Error:", err);
          return { success: false, error: err.message };
        }),
    };
  },

  /** CALLBACKS */
  onComplete: (context) => {
    console.log("AppStart complete for:", context.collegeName);
  }
};

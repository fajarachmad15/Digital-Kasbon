/**
 * ============================================================================
 * MODULE: Config.gs
 * PROJECT: Digital Kasbon Petty Cash - Enterprise GAS Suite
 * DESCRIPTION: Centralized configuration, environment variables, & portal URLs
 * ============================================================================
 */

const AppConfig = (function () {
  const DEFAULT_CONFIG = {
    SUPABASE_URL: 'https://uvjtzkpyunfuzcbhubpg.supabase.co',
    SUPABASE_KEY: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InV2anR6a3B5dW5mdXpjYmh1YnBnIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc4MDMxMjkxNCwiZXhwIjoyMDk1ODg4OTE0fQ.khqhowBMuwRSkHQPJaR0bd_ZeCZ5QuIjC3a50XTxaQo',
    STORAGE_PROVIDER: 'DRIVE', // 'DRIVE' (Google Drive) | 'SUPABASE'
    DRIVE_ROOT_FOLDER_NAME: 'Digital Kasbon Bukti',
    DRIVE_ROOT_FOLDER_ID: '',
    DRIVE_MAX_FOLDER_BYTES: 950 * 1024 * 1024, // Jaga kapasitas folder tetap di bawah 1 GB (950 MB)
    STORAGE_BUCKET: 'kasbon-attachments',
    MAX_UPLOAD_BYTES: 1 * 1024 * 1024, // Maksimal 1 MB per file (setelah kompresi otomatis ~150-220 KB)
    TIMEZONE: 'Asia/Jakarta',

    // Interconnected Web App URLs (Live Deployed)
    URL_DASHBOARD: 'https://script.google.com/macros/s/AKfycbyczT2ZiO4w0RrBO11oykTClTBRzMDRu5QXSltWJgzZALGABQIcIJZeV9xckEZ7lHc/exec',
    URL_FORM_PENGAJUAN: 'https://script.google.com/macros/s/AKfycbxIIwX65loQQc3PuYoEU4Vgm0vqwwTOAmjUrLtSfWNB54rt3QyBs7EPjkayqW7eQWg/exec',
    URL_PORTAL_MANAGER: 'https://script.google.com/macros/s/AKfycbwhF50hYPVveIRiVRrN-xPsqXQEX8ca8S8qhYPejc0simgjOU-YyQf7gD2wngUcIiA/exec',
    URL_PORTAL_CASHIER: 'https://script.google.com/macros/s/AKfycbwA4RkZ2un-7q3kjkKWItYRon3OYrGftIHLP89oIn7SuaRjTBWRnu5vhA72_aLimXw/exec',
    URL_PORTAL_REQUESTER: 'https://script.google.com/macros/s/AKfycbxvcV2WMj4f-lGCpN6qQhs-zbZzvsnQ_WguRIA8moZHkFWgr-1NsFtUGFPR2yl5Kt0P/exec',
    URL_PORTAL_ADMIN: 'https://script.google.com/macros/s/AKfycbwLnkSpl8bmr1cQOgJmPsuUSRLtOCl51qOIOq2Y-iMGTke5zBFyy4KsMVYk9zhkzoVa/exec'
  };

  /**
   * Reads a configuration key from ScriptProperties first, falling back to DEFAULT_CONFIG.
   * @param {string} key
   * @return {string}
   */
  function get(key) {
    try {
      const props = PropertiesService.getScriptProperties();
      const val = props.getProperty(key);
      if (val !== null && val !== undefined && val !== '') {
        return val;
      }
    } catch (e) {
      // Fallback to default config if ScriptProperties unavailable
    }
    return DEFAULT_CONFIG[key] !== undefined ? DEFAULT_CONFIG[key] : '';
  }

  /**
   * Returns all interconnected portal URLs for cross-app navigation.
   */
  function getPortalUrls() {
    let currentAppUrl = '';
    try {
      currentAppUrl = ScriptApp.getService().getUrl() || '';
    } catch (e) {}

    return {
      dashboard: get('URL_DASHBOARD') || currentAppUrl,
      pengajuan: get('URL_FORM_PENGAJUAN') || currentAppUrl,
      manager: get('URL_PORTAL_MANAGER') || currentAppUrl,
      cashier: get('URL_PORTAL_CASHIER') || currentAppUrl,
      requester: get('URL_PORTAL_REQUESTER') || currentAppUrl,
      admin: get('URL_PORTAL_ADMIN') || currentAppUrl
    };
  }

  /**
   * Helper to save Web App URLs into ScriptProperties from the Apps Script Editor.
   * @param {Object} urlsObj
   */
  function savePortalUrls(urlsObj) {
    const props = PropertiesService.getScriptProperties();
    const payload = {};
    if (urlsObj.dashboard) payload.URL_DASHBOARD = urlsObj.dashboard.trim();
    if (urlsObj.pengajuan) payload.URL_FORM_PENGAJUAN = urlsObj.pengajuan.trim();
    if (urlsObj.manager) payload.URL_PORTAL_MANAGER = urlsObj.manager.trim();
    if (urlsObj.cashier) payload.URL_PORTAL_CASHIER = urlsObj.cashier.trim();
    if (urlsObj.requester) payload.URL_PORTAL_REQUESTER = urlsObj.requester.trim();
    if (urlsObj.admin) payload.URL_PORTAL_ADMIN = urlsObj.admin.trim();
    props.setProperties(payload, false);
    return payload;
  }

  return {
    get: get,
    getPortalUrls: getPortalUrls,
    savePortalUrls: savePortalUrls
  };
})();

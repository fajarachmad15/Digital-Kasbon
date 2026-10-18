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
    STORAGE_BUCKET: 'kasbon-attachments',
    MAX_UPLOAD_BYTES: 5 * 1024 * 1024, // 5 MB
    TIMEZONE: 'Asia/Jakarta',

    // Interconnected Web App URLs (Update via Script Properties or here after deployment)
    URL_DASHBOARD: '',
    URL_FORM_PENGAJUAN: '',
    URL_PORTAL_MANAGER: '',
    URL_PORTAL_CASHIER: '',
    URL_PORTAL_REQUESTER: ''
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
      requester: get('URL_PORTAL_REQUESTER') || currentAppUrl
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
    props.setProperties(payload, false);
    return payload;
  }

  return {
    get: get,
    getPortalUrls: getPortalUrls,
    savePortalUrls: savePortalUrls
  };
})();

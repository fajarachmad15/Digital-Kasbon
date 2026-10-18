/**
 * ============================================================================
 * MODULE: Controller.gs (PROJECT 1: GAS-DASHBOARD)
 * PROJECT: Digital Kasbon Petty Cash - Enterprise Control Center
 * DESCRIPTION: Backend controller for the Central Monitoring Dashboard
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.activeEmail = AuthService.getActiveUserEmail();
  template.initialId = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim() : '';
  return template.evaluate()
    .setTitle('Dashboard Monitoring E-Kasbon | Petty Cash')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches initial dashboard payload: stores, kasbon list from v_kasbon_dashboard, and portal URLs.
 * @param {Object=} filters - { kodeStore, roleFilter, search }
 */
function apiGetDashboardData(filters) {
  filters = filters || {};
  const stores = SupabaseClient.select('kasbon_stores', 'is_active=eq.true&order=kode_store.asc');

  const queryParts = ['order=timestamp_req.desc', 'limit=250'];
  if (filters.kodeStore && filters.kodeStore !== 'ALL') {
    queryParts.push('kode_store=eq.' + encodeURIComponent(filters.kodeStore));
  }
  if (filters.roleFilter && filters.roleFilter !== 'ALL') {
    if (filters.roleFilter === 'COMPLETED') {
      queryParts.push('overall_status=eq.COMPLETED');
    } else if (filters.roleFilter === 'REJECTED') {
      queryParts.push('is_rejected=eq.true');
    } else {
      queryParts.push('next_action_role=eq.' + encodeURIComponent(filters.roleFilter));
    }
  }

  const items = SupabaseClient.select('v_kasbon_dashboard', queryParts.join('&'));

  return {
    status: 'success',
    activeEmail: AuthService.getActiveUserEmail(),
    portalUrls: AppConfig.getPortalUrls(),
    stores: stores || [],
    items: items || []
  };
}

/**
 * Fetches detailed single kasbon record + audit logs for the 7-Step Tracker Modal.
 * @param {string} noPengajuan
 */
function apiGetKasbonDetail(noPengajuan) {
  const cleanId = String(noPengajuan || '').trim().toUpperCase();
  if (!cleanId) {
    throw new Error('Nomor pengajuan wajib diisi.');
  }

  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(cleanId)
  );
  if (!record) {
    throw new Error('Data pengajuan ' + cleanId + ' tidak ditemukan.');
  }

  let auditLogs = [];
  try {
    auditLogs = SupabaseClient.select(
      'kasbon_audit_logs',
      'no_pengajuan=eq.' + encodeURIComponent(cleanId) + '&order=created_at.asc'
    );
  } catch (e) {}

  return {
    status: 'success',
    record: record,
    auditLogs: auditLogs,
    portalUrls: AppConfig.getPortalUrls()
  };
}

/**
 * Saves interconnected Web App URLs into ScriptProperties.
 */
function apiSavePortalUrls(urlsObj) {
  const saved = AppConfig.savePortalUrls(urlsObj || {});
  return {
    status: 'success',
    saved: saved,
    portalUrls: AppConfig.getPortalUrls()
  };
}

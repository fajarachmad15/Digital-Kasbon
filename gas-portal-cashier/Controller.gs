/**
 * ============================================================================
 * MODULE: Controller.gs (PROJECT 4: GAS-PORTAL-CASHIER)
 * PROJECT: Digital Kasbon Petty Cash - Portal Verifikasi Senior Cashier
 * DESCRIPTION: Handles Step 3 (Verifikasi Kasbon) & Step 6 (Verifikasi Realisasi)
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.activeEmail = AuthService.getActiveUserEmail();
  template.queryId = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim().toUpperCase() : '';
  return template.evaluate()
    .setTitle('Portal Senior Cashier (Verifikasi Dana & Realisasi) | E-Kasbon')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches either a specific Kasbon by ID or the list of pending Senior Cashier tasks.
 */
function apiGetCashierTask(noPengajuan) {
  const cleanId = String(noPengajuan || '').trim().toUpperCase();
  const portalUrls = AppConfig.getPortalUrls();

  if (cleanId) {
    const record = SupabaseClient.selectSingle(
      'v_kasbon_dashboard',
      'no_pengajuan=eq.' + encodeURIComponent(cleanId)
    );
    if (!record) {
      throw new Error('Nomor Pengajuan ' + cleanId + ' tidak ditemukan di database.');
    }
    return {
      status: 'success',
      mode: 'DETAIL',
      record: record,
      portalUrls: portalUrls
    };
  }

  const queue = SupabaseClient.select(
    'v_kasbon_dashboard',
    'next_action_role=eq.' + encodeURIComponent('Senior Cashier') + '&order=timestamp_req.desc&limit=100'
  );
  return {
    status: 'success',
    mode: 'QUEUE',
    queue: queue || [],
    portalUrls: portalUrls
  };
}

/**
 * Verifies Senior Cashier NIK & Password for the assigned Kasbon request.
 */
function apiVerifyCashier(params) {
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(String(params.noPengajuan || '').trim().toUpperCase())
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');

  const verifiedUser = AuthService.verifyOfficerLogin({
    nik: params.nik,
    password: params.password,
    expectedRole: 'Senior Cashier',
    expectedStore: record.kode_store,
    assignedNik: record.senior_cashier_nik
  });

  return {
    status: 'success',
    user: verifiedUser
  };
}

/**
 * Submits Step 3 Cashier Verification (Verifikasi Approved / Verifikasi Reject).
 */
function apiSubmitCashierVerif(payload) {
  const noPengajuan = String(payload.noPengajuan || '').trim().toUpperCase();
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan)
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');
  if (record.overall_status !== 'WAITING_CASHIER_VERIF') {
    throw new Error('Pengajuan ini tidak berada pada tahap Verifikasi Cashier (Status: ' + record.status_label + ').');
  }

  const verifiedUser = AuthService.verifyOfficerLogin({
    nik: payload.nik,
    password: payload.password,
    expectedRole: 'Senior Cashier',
    expectedStore: record.kode_store,
    assignedNik: record.senior_cashier_nik
  });

  const decision = payload.decision === 'Verifikasi Approved' ? 'Verifikasi Approved' : 'Verifikasi Reject';
  const reason = decision === 'Verifikasi Reject' ? String(payload.reason || '').trim() : '-';
  if (decision === 'Verifikasi Reject' && !reason) {
    throw new Error('Alasan Reject wajib diisi jika menolak verifikasi.');
  }

  const nowIso = KasbonHelper.getNowWibIso();
  SupabaseClient.update(
    'kasbon_requests',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan),
    {
      timestamp_ver: nowIso,
      nip_ver: verifiedUser.nik,
      status_verifikasi: decision,
      reason_reject_ver: reason
    }
  );

  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 3,
    action_type: decision === 'Verifikasi Approved' ? 'CASHIER_VERIFY_APPROVE' : 'CASHIER_VERIFY_REJECT',
    actor_nik: verifiedUser.nik,
    actor_name: verifiedUser.nama_lengkap,
    actor_email: verifiedUser.email,
    actor_role: 'Senior Cashier',
    previous_status: 'WAITING_CASHIER_VERIF',
    new_status: decision === 'Verifikasi Approved' ? 'WAITING_MONEY_RECEIPT' : 'REJECTED_BY_CASHIER',
    notes: reason
  });

  return {
    status: 'success',
    decision: decision,
    noPengajuan: noPengajuan
  };
}

/**
 * Submits Step 6 Senior Cashier Realization Verification.
 */
function apiSubmitCashierRealizationVerif(payload) {
  const noPengajuan = String(payload.noPengajuan || '').trim().toUpperCase();
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan)
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');
  if (record.overall_status !== 'WAITING_CASHIER_REA_VERIF') {
    throw new Error('Pengajuan ini tidak berada pada tahap Verifikasi Realisasi Cashier.');
  }

  const verifiedUser = AuthService.verifyOfficerLogin({
    nik: payload.nik,
    password: payload.password,
    expectedRole: 'Senior Cashier',
    expectedStore: record.kode_store,
    assignedNik: record.senior_cashier_nik
  });

  const statusPilihan = payload.statusPilihan === 'Tidak Sesuai' ? 'Tidak Sesuai' : 'Ya, Sesuai';
  const reasonVerif = statusPilihan === 'Tidak Sesuai' ? String(payload.reasonVerif || '').trim() : '-';
  if (statusPilihan === 'Tidak Sesuai' && !reasonVerif) {
    throw new Error('Alasan (Reason) wajib diisi jika status realisasi Tidak Sesuai.');
  }

  const uKembali = Math.max(0, Math.floor(Number(payload.uangDikembalikan) || 0));
  const uTerima = Math.max(0, Math.floor(Number(payload.uangDiterima) || 0));
  const txtKembali = KasbonHelper.formatTerbilang(uKembali);
  const txtTerima = KasbonHelper.formatTerbilang(uTerima);
  const nowIso = KasbonHelper.getNowWibIso();

  SupabaseClient.update(
    'kasbon_requests',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan),
    {
      timestamp_ver_rea: nowIso,
      nip_ver_rea: verifiedUser.nik,
      uang_dikembalikan_ver_rea: uKembali,
      terbilang_ver_rea1: txtKembali,
      uang_diterima_ver_rea: uTerima,
      terbilang_ver_rea2: txtTerima,
      status_verifikasi_rea: statusPilihan,
      reason_ver_rea: reasonVerif
    }
  );

  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 6,
    action_type: 'CASHIER_REALIZATION_VERIFY',
    actor_nik: verifiedUser.nik,
    actor_name: verifiedUser.nama_lengkap,
    actor_email: verifiedUser.email,
    actor_role: 'Senior Cashier',
    previous_status: 'WAITING_CASHIER_REA_VERIF',
    new_status: 'WAITING_MGR_FINAL_CHECK',
    notes: statusPilihan + ' | Kembali: Rp ' + uKembali.toLocaleString('id-ID') + ' | Reimburse: Rp ' + uTerima.toLocaleString('id-ID') + ' | Reason: ' + reasonVerif
  });

  return {
    status: 'success',
    noPengajuan: noPengajuan
  };
}

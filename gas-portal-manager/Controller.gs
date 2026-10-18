/**
 * ============================================================================
 * MODULE: Controller.gs (PROJECT 3: GAS-PORTAL-MANAGER)
 * PROJECT: Digital Kasbon Petty Cash - Portal Approval & Final Cek Manager
 * DESCRIPTION: Handles Step 2 (Approval Pengajuan) & Step 7 (Final Cek Realisasi)
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.activeEmail = AuthService.getActiveUserEmail();
  template.queryId = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim().toUpperCase() : '';
  return template.evaluate()
    .setTitle('Portal Manager (Approval & Final Cek) | E-Kasbon')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches either a specific Kasbon by ID or the list of pending Manager tasks.
 * @param {string=} noPengajuan
 */
function apiGetManagerTask(noPengajuan) {
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

  // If no ID specified, return pending queue for Manager (Step 2 and Step 7)
  const queue = SupabaseClient.select(
    'v_kasbon_dashboard',
    'next_action_role=eq.Manager&order=timestamp_req.desc&limit=100'
  );
  return {
    status: 'success',
    mode: 'QUEUE',
    queue: queue || [],
    portalUrls: portalUrls
  };
}

/**
 * Verifies Manager NIK & Password for a specific Kasbon request.
 */
function apiVerifyManager(params) {
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(String(params.noPengajuan || '').trim().toUpperCase())
  );
  if (!record) {
    throw new Error('Data pengajuan tidak ditemukan.');
  }

  const verifiedUser = AuthService.verifyOfficerLogin({
    nik: params.nik,
    password: params.password,
    expectedRole: 'Manager',
    expectedStore: record.kode_store,
    assignedNik: record.manager_incharge_nik
  });

  return {
    status: 'success',
    user: verifiedUser
  };
}

/**
 * Submits Step 2 Manager Approval / Rejection.
 * @param {Object} payload - { noPengajuan, nik, password, decision: 'Approved'|'Reject', reason }
 */
function apiSubmitManagerApproval(payload) {
  const noPengajuan = String(payload.noPengajuan || '').trim().toUpperCase();
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan)
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');
  if (record.overall_status !== 'WAITING_MGR_APPROVAL') {
    throw new Error('Pengajuan ini sudah diproses sebelumnya (Status: ' + record.status_label + ').');
  }

  const verifiedUser = AuthService.verifyOfficerLogin({
    nik: payload.nik,
    password: payload.password,
    expectedRole: 'Manager',
    expectedStore: record.kode_store,
    assignedNik: record.manager_incharge_nik
  });

  const decision = payload.decision === 'Approved' ? 'Approved' : 'Reject';
  const reason = decision === 'Reject' ? String(payload.reason || '').trim() : '-';
  if (decision === 'Reject' && !reason) {
    throw new Error('Alasan Reject wajib diisi jika menolak pengajuan.');
  }

  const nowIso = KasbonHelper.getNowWibIso();
  SupabaseClient.update(
    'kasbon_requests',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan),
    {
      timestamp_app: nowIso,
      nip_app: verifiedUser.nik,
      status_approval: decision,
      reason_reject_app: reason
    }
  );

  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 2,
    action_type: decision === 'Approved' ? 'MANAGER_APPROVE' : 'MANAGER_REJECT',
    actor_nik: verifiedUser.nik,
    actor_name: verifiedUser.nama_lengkap,
    actor_email: verifiedUser.email,
    actor_role: 'Manager',
    previous_status: 'WAITING_MGR_APPROVAL',
    new_status: decision === 'Approved' ? 'WAITING_CASHIER_VERIF' : 'REJECTED_BY_MGR',
    notes: reason
  });

  return {
    status: 'success',
    decision: decision,
    noPengajuan: noPengajuan
  };
}

/**
 * Submits Step 7 Manager Final Check on Realization.
 * @param {Object} payload
 */
function apiSubmitManagerFinalCheck(payload) {
  const noPengajuan = String(payload.noPengajuan || '').trim().toUpperCase();
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan)
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');
  if (record.overall_status !== 'WAITING_MGR_FINAL_CHECK') {
    throw new Error('Pengajuan ini belum berada pada tahap Final Cek Manager atau sudah selesai.');
  }

  const verifiedUser = AuthService.verifyOfficerLogin({
    nik: payload.nik,
    password: payload.password,
    expectedRole: 'Manager',
    expectedStore: record.kode_store,
    assignedNik: record.manager_incharge_nik
  });

  const q1Ans = payload.q1Ans === 'Tidak Sesuai' ? 'Tidak Sesuai' : 'Ya, Sesuai';
  const q1Reason = q1Ans === 'Tidak Sesuai' ? String(payload.q1Reason || '').trim() : '-';
  const q2Ans = payload.q2Ans === 'Tidak Sesuai' ? 'Tidak Sesuai' : 'Ya, Sesuai';
  const q2Reason = q2Ans === 'Tidak Sesuai' ? String(payload.q2Reason || '').trim() : '-';

  if ((q1Ans === 'Tidak Sesuai' && !q1Reason) || (q2Ans === 'Tidak Sesuai' && !q2Reason)) {
    throw new Error('Harap isi alasan (Reason) untuk poin yang dipilih Tidak Sesuai.');
  }

  const needsRevisionFile = (q1Ans === 'Tidak Sesuai' || q2Ans === 'Tidak Sesuai');
  if (needsRevisionFile && (!payload.filePayload || !payload.filePayload.base64)) {
    throw new Error('Karena terdapat ketidaksesuaian, Anda wajib mengunggah Bukti Revisi Realisasi.');
  }

  let linkRevisi = '-';
  if (payload.filePayload && payload.filePayload.base64) {
    linkRevisi = StorageService.uploadAttachment(
      payload.filePayload,
      'revisi',
      noPengajuan
    );
  }

  const nowIso = KasbonHelper.getNowWibIso();
  SupabaseClient.update(
    'kasbon_requests',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan),
    {
      timestamp_fin_cek: nowIso,
      nip_fin_cek: verifiedUser.nik,
      status_qty_value_nota: q1Ans,
      reason_fin_cek_1: q1Reason,
      status_item: q2Ans,
      reason_fin_cek_2: q2Reason,
      revisi_bukti_realisasi: linkRevisi
    }
  );

  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 7,
    action_type: 'MANAGER_FINAL_CHECK',
    actor_nik: verifiedUser.nik,
    actor_name: verifiedUser.nama_lengkap,
    actor_email: verifiedUser.email,
    actor_role: 'Manager',
    previous_status: 'WAITING_MGR_FINAL_CHECK',
    new_status: 'COMPLETED',
    notes: 'Nota: ' + q1Ans + ' (' + q1Reason + ') | Item: ' + q2Ans + ' (' + q2Reason + ')'
  });

  return {
    status: 'success',
    noPengajuan: noPengajuan
  };
}

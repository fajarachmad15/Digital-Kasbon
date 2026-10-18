/**
 * ============================================================================
 * MODULE: Controller.gs (PROJECT 5: GAS-PORTAL-REQUESTER)
 * PROJECT: Digital Kasbon Petty Cash - Portal Konfirmasi & Realisasi Pemohon
 * DESCRIPTION: Handles Step 4 (Konfirmasi Uang Diterima) & Step 5 (Input Realisasi)
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.activeEmail = AuthService.getActiveUserEmail();
  template.queryId = (e && e.parameter && e.parameter.id) ? String(e.parameter.id).trim().toUpperCase() : '';
  return template.evaluate()
    .setTitle('Portal Pemohon (Konfirmasi Uang & Realisasi) | E-Kasbon')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches either a specific Kasbon by ID or the list of pending Requester tasks.
 */
function apiGetRequesterTask(noPengajuan) {
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
    'next_action_role=eq.Requester&order=timestamp_req.desc&limit=100'
  );
  return {
    status: 'success',
    mode: 'QUEUE',
    queue: queue || [],
    portalUrls: portalUrls
  };
}

/**
 * Verifies Requester NIP + 6 first characters of requester email.
 */
function apiVerifyRequester(params) {
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(String(params.noPengajuan || '').trim().toUpperCase())
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');

  const verified = AuthService.verifyRequesterIdentity({
    nipInput: params.nip,
    passInput: params.password,
    expectedNip: record.nip_req,
    expectedEmail: record.email_request
  });

  return {
    status: 'success',
    verified: verified
  };
}

/**
 * Submits Step 4 Confirmation that physical cash has been received by Requester.
 */
function apiConfirmMoneyReceived(payload) {
  const noPengajuan = String(payload.noPengajuan || '').trim().toUpperCase();
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan)
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');
  if (record.overall_status !== 'WAITING_MONEY_RECEIPT') {
    throw new Error('Pengajuan ini tidak berada pada tahap Konfirmasi Penerimaan Uang.');
  }

  AuthService.verifyRequesterIdentity({
    nipInput: payload.nip,
    passInput: payload.password,
    expectedNip: record.nip_req,
    expectedEmail: record.email_request
  });

  const nowIso = KasbonHelper.getNowWibIso();
  const nipClean = AuthService.padNik(record.nip_req);

  SupabaseClient.update(
    'kasbon_requests',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan),
    {
      timestamp_rec: nowIso,
      nip_rec: nipClean,
      status_uang: 'Sudah diterima'
    }
  );

  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 4,
    action_type: 'REQUESTER_CONFIRM_CASH_RECEIVED',
    actor_nik: nipClean,
    actor_name: record.dibayarkan_kepada,
    actor_email: record.email_request,
    actor_role: 'Requester',
    previous_status: 'WAITING_MONEY_RECEIPT',
    new_status: 'WAITING_REALIZATION_INPUT',
    notes: 'Uang fisik kasbon telah diterima sesuai nominal Rp ' + Number(record.nominal).toLocaleString('id-ID')
  });

  return {
    status: 'success',
    noPengajuan: noPengajuan
  };
}

/**
 * Submits Step 5 Realization Report & Uploads Receipt to Supabase Storage.
 */
function apiSubmitRealization(payload) {
  const noPengajuan = String(payload.noPengajuan || '').trim().toUpperCase();
  const record = SupabaseClient.selectSingle(
    'v_kasbon_dashboard',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan)
  );
  if (!record) throw new Error('Data pengajuan tidak ditemukan.');
  if (record.overall_status !== 'WAITING_REALIZATION_INPUT') {
    throw new Error('Pengajuan ini belum berada pada tahap Input Realisasi atau sudah direalisasikan.');
  }

  AuthService.verifyRequesterIdentity({
    nipInput: payload.nip,
    passInput: payload.password,
    expectedNip: record.nip_req,
    expectedEmail: record.email_request
  });

  if (!payload.filePayload || !payload.filePayload.base64) {
    throw new Error('Bukti Lampiran Realisasi (Foto Nota & Item) wajib diunggah.');
  }

  const nominalAwal = Math.floor(Number(record.nominal) || 0);
  const uangDigunakan = Math.max(0, Math.floor(Number(payload.uangDigunakan) || 0));
  const selisih = nominalAwal - uangDigunakan;

  let valKembali = 0;
  let valTerima = 0;
  if (uangDigunakan < nominalAwal) {
    valKembali = selisih;
  } else if (uangDigunakan > nominalAwal) {
    valTerima = Math.abs(selisih);
  }

  const terbilangGuna = KasbonHelper.formatTerbilang(uangDigunakan);
  const txtKembali = KasbonHelper.formatTerbilang(valKembali);
  const txtTerima = KasbonHelper.formatTerbilang(valTerima);

  const linkBuktiRea = StorageService.uploadAttachment(
    payload.filePayload,
    'realisasi',
    noPengajuan
  );

  const nowIso = KasbonHelper.getNowWibIso();
  const nipClean = AuthService.padNik(record.nip_req);

  SupabaseClient.update(
    'kasbon_requests',
    'no_pengajuan=eq.' + encodeURIComponent(noPengajuan),
    {
      timestamp_rea: nowIso,
      nip_rea: nipClean,
      nominal_pembelanjaan_rea: uangDigunakan,
      terbilang_rea1: terbilangGuna,
      uang_dikembalikan_rea: valKembali,
      terbilang_rea2: txtKembali,
      uang_diterima_rea: valTerima,
      terbilang_rea3: txtTerima,
      bukti_lampiran_rea: linkBuktiRea,
      status_realisasi: 'Terrealisasi'
    }
  );

  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 5,
    action_type: 'REQUESTER_SUBMIT_REALIZATION',
    actor_nik: nipClean,
    actor_name: record.dibayarkan_kepada,
    actor_email: record.email_request,
    actor_role: 'Requester',
    previous_status: 'WAITING_REALIZATION_INPUT',
    new_status: 'WAITING_CASHIER_REA_VERIF',
    notes: 'Belanja: Rp ' + uangDigunakan.toLocaleString('id-ID') + ' | Kembali: Rp ' + valKembali.toLocaleString('id-ID') + ' | Reimburse: Rp ' + valTerima.toLocaleString('id-ID')
  });

  return {
    status: 'success',
    noPengajuan: noPengajuan
  };
}

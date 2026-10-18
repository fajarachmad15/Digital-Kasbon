/**
 * ============================================================================
 * MODULE: Controller.gs (PROJECT 2: GAS-FORM-PENGAJUAN)
 * PROJECT: Digital Kasbon Petty Cash - Form Pengajuan Kasbon Baru (Step 1)
 * DESCRIPTION: Backend controller for creating new Kasbon requests
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.activeEmail = AuthService.getActiveUserEmail();
  return template.evaluate()
    .setTitle('Form Pengajuan Kasbon Baru | Petty Cash')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Loads store master and officer lists for the request form.
 */
function apiGetFormInitData() {
  const master = AuthService.getStoresAndOfficers();
  return {
    status: 'success',
    activeEmail: AuthService.getActiveUserEmail(),
    stores: master.stores,
    officers: master.officers,
    portalUrls: AppConfig.getPortalUrls()
  };
}

/**
 * Validates and submits a new Kasbon request to Supabase Database & Storage.
 * @param {Object} formPayload
 */
function apiSubmitPengajuan(formPayload) {
  if (!formPayload) {
    throw new Error('Data form pengajuan kosong.');
  }

  const kodeStore = String(formPayload.kodeStore || '').trim().toUpperCase();
  const emailRequest = String(formPayload.emailRequest || AuthService.getActiveUserEmail() || '').trim().toLowerCase();
  const namaPemohon = String(formPayload.namaPemohon || '').trim();
  const nipPemohon = AuthService.padNik(formPayload.nipPemohon);
  const departemen = String(formPayload.departemen || '').trim();
  const nominal = Math.floor(Number(formPayload.nominal) || 0);
  const keperluan = String(formPayload.keperluan || '').trim();
  const janjiRaw = String(formPayload.janjiPenyelesaian || '').trim(); // YYYY-MM-DD
  const managerIncharge = String(formPayload.managerIncharge || '').trim();
  const seniorCashier = String(formPayload.seniorCashier || '').trim();

  // Corporate Server-Side Validations
  if (!kodeStore) throw new Error('Kode Store wajib dipilih.');
  if (!emailRequest || emailRequest.indexOf('@') === -1) throw new Error('Email Pemohon wajib diisi dengan format yang valid.');
  if (!namaPemohon || /\d/.test(namaPemohon)) throw new Error('Nama Lengkap wajib diisi dan tidak boleh mengandung angka.');
  if (!nipPemohon || !/^\d{6}$/.test(nipPemohon)) throw new Error('NIP wajib terdiri dari 6 digit angka.');
  if (!departemen || departemen === '-') throw new Error('Departemen wajib dipilih.');
  if (nominal <= 0) throw new Error('Nominal kasbon harus lebih besar dari 0.');
  if (keperluan.split(/\s+/).filter(Boolean).length < 2) throw new Error('Keperluan wajib diisi minimal 2 kata.');
  if (!janjiRaw) throw new Error('Tanggal Janji Penyelesaian wajib diisi.');
  if (!managerIncharge || managerIncharge === '-') throw new Error('Manager Incharge wajib dipilih.');
  if (!seniorCashier || seniorCashier === '-') throw new Error('Senior Cashier Incharge wajib dipilih.');

  // Format date from YYYY-MM-DD to DD/MM/YYYY
  let janjiFormatted = janjiRaw;
  if (/^\d{4}-\d{2}-\d{2}$/.test(janjiRaw)) {
    const p = janjiRaw.split('-');
    janjiFormatted = p[2] + '/' + p[1] + '/' + p[0];
  }

  // Generate Sequential Kasbon ID (e.g. KBA644-1026-001)
  const noPengajuan = KasbonHelper.generateNoPengajuan(kodeStore);
  const terbilangStr = KasbonHelper.formatTerbilang(nominal);
  const nowIso = KasbonHelper.getNowWibIso();

  // Upload optional attachment to Supabase Storage ('kasbon-attachments')
  let linkLampiran = '-';
  if (formPayload.filePayload && formPayload.filePayload.base64) {
    linkLampiran = StorageService.uploadAttachment(
      formPayload.filePayload,
      'pengajuan',
      noPengajuan
    );
  }

  // Insert into kasbon_requests (Trigger automatically populates state machine fields)
  const recordToInsert = {
    no_pengajuan: noPengajuan,
    kode_store: kodeStore,
    timestamp_req: nowIso,
    email_request: emailRequest,
    dibayarkan_kepada: namaPemohon,
    nip_req: nipPemohon,
    departemen: departemen,
    nominal: nominal,
    terbilang_req: terbilangStr,
    keperluan: keperluan,
    bukti_lampiran_req: linkLampiran,
    janji_penyelesaian: janjiFormatted,
    senior_cashier: seniorCashier,
    manager_incharge: managerIncharge,
    status_approval: 'Pending',
    status_verifikasi: 'Pending',
    status_uang: 'Pending',
    status_realisasi: 'Pending',
    status_verifikasi_rea: 'Pending'
  };

  SupabaseClient.insert('kasbon_requests', recordToInsert);

  // Record Audit Trail
  SupabaseClient.logAudit({
    no_pengajuan: noPengajuan,
    step_number: 1,
    action_type: 'SUBMIT_PENGAJUAN',
    actor_nik: nipPemohon,
    actor_name: namaPemohon,
    actor_email: emailRequest,
    actor_role: 'Requester',
    previous_status: 'NEW',
    new_status: 'WAITING_MGR_APPROVAL',
    notes: 'Pengajuan kasbon baru sebesar Rp ' + nominal.toLocaleString('id-ID')
  });

  return {
    status: 'success',
    summary: {
      no_pengajuan: noPengajuan,
      tgl_pengajuan: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm'),
      kode_store: kodeStore,
      email_request: emailRequest,
      nama: namaPemohon,
      nip: nipPemohon,
      departemen: departemen,
      nominal: nominal,
      terbilang: terbilangStr,
      keperluan: keperluan,
      janji: janjiFormatted,
      link_lampiran: linkLampiran,
      manager_incharge: managerIncharge,
      senior_cashier: seniorCashier
    },
    portalUrls: AppConfig.getPortalUrls()
  };
}

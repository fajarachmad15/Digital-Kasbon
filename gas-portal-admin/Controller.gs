/**
 * ============================================================================
 * MODULE: Controller.gs (PROJECT 6: GAS-PORTAL-ADMIN)
 * PROJECT: Digital Kasbon Petty Cash - Portal Super Admin (Master Store & Role)
 * DESCRIPTION: Backend controller for managing Master Store, User Roles, & PINs
 * ============================================================================
 */

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  template.activeEmail = AuthService.getActiveUserEmail();
  template.initialTab = (e && e.parameter && e.parameter.tab) ? String(e.parameter.tab).trim() : 'stores';
  return template.evaluate()
    .setTitle('Portal Admin Master Store & Role | E-Kasbon Petty Cash')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Loads initial public portal configuration before login.
 */
function apiGetAdminInitData() {
  return {
    status: 'success',
    activeEmail: AuthService.getActiveUserEmail(),
    portalUrls: AppConfig.getPortalUrls()
  };
}

/**
 * Internal helper to strictly verify Super Admin credentials on every request.
 * @param {Object} auth - { nik, password }
 * @return {Object} Verified admin user record
 */
function verifyAdminSession_(auth) {
  if (!auth || !auth.nik || !auth.password) {
    throw new Error('Sesi Admin tidak valid. Silakan login kembali.');
  }
  return AuthService.verifyOfficerLogin({
    nik: auth.nik,
    password: auth.password,
    expectedRole: 'Admin',
    expectedStore: '',
    assignedNik: ''
  });
}

/**
 * Fetches complete master dataset (stores, users, audit logs) after Admin verification.
 */
function fetchMasterDataset_(adminUser) {
  const stores = SupabaseClient.select(
    'kasbon_stores',
    'order=kode_store.asc'
  );
  const users = SupabaseClient.select(
    'kasbon_users',
    'order=kode_store.asc,role.asc,nama_lengkap.asc'
  );
  let adminLogs = [];
  try {
    adminLogs = SupabaseClient.select(
      'kasbon_admin_logs',
      'order=created_at.desc&limit=100'
    );
  } catch (e) {}

  return {
    status: 'success',
    adminUser: adminUser,
    stores: stores || [],
    users: users || [],
    adminLogs: adminLogs || [],
    portalUrls: AppConfig.getPortalUrls()
  };
}

/**
 * Verifies Admin login and returns all Master Stores, Users, and Admin Audit Logs.
 * @param {Object} auth - { nik, password }
 */
function apiLoginAndLoadMaster(auth) {
  const adminUser = verifyAdminSession_(auth);
  return fetchMasterDataset_(adminUser);
}

/**
 * Creates a new store or updates an existing store's name and status.
 * @param {Object} payload - { auth, mode ('create'|'edit'), kodeStore, namaStore, isActive }
 */
function apiSaveStore(payload) {
  if (!payload) throw new Error('Payload store kosong.');
  const adminUser = verifyAdminSession_(payload.auth);

  const mode = String(payload.mode || 'create').trim().toLowerCase();
  const kodeStore = String(payload.kodeStore || '').trim().toUpperCase();
  const namaStore = String(payload.namaStore || '').trim();
  const isActive = payload.isActive !== false && String(payload.isActive) !== 'false';
  const nowIso = KasbonHelper.getNowWibIso();

  if (!kodeStore || !/^[A-Z0-9_-]{2,16}$/.test(kodeStore)) {
    throw new Error('Kode Store wajib terdiri dari 2-16 karakter huruf/angka (contoh: A300, A644, TK01).');
  }
  if (!namaStore || namaStore.length < 3) {
    throw new Error('Nama Store wajib diisi minimal 3 karakter.');
  }

  const existing = SupabaseClient.selectSingle(
    'kasbon_stores',
    'kode_store=eq.' + encodeURIComponent(kodeStore)
  );

  if (mode === 'create') {
    if (existing) {
      throw new Error('Kode Store ' + kodeStore + ' (' + existing.nama_store + ') sudah terdaftar. Gunakan tombol Edit pada tabel untuk mengubahnya.');
    }
    SupabaseClient.insert('kasbon_stores', {
      kode_store: kodeStore,
      nama_store: namaStore,
      is_active: isActive,
      created_at: nowIso,
      updated_at: nowIso
    });

    SupabaseClient.logAdminAudit({
      action_type: 'CREATE_STORE',
      target_entity: 'STORE',
      target_id: kodeStore,
      actor_nik: adminUser.nik,
      actor_name: adminUser.nama_lengkap,
      actor_email: adminUser.email,
      details: 'Menambahkan unit store baru: ' + kodeStore + ' - ' + namaStore + ' (Status: ' + (isActive ? 'Aktif' : 'Nonaktif') + ')'
    });
  } else {
    if (!existing) {
      throw new Error('Data Store ' + kodeStore + ' tidak ditemukan.');
    }
    SupabaseClient.update(
      'kasbon_stores',
      'kode_store=eq.' + encodeURIComponent(kodeStore),
      {
        nama_store: namaStore,
        is_active: isActive,
        updated_at: nowIso
      }
    );

    SupabaseClient.logAdminAudit({
      action_type: 'UPDATE_STORE',
      target_entity: 'STORE',
      target_id: kodeStore,
      actor_nik: adminUser.nik,
      actor_name: adminUser.nama_lengkap,
      actor_email: adminUser.email,
      details: 'Memperbarui store ' + kodeStore + ': "' + existing.nama_store + '" -> "' + namaStore + '" (Status: ' + (isActive ? 'Aktif' : 'Nonaktif') + ')'
    });
  }

  const res = fetchMasterDataset_(adminUser);
  res.message = mode === 'create'
    ? ('Unit Store ' + kodeStore + ' (' + namaStore + ') berhasil ditambahkan.')
    : ('Unit Store ' + kodeStore + ' (' + namaStore + ') berhasil diperbarui.');
  return res;
}

/**
 * Toggles a store's is_active status.
 * @param {Object} payload - { auth, kodeStore, isActive }
 */
function apiToggleStoreStatus(payload) {
  if (!payload) throw new Error('Payload kosong.');
  const adminUser = verifyAdminSession_(payload.auth);

  const kodeStore = String(payload.kodeStore || '').trim().toUpperCase();
  const nextStatus = Boolean(payload.isActive);
  const nowIso = KasbonHelper.getNowWibIso();

  const existing = SupabaseClient.selectSingle(
    'kasbon_stores',
    'kode_store=eq.' + encodeURIComponent(kodeStore)
  );
  if (!existing) {
    throw new Error('Unit Store ' + kodeStore + ' tidak ditemukan.');
  }

  SupabaseClient.update(
    'kasbon_stores',
    'kode_store=eq.' + encodeURIComponent(kodeStore),
    {
      is_active: nextStatus,
      updated_at: nowIso
    }
  );

  SupabaseClient.logAdminAudit({
    action_type: 'TOGGLE_STORE',
    target_entity: 'STORE',
    target_id: kodeStore,
    actor_nik: adminUser.nik,
    actor_name: adminUser.nama_lengkap,
    actor_email: adminUser.email,
    details: 'Mengubah status store ' + kodeStore + ' (' + existing.nama_store + ') menjadi ' + (nextStatus ? 'AKTIF' : 'NONAKTIF')
  });

  const res = fetchMasterDataset_(adminUser);
  res.message = 'Status Store ' + kodeStore + ' diubah menjadi ' + (nextStatus ? 'Aktif.' : 'Nonaktif.');
  return res;
}

/**
 * Creates or updates a user and their role assignment in kasbon_users.
 * @param {Object} payload - { auth, id, nik, namaLengkap, email, kodeStore, role, passwordPin, isActive }
 */
function apiSaveUser(payload) {
  if (!payload) throw new Error('Payload user kosong.');
  const adminUser = verifyAdminSession_(payload.auth);

  const userId = String(payload.id || '').trim();
  const nik = AuthService.padNik(payload.nik);
  const namaLengkap = String(payload.namaLengkap || '').trim();
  const email = String(payload.email || '').trim().toLowerCase();
  const kodeStore = String(payload.kodeStore || '').trim().toUpperCase();
  const role = String(payload.role || '').trim();
  const passwordPin = String(payload.passwordPin || '').trim();
  const isActive = payload.isActive !== false && String(payload.isActive) !== 'false';
  const nowIso = KasbonHelper.getNowWibIso();

  const validRoles = ['Manager', 'Senior Cashier', 'Requester', 'Admin'];

  if (!nik || !/^\d{6}$/.test(nik)) {
    throw new Error('NIK wajib berupa 6 digit angka (contoh: 121014).');
  }
  if (!namaLengkap || /\d/.test(namaLengkap)) {
    throw new Error('Nama Lengkap wajib diisi dan tidak boleh mengandung angka.');
  }
  if (!email || email.indexOf('@') === -1) {
    throw new Error('Email wajib diisi dengan format yang valid.');
  }
  if (!kodeStore) {
    throw new Error('Kode Store wajib dipilih.');
  }
  if (validRoles.indexOf(role) === -1) {
    throw new Error('Role tidak valid. Pilihan: Manager, Senior Cashier, Requester, atau Admin.');
  }

  // Verify store exists
  const storeObj = SupabaseClient.selectSingle(
    'kasbon_stores',
    'kode_store=eq.' + encodeURIComponent(kodeStore)
  );
  if (!storeObj) {
    throw new Error('Kode Store ' + kodeStore + ' tidak ditemukan di database.');
  }

  // Check for duplicate (nik, kode_store, role)
  const dupCheck = SupabaseClient.select(
    'kasbon_users',
    'nik=eq.' + encodeURIComponent(nik)
    + '&kode_store=eq.' + encodeURIComponent(kodeStore)
    + '&role=eq.' + encodeURIComponent(role)
  );
  if (dupCheck && dupCheck.length > 0) {
    for (let i = 0; i < dupCheck.length; i++) {
      if (String(dupCheck[i].id) !== userId) {
        throw new Error('NIK ' + nik + ' (' + dupCheck[i].nama_lengkap + ') sudah terdaftar dengan Role "' + role + '" di Store ' + kodeStore + '.');
      }
    }
  }

  let updatedOwnPassword = null;

  if (!userId) {
    // CREATE NEW USER ROLE
    const finalPin = passwordPin || '123123';
    if (finalPin.length < 4) {
      throw new Error('PIN / Password minimal 4 karakter.');
    }

    SupabaseClient.insert('kasbon_users', {
      nik: nik,
      nama_lengkap: namaLengkap,
      email: email,
      role: role,
      kode_store: kodeStore,
      password_pin: finalPin,
      is_active: isActive,
      created_at: nowIso,
      updated_at: nowIso
    });

    SupabaseClient.logAdminAudit({
      action_type: 'CREATE_USER',
      target_entity: 'USER',
      target_id: nik + '@' + kodeStore + ':' + role,
      actor_nik: adminUser.nik,
      actor_name: adminUser.nama_lengkap,
      actor_email: adminUser.email,
      details: 'Menambahkan user ' + nik + ' - ' + namaLengkap + ' (' + email + ') sebagai ' + role + ' di Store ' + kodeStore
    });
  } else {
    // UPDATE EXISTING USER ROLE
    const currentRecord = SupabaseClient.selectSingle(
      'kasbon_users',
      'id=eq.' + encodeURIComponent(userId)
    );
    if (!currentRecord) {
      throw new Error('Data user tidak ditemukan.');
    }

    // Prevent locking out the last active Admin
    if (currentRecord.role === 'Admin' && currentRecord.is_active && (role !== 'Admin' || !isActive)) {
      const activeAdmins = SupabaseClient.select(
        'kasbon_users',
        'role=eq.Admin&is_active=eq.true'
      );
      if (!activeAdmins || activeAdmins.length <= 1) {
        throw new Error('Tidak dapat menghapus atau menonaktifkan satu-satunya akun Super Admin yang aktif.');
      }
    }

    const updateData = {
      nik: nik,
      nama_lengkap: namaLengkap,
      email: email,
      role: role,
      kode_store: kodeStore,
      is_active: isActive,
      updated_at: nowIso
    };

    if (passwordPin) {
      if (passwordPin.length < 4) {
        throw new Error('PIN / Password baru minimal 4 karakter.');
      }
      updateData.password_pin = passwordPin;
      if (nik === adminUser.nik && role === 'Admin') {
        updatedOwnPassword = passwordPin;
      }
    }

    SupabaseClient.update(
      'kasbon_users',
      'id=eq.' + encodeURIComponent(userId),
      updateData
    );

    SupabaseClient.logAdminAudit({
      action_type: 'UPDATE_USER',
      target_entity: 'USER',
      target_id: nik + '@' + kodeStore + ':' + role,
      actor_nik: adminUser.nik,
      actor_name: adminUser.nama_lengkap,
      actor_email: adminUser.email,
      details: 'Memperbarui user ' + nik + ' (' + namaLengkap + '): Role ' + currentRecord.role + ' -> ' + role + ', Store ' + currentRecord.kode_store + ' -> ' + kodeStore + (passwordPin ? ' [+Ubah PIN]' : '')
    });
  }

  const res = fetchMasterDataset_(adminUser);
  if (updatedOwnPassword) {
    res.updatedOwnPassword = updatedOwnPassword;
  }
  res.message = !userId
    ? ('User ' + namaLengkap + ' (NIK: ' + nik + ') berhasil ditambahkan sebagai ' + role + ' di Store ' + kodeStore + '.')
    : ('Data & Role untuk ' + namaLengkap + ' (NIK: ' + nik + ') berhasil diperbarui.');
  return res;
}

/**
 * Resets a user's password_pin.
 * @param {Object} payload - { auth, userId, newPin }
 */
function apiResetUserPin(payload) {
  if (!payload) throw new Error('Payload kosong.');
  const adminUser = verifyAdminSession_(payload.auth);

  const userId = String(payload.userId || '').trim();
  const newPin = String(payload.newPin || '123123').trim();
  const nowIso = KasbonHelper.getNowWibIso();

  if (!userId) throw new Error('ID User wajib diisi.');
  if (newPin.length < 4) throw new Error('PIN baru minimal 4 karakter.');

  const currentRecord = SupabaseClient.selectSingle(
    'kasbon_users',
    'id=eq.' + encodeURIComponent(userId)
  );
  if (!currentRecord) {
    throw new Error('Data user tidak ditemukan.');
  }

  SupabaseClient.update(
    'kasbon_users',
    'id=eq.' + encodeURIComponent(userId),
    {
      password_pin: newPin,
      updated_at: nowIso
    }
  );

  SupabaseClient.logAdminAudit({
    action_type: 'RESET_PIN',
    target_entity: 'USER',
    target_id: currentRecord.nik + '@' + currentRecord.kode_store + ':' + currentRecord.role,
    actor_nik: adminUser.nik,
    actor_name: adminUser.nama_lengkap,
    actor_email: adminUser.email,
    details: 'Reset PIN untuk ' + currentRecord.nik + ' - ' + currentRecord.nama_lengkap + ' (' + currentRecord.role + ' - ' + currentRecord.kode_store + ')'
  });

  const res = fetchMasterDataset_(adminUser);
  if (currentRecord.nik === adminUser.nik && currentRecord.role === 'Admin') {
    res.updatedOwnPassword = newPin;
  }
  res.message = 'PIN untuk ' + currentRecord.nama_lengkap + ' (NIK: ' + currentRecord.nik + ') berhasil direset.';
  return res;
}

/**
 * Toggles a user's active status (is_active).
 * @param {Object} payload - { auth, userId, isActive }
 */
function apiToggleUserStatus(payload) {
  if (!payload) throw new Error('Payload kosong.');
  const adminUser = verifyAdminSession_(payload.auth);

  const userId = String(payload.userId || '').trim();
  const nextStatus = Boolean(payload.isActive);
  const nowIso = KasbonHelper.getNowWibIso();

  const currentRecord = SupabaseClient.selectSingle(
    'kasbon_users',
    'id=eq.' + encodeURIComponent(userId)
  );
  if (!currentRecord) {
    throw new Error('Data user tidak ditemukan.');
  }

  if (currentRecord.role === 'Admin' && !nextStatus) {
    const activeAdmins = SupabaseClient.select(
      'kasbon_users',
      'role=eq.Admin&is_active=eq.true'
    );
    if (!activeAdmins || activeAdmins.length <= 1) {
      throw new Error('Tidak dapat menonaktifkan satu-satunya Super Admin yang aktif.');
    }
  }

  SupabaseClient.update(
    'kasbon_users',
    'id=eq.' + encodeURIComponent(userId),
    {
      is_active: nextStatus,
      updated_at: nowIso
    }
  );

  SupabaseClient.logAdminAudit({
    action_type: 'TOGGLE_USER',
    target_entity: 'USER',
    target_id: currentRecord.nik + '@' + currentRecord.kode_store + ':' + currentRecord.role,
    actor_nik: adminUser.nik,
    actor_name: adminUser.nama_lengkap,
    actor_email: adminUser.email,
    details: 'Mengubah status user ' + currentRecord.nik + ' - ' + currentRecord.nama_lengkap + ' (' + currentRecord.role + ') menjadi ' + (nextStatus ? 'AKTIF' : 'NONAKTIF')
  });

  const res = fetchMasterDataset_(adminUser);
  res.message = 'Status user ' + currentRecord.nama_lengkap + ' diubah menjadi ' + (nextStatus ? 'Aktif.' : 'Nonaktif.');
  return res;
}

/**
 * Deletes a user role entry from kasbon_users.
 * @param {Object} payload - { auth, userId }
 */
function apiDeleteUser(payload) {
  if (!payload) throw new Error('Payload kosong.');
  const adminUser = verifyAdminSession_(payload.auth);

  const userId = String(payload.userId || '').trim();
  if (!userId) throw new Error('ID User wajib diisi.');

  const currentRecord = SupabaseClient.selectSingle(
    'kasbon_users',
    'id=eq.' + encodeURIComponent(userId)
  );
  if (!currentRecord) {
    throw new Error('Data user tidak ditemukan.');
  }

  if (currentRecord.nik === adminUser.nik && currentRecord.role === 'Admin') {
    throw new Error('Demi keamanan, Anda tidak dapat menghapus akun Admin yang sedang Anda gunakan login.');
  }

  if (currentRecord.role === 'Admin') {
    const activeAdmins = SupabaseClient.select(
      'kasbon_users',
      'role=eq.Admin&is_active=eq.true'
    );
    if (!activeAdmins || activeAdmins.length <= 1) {
      throw new Error('Tidak dapat menghapus satu-satunya akun Super Admin yang aktif.');
    }
  }

  SupabaseClient.remove(
    'kasbon_users',
    'id=eq.' + encodeURIComponent(userId)
  );

  SupabaseClient.logAdminAudit({
    action_type: 'DELETE_USER',
    target_entity: 'USER',
    target_id: currentRecord.nik + '@' + currentRecord.kode_store + ':' + currentRecord.role,
    actor_nik: adminUser.nik,
    actor_name: adminUser.nama_lengkap,
    actor_email: adminUser.email,
    details: 'Menghapus data user ' + currentRecord.nik + ' - ' + currentRecord.nama_lengkap + ' (Role: ' + currentRecord.role + ', Store: ' + currentRecord.kode_store + ')'
  });

  const res = fetchMasterDataset_(adminUser);
  res.message = 'Data user ' + currentRecord.nama_lengkap + ' (' + currentRecord.role + ' - ' + currentRecord.kode_store + ') berhasil dihapus.';
  return res;
}

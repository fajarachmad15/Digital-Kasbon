/**
 * ============================================================================
 * MODULE: AuthService.gs
 * PROJECT: Digital Kasbon Petty Cash - Enterprise GAS Suite
 * DESCRIPTION: User authentication, role verification, and store/officer lookup
 * ============================================================================
 */

const AuthService = (function () {
  function padNik(nik) {
    const clean = String(nik || '').trim().toUpperCase();
    if (!clean) return '';
    if (/^\d+$/.test(clean) && clean.length < 6) {
      return ('000000' + clean).slice(-6);
    }
    return clean;
  }

  /**
   * Validates corporate NIK format:
   * Exactly 6 characters, starts with 0 or more uppercase letters, followed only by digits.
   * e.g. "E12345" is valid, "121014" is valid, "12e456" is invalid.
   */
  function isValidNik(nik) {
    const clean = padNik(nik);
    return clean.length === 6 && /^[A-Z]*\d+$/.test(clean);
  }

  /**
   * Gets the currently logged-in Google user email (if available in Google Workspace/Session).
   */
  function getActiveUserEmail() {
    try {
      const email = Session.getActiveUser().getEmail();
      return email ? String(email).trim().toLowerCase() : '';
    } catch (e) {
      return '';
    }
  }

  /**
   * Fetches all active stores and officers from Supabase.
   */
  function getStoresAndOfficers() {
    const stores = SupabaseClient.select(
      'kasbon_stores',
      'is_active=eq.true&order=kode_store.asc'
    );
    const officers = SupabaseClient.select(
      'kasbon_users',
      'is_active=eq.true&select=nik,nama_lengkap,email,role,kode_store&order=nama_lengkap.asc'
    );
    return {
      stores: stores || [],
      officers: officers || []
    };
  }

  /**
   * Verifies Manager or Senior Cashier credentials against kasbon_users and assignment.
   * @param {Object} params - { nik, password, expectedRole, expectedStore, assignedNik }
   * @return {Object} Verified user record
   */
  function verifyOfficerLogin(params) {
    const inputNik = padNik(params.nik);
    const inputPass = String(params.password || '').trim();
    const expectedRole = String(params.expectedRole || '').trim();
    const expectedStore = String(params.expectedStore || '').trim().toUpperCase();
    const assignedNik = padNik(params.assignedNik);

    if (!inputNik || !isValidNik(inputNik)) {
      throw new Error('NIK wajib 6 karakter (angka, atau huruf di depan contoh: E12345).');
    }
    if (!inputPass) {
      throw new Error('Password wajib diisi.');
    }
    if (assignedNik && inputNik !== assignedNik) {
      throw new Error('Akses Ditolak: NIK (' + inputNik + ') tidak sesuai dengan petugas yang ditugaskan (' + assignedNik + ').');
    }

    const query = 'nik=eq.' + encodeURIComponent(inputNik)
      + '&role=eq.' + encodeURIComponent(expectedRole)
      + '&is_active=eq.true';

    const users = SupabaseClient.select('kasbon_users', query);
    if (!users || users.length === 0) {
      throw new Error('Data petugas dengan NIK ' + inputNik + ' dan Role ' + expectedRole + ' tidak ditemukan.');
    }

    // Match store and password
    let matchedUser = null;
    for (let i = 0; i < users.length; i++) {
      const u = users[i];
      if (String(u.password_pin) === inputPass) {
        if (!expectedStore || String(u.kode_store).toUpperCase() === expectedStore) {
          matchedUser = u;
          break;
        }
      }
    }

    if (!matchedUser) {
      throw new Error('NIK, Password, atau Kode Store tidak sesuai.');
    }

    return {
      nik: matchedUser.nik,
      nama_lengkap: matchedUser.nama_lengkap,
      email: matchedUser.email,
      role: matchedUser.role,
      kode_store: matchedUser.kode_store
    };
  }

  /**
   * Verifies Requester identity using 6-digit NIP + 6 first characters of requester email.
   * @param {Object} params - { nipInput, passInput, expectedNip, expectedEmail }
   */
  function verifyRequesterIdentity(params) {
    const nipInput = padNik(params.nipInput);
    const passInput = String(params.passInput || '').trim().toLowerCase();
    const expectedNip = padNik(params.expectedNip);
    const expectedEmail = String(params.expectedEmail || '').trim().toLowerCase();
    const expectedPass = expectedEmail.slice(0, 6).toLowerCase();

    if (!nipInput || nipInput !== expectedNip) {
      throw new Error('Validasi Gagal: NIP tidak sesuai dengan data pemohon pada pengajuan ini.');
    }
    if (!passInput || passInput !== expectedPass) {
      throw new Error('Validasi Gagal: Password harus berupa 6 karakter pertama dari email pengajuan (' + expectedEmail + ').');
    }

    return {
      nip: expectedNip,
      email: expectedEmail,
      verified: true
    };
  }

  return {
    padNik: padNik,
    isValidNik: isValidNik,
    getActiveUserEmail: getActiveUserEmail,
    getStoresAndOfficers: getStoresAndOfficers,
    verifyOfficerLogin: verifyOfficerLogin,
    verifyRequesterIdentity: verifyRequesterIdentity
  };
})();

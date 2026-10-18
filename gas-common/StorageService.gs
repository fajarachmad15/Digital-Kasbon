/**
 * ============================================================================
 * MODULE: StorageService.gs
 * PROJECT: Digital Kasbon Petty Cash - Enterprise GAS Suite
 * DESCRIPTION: Attachment upload handler to Google Drive (structured folders:
 *              Digital Kasbon Bukti / YYYY-MM / <No_Pengajuan>) so Supabase
 *              remains 100% text-only and free-tier friendly.
 * ============================================================================
 */

const StorageService = (function () {
  const ALLOWED_MIMES = [
    'image/jpeg',
    'image/jpg',
    'image/png',
    'image/webp',
    'application/pdf'
  ];

  function getExtensionFromMime(mimeType, fallbackName) {
    if (mimeType === 'application/pdf') return 'pdf';
    if (mimeType === 'image/png') return 'png';
    if (mimeType === 'image/webp') return 'webp';
    if (mimeType === 'image/jpeg' || mimeType === 'image/jpg') return 'jpg';
    if (fallbackName && fallbackName.indexOf('.') !== -1) {
      return fallbackName.split('.').pop().toLowerCase();
    }
    return 'jpg';
  }

  function getCategoryPrefix(category) {
    const c = String(category || '').toLowerCase();
    if (c === 'pengajuan') return 'Pengajuan_Bukti';
    if (c === 'realisasi') return 'Realisasi_Nota';
    if (c === 'revisi') return 'Revisi_Final';
    return 'Lampiran';
  }

  /**
   * Finds or creates a child folder by name inside a parent Drive folder.
   */
  function getOrCreateSubfolder_(parentFolder, folderName) {
    const iter = parentFolder.getFoldersByName(folderName);
    if (iter.hasNext()) {
      return iter.next();
    }
    return parentFolder.createFolder(folderName);
  }

  /**
   * Resolves or creates the root Google Drive folder ("Digital Kasbon Bukti")
   * and caches its folder ID in ScriptProperties for fast subsequent access.
   */
  function getOrCreateRootDriveFolder_() {
    const rootFolderName = AppConfig.get('DRIVE_ROOT_FOLDER_NAME') || 'Digital Kasbon Bukti';
    const configuredId = AppConfig.get('DRIVE_ROOT_FOLDER_ID');

    if (configuredId) {
      try {
        const f = DriveApp.getFolderById(configuredId);
        if (f && !f.isTrashed()) {
          return f;
        }
      } catch (e) {
        // Fallback to search/create if cached ID is invalid
      }
    }

    const existingIter = DriveApp.getFoldersByName(rootFolderName);
    let rootFolder = null;
    while (existingIter.hasNext()) {
      const candidate = existingIter.next();
      if (!candidate.isTrashed()) {
        rootFolder = candidate;
        break;
      }
    }

    if (!rootFolder) {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    try {
      PropertiesService.getScriptProperties().setProperty('DRIVE_ROOT_FOLDER_ID', rootFolder.getId());
    } catch (e) {}

    return rootFolder;
  }

  /**
   * Ensures total usage of the "Digital Kasbon Bukti" folder stays strictly below 1 GB (950 MB threshold).
   * Tracks cumulative bytes in ScriptProperties for zero-latency checks on every upload, and automatically
   * rotates/trashes the oldest month subfolder files if total usage approaches 950 MB.
   */
  function enforceDriveQuotaUnder1GB_(rootFolder, incomingBytes) {
    try {
      const maxFolderBytes = Number(AppConfig.get('DRIVE_MAX_FOLDER_BYTES')) || (950 * 1024 * 1024);
      const props = PropertiesService.getScriptProperties();
      let usedBytes = Number(props.getProperty('DRIVE_FOLDER_USED_BYTES') || '0');

      if (usedBytes + incomingBytes > maxFolderBytes) {
        // Scan month folders sorted ascending (oldest YYYY-MM first) and prune oldest files until < 850 MB
        const targetBytes = Math.floor(maxFolderBytes * 0.88);
        const monthFolders = [];
        const mIter = rootFolder.getFolders();
        while (mIter.hasNext()) {
          const mf = mIter.next();
          if (!mf.isTrashed()) monthFolders.push(mf);
        }
        monthFolders.sort(function (a, b) {
          return a.getName().localeCompare(b.getName());
        });

        let recalculatedBytes = 0;
        const allFiles = [];
        for (let i = 0; i < monthFolders.length; i++) {
          const mf = monthFolders[i];
          // Direct files or nested <No_Pengajuan> folders
          const subIter = mf.getFolders();
          while (subIter.hasNext()) {
            const sf = subIter.next();
            if (sf.isTrashed()) continue;
            const fIter = sf.getFiles();
            while (fIter.hasNext()) {
              const f = fIter.next();
              if (!f.isTrashed()) {
                const sz = f.getSize();
                recalculatedBytes += sz;
                allFiles.push({ file: f, created: f.getDateCreated().getTime(), size: sz });
              }
            }
          }
        }

        usedBytes = recalculatedBytes;
        if (usedBytes + incomingBytes > maxFolderBytes) {
          allFiles.sort(function (a, b) { return a.created - b.created; });
          for (let j = 0; j < allFiles.length && (usedBytes + incomingBytes > targetBytes); j++) {
            allFiles[j].file.setTrashed(true);
            usedBytes = Math.max(0, usedBytes - allFiles[j].size);
          }
        }
      }

      props.setProperty('DRIVE_FOLDER_USED_BYTES', String(usedBytes + incomingBytes));
    } catch (quotaErr) {
      console.warn('Quota guard warning:', quotaErr.message);
    }
  }

  /**
   * Uploads file bytes into Google Drive under:
   * Digital Kasbon Bukti / YYYY-MM / <noPengajuan> / <CategoryPrefix>_<noPengajuan>_<timestamp>.<ext>
   */
  function uploadToGoogleDrive_(bytes, mimeType, fileName, noPengajuan) {
    const rootFolder = getOrCreateRootDriveFolder_();
    enforceDriveQuotaUnder1GB_(rootFolder, bytes.length);

    const monthFolderName = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM');
    const monthFolder = getOrCreateSubfolder_(rootFolder, monthFolderName);

    const safeKasbonFolder = String(noPengajuan || 'UMUM').replace(/[^a-zA-Z0-9_-]/g, '');
    const kasbonFolder = getOrCreateSubfolder_(monthFolder, safeKasbonFolder);

    const blob = Utilities.newBlob(bytes, mimeType, fileName);
    const driveFile = kasbonFolder.createFile(blob);

    // Grant view access via link so Manager, Cashier, Requester & Admin can open without permission blocks
    try {
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (shareErr) {
      try {
        driveFile.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (domainErr) {
        console.warn('Drive sharing warning:', domainErr.message);
      }
    }

    return 'https://drive.google.com/file/d/' + driveFile.getId() + '/view?usp=sharing';
  }

  /**
   * Optional backup uploader to Supabase Storage if STORAGE_PROVIDER is explicitly set to 'SUPABASE'.
   */
  function uploadToSupabaseStorage_(bytes, mimeType, category, safeId, timestamp, ext) {
    const objectPath = category + '/' + safeId + '_' + timestamp + '.' + ext;
    const baseUrl = AppConfig.get('SUPABASE_URL').replace(/\/$/, '');
    const bucket = AppConfig.get('STORAGE_BUCKET') || 'kasbon-attachments';
    const apiKey = AppConfig.get('SUPABASE_KEY');
    const uploadUrl = baseUrl + '/storage/v1/object/' + bucket + '/' + objectPath;

    const response = UrlFetchApp.fetch(uploadUrl, {
      method: 'post',
      headers: {
        'apikey': apiKey,
        'Authorization': 'Bearer ' + apiKey,
        'Content-Type': mimeType,
        'x-upsert': 'true'
      },
      payload: bytes,
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code >= 200 && code < 300) {
      return baseUrl + '/storage/v1/object/public/' + bucket + '/' + objectPath;
    }
    throw new Error('Gagal mengunggah lampiran ke Supabase Storage (' + code + '): ' + response.getContentText());
  }

  /**
   * Uploads a base64-encoded file to Google Drive (default) and returns its shareable view URL.
   * @param {Object} fileObj - { base64: string, mimeType: string, fileName: string }
   * @param {string} category - 'pengajuan' | 'realisasi' | 'revisi'
   * @param {string} noPengajuan - e.g. 'KBA644-1026-001'
   * @return {string} Shareable URL of the uploaded file
   */
  function uploadAttachment(fileObj, category, noPengajuan) {
    if (!fileObj || !fileObj.base64) {
      return '-';
    }

    const mimeType = (fileObj.mimeType || 'image/jpeg').toLowerCase();
    if (ALLOWED_MIMES.indexOf(mimeType) === -1) {
      throw new Error('Format file tidak didukung (' + mimeType + '). Gunakan JPG, PNG, WEBP, atau PDF.');
    }

    // Strip data URL prefix if present (e.g., "data:image/jpeg;base64,...")
    const cleanBase64 = fileObj.base64.indexOf(',') !== -1
      ? fileObj.base64.split(',')[1]
      : fileObj.base64;

    const bytes = Utilities.base64Decode(cleanBase64);
    const maxBytes = AppConfig.get('MAX_UPLOAD_BYTES') || (1 * 1024 * 1024);
    if (bytes.length > maxBytes) {
      throw new Error('Ukuran file setelah kompresi (' + Math.round(bytes.length / 1024) + ' KB) melebihi batas maksimum 1 MB.');
    }

    const ext = getExtensionFromMime(mimeType, fileObj.fileName);
    const safeId = String(noPengajuan || 'KASBON').replace(/[^a-zA-Z0-9_-]/g, '');
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
    const prefix = getCategoryPrefix(category);
    const finalFileName = prefix + '_' + safeId + '_' + timestamp + '.' + ext;

    const provider = String(AppConfig.get('STORAGE_PROVIDER') || 'DRIVE').toUpperCase();
    if (provider === 'SUPABASE') {
      return uploadToSupabaseStorage_(bytes, mimeType, category, safeId, timestamp, ext);
    }

    return uploadToGoogleDrive_(bytes, mimeType, finalFileName, safeId);
  }

  return {
    uploadAttachment: uploadAttachment
  };
})();

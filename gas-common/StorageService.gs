/**
 * ============================================================================
 * MODULE: StorageService.gs
 * PROJECT: Digital Kasbon Petty Cash - Enterprise GAS Suite
 * DESCRIPTION: Direct file upload handler to Supabase Storage ('kasbon-attachments')
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

  /**
   * Uploads a base64-encoded file to Supabase Storage and returns its public URL.
   * @param {Object} fileObj - { base64: string, mimeType: string, fileName: string }
   * @param {string} category - 'pengajuan' | 'realisasi' | 'revisi'
   * @param {string} noPengajuan - e.g. 'KBA644-1026-001'
   * @return {string} Public URL of the uploaded file
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
    const maxBytes = AppConfig.get('MAX_UPLOAD_BYTES') || (5 * 1024 * 1024);
    if (bytes.length > maxBytes) {
      throw new Error('Ukuran file melebihi batas maksimum 5 MB.');
    }

    const ext = getExtensionFromMime(mimeType, fileObj.fileName);
    const safeId = String(noPengajuan).replace(/[^a-zA-Z0-9_-]/g, '');
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
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

  return {
    uploadAttachment: uploadAttachment
  };
})();

/**
 * ============================================================================
 * MODULE: KasbonHelper.gs
 * PROJECT: Digital Kasbon Petty Cash - Enterprise GAS Suite
 * DESCRIPTION: Business helpers (Terbilang, ID Generator, Date WIB, HTML Include)
 * ============================================================================
 */

const KasbonHelper = (function () {
  const UNITS = [
    '', 'Satu', 'Dua', 'Tiga', 'Empat', 'Lima',
    'Enam', 'Tujuh', 'Delapan', 'Sembilan', 'Sepuluh', 'Sebelas'
  ];

  function terbilangCore(n) {
    n = Math.floor(Math.abs(Number(n) || 0));
    if (n === 0) return '';
    if (n < 12) return UNITS[n];
    if (n < 20) return (terbilangCore(n - 10) + ' Belas').trim();
    if (n < 100) return (UNITS[Math.floor(n / 10)] + ' Puluh ' + terbilangCore(n % 10)).trim();
    if (n < 200) return ('Seratus ' + terbilangCore(n - 100)).trim();
    if (n < 1000) return (UNITS[Math.floor(n / 100)] + ' Ratus ' + terbilangCore(n % 100)).trim();
    if (n < 2000) return ('Seribu ' + terbilangCore(n - 1000)).trim();
    if (n < 1000000) return (terbilangCore(Math.floor(n / 1000)) + ' Ribu ' + terbilangCore(n % 1000)).trim();
    if (n < 1000000000) return (terbilangCore(Math.floor(n / 1000000)) + ' Juta ' + terbilangCore(n % 1000000)).trim();
    if (n < 1000000000000) return (terbilangCore(Math.floor(n / 1000000000)) + ' Miliar ' + terbilangCore(n % 1000000000)).trim();
    return String(n);
  }

  /**
   * Converts integer amount to Indonesian words with "Rupiah" suffix.
   * Properly returns "Nol Rupiah" when amount is 0.
   */
  function formatTerbilang(amount) {
    const val = Math.floor(Number(amount) || 0);
    if (val <= 0) return 'Nol Rupiah';
    return terbilangCore(val).replace(/\s+/g, ' ').trim() + ' Rupiah';
  }

  /**
   * Generates next sequential Kasbon ID: e.g. KBA644-1026-001
   * @param {string} kodeStore
   * @return {string}
   */
  function generateNoPengajuan(kodeStore) {
    const storeClean = String(kodeStore || '').trim().toUpperCase();
    const mmyy = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'MMyy');
    const prefix = 'KB' + storeClean + '-' + mmyy + '-';

    const existing = SupabaseClient.select(
      'kasbon_requests',
      'select=no_pengajuan&no_pengajuan=like.' + encodeURIComponent(prefix + '*')
    );

    let maxSeq = 0;
    if (existing && existing.length > 0) {
      for (let i = 0; i < existing.length; i++) {
        const idStr = existing[i].no_pengajuan || '';
        const parts = idStr.split('-');
        const seqNum = parseInt(parts[parts.length - 1], 10);
        if (!isNaN(seqNum) && seqNum > maxSeq) {
          maxSeq = seqNum;
        }
      }
    }

    const nextSeq = ('000' + (maxSeq + 1)).slice(-3);
    return prefix + nextSeq;
  }

  /**
   * Returns current timestamp in ISO-8601 WIB (+07:00) format for PostgreSQL TIMESTAMPTZ.
   */
  function getNowWibIso() {
    return Utilities.formatDate(new Date(), 'Asia/Jakarta', "yyyy-MM-dd'T'HH:mm:ss'+07:00'");
  }

  return {
    formatTerbilang: formatTerbilang,
    generateNoPengajuan: generateNoPengajuan,
    getNowWibIso: getNowWibIso
  };
})();

/**
 * Global HTML template include helper for modular .html partials in Apps Script.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

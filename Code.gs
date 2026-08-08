function verifyAdminPin(pin) {
  try {
    // Ambil PIN yang diharapkan — fallback ke ADMIN_PIN constant kalau PropertiesService gagal
    var expectedPin;
    try {
      expectedPin = getAdminPin_();
    } catch (propErr) {
      Logger.log('getAdminPin_ error, fallback ke ADMIN_PIN: ' + propErr);
      expectedPin = String(ADMIN_PIN || '1234');
    }

    // CacheService bisa gagal untuk user anonim di sebagian akun — bungkus try/catch
    var cache = null;
    var attempts = 0;
    try {
      cache = CacheService.getUserCache();
      attempts = Number(cache.get(ADMIN_ATTEMPT_KEY) || 0);
    } catch (cacheErr) {
      Logger.log('CacheService tidak tersedia: ' + cacheErr);
    }

    if (attempts >= 5) {
      return { success: false, message: 'Terlalu banyak percobaan. Coba lagi dalam beberapa menit.' };
    }

    if (!constantTimeEquals_(pin, expectedPin)) {
      if (cache) { try { cache.put(ADMIN_ATTEMPT_KEY, String(attempts + 1), 300); } catch (e) {} }
      Logger.log('Verifikasi PIN gagal (input tidak cocok).');
      return { success: false, message: 'PIN salah. Coba lagi.' };
    }

    var token = Utilities.getUuid();
    if (cache) {
      try {
        cache.remove(ADMIN_ATTEMPT_KEY);
        cache.put(ADMIN_SESSION_PREFIX + token, '1', ADMIN_SESSION_TTL_SECONDS);
      } catch (e) {
        Logger.log('Gagal menyimpan sesi ke cache: ' + e);
      }
    }
    Logger.log('Sesi admin berhasil dibuat.');
    return { success: true, token: token };
  } catch (err) {
    Logger.log('verifyAdminPin error: ' + err);
    return { success: false, message: 'Verifikasi PIN error: ' + (err && err.message ? err.message : String(err)) };
  }
}

/**
 * Endpoint ping ringan — dipakai frontend untuk mengecek apakah Web App
 * yang di-deploy sudah versi terbaru (mengandung fungsi ini).
 * Kalau ini gagal dipanggil dari client, artinya user perlu redeploy.
 */
function pingBackend() {
  return { success: true, version: 'v3-pin-fix', time: new Date().toISOString() };
}

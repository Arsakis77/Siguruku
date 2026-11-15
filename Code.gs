// ==========================================
// BACKEND: GOOGLE APPS SCRIPT (Code.gs)  -  versi 3
// SIGURUKU (SD/MI) + Role (superadmin/admin/user) + Rombel + Tugas Online (Gemini)
// ==========================================
//
// DEPLOY: Deploy > New deployment > Web app
//   Execute as : Me      |  Who has access : Anyone
//
// SETELAH MENIMPA KODE LAMA:
//   1. Pilih fungsi izinkanAkses lalu Run (sekali). Ini memigrasi data lama, tanpa menghapus apa pun.
//   2. Deploy > Manage deployments > Edit > New version > Deploy.
//   (izinkanAkses juga memasang pemicu otomatis: segarkan cache saat sheet diedit tangan + arsip harian jam 02.00.)
//
// BARU DI VERSI 5 (SIGURUKU)
//   - Nama aplikasi: Siguruku. Superadmin bisa mengunggah LOGO di Pengaturan (tampil di halaman masuk, menu, dan ikon tab).
//   - 6 bentuk soal. Pilihan ganda dan pilihan ganda kompleks maksimal 30 soal per bentuk; menjodohkan, benar/salah, isian, uraian maksimal 10 per bentuk.
//   - Pilihan ganda dinilai otomatis. Menjodohkan, benar/salah, isian singkat, dan uraian DINILAI GURU dengan klik Benar/Salah;
//     nilai otomatis dikonversi ke semua soal: tiap soal = 100 / jumlah seluruh soal (maksimal 100).
//   - Nilai maksimal 100: tiap soal bernilai 100 dibagi jumlah soal. Jawaban dicocokkan otomatis dengan kunci;
//     guru bisa menekan tombol Benar/Salah per soal di halaman Hasil (wajib untuk uraian) dan nilai otomatis dihitung ulang.
//   - Kolom baru "Koreksi" di sheet Jawaban dibuat otomatis saat login pertama setelah deploy versi ini.
//
// BARU DI VERSI 4
//   - Soal & buku mendukung Arab dan Inggris (pilih bahasa saat generate, atau ketik manual).
//   - Mapel MI pilihan + input manual, kelas 1-6. Tugas bisa ditulis manual tanpa AI.
//   - Impor/ekspor data siswa lewat Excel. Rilis tugas dapat memakai TOKEN dan DURASI pengerjaan.
//   - Setelah menimpa kode: jalankan izinkanAkses SEKALI (menambah kolom Token & Durasi), lalu New version pada deployment.
//
// BARU DI VERSI 3
//   - Generate soal otomatis membuat BUKU MATERI 5 halaman (rangkuman, mind map, poin penting, glosarium; tanpa soal) di Perpustakaan.
//   - 6 bentuk soal: I pilihan ganda, II pilihan ganda kompleks, III menjodohkan, IV benar-salah, V isian singkat, VI uraian.
//   - Nilai TIDAK tampil langsung ke siswa. Guru merilis nilai per tugas (kolom NilaiRilis di sheet Tugas).
//   - Setelah menimpa kode: jalankan izinkanAkses SEKALI (menambah kolom NilaiRilis), lalu New version pada deployment.
//
// LOGIN AWAL (hanya jika spreadsheet masih kosong): admin / admin123 (superadmin). Segera ganti.
// Akun "admin" lama otomatis menjadi SUPERADMIN saat migrasi.
//
// ROLE
//   superadmin : akses penuh (Pengaturan, tambah admin, kelola siswa & rombel, pemeliharaan data)
//   admin      : Dashboard, Tugas, Perpustakaan
//   user       : siswa (Perpustakaan + mengerjakan Tugas untuk rombelnya)
//
// PRINSIP PERFORMA (agar tetap ringan saat data menumpuk)
//   - Tidak ada getDataRange() pada sheet besar. Hanya kolom yang dibutuhkan yang dibaca.
//   - Pencarian per-ID memakai TextFinder (tidak memuat seluruh sheet).
//   - Daftar materi/tugas di-cache dan dipaginasi; isi materi & soal hanya diambil saat dibuka.
//   - Penghapusan massal memakai penghapusan blok baris, bukan baris per baris.
//   - Arsip: tugas lama yang sudah ditutup bisa dipindah ke spreadsheet arsip (menu Pengaturan > Pemeliharaan).
// ==========================================

const SHEET_NAME   = "BankMateri";
const SHEET_USERS  = "Users";
const SHEET_TUGAS  = "Tugas";
const SHEET_JAWAB  = "Jawaban";
const SHEET_ROMBEL = "Rombel";

const ROLE_SUPER = "superadmin";
const ROLE_ADMIN = "admin";
const ROLE_USER  = "user";
const STAFF = [ROLE_SUPER, ROLE_ADMIN];
const SUPER = [ROLE_SUPER];
const STUDENT = [ROLE_USER];

const SESSION_TTL   = 21600;                // 6 jam (maksimum CacheService)
const DEFAULT_MODEL = "gemini-3.5-flash";   // bisa diganti di Pengaturan
const PAGE_SIZE     = 24;
const MAX_CELL      = 49000;                // batas aman 50.000 karakter per sel
const SCHEMA_VERSION = "5";
// Mapel baku (nilai harus sama dengan daftar MAPEL di Index.html). Mapel di luar daftar = "Lainnya" (input manual).
const MAPEL_BAKU = ["Matematika", "B.Indonesia", "IPAS", "PKN", "PJOK", "SBdP", "Bahasa Inggris", "Bahasa Jawa", "Al-Qur'an Hadis", "Akidah Akhlak", "Fikih", "SKI", "Bahasa Arab"];

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Siguruku')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// IZIN & SETUP  (jalankan SEKALI dari editor)
// ==========================================
function izinkanAkses() {
  UrlFetchApp.fetch("https://generativelanguage.googleapis.com/", { muteHttpExceptions: true });
  ss_();
  migrasi_();
  pasangTriggerOtomatis();
  Logger.log("Izin dan migrasi selesai. Sekarang buat New version pada deployment Web app.");
}

// Hanya perlu jika skrip MANDIRI (bukan dari Extensions > Apps Script di Google Sheets).
function setSpreadsheetId() {
  const ID = "ISI_ID_SPREADSHEET_DI_SINI";
  if (ID.indexOf("ISI_ID") === 0) throw new Error("Isi dulu ID spreadsheet pada variabel ID.");
  PropertiesService.getScriptProperties().setProperty("SPREADSHEET_ID", ID);
  Logger.log("Tersimpan: " + ID);
}

function ss_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID");
  if (!id) throw new Error("Spreadsheet tidak ditemukan. Buat skrip dari Extensions > Apps Script di Google Sheets, atau jalankan setSpreadsheetId().");
  return SpreadsheetApp.openById(id);
}

// ==========================================
// UTIL UMUM
// ==========================================
function ok_(o) { o = o || {}; o.success = true; return o; }
function fail_(e, extra) {
  const m = String((e && e.message) || e);
  const o = { success: false, message: /Could not obtain lock/i.test(m) ? "Server sedang sibuk. Coba lagi sebentar." : m };
  if (extra) Object.keys(extra).forEach(k => o[k] = extra[k]);
  return o;
}
function withLock_(fn, waitMs) {
  const lock = LockService.getScriptLock();
  lock.waitLock(waitMs || 30000);
  try { return fn(); } finally { try { lock.releaseLock(); } catch (x) {} }
}
function toIso_(v) { return v instanceof Date ? v.toISOString() : String(v == null ? "" : v); }
function tingkat_(kelas) { const m = String(kelas || "").match(/^(\d{1,2})/); return m ? m[1] : ""; }
function int_(v, d) { const n = parseInt(v, 10); return isNaN(n) ? d : n; }

// kelas target (tugas) cocok dengan kelas siswa? target "5" = semua rombel 5, "5A" = rombel 5A saja.
function cocokKelas_(target, kelasSiswa) {
  target = String(target).trim(); kelasSiswa = String(kelasSiswa || "").trim();
  if (!kelasSiswa) return false;   // siswa tanpa rombel tidak boleh melihat apa pun
  return target === kelasSiswa || target === tingkat_(kelasSiswa);
}
// filter admin: filter "5A" cocok dengan tugas "5A" dan tugas "5"; filter "5" cocok dengan semua yang bertingkat 5
function cocokFilter_(kelasData, filter) {
  if (!filter || filter === "all") return true;
  kelasData = String(kelasData);
  if (kelasData === filter) return true;
  if (/^\d+$/.test(filter)) return tingkat_(kelasData) === filter;
  return kelasData === tingkat_(filter);
}

// filter untuk baris siswa/jawaban: "5A" = hanya rombel 5A, "5" = semua rombel kelas 5
function cocokKelasSiswa_(kelasSiswa, filter) {
  if (!filter || filter === "all") return true;
  kelasSiswa = String(kelasSiswa);
  return kelasSiswa === filter || (/^\d+$/.test(filter) && tingkat_(kelasSiswa) === filter);
}

// Baca kolom c1..c2 (1-based) untuk semua baris data (baris 2 sampai terakhir).
function rd_(sh, c1, c2) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, c1, last - 1, c2 - c1 + 1).getValues();
}

// Cari baris yang sel di kolom `col` sama persis dengan value. Tidak memuat seluruh sheet.
function findRow_(sh, col, value, matchCase) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const r = sh.getRange(2, col, last - 1, 1).createTextFinder(String(value))
    .matchEntireCell(true).matchCase(!!matchCase).findNext();
  return r ? r.getRow() : 0;
}

// Semua nomor baris (1-based, terurut) yang sel di kolom `col` sama persis dengan value. Hanya memindai satu kolom.
function findAllRows_(sh, col, value) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  const found = sh.getRange(2, col, last - 1, 1).createTextFinder(String(value)).matchEntireCell(true).matchCase(true).findAll();
  return found.map(r => r.getRow()).sort((a, b) => a - b);
}

// Baca kolom c1..c2 hanya untuk baris tertentu (terurut). Baris yang berdekatan (selisih <= gap) dibaca satu blok.
// Hasil: { nomorBaris: [nilai kolom c1..c2] }
function readRows_(sh, rows, c1, c2, gap) {
  const out = {}; gap = gap || 50;
  let i = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] - rows[j] <= gap) j++;
    const lo = rows[i], hi = rows[j];
    const vals = sh.getRange(lo, c1, hi - lo + 1, c2 - c1 + 1).getValues();
    for (let k = i; k <= j; k++) out[rows[k]] = vals[rows[k] - lo];
    i = j + 1;
  }
  return out;
}

// Hapus baris (nomor baris 1-based) dari bawah ke atas dalam blok berurutan.
function deleteRowsBottomUp_(sh, rowNums, t0) {
  const rows = rowNums.slice().sort((a, b) => b - a);
  let i = 0, deleted = 0;
  while (i < rows.length) {
    let j = i;
    while (j + 1 < rows.length && rows[j + 1] === rows[j] - 1) j++;
    const count = j - i + 1;
    sh.deleteRows(rows[j], count);
    deleted += count; i = j + 1;
    if (t0 && Date.now() - t0 > 240000) break;
  }
  return deleted;
}

// Tambah banyak baris sekaligus dengan satu setValues
function appendRows_(sh, rows) {
  if (!rows.length) return;
  const cols = rows[0].length;
  const start = sh.getLastRow() + 1;
  const need = start + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  if (sh.getMaxColumns() < cols) sh.insertColumnsAfter(sh.getMaxColumns(), cols - sh.getMaxColumns());
  sh.getRange(start, 1, rows.length, cols).setValues(rows);
}

// ---- Cache JSON bertahap (chunk) ----
function cacheGetJson_(key) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(key + "#n") || 0);
  if (!n) return null;
  const keys = []; for (let i = 0; i < n; i++) keys.push(key + "#" + i);
  const all = c.getAll(keys);
  let s = "";
  for (let i = 0; i < n; i++) { const p = all[key + "#" + i]; if (p == null) return null; s += p; }
  try { return JSON.parse(s); } catch (e) { return null; }
}
function cachePutJson_(key, obj, ttl) {
  try {
    const s = JSON.stringify(obj), SZ = 30000, n = Math.ceil(s.length / SZ) || 1;
    if (n > 20) return false;
    const put = {}; put[key + "#n"] = String(n);
    for (let i = 0; i < n; i++) put[key + "#" + i] = s.substr(i * SZ, SZ);
    CacheService.getScriptCache().putAll(put, ttl || 600);
    return true;
  } catch (e) { return false; }
}
function cacheRemove_(key) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(key + "#n") || 0);
  const keys = [key + "#n"]; for (let i = 0; i < n; i++) keys.push(key + "#" + i);
  c.removeAll(keys);
}
// Versi data (disimpan di Properties agar tidak hilang bila cache dibersihkan)
function getVer_(name) { return PropertiesService.getScriptProperties().getProperty("v_" + name) || "0"; }
function bumpVer_(name) { PropertiesService.getScriptProperties().setProperty("v_" + name, String(Date.now())); }

// ==========================================
// SHEET & SKEMA
// ==========================================
function getSheet_(name, headers) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#d0e0e3");
    sh.setFrozenRows(1);
  }
  return sh;
}
function initSheet()   { return getSheet_(SHEET_NAME, ["ID", "Tanggal", "Kelas", "Semester", "MataPelajaran", "Bab", "Judul", "Konten"]); }
function tugasSheet_() { const sh = tugasSheet0_(); if (sh.getMaxColumns() < 15) sh.insertColumnsAfter(sh.getMaxColumns(), 15 - sh.getMaxColumns()); return sh; }
function tugasSheet0_() { return getSheet_(SHEET_TUGAS, ["ID", "Tanggal", "Kelas", "Mapel", "Judul", "MateriId", "Instruksi", "Soal", "Deadline", "Status", "JumlahSoal", "Terkumpul", "NilaiRilis", "Token", "Durasi"]); }
function jawabSheet_() {
  const sh = getSheet_(SHEET_JAWAB, ["ID", "TugasId", "Username", "Nama", "Kelas", "Waktu", "Jawaban", "Skor", "Benar", "TotalPG", "Koreksi"]);
  if (sh.getMaxColumns() < 11) sh.insertColumnsAfter(sh.getMaxColumns(), 11 - sh.getMaxColumns());
  return sh;
}
function usersSheet_() {
  const existed = !!ss_().getSheetByName(SHEET_USERS);
  const sh = getSheet_(SHEET_USERS, ["Username", "PasswordHash", "Nama", "Role", "Kelas", "Dibuat", "WajibGanti"]);
  if (!existed) sh.appendRow(["admin", hashPassword_("admin123"), "Superadmin", ROLE_SUPER, "", new Date().toISOString(), "TRUE"]);
  return sh;
}
function rombelSheet_() {
  const existed = !!ss_().getSheetByName(SHEET_ROMBEL);
  const sh = getSheet_(SHEET_ROMBEL, ["Kode", "Tingkat", "Rombel", "Dibuat"]);
  if (!existed) {
    const now = new Date().toISOString(), rows = [];
    for (let t = 1; t <= 6; t++) ["A", "B", "C"].forEach(r => rows.push([t + r, t, r, now]));
    appendRows_(sh, rows);
  }
  return sh;
}

// Migrasi satu kali dari versi lama (aman dijalankan ulang)
function migrasi_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty("SCHEMA_V") === SCHEMA_VERSION) return;
  withLock_(function () {
    if (p.getProperty("SCHEMA_V") === SCHEMA_VERSION) return;
    initSheet(); jawabSheet_(); rombelSheet_();
    const us = usersSheet_(), ts = tugasSheet_(), js = jawabSheet_();

    js.getRange(1, 11).setValue("Koreksi").setFontWeight("bold").setBackground("#d0e0e3");
    if (ts.getMaxColumns() < 15) ts.insertColumnsAfter(ts.getMaxColumns(), 15 - ts.getMaxColumns());
    ts.getRange(1, 11, 1, 5).setValues([["JumlahSoal", "Terkumpul", "NilaiRilis", "Token", "Durasi"]]).setFontWeight("bold").setBackground("#d0e0e3");
    const last = ts.getLastRow();
    if (last >= 2) {
      const ids = ts.getRange(2, 1, last - 1, 1).getValues();
      const soal = ts.getRange(2, 8, last - 1, 1).getValues();
      const cnt = {};
      rd_(js, 2, 2).forEach(r => { cnt[r[0]] = (cnt[r[0]] || 0) + 1; });
      const out = soal.map((r, i) => { let n = 0; try { n = JSON.parse(r[0]).length; } catch (e) {} return [n, cnt[ids[i][0]] || 0]; });
      ts.getRange(2, 11, out.length, 2).setValues(out);
      // NilaiRilis: tugas lama yang sudah dikumpulkan siswa dianggap sudah dirilis (siswa sudah pernah melihat skornya); sisanya disembunyikan.
      const cur = ts.getRange(2, 13, last - 1, 1).getValues();
      const rl = cur.map((r, i) => [String(r[0]) !== "" ? r[0] : (out[i][1] > 0 ? "TRUE" : "FALSE")]);
      ts.getRange(2, 13, rl.length, 1).setValues(rl);
    }

    const u = rd_(us, 1, 7);
    let hasSuper = false, promote = -1;
    u.forEach((r, i) => {
      if (String(r[3]) === ROLE_SUPER) hasSuper = true;
      if (promote < 0 && String(r[3]) === ROLE_ADMIN && String(r[0]).toLowerCase() === "admin") promote = i;
    });
    if (!hasSuper) {
      if (promote < 0) promote = u.findIndex(r => String(r[3]) === ROLE_ADMIN);
      if (promote >= 0) us.getRange(promote + 2, 4).setValue(ROLE_SUPER);
    }
    ["materi", "tugas", "rombel"].forEach(bumpVer_);
    p.setProperty("SCHEMA_V", SCHEMA_VERSION);
  }, 60000);
}

// ==========================================
// AUTENTIKASI & SESI
// ==========================================
function hashPassword_(plain, salt) {
  salt = salt || Utilities.getUuid().replace(/-/g, '').substring(0, 16);
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + plain, Utilities.Charset.UTF_8);
  return salt + "$" + Utilities.base64Encode(digest);
}
function verifyPassword_(plain, stored) {
  if (!stored || String(stored).indexOf("$") < 0) return false;
  const salt = String(stored).split("$")[0];
  return hashPassword_(plain, salt) === String(stored);
}

function login(username, password) {
  try {
    migrasi_();
    username = String(username || "").trim().toLowerCase();
    password = String(password || "");
    if (!username || !password) return { success: false, message: "Username dan password wajib diisi." };

    const cache = CacheService.getScriptCache();
    const failKey = "fail_" + username;
    const fails = Number(cache.get(failKey) || 0);
    if (fails >= 5) return { success: false, message: "Terlalu banyak percobaan gagal. Coba lagi dalam 5 menit." };

    const sh = usersSheet_();
    const row = findRow_(sh, 1, username, false);
    if (row) {
      const r = sh.getRange(row, 1, 1, 7).getValues()[0];
      if (String(r[0]).toLowerCase() === username && verifyPassword_(password, r[1])) {
        const user = {
          username: String(r[0]), nama: String(r[2]), role: String(r[3]), kelas: String(r[4] || ""),
          wajibGanti: String(r[6]).toUpperCase() === "TRUE", iat: Date.now()
        };
        const token = Utilities.getUuid() + Utilities.getUuid();
        cache.put("sess_" + token, JSON.stringify(user), SESSION_TTL);
        cache.remove(failKey);
        return { success: true, token: token, user: user };
      }
    }
    cache.put(failKey, String(fails + 1), 300);
    return { success: false, message: "Username atau password salah." };
  } catch (e) {
    return fail_(e);
  }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove("sess_" + token);
  return { success: true };
}

function getSession(token) {
  const s = auth_(token, null, true);
  return s ? { success: true, user: s } : { success: false };
}

// roles: array role yang diizinkan, atau null (siapa saja yang login)
function auth_(token, roles, soft, allowForce) {
  const cache = CacheService.getScriptCache();
  const raw = token ? cache.get("sess_" + token) : null;
  let user = null;
  if (raw) {
    user = JSON.parse(raw);
    const rev = cache.get("rev_" + user.username);
    if (rev && Number(rev) > Number(user.iat || 0)) { cache.remove("sess_" + token); user = null; }
  }
  if (!user) { if (soft) return null; throw new Error("Sesi berakhir. Silakan login kembali."); }
  if (user.wajibGanti && !soft && !allowForce) throw new Error("Anda harus mengganti password terlebih dahulu.");
  if (roles && roles.indexOf(user.role) < 0) { if (soft) return null; throw new Error("Anda tidak memiliki akses untuk tindakan ini."); }
  return user;
}
// paksa semua sesi milik username agar login ulang (dipakai saat hapus akun / pindah kelas)
function revoke_(username) { CacheService.getScriptCache().put("rev_" + username, String(Date.now()), SESSION_TTL); }

function changePassword(token, oldPass, newPass) {
  try {
    const me = auth_(token, null, false, true);
    if (!newPass || String(newPass).length < 6) return { success: false, message: "Password baru minimal 6 karakter." };
    const sh = usersSheet_();
    const row = findRow_(sh, 1, me.username, true);
    if (!row) return { success: false, message: "Pengguna tidak ditemukan." };
    if (!verifyPassword_(oldPass, sh.getRange(row, 2).getValue())) return { success: false, message: "Password lama salah." };
    sh.getRange(row, 2).setValue(hashPassword_(newPass));
    sh.getRange(row, 7).setValue("FALSE");
    me.wajibGanti = false;
    CacheService.getScriptCache().put("sess_" + token, JSON.stringify(me), SESSION_TTL);
    return { success: true, message: "Password berhasil diganti." };
  } catch (e) { return fail_(e); }
}

// ==========================================
// ROMBEL (kelas 1A, 1B, 1C, ...)
// ==========================================
function rombel_() {
  const key = "rb_" + getVer_("rombel");
  let hit = cacheGetJson_(key);
  if (hit) return hit;
  const out = rd_(rombelSheet_(), 1, 3).map(r => ({ kode: String(r[0]), tingkat: String(r[1]), rombel: String(r[2]) }))
    .sort((a, b) => a.kode.localeCompare(b.kode, "id", { numeric: true }));
  cachePutJson_(key, out, 3600);
  return out;
}
function rombelSet_() { const s = {}; rombel_().forEach(r => s[r.kode] = true); return s; }

function getRombel(token) {
  try { auth_(token); return ok_({ data: rombel_() }); } catch (e) { return fail_(e, { data: [] }); }
}

// input: "1A, 1B, 1C" atau satu per baris. Hanya format <tingkat 1-6><huruf>
function addRombel(token, kodeList) {
  try {
    auth_(token, SUPER);
    const have = rombelSet_();
    const kodes = String(kodeList || "").toUpperCase().split(/[\s,;]+/).filter(Boolean);
    const rows = [], skipped = [], now = new Date().toISOString();
    kodes.forEach(k => {
      if (!/^[1-6][A-Z]$/.test(k)) { skipped.push(k + " (format salah)"); return; }
      if (have[k]) { skipped.push(k + " (sudah ada)"); return; }
      have[k] = true; rows.push([k, Number(k.charAt(0)), k.charAt(1), now]);
    });
    if (rows.length) withLock_(function () { appendRows_(rombelSheet_(), rows); bumpVer_("rombel"); });
    return ok_({ message: rows.length + " rombel ditambahkan." + (skipped.length ? " Dilewati: " + skipped.join(", ") : "") });
  } catch (e) { return fail_(e); }
}

function deleteRombel(token, kode) {
  try {
    auth_(token, SUPER);
    kode = String(kode || "").toUpperCase();
    const pakai = rd_(usersSheet_(), 4, 5).filter(r => String(r[0]) === ROLE_USER && String(r[1]).toUpperCase() === kode).length;
    if (pakai) return { success: false, message: "Rombel " + kode + " masih dipakai " + pakai + " siswa. Pindahkan siswanya dulu." };
    return withLock_(function () {
      const sh = rombelSheet_(), row = findRow_(sh, 1, kode, false);
      if (!row) return { success: false, message: "Rombel tidak ditemukan." };
      sh.deleteRow(row); bumpVer_("rombel");
      return ok_({ message: "Rombel " + kode + " dihapus." });
    });
  } catch (e) { return fail_(e); }
}

// ==========================================
// MANAJEMEN PENGGUNA (SUPERADMIN)
// ==========================================
function listUsers(token, opt) {
  try {
    auth_(token, SUPER);
    opt = opt || {};
    const role = String(opt.role || "all"), kelas = String(opt.kelas || "all"), q = String(opt.q || "").toLowerCase().trim();
    const size = Math.min(100, int_(opt.size, 50)), page = Math.max(1, int_(opt.page, 1));
    const order = { superadmin: 0, admin: 1, user: 2 };
    let list = rd_(usersSheet_(), 1, 5).map(r => ({ username: String(r[0]), nama: String(r[2]), role: String(r[3]), kelas: String(r[4] || "") }))
      .filter(u => (role === "all" || u.role === role) && (kelas === "all" || u.kelas === kelas || (/^\d+$/.test(kelas) && tingkat_(u.kelas) === kelas)) &&
                   (!q || u.nama.toLowerCase().indexOf(q) >= 0 || u.username.toLowerCase().indexOf(q) >= 0));
    list.sort((a, b) => (order[a.role] - order[b.role]) || a.kelas.localeCompare(b.kelas, "id", { numeric: true }) || a.nama.localeCompare(b.nama, "id"));
    return ok_({ data: list.slice((page - 1) * size, page * size), total: list.length, page: page, hasMore: page * size < list.length });
  } catch (e) { return fail_(e, { data: [] }); }
}

// entries: [{nama, username, password, kelas}], role: 'user' (default) | 'admin'
function addUsers(token, entries, role) {
  try {
    auth_(token, SUPER);
    role = role === ROLE_ADMIN ? ROLE_ADMIN : ROLE_USER;
    entries = (entries || []).slice(0, 300);
    const valid = rombelSet_();
    return withLock_(function () {
      const sh = usersSheet_();
      const existing = {};
      rd_(sh, 1, 1).forEach(r => existing[String(r[0]).toLowerCase()] = true);
      const rows = [], skipped = [], hasil = [], now = new Date().toISOString();
      entries.forEach(en => {
        const nama = String(en.nama || "").trim();
        const uname = String(en.username || "").trim().toLowerCase();
        const pass = String(en.password || "");
        const kelas = String(en.kelas || "").trim().toUpperCase();
        const minPass = role === ROLE_ADMIN ? 6 : 4;
        const skip = alasan => { skipped.push((uname || nama || "(kosong)") + " (" + alasan + ")"); hasil.push({ username: uname, status: "Dilewati", alasan: alasan }); };
        if (!nama) return skip("nama kosong");
        if (!uname || !/^[a-z0-9._-]+$/.test(uname)) return skip("username kosong/tidak valid");
        if (pass.length < minPass) return skip("password minimal " + minPass + " karakter");
        if (role === ROLE_USER && !valid[kelas]) return skip("rombel tidak ada");
        if (existing[uname]) return skip("username sudah ada");
        existing[uname] = true;
        rows.push([uname, hashPassword_(pass), nama, role, role === ROLE_USER ? kelas : "", now, role === ROLE_ADMIN ? "TRUE" : "FALSE"]);
        hasil.push({ username: uname, status: "Berhasil", alasan: "" });
      });
      appendRows_(sh, rows);
      cacheRemove_("stats");
      return ok_({ hasil: hasil, message: rows.length + (role === ROLE_ADMIN ? " admin" : " siswa") + " ditambahkan." + (skipped.length ? " Dilewati: " + skipped.join(", ") : "") });
    });
  } catch (e) { return fail_(e); }
}

function resetUserPassword(token, username, newPass) {
  try {
    auth_(token, SUPER);
    if (!newPass || String(newPass).length < 4) return { success: false, message: "Password minimal 4 karakter." };
    return withLock_(function () {
      const sh = usersSheet_(), row = findRow_(sh, 1, username, true);
      if (!row) return { success: false, message: "Pengguna tidak ditemukan." };
      const role = String(sh.getRange(row, 4).getValue());
      if (role !== ROLE_USER && String(newPass).length < 6) return { success: false, message: "Password admin minimal 6 karakter." };
      sh.getRange(row, 2).setValue(hashPassword_(newPass));
      sh.getRange(row, 7).setValue(role === ROLE_USER ? "FALSE" : "TRUE");
      revoke_(username);
      return ok_({ message: "Password " + username + " direset." });
    });
  } catch (e) { return fail_(e); }
}

function updateUserKelas(token, username, kelas) {
  try {
    auth_(token, SUPER);
    kelas = String(kelas || "").trim().toUpperCase();
    if (!rombelSet_()[kelas]) return { success: false, message: "Rombel " + kelas + " tidak ada." };
    return withLock_(function () {
      const sh = usersSheet_(), row = findRow_(sh, 1, username, true);
      if (!row) return { success: false, message: "Pengguna tidak ditemukan." };
      if (String(sh.getRange(row, 4).getValue()) !== ROLE_USER) return { success: false, message: "Hanya siswa yang punya rombel." };
      sh.getRange(row, 5).setValue(kelas);
      revoke_(username);
      return ok_({ message: username + " dipindah ke " + kelas + "." });
    });
  } catch (e) { return fail_(e); }
}

function deleteUser(token, username) {
  try {
    const me = auth_(token, SUPER);
    if (username === me.username) return { success: false, message: "Tidak bisa menghapus akun yang sedang dipakai." };
    return withLock_(function () {
      const sh = usersSheet_(), row = findRow_(sh, 1, username, true);
      if (!row) return { success: false, message: "Pengguna tidak ditemukan." };
      if (String(sh.getRange(row, 4).getValue()) === ROLE_SUPER) return { success: false, message: "Akun superadmin tidak bisa dihapus." };
      sh.deleteRow(row); revoke_(username); cacheRemove_("stats");
      return ok_({ message: "Pengguna dihapus." });
    });
  } catch (e) { return fail_(e); }
}

// ==========================================
// MATERI
// ==========================================
// Metadata materi (tanpa isi) di-cache. Terbaru di depan.
function materiMeta_() {
  const key = "mm_" + getVer_("materi");
  const hit = cacheGetJson_(key);
  if (hit) return hit;
  const out = rd_(initSheet(), 1, 7).filter(r => r[0] !== "" && r[6] !== "").map(r => ({
    id: String(r[0]), tanggal: toIso_(r[1]), kelas: String(r[2]).trim(), semester: String(r[3]),
    mapel: String(r[4]), bab: String(r[5]), judul: String(r[6])
  })).reverse();
  cachePutJson_(key, out, 1800);
  return out;
}

// opt: {kelas:'all'|'1'..'6', mapel:'all'|..., q, page, size}
function getMateriList(token, opt) {
  try {
    const me = auth_(token);
    opt = opt || {};
    let kelas;
    if (me.role === ROLE_USER) {
      kelas = tingkat_(me.kelas);   // siswa hanya melihat buku kelasnya sendiri, apa pun yang dikirim browser
      if (!kelas) return ok_({ data: [], total: 0, page: 1, hasMore: false });
    } else {
      kelas = String(opt.kelas || "all");
    }
    const mapel = String(opt.mapel || "all"), q = String(opt.q || "").toLowerCase().trim();
    const size = Math.min(60, int_(opt.size, PAGE_SIZE)), page = Math.max(1, int_(opt.page, 1));
    const list = materiMeta_().filter(d =>
      (kelas === "all" || d.kelas === kelas) &&
      (mapel === "all" || (mapel === "Lainnya" ? MAPEL_BAKU.indexOf(d.mapel) < 0 : d.mapel === mapel)) &&
      (!q || d.judul.toLowerCase().indexOf(q) >= 0 || d.bab.toLowerCase().indexOf(q) >= 0));
    return ok_({ data: list.slice((page - 1) * size, page * size), total: list.length, page: page, hasMore: page * size < list.length });
  } catch (e) { return fail_(e, { data: [] }); }
}

// Pilihan materi ringkas untuk form buat tugas
function getMateriOpsi(token, kelas, mapel) {
  try {
    auth_(token, STAFF);
    const t = tingkat_(kelas);
    const out = materiMeta_().filter(d => (!t || d.kelas === t) && (!mapel || d.mapel === mapel)).slice(0, 200)
      .map(d => ({ id: d.id, bab: d.bab, judul: d.judul }));
    return ok_({ data: out });
  } catch (e) { return fail_(e, { data: [] }); }
}

// Isi lengkap satu materi, hanya saat dibuka
function getMateri(token, id) {
  try {
    const me = auth_(token);
    const key = "md_" + id;
    let m = cacheGetJson_(key);
    if (!m) {
      const sh = initSheet(), row = findRow_(sh, 1, id, true);
      if (!row) return { success: false, message: "Buku tidak ditemukan." };
      const v = sh.getRange(row, 1, 1, 8).getValues()[0];
      m = { id: String(v[0]), tanggal: toIso_(v[1]), kelas: String(v[2]), semester: String(v[3]), mapel: String(v[4]), bab: String(v[5]), judul: String(v[6]), konten: String(v[7]) };
      cachePutJson_(key, m, 900);
    }
    if (me.role === ROLE_USER && (!tingkat_(me.kelas) || String(m.kelas).trim() !== tingkat_(me.kelas))) return { success: false, message: "Buku ini bukan untuk kelasmu." };
    return ok_({ data: m });
  } catch (e) { return fail_(e); }
}

// Tidak ada menu "Tulis Materi" di UI. Buku ditambah langsung di sheet BankMateri (ID & Tanggal terisi otomatis). Fungsi ini hanya cadangan.
function saveMateri(token, data) {
  try {
    auth_(token, STAFF);
    if (!/^[1-6]$/.test(String(data.kelas))) return { success: false, message: "Kelas harus 1 sampai 6." };
    if (!String(data.judul || "").trim() || !String(data.mapel || "").trim()) return { success: false, message: "Judul dan mapel wajib diisi." };
    if (String(data.konten || "").length > MAX_CELL) return { success: false, message: "Isi materi terlalu panjang (maks " + MAX_CELL + " karakter). Pecah menjadi beberapa buku." };
    return withLock_(function () {
      const sheet = initSheet(), date = new Date().toISOString();
      const id = data.id || Utilities.getUuid();
      const row = data.id ? findRow_(sheet, 1, data.id, true) : 0;
      if (row) {
        sheet.getRange(row, 2, 1, 7).setValues([[date, data.kelas, data.semester, data.mapel, data.bab, data.judul, data.konten]]);
      } else {
        sheet.appendRow([id, date, data.kelas, data.semester, data.mapel, data.bab, data.judul, data.konten]);
      }
      bumpVer_("materi"); cacheRemove_("md_" + id); cacheRemove_("stats");
      return ok_({ message: row ? "Materi berhasil diperbarui!" : "Materi berhasil disimpan!", id: id });
    });
  } catch (e) { return fail_(e); }
}

function deleteMateri(token, id) {
  try {
    auth_(token, STAFF);
    return withLock_(function () {
      const sheet = initSheet(), row = findRow_(sheet, 1, id, true);
      if (!row) return { success: false, message: "Data tidak ditemukan." };
      sheet.deleteRow(row);
      bumpVer_("materi"); cacheRemove_("md_" + id); cacheRemove_("stats");
      return ok_({ message: "Buku materi berhasil dihapus." });
    });
  } catch (e) { return fail_(e); }
}

// ==========================================
// PENGATURAN GEMINI (SUPERADMIN, API key hanya di server)
// ==========================================
function getSettings(token) {
  try {
    auth_(token, SUPER);
    const p = PropertiesService.getScriptProperties();
    const key = p.getProperty('GEMINI_API_KEY') || "";
    return ok_({ hasKey: !!key, keyHint: key ? "••••" + key.slice(-4) : "", model: p.getProperty('GEMINI_MODEL') || DEFAULT_MODEL });
  } catch (e) { return fail_(e); }
}
function saveSettings(token, apiKey, model) {
  try {
    auth_(token, SUPER);
    const p = PropertiesService.getScriptProperties();
    if (apiKey && String(apiKey).trim()) p.setProperty('GEMINI_API_KEY', String(apiKey).trim());
    if (model && String(model).trim()) p.setProperty('GEMINI_MODEL', String(model).trim());
    return ok_({ message: "Pengaturan disimpan." });
  } catch (e) { return fail_(e); }
}
function deleteApiKey(token) {
  try {
    auth_(token, SUPER);
    PropertiesService.getScriptProperties().deleteProperty('GEMINI_API_KEY');
    return ok_({ message: "API key dihapus." });
  } catch (e) { return fail_(e); }
}
function listGeminiModels(token) {
  try {
    auth_(token, SUPER);
    const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return { success: false, message: "Simpan API key terlebih dahulu." };
    const res = UrlFetchApp.fetch("https://generativelanguage.googleapis.com/v1beta/models?pageSize=200", {
      method: "get", headers: { "x-goog-api-key": key }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return { success: false, message: geminiError_(res) };
    const models = (JSON.parse(res.getContentText()).models || [])
      .filter(m => (m.supportedGenerationMethods || []).indexOf("generateContent") >= 0)
      .map(m => ({ id: String(m.name).replace("models/", ""), nama: m.displayName || m.name }))
      .filter(m => /^gemini-/.test(m.id) && !/(tts|image|embedding|live|audio|robotics|computer-use)/.test(m.id))
      .sort((a, b) => b.id.localeCompare(a.id));
    return ok_({ data: models });
  } catch (e) { return fail_(e); }
}
function geminiError_(res) {
  const code = res.getResponseCode();
  let detail = "";
  try { detail = JSON.parse(res.getContentText()).error.message; } catch (e) { detail = res.getContentText().substring(0, 200); }
  if (code === 400 && /API key/i.test(detail)) return "API key tidak valid.";
  if (code === 403) return "API key ditolak/tidak punya akses: " + detail;
  if (code === 404) return "Model tidak ditemukan. Pilih model lain di Pengaturan.";
  if (code === 429) return "Kuota Gemini habis/terlalu sering dipakai. Coba lagi sebentar.";
  return "Gemini error " + code + ": " + detail;
}

// ==========================================
// GENERATE SOAL + BUKU MATERI DENGAN GEMINI (ADMIN & SUPERADMIN)
// ==========================================
// Satu kali "Buat soal" menghasilkan DUA hal:
//   1) draf soal 6 bentuk (diperiksa guru dulu sebelum disimpan sebagai tugas)
//   2) buku materi 5 halaman (tanpa soal) yang LANGSUNG tersimpan di Perpustakaan
//      sesuai kelas dan materi. Generate ulang untuk materi yang sama memperbarui buku yang sama (tidak dobel).
const TIPE_ORDER = ["pg", "pgk", "jodoh", "bs", "isian", "uraian"];   // semua bentuk (tugas lama masih bisa dibaca)
const TIPE_BARU  = ["pg", "pgk", "jodoh", "bs", "isian", "uraian"];     // bentuk yang boleh dibuat (6 bentuk, urutan tetap)
const MAX_TIPE = { pg: 30, pgk: 30, jodoh: 10, bs: 10, isian: 10, uraian: 10 };   // maksimal soal per bentuk
const MAX_SOAL = 100;                                                   // maksimal soal per tugas; nilai maksimal tetap 100
const TIPE_GURU = ["jodoh", "bs", "isian", "uraian"];                   // dinilai guru (klik Benar/Salah)

function clampInt_(v, lo, hi) { return Math.max(lo, Math.min(hi, int_(v, 0))); }
function shuffle_(a) {
  for (let i = a.length - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); const t = a[i]; a[i] = a[j]; a[j] = t; }
  return a;
}

// Panggilan generik ke Gemini. Hasil: {ok:true, data} atau {ok:false, message}
function callGemini_(key, model, prompt, schema, maxTokens, temperature) {
  const url = "https://generativelanguage.googleapis.com/v1beta/models/" + encodeURIComponent(model) + ":generateContent";
  const payload = {
    contents: [{ role: "user", parts: [{ text: prompt }] }],
    generationConfig: { responseMimeType: "application/json", responseSchema: schema, temperature: temperature, maxOutputTokens: maxTokens }
  };
  let res, attempt = 0;
  while (true) {
    res = UrlFetchApp.fetch(url, { method: "post", contentType: "application/json", headers: { "x-goog-api-key": key }, payload: JSON.stringify(payload), muteHttpExceptions: true });
    const c = res.getResponseCode();
    if ((c === 429 || c === 500 || c === 503) && attempt < 2) { attempt++; Utilities.sleep(2500 * attempt); continue; }
    break;
  }
  if (res.getResponseCode() !== 200) return { ok: false, message: geminiError_(res) };
  const body = JSON.parse(res.getContentText());
  if (body.promptFeedback && body.promptFeedback.blockReason) return { ok: false, message: "Permintaan diblokir filter Gemini. Ubah topik atau catatan." };
  const cand = body.candidates && body.candidates[0];
  const parts = (cand && cand.content && cand.content.parts) || [];
  const text = parts.filter(x => x.text && !x.thought).map(x => x.text).join("");
  if (!text) return { ok: false, message: "Gemini tidak mengembalikan hasil (kemungkinan diblokir filter). Ubah topik lalu coba lagi." };
  try { return { ok: true, data: JSON.parse(text.replace(/```json|```/g, "").trim()) }; }
  catch (pe) { return { ok: false, message: "Jawaban Gemini terpotong/tidak valid. Kurangi jumlah soal atau coba lagi." }; }
}

function generateSoal(token, opt) {
  try {
    const me = auth_(token, STAFF);
    opt = opt || {};
    opt.mapel = cut_(opt.mapel, 60);
    if (!opt.mapel) return { success: false, message: "Mata pelajaran wajib diisi." };
    const p = PropertiesService.getScriptProperties();
    const key = p.getProperty('GEMINI_API_KEY');
    const model = p.getProperty('GEMINI_MODEL') || DEFAULT_MODEL;
    if (!key) return { success: false, message: "API key Gemini belum diatur. Minta superadmin mengisinya di menu Pengaturan." };

    const cache = CacheService.getScriptCache(), gk = "gen_" + me.username, used = Number(cache.get(gk) || 0);
    if (used >= 20) return { success: false, message: "Batas 20 kali buat soal per jam tercapai. Coba lagi nanti." };

    const jml = {
      pg: clampInt_(opt.jumlahPG, 0, MAX_TIPE.pg), pgk: clampInt_(opt.jumlahPGK, 0, MAX_TIPE.pgk),
      jodoh: clampInt_(opt.jumlahJodoh, 0, MAX_TIPE.jodoh), bs: clampInt_(opt.jumlahBS, 0, MAX_TIPE.bs),
      isian: clampInt_(opt.jumlahIsian, 0, MAX_TIPE.isian), uraian: clampInt_(opt.jumlahUraian, 0, MAX_TIPE.uraian)
    };
    const total = TIPE_BARU.reduce((a, t) => a + jml[t], 0);
    if (total === 0) return { success: false, message: "Jumlah soal tidak boleh nol." };
    if (total > MAX_SOAL) return { success: false, message: "Maksimal " + MAX_SOAL + " soal sekali buat. Sekarang " + total + " soal." };

    let sumber = "", src = null;
    if (opt.materiId) {
      const sh = initSheet(), row = findRow_(sh, 1, opt.materiId, true);
      if (row) {
        const v = sh.getRange(row, 1, 1, 8).getValues()[0];
        src = { judul: String(v[6]), bab: String(v[5]), semester: String(v[3]) };
        sumber = "Judul: " + v[6] + "\nBab: " + v[5] + "\n\n" + kontenKeTeks_(v[7]).substring(0, 12000);
      }
    }
    if (!sumber && !String(opt.topik || "").trim()) return { success: false, message: "Pilih materi atau isi topik." };
    cache.put(gk, String(used + 1), 3600);

    const kls = tingkat_(opt.kelas) || opt.kelas;
    const kom = [];
    if (jml.pg) kom.push(jml.pg + ' soal tipe "pg" (pilihan ganda): "opsi" berisi tepat 4 teks tanpa awalan huruf, hanya 1 jawaban benar; "kunci" satu huruf A, B, C, atau D.');
    if (jml.pgk) kom.push(jml.pgk + ' soal tipe "pgk" (pilihan ganda kompleks): "opsi" berisi 4 atau 5 teks tanpa awalan huruf dengan 2 atau 3 jawaban benar; "kunci" berisi huruf yang benar dipisah koma, contoh "A,C"; akhiri pertanyaan dengan "(pilih lebih dari satu jawaban)".');
    if (jml.jodoh) kom.push(jml.jodoh + ' soal tipe "jodoh" (menjodohkan): tiap soal berisi 4 pasangan pada "pasangan" (objek "kiri" dan "kanan" yang berpasangan benar; teks "kanan" harus berbeda satu sama lain); "pertanyaan" berisi petunjuk singkat; "opsi" array kosong; "kunci" string kosong.');
    if (jml.bs) kom.push(jml.bs + ' soal tipe "bs" (benar-salah): "pertanyaan" berisi SATU pernyataan; "opsi" array kosong; "kunci" berisi "Benar" atau "Salah". Usahakan jumlah pernyataan benar dan salah seimbang.');
    if (jml.isian) kom.push(jml.isian + ' soal tipe "isian" (isian singkat): "pertanyaan" berupa kalimat dengan titik-titik (......) yang harus diisi; "opsi" array kosong; "kunci" berisi jawaban singkat (1 sampai 3 kata); bila ada penulisan lain yang sama benar, pisahkan dengan tanda " | ".');
    if (jml.uraian) kom.push(jml.uraian + ' soal tipe "uraian": "opsi" array kosong; "kunci" berisi contoh jawaban ideal singkat.');

    const prompt =
      "Kamu adalah guru SD/MI di Indonesia yang ahli menyusun soal sesuai Kurikulum Merdeka.\n" +
      "Buat soal untuk siswa Kelas " + kls + ", mata pelajaran " + opt.mapel + ".\n" +
      "Tingkat kesulitan: " + (opt.tingkat || "sedang") + ".\n" +
      "Komposisi soal, DIURUTKAN persis seperti ini (semua soal sejenis dikelompokkan berurutan):\n- " + kom.join("\n- ") + "\n" +
      bahasaInstruksi_(opt.bahasa) + " Pastikan jawaban benar dan tidak ambigu.\n" +
      "Untuk field yang tidak dipakai suatu soal, isi array kosong atau string kosong. 'pembahasan' berisi penjelasan singkat 1-2 kalimat.\n\n" +
      (sumber ? "Dasarkan soal HANYA pada materi berikut:\n\"\"\"\n" + sumber + "\n\"\"\"" : "Topik: " + opt.topik) +
      (opt.catatan ? "\n\nCatatan tambahan dari guru: " + String(opt.catatan).substring(0, 500) : "");

    const schema = {
      type: "OBJECT",
      properties: { soal: { type: "ARRAY", items: { type: "OBJECT", properties: {
        tipe: { type: "STRING", enum: TIPE_BARU }, pertanyaan: { type: "STRING" },
        opsi: { type: "ARRAY", items: { type: "STRING" } }, kunci: { type: "STRING" }, pembahasan: { type: "STRING" },
        pasangan: { type: "ARRAY", items: { type: "OBJECT", properties: { kiri: { type: "STRING" }, kanan: { type: "STRING" } }, required: ["kiri", "kanan"] } }
      }, required: ["tipe", "pertanyaan", "opsi", "kunci", "pembahasan", "pasangan"] } } },
      required: ["soal"]
    };

    const r1 = callGemini_(key, model, prompt, schema, 32768, 0.7);
    if (!r1.ok) return { success: false, message: r1.message };
    const soal = normalizeSoal_(r1.data.soal || []);
    if (!soal.length) return { success: false, message: "Format soal dari Gemini tidak valid. Coba generate ulang." };

    // Buku materi otomatis. Kegagalan buku tidak menggagalkan soal.
    let buku = null, bukuError = "";
    try { buku = buatBukuOtomatis_(key, model, opt, src, sumber, soal); }
    catch (be) { bukuError = String((be && be.message) || be); }

    return ok_({ soal: soal, model: model, buku: buku, bukuError: bukuError });
  } catch (e) {
    const m = String((e && e.message) || e);
    if (/script\.external_request|UrlFetchApp/.test(m)) return { success: false, message: "Izin belum diberikan. Jalankan fungsi izinkanAkses di editor, lalu buat New version pada deployment." };
    return { success: false, message: "Gagal membuat soal: " + m };
  }
}

// ---------- SOAL: normalisasi 6 bentuk ----------
// Dipakai untuk hasil Gemini DAN untuk draf yang sudah diedit guru (jodoh sudah berbentuk kiri/kanan/kunci).
function normalizeSoal_(arr) {
  const out = [];
  const clean = o => String(o == null ? "" : o).replace(/^\s*[A-Ea-e][\.\)]\s+/, "").trim();
  (arr || []).forEach(s => {
    const tipe = TIPE_ORDER.indexOf(s.tipe) >= 0 ? s.tipe : "pg";
    const q = String(s.pertanyaan || "").trim();
    if (!q) return;
    const pemb = String(s.pembahasan || "");

    if (tipe === "pg") {
      const opsi = (s.opsi || []).map(clean).filter(Boolean);
      const kunci = String(s.kunci || "").trim().toUpperCase().charAt(0);
      if (opsi.length !== 4 || "ABCD".indexOf(kunci) < 0) return;
      out.push({ tipe: "pg", pertanyaan: q, opsi: opsi, kunci: kunci, pembahasan: pemb });

    } else if (tipe === "pgk") {
      const opsi = (s.opsi || []).map(clean).filter(Boolean);
      if (opsi.length < 4 || opsi.length > 5) return;
      let ks = String(s.kunci || "").toUpperCase().match(/[A-E]/g) || [];
      ks = ks.filter((v, i) => ks.indexOf(v) === i && "ABCDE".indexOf(v) < opsi.length).sort();
      if (ks.length < 2 || ks.length >= opsi.length) return;
      out.push({ tipe: "pgk", pertanyaan: q, opsi: opsi, kunci: ks.join(","), pembahasan: pemb });

    } else if (tipe === "jodoh") {
      let kiri, kanan, kunci;
      if (Array.isArray(s.kiri) && Array.isArray(s.kanan) && s.kiri.length && s.kunci) {     // sudah final (hasil editor)
        kiri = s.kiri.map(x => String(x).trim()); kanan = s.kanan.map(x => String(x).trim());
        kunci = String(s.kunci).toUpperCase().split(",").map(x => x.trim());
        const uniq = kunci.filter((v, i) => kunci.indexOf(v) === i).length === kunci.length;
        if (kiri.length < 3 || kiri.length > 6 || kanan.length !== kiri.length || kunci.length !== kiri.length || !uniq) return;
        if (kiri.some(x => !x) || kanan.some(x => !x) || kunci.some(k => "ABCDEF".indexOf(k) < 0 || "ABCDEF".indexOf(k) >= kanan.length)) return;
      } else {                                                                                  // dari Gemini: pasangan benar
        const pr = (s.pasangan || []).map(x => ({ k: String(x.kiri || "").trim(), r: String(x.kanan || "").trim() })).filter(x => x.k && x.r).slice(0, 6);
        if (pr.length < 3) return;
        const rs = pr.map(x => x.r.toLowerCase());
        if (rs.filter((v, i) => rs.indexOf(v) === i).length !== rs.length) return;
        const order = pr.map((_, i) => i);
        for (let t = 0; t < 6; t++) { shuffle_(order); if (order.some((v, i) => v !== i)) break; }
        kiri = pr.map(x => x.k);
        kanan = order.map(i => pr[i].r);
        kunci = pr.map((_, i) => "ABCDEF".charAt(order.indexOf(i)));
      }
      out.push({ tipe: "jodoh", pertanyaan: q, opsi: [], kiri: kiri, kanan: kanan, kunci: kunci.join(","), pembahasan: pemb });

    } else if (tipe === "bs") {
      const k = String(s.kunci || "").trim();
      if (!k) return;
      out.push({ tipe: "bs", pertanyaan: q, opsi: [], kunci: /^(b|t)/i.test(k) ? "B" : "S", pembahasan: pemb });

    } else if (tipe === "isian") {
      const k = String(s.kunci || "").trim();
      if (!k) return;
      out.push({ tipe: "isian", pertanyaan: q, opsi: [], kunci: k, pembahasan: pemb });

    } else {
      out.push({ tipe: "uraian", pertanyaan: q, opsi: [], kunci: String(s.kunci || ""), pembahasan: pemb });
    }
  });
  // kelompokkan per bentuk (urutan tetap I..VI), urutan dalam bentuk dipertahankan
  return out.map((s, i) => ({ s: s, i: i })).sort((a, b) => (TIPE_ORDER.indexOf(a.s.tipe) - TIPE_ORDER.indexOf(b.s.tipe)) || (a.i - b.i)).map(x => x.s);
}

// ---------- PENILAIAN OTOMATIS ----------
function kn_(x) { return String(x == null ? "" : x).replace(/[\u0640\u064B-\u065F\u0670]/g, "").replace(/[\u0622\u0623\u0625\u0671]/g, "\u0627").replace(/\u0649/g, "\u064A").toLowerCase().replace(/[^\p{L}\p{N}\s]/gu, " ").replace(/\s+/g, "").trim(); }

// Nilai satu soal: 0..1, atau null bila tidak dinilai otomatis (uraian). Menjodohkan memberi nilai sebagian per pasangan.
function nilaiSoal_(s, a) {
  a = String(a == null ? "" : a).trim();
  switch (s.tipe) {
    case "pg": return a.toUpperCase() === String(s.kunci).toUpperCase() ? 1 : 0;
    case "pgk": {
      if (!a) return 0;
      const m = a.toUpperCase().match(/[A-E]/g) || [];
      const mine = m.filter((v, i) => m.indexOf(v) === i).sort().join(",");
      const keyS = String(s.kunci).toUpperCase().split(",").map(x => x.trim()).filter(Boolean).sort().join(",");
      return mine === keyS ? 1 : 0;
    }
    case "jodoh": {
      const keyA = String(s.kunci).toUpperCase().split(",").map(x => x.trim()), mine = a.toUpperCase().split(",");
      let b = 0; keyA.forEach((k, i) => { if (k && String(mine[i] || "").trim() === k) b++; });
      return keyA.length ? b / keyA.length : 0;
    }
    case "bs": return a && a.toUpperCase().charAt(0) === String(s.kunci).toUpperCase().charAt(0) ? 1 : 0;
    case "isian": {
      if (!a) return 0;
      const mine = kn_(a);
      return String(s.kunci).split("|").some(k => kn_(k) && kn_(k) === mine) ? 1 : 0;
    }
    default: return null;
  }
}
// ---- Koreksi guru: array per soal berisi 1 (dibenarkan), 0 (disalahkan), atau "" (ikut pencocokan otomatis) ----
function korArr_(raw) {
  try { const a = JSON.parse(raw || "[]"); return Array.isArray(a) ? a : []; } catch (e) { return []; }
}
// Nilai akhir satu soal 0..1. null = soal yang dinilai guru dan belum diperiksa (dihitung 0 pada skor).
function nilaiFinal_(s, a, k) {
  if (k === 0 || k === 1) return k;
  if (s.tipe === "pg" || s.tipe === "pgk") return nilaiSoal_(s, a);   // pilihan ganda: otomatis
  return null;                                                         // lainnya: menunggu klik Benar/Salah guru (dihitung 0)
}
function nilaiList_(soal, ans, kor) { return soal.map((s, i) => nilaiFinal_(s, ans[i], (kor || [])[i])); }

// Skor 0..100: tiap soal bernilai 100 / jumlah soal. Menjodohkan memberi nilai sebagian per pasangan.
function skorTugas_(soal, ans, kor) {
  const N = soal.length;
  let poin = 0;
  nilaiList_(soal, ans, kor).forEach(n => { if (n != null) poin += n; });
  return { total: N, poin: Math.round(poin * 10) / 10, skor: N ? Math.min(100, Math.round(poin / N * 1000) / 10) : 0 };
}

// Saran untuk uraian: persentase kata kunci dari contoh jawaban yang muncul di jawaban siswa.
function tokens_(x) {
  return String(x == null ? "" : x).replace(/[\u0640\u064B-\u065F\u0670]/g, "").toLowerCase().replace(/[^\p{L}\p{N}\s]/gu, " ").split(/\s+/).filter(w => w.length >= 4);
}
function kecocokan_(kunci, jawab) {
  const kt = tokens_(kunci).filter((v, i, a) => a.indexOf(v) === i), at = tokens_(jawab);
  if (!kt.length || !at.length) return 0;
  let c = 0;
  kt.forEach(w => {
    if (at.some(x => x === w || (w.length >= 6 && x.length >= 6 && (x.indexOf(w.substring(0, w.length - 2)) === 0 || w.indexOf(x.substring(0, x.length - 2)) === 0)))) c++;
  });
  return Math.round(c / kt.length * 100);
}

function reviewSoal_(soal, ans, kor) {
  const nilai = nilaiList_(soal, ans, kor);
  return soal.map((s, i) => ({
    tipe: s.tipe, pertanyaan: s.pertanyaan, opsi: s.opsi || [], kiri: s.kiri || [], kanan: s.kanan || [],
    jawabanmu: String(ans[i] == null ? "" : ans[i]), kunci: s.kunci, nilai: nilai[i], pembahasan: s.pembahasan || ""
  }));
}

// ---------- BUKU MATERI OTOMATIS (5 halaman, tanpa soal) ----------
// Isi buku lama (teks biasa) maupun buku baru (JSON v2) diubah menjadi teks polos sebagai sumber soal/buku.
function kontenKeTeks_(konten) {
  const s = String(konten || "");
  if (s.charAt(0) === "{") {
    try {
      const d = JSON.parse(s);
      if (d && (d.v === 2 || d.v === 3)) {
        const t = [];
        (d.rangkuman || []).forEach(r => t.push(r.subjudul + ": " + r.isi));
        (d.poin || []).forEach(r => t.push("- " + r.judul + ": " + r.isi));
        (d.tambahan || []).forEach(r => t.push(r.judul + ": " + r.isi));
        (d.glosarium || []).forEach(g => t.push(g.istilah + " = " + g.arti));
        return t.join("\n");
      }
    } catch (e) {}
  }
  return s;
}

function cut_(v, n) { return String(v == null ? "" : v).replace(/\s+/g, " ").trim().substring(0, n); }

// Buang kalimat tanya (? atau tanda tanya Arab) agar buku tidak pernah memuat soal, lalu potong ke n karakter.
function bersih_(v, n) { return cut_(String(v == null ? "" : v).replace(/[^.!?\u061F]*[?\u061F]/g, ""), n); }

function normalizeBuku_(d) {
  d = d || {};
  const rangkuman = (d.rangkuman || []).map(r => ({ subjudul: bersih_(r.subjudul, 60), isi: bersih_(r.isi, 800) })).filter(r => r.subjudul && r.isi).slice(0, 6);
  const cabang = ((d.mindmap && d.mindmap.cabang) || []).map(c => ({
    judul: bersih_(c.judul, 40), anak: (c.anak || []).map(a => bersih_(a, 60)).filter(Boolean).slice(0, 4)
  })).filter(c => c.judul && c.anak.length).slice(0, 6);
  const poin = (d.poin || []).map(r => ({ judul: bersih_(r.judul, 60), isi: bersih_(r.isi, 300) })).filter(r => r.judul && r.isi).slice(0, 8);
  const tambahan = (d.tambahan || []).map(r => ({ judul: bersih_(r.judul, 70), isi: bersih_(r.isi, 700) })).filter(r => r.judul && r.isi).slice(0, 4);
  const glosarium = (d.glosarium || []).map(g => ({ istilah: bersih_(g.istilah, 40), arti: bersih_(g.arti, 180) })).filter(g => g.istilah && g.arti).slice(0, 12);
  if (rangkuman.length < 4 || cabang.length < 3 || poin.length < 4 || tambahan.length < 2 || glosarium.length < 3) return null;
  return {
    v: 3,
    tujuan: (d.tujuan || []).map(x => bersih_(x, 140)).filter(Boolean).slice(0, 4),
    rangkuman: rangkuman,
    mindmap: { pusat: bersih_(d.mindmap && d.mindmap.pusat, 50), cabang: cabang },
    poin: poin,
    tambahan: tambahan,
    fakta: bersih_(d.fakta, 300),
    glosarium: glosarium
  };
}

// Daftar konsep yang diuji soal (hanya untuk memandu isi buku; tidak ditampilkan dan tidak disalin ke buku).
function cakupanSoal_(soal) {
  const L = "ABCDE";
  return (soal || []).map(s => {
    let j = "";
    if (s.tipe === "pg" || s.tipe === "pgk") j = String(s.kunci).split(",").map(k => (s.opsi || [])[L.indexOf(k.trim())] || "").filter(Boolean).join("; ");
    else if (s.tipe === "jodoh") j = (s.kiri || []).map((k, i) => k + " = " + (s.kanan || [])["ABCDEF".indexOf(String(s.kunci).split(",")[i])]).join("; ");
    else if (s.tipe === "bs") j = s.kunci === "B" ? "pernyataan benar" : "pernyataan salah";
    else j = s.kunci || "";
    return "- " + cut_(s.pertanyaan, 160) + (j ? " => " + cut_(j, 120) : "");
  }).join("\n").substring(0, 6000);
}

function buatBukuOtomatis_(key, model, opt, src, sumber, soal) {
  const kelas = tingkat_(opt.kelas) || String(opt.kelas || "");
  if (!/^[1-6]$/.test(kelas)) throw new Error("Kelas tidak valid untuk buku.");
  const topik = String(opt.topik || "").trim();
  const dasar = String((src && src.judul) || topik || "Materi").replace(/^Buku Materi:\s*/i, "").trim();
  const meta = {
    kelas: kelas, mapel: String(opt.mapel || ""), judul: "Buku Materi: " + dasar,
    bab: (src && src.bab) || dasar,
    semester: (src && src.semester) || (opt.semester === "Genap" ? "Genap" : "Ganjil")
  };
  const cakupan = cakupanSoal_(soal);

  const prompt =
    "Kamu adalah penulis buku pelajaran SD/MI di Indonesia (Kurikulum Merdeka).\n" +
    "Buat BUKU MATERI 7 halaman untuk siswa Kelas " + kelas + ", mata pelajaran " + meta.mapel + ".\n" +
    "Buku harus LENGKAP dan menjadi bahan belajar utama: seluruh konsep, fakta, istilah, dan langkah yang dibutuhkan untuk menguasai materi, termasuk yang menjadi dasar daftar cakupan soal di bawah, harus dijelaskan runtut di rangkuman.\n" +
    "Isi buku HANYA: tujuan belajar, rangkuman materi, peta pikiran (mind map), poin-poin penting, materi tambahan (pengayaan), fakta menarik, dan glosarium.\n" +
    "DILARANG KERAS menyertakan soal, latihan, kuis, teka-teki, perintah mengerjakan, atau kalimat tanya apa pun. Jangan menyalin kalimat soal; tulis semuanya sebagai penjelasan berupa pernyataan.\n" +
    bahasaBuku_(opt.bahasa) + " Semua informasi harus benar.\n" +
    "Ketentuan field:\n" +
    "- 'tujuan': 3 tujuan belajar singkat diawali kata kerja (contoh: Menjelaskan ...).\n" +
    "- 'rangkuman': 6 bagian berurutan dari konsep dasar ke lanjutan; tiap bagian punya 'subjudul' (maks 6 kata) dan 'isi' 4 sampai 6 kalimat (maks 650 karakter) berisi definisi, penjelasan sebab-akibat, dan contoh nyata.\n" +
    "- 'mindmap': 'pusat' (maks 5 kata) dan 5 'cabang'; tiap cabang punya 'judul' (maks 4 kata) dan 2 sampai 3 'anak' berupa frasa pendek (maks 6 kata).\n" +
    "- 'poin': 8 poin penting; 'judul' maks 5 kata, 'isi' 1 sampai 2 kalimat (maks 220 karakter).\n" +
    "- 'tambahan': 4 bagian materi tambahan di luar rangkuman (misalnya penerapan dalam kehidupan sehari-hari, kisah atau tokoh, kegiatan sederhana yang bisa dicoba di rumah, tips mengingat); 'judul' maks 6 kata, 'isi' 3 sampai 5 kalimat (maks 550 karakter), bukan kalimat tanya.\n" +
    "- 'fakta': 1 sampai 2 kalimat fakta menarik yang berkaitan dengan materi.\n" +
    "- 'glosarium': 12 istilah penting; 'istilah' singkat dan 'arti' maks 140 karakter.\n\n" +
    (sumber ? "Dasarkan buku pada materi berikut:\n\"\"\"\n" + sumber + "\n\"\"\"" : "Topik: " + topik) +
    (cakupan ? "\n\nCakupan soal (hanya untuk memastikan materinya terjelaskan di buku; JANGAN disalin sebagai soal):\n" + cakupan : "") +
    (opt.catatan ? "\n\nCatatan tambahan dari guru: " + String(opt.catatan).substring(0, 500) : "");

  const S = { type: "STRING" };
  const arrS = { type: "ARRAY", items: S };
  const schema = {
    type: "OBJECT",
    properties: {
      tujuan: arrS,
      rangkuman: { type: "ARRAY", items: { type: "OBJECT", properties: { subjudul: S, isi: S }, required: ["subjudul", "isi"] } },
      mindmap: { type: "OBJECT", properties: { pusat: S, cabang: { type: "ARRAY", items: { type: "OBJECT", properties: { judul: S, anak: arrS }, required: ["judul", "anak"] } } }, required: ["pusat", "cabang"] },
      poin: { type: "ARRAY", items: { type: "OBJECT", properties: { judul: S, isi: S }, required: ["judul", "isi"] } },
      tambahan: { type: "ARRAY", items: { type: "OBJECT", properties: { judul: S, isi: S }, required: ["judul", "isi"] } },
      fakta: S,
      glosarium: { type: "ARRAY", items: { type: "OBJECT", properties: { istilah: S, arti: S }, required: ["istilah", "arti"] } }
    },
    required: ["tujuan", "rangkuman", "mindmap", "poin", "tambahan", "fakta", "glosarium"]
  };

  const r = callGemini_(key, model, prompt, schema, 14000, 0.6);
  if (!r.ok) throw new Error(r.message);
  const buku = normalizeBuku_(r.data);
  if (!buku) throw new Error("Isi buku dari Gemini kurang lengkap.");
  if (!buku.mindmap.pusat) buku.mindmap.pusat = dasar.substring(0, 50);
  return simpanBukuOtomatis_(meta, buku);
}

function simpanBukuOtomatis_(meta, buku) {
  const konten = JSON.stringify(buku);
  if (konten.length > MAX_CELL) throw new Error("Isi buku terlalu panjang.");
  return withLock_(function () {
    const sh = initSheet(), date = new Date().toISOString();
    const lama = materiMeta_().filter(m => m.kelas === meta.kelas && m.mapel === meta.mapel && m.judul === meta.judul)[0];
    let id = "", row = 0;
    if (lama) { id = lama.id; row = findRow_(sh, 1, id, true); }
    if (row) {
      sh.getRange(row, 2, 1, 7).setValues([[date, meta.kelas, meta.semester, meta.mapel, meta.bab, meta.judul, konten]]);
    } else {
      id = Utilities.getUuid();
      sh.appendRow([id, date, meta.kelas, meta.semester, meta.mapel, meta.bab, meta.judul, konten]);
    }
    bumpVer_("materi"); cacheRemove_("md_" + id); cacheRemove_("stats");
    return { id: id, judul: meta.judul, kelas: meta.kelas, diperbarui: !!row };
  }, 30000);
}

// ==========================================
// TUGAS (ADMIN & SUPERADMIN)
// ==========================================
// Metadata tugas tanpa soal & tanpa jumlah terkumpul. Terbaru di depan. Di-cache.
function tugasMeta_() {
  const key = "tm_" + getVer_("tugas");
  const hit = cacheGetJson_(key);
  if (hit) return hit;
  const sh = tugasSheet_(), last = sh.getLastRow();
  let out = [];
  if (last >= 2) {
    const a = sh.getRange(2, 1, last - 1, 7).getValues();    // A:G
    const b = sh.getRange(2, 9, last - 1, 3).getValues();    // I:K
    const c = sh.getMaxColumns() >= 13 ? sh.getRange(2, 13, last - 1, 1).getValues() : null;   // M: NilaiRilis
    const e = sh.getMaxColumns() >= 15 ? sh.getRange(2, 14, last - 1, 2).getValues() : null;   // N:O Token, Durasi
    out = a.map((r, i) => ({
      id: String(r[0]), tanggal: toIso_(r[1]), kelas: String(r[2]), mapel: String(r[3]), judul: String(r[4]),
      materiId: String(r[5]), instruksi: String(r[6]).substring(0, 160), deadline: toIso_(b[i][0]), status: String(b[i][1]), jumlahSoal: Number(b[i][2]) || 0,
      rilis: c ? String(c[i][0]).toUpperCase() === "TRUE" : false,
      token: e ? String(e[i][0] || "") : "", pakaiToken: !!(e && String(e[i][0] || "")), durasi: e ? (Number(e[i][1]) || 0) : 0
    })).reverse();
  }
  cachePutJson_(key, out, 1800);
  return out;
}

function kelasTugasValid_(k) { k = String(k || ""); return /^[1-6]$/.test(k) || !!rombelSet_()[k]; }

function saveTugas(token, data) {
  try {
    auth_(token, STAFF);
    const soal = normalizeSoal_(data.soal || []);
    if (!soal.length) return { success: false, message: "Tugas harus punya minimal satu soal." };
    const perTipe = {}; (data.soal || []).forEach(q => { perTipe[q.tipe] = (perTipe[q.tipe] || 0) + 1; });
    for (const t in MAX_TIPE) if ((perTipe[t] || 0) > MAX_TIPE[t]) return { success: false, message: "Maksimal " + MAX_TIPE[t] + " soal untuk satu bentuk soal (" + t + "). Kurangi jumlahnya." };
    if ((data.soal || []).length > MAX_SOAL) return { success: false, message: "Maksimal " + MAX_SOAL + " soal per tugas. Sekarang " + (data.soal || []).length + " soal." };
    if (soal.length !== (data.soal || []).length) return { success: false, message: "Ada soal yang tidak valid (opsi kosong, kunci belum dipilih, atau pasangan tidak lengkap). Periksa kembali soalnya." };
    if (!data.judul || !data.kelas || !data.mapel) return { success: false, message: "Judul, kelas, dan mapel wajib diisi." };
    if (!kelasTugasValid_(data.kelas)) return { success: false, message: "Kelas/rombel tidak valid." };
    const json = JSON.stringify(soal);
    if (json.length > MAX_CELL) return { success: false, message: "Soal terlalu panjang untuk satu tugas. Kurangi jumlah soal atau persingkat teksnya." };
    const status = ["Aktif", "Draft", "Tutup"].indexOf(data.status) >= 0 ? data.status : "Aktif";
    const ak = aksesValid_(data);
    if (ak.error) return { success: false, message: ak.error };
    const tokenVal = ak.token, durasi = ak.durasi;

    return withLock_(function () {
      const sh = tugasSheet_();
      const id = data.id || Utilities.getUuid();
      const row = [id, new Date().toISOString(), String(data.kelas), data.mapel, data.judul, data.materiId || "",
                   data.instruksi || "", json, data.deadline || "", status, soal.length];
      const r = data.id ? findRow_(sh, 1, data.id, true) : 0;
      if (r) {
        sh.getRange(r, 1, 1, row.length).setValues([row]);
        sh.getRange(r, 14, 1, 2).setValues([[tokenVal, durasi]]);
      } else {
        row.push(0, "FALSE", tokenVal, durasi);   // Terkumpul, NilaiRilis (disembunyikan sampai dirilis), Token, Durasi
        sh.appendRow(row);
      }
      bumpVer_("tugas"); cacheRemove_("stats");
      return ok_({ message: r ? "Tugas diperbarui." : "Tugas berhasil disimpan.", id: id, token: tokenVal });
    });
  } catch (e) { return fail_(e); }
}

function setTugasStatus(token, id, status) {
  try {
    auth_(token, STAFF);
    if (["Aktif", "Draft", "Tutup"].indexOf(status) < 0) return { success: false, message: "Status tidak valid." };
    return withLock_(function () {
      const sh = tugasSheet_(), row = findRow_(sh, 1, id, true);
      if (!row) return { success: false, message: "Tugas tidak ditemukan." };
      sh.getRange(row, 10).setValue(status); bumpVer_("tugas");
      return ok_({ message: "Status tugas: " + status });
    });
  } catch (e) { return fail_(e); }
}

// Guru/admin menampilkan atau menyembunyikan nilai (skor, kunci, pembahasan) untuk siswa pada satu tugas.
function setNilaiRilis(token, id, rilis) {
  try {
    auth_(token, STAFF);
    return withLock_(function () {
      const sh = tugasSheet_(), row = findRow_(sh, 1, id, true);
      if (!row) return { success: false, message: "Tugas tidak ditemukan." };
      sh.getRange(row, 13).setValue(rilis ? "TRUE" : "FALSE"); bumpVer_("tugas");
      return ok_({ message: rilis ? "Nilai dirilis. Siswa sekarang bisa melihat skor dan pembahasan." : "Nilai disembunyikan dari siswa." });
    });
  } catch (e) { return fail_(e); }
}

function deleteTugas(token, id) {
  try {
    auth_(token, STAFF);
    return withLock_(function () {
      const sh = tugasSheet_(), row = findRow_(sh, 1, id, true);
      if (!row) return { success: false, message: "Tugas tidak ditemukan." };
      sh.deleteRow(row);
      // hapus jawaban terkait: baca hanya kolom TugasId, hapus per blok
      const js = jawabSheet_(), rows = findAllRows_(js, 2, id);
      deleteRowsBottomUp_(js, rows, Date.now());
      bumpVer_("tugas"); cacheRemove_("stats");
      return ok_({ message: "Tugas dan jawabannya dihapus." });
    }, 60000);
  } catch (e) { return fail_(e); }
}

// opt: {kelas:'all'|'5'|'5A', page}
function getTugasAdmin(token, opt) {
  try {
    auth_(token, STAFF);
    opt = opt || {};
    const filter = String(opt.kelas || "all"), size = 20, page = Math.max(1, int_(opt.page, 1));
    const list = tugasMeta_().filter(t => cocokFilter_(t.kelas, filter));
    const slice = list.slice((page - 1) * size, page * size);
    // jumlah terkumpul: satu kolom saja, hanya untuk baris yang tampil
    const sh = tugasSheet_(), last = sh.getLastRow();
    const ids = last >= 2 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    const cnt = last >= 2 ? sh.getRange(2, 12, last - 1, 1).getValues() : [];
    const map = {}; ids.forEach((r, i) => map[String(r[0])] = Number(cnt[i][0]) || 0);
    const data = slice.map(t => Object.assign({}, t, { terkumpul: map[t.id] || 0 }));
    return ok_({ data: data, total: list.length, page: page, hasMore: page * size < list.length });
  } catch (e) { return fail_(e, { data: [] }); }
}

function getTugasDetailAdmin(token, id) {
  try {
    auth_(token, STAFF);
    const sh = tugasSheet_(), row = findRow_(sh, 1, id, true);
    if (!row) return { success: false, message: "Tugas tidak ditemukan." };
    const v = sh.getRange(row, 1, 1, 10).getValues()[0];
    return ok_({ data: { id: String(v[0]), kelas: String(v[2]), mapel: String(v[3]), judul: String(v[4]), instruksi: String(v[6]),
      soal: JSON.parse(v[7]), deadline: toIso_(v[8]), status: String(v[9]) } });
  } catch (e) { return fail_(e); }
}

// Hasil pengerjaan satu tugas. kelasFilter opsional ('5A' atau '5').
// Soal (dengan kunci) dikirim SEKALI; tiap siswa hanya membawa jawaban, koreksi, nilai per soal, dan saran uraian.
function getHasilTugas(token, tugasId, kelasFilter) {
  try {
    auth_(token, STAFF);
    const ts = tugasSheet_(), trow = findRow_(ts, 1, tugasId, true);
    if (!trow) return { success: false, message: "Tugas tidak ditemukan.", data: [] };
    const soal = JSON.parse(ts.getRange(trow, 8).getValue() || "[]");
    const js = jawabSheet_();
    const rows = findAllRows_(js, 2, tugasId);             // hanya memindai kolom TugasId
    if (!rows.length) return ok_({ data: [], soal: soal });
    const A = readRows_(js, rows, 1, 6);                   // ID, TugasId, Username, Nama, Kelas, Waktu
    const filter = String(kelasFilter || "all");
    const sel = rows.filter(r => cocokKelasSiswa_(A[r][4], filter));
    const B = sel.length ? readRows_(js, sel, 7, 11, 10) : {};   // Jawaban, Skor, Benar, TotalPG, Koreksi
    const out = sel.map(r => {
      let ans = []; try { ans = JSON.parse((B[r] && B[r][0]) || "[]"); } catch (e) {}
      const kor = korArr_(B[r] && B[r][4]);
      while (kor.length < soal.length) kor.push("");
      const sc = skorTugas_(soal, ans, kor);               // dihitung ulang agar tugas lama ikut aturan nilai maks 100
      return {
        id: String(A[r][0]), username: String(A[r][2]), nama: String(A[r][3]), kelas: String(A[r][4]), waktu: toIso_(A[r][5]),
        skor: sc.skor, benar: sc.poin, totalPG: sc.total,
        jawaban: soal.map((_, i) => String(ans[i] == null ? "" : ans[i])),
        kor: kor.slice(0, soal.length),
        nilai: nilaiList_(soal, ans, kor),
        saran: soal.map((q, i) => q.tipe === "uraian" ? kecocokan_(q.kunci, ans[i]) : (TIPE_GURU.indexOf(q.tipe) >= 0 ? Math.round((nilaiSoal_(q, ans[i]) || 0) * 100) : null))
      };
    });
    out.sort((x, y) => y.skor - x.skor);
    return ok_({ data: out, soal: soal });
  } catch (e) { return fail_(e, { data: [] }); }
}

// Guru menekan Benar / Salah pada satu soal milik satu jawaban siswa. val: 1 = Benar, 0 = Salah, "" = kembali ke otomatis.
// Skor dihitung ulang dan disimpan (maksimal 100).
function setKoreksi(token, jawabId, idx, val) {
  try {
    auth_(token, STAFF);
    idx = int_(idx, -1);
    return withLock_(function () {
      const js = jawabSheet_(), row = findRow_(js, 1, jawabId, true);
      if (!row) return { success: false, message: "Jawaban siswa tidak ditemukan." };
      const r = js.getRange(row, 1, 1, 11).getValues()[0];
      const ts = tugasSheet_(), trow = findRow_(ts, 1, String(r[1]), true);
      if (!trow) return { success: false, message: "Tugas tidak ditemukan." };
      const soal = JSON.parse(ts.getRange(trow, 8).getValue() || "[]");
      if (idx < 0 || idx >= soal.length) return { success: false, message: "Nomor soal tidak valid." };
      let ans = []; try { ans = JSON.parse(r[6] || "[]"); } catch (e) {}
      const kor = korArr_(r[10]);
      while (kor.length < soal.length) kor.push("");
      kor[idx] = (val === "" || val == null) ? "" : (Number(val) >= 1 ? 1 : 0);
      const sc = skorTugas_(soal, ans, kor);
      js.getRange(row, 8, 1, 4).setValues([[sc.skor, sc.poin, sc.total, JSON.stringify(kor.slice(0, soal.length))]]);
      cacheRemove_("done_" + String(r[2]));
      return ok_({ skor: sc.skor, benar: sc.poin, totalPG: sc.total, kor: kor.slice(0, soal.length), nilai: nilaiList_(soal, ans, kor) });
    }, 30000);
  } catch (e) { return fail_(e); }
}

// ==========================================
// TUGAS (SISWA) - kunci jawaban TIDAK dikirim sebelum dikumpulkan
// ==========================================
// Peta tugas yang sudah dikumpulkan siswa: {tugasId: {skor, waktu, row}}. Di-cache per siswa.
function doneMap_(username, force) {
  const key = "done_" + username;
  if (!force) { const hit = cacheGetJson_(key); if (hit) return hit; }
  const sh = jawabSheet_(), map = {};
  const rows = findAllRows_(sh, 3, username);              // hanya memindai kolom Username
  if (rows.length) {
    const bf = readRows_(sh, rows, 2, 6);                  // TugasId .. Waktu
    const h = readRows_(sh, rows, 8, 8);                   // Skor
    rows.forEach(r => { map[bf[r][0]] = { skor: Number(h[r][0]), waktu: toIso_(bf[r][4]), row: r }; });
  }
  cachePutJson_(key, map, SESSION_TTL);
  return map;
}

function deadlinePassed_(deadline) {
  if (!deadline) return false;
  const d = new Date(deadline);
  return !isNaN(d.getTime()) && new Date().getTime() > d.getTime();
}

// opt: {page}
function getTugasSiswa(token, opt) {
  try {
    const me = auth_(token, STUDENT);
    opt = opt || {};
    const size = 20, page = Math.max(1, int_(opt.page, 1));
    const done = doneMap_(me.username);
    const list = tugasMeta_().filter(t => t.status !== "Draft" && cocokKelas_(t.kelas, me.kelas));
    const data = list.slice((page - 1) * size, page * size).map(t => {
      const d = done[t.id];
      // skor hanya dikirim bila guru sudah merilis nilai tugas ini
      return { id: t.id, kelas: t.kelas, mapel: t.mapel, judul: t.judul, instruksi: t.instruksi, deadline: t.deadline, status: t.status,
               jumlahSoal: t.jumlahSoal, pakaiToken: t.pakaiToken, durasi: t.durasi, selesai: !!d, rilis: !!t.rilis, skor: (d && t.rilis) ? d.skor : null, waktu: d ? d.waktu : null };
    });
    return ok_({ data: data, total: list.length, page: page, hasMore: page * size < list.length });
  } catch (e) { return fail_(e, { data: [] }); }
}

// tokenInput: wajib bila tugas memakai token. Timer durasi dicatat di server (cache) saat siswa pertama kali mulai;
// membuka ulang tugas yang sedang berjalan melanjutkan sisa waktu yang sama.
function mulaiTugas(token, id, tokenInput) {
  try {
    const me = auth_(token, STUDENT);
    const sh = tugasSheet_(), row = findRow_(sh, 1, id, true);
    if (!row) return { success: false, message: "Tugas tidak ditemukan." };
    const v = sh.getRange(row, 1, 1, 15).getValues()[0];
    if (String(v[9]) !== "Aktif") return { success: false, message: "Tugas ini belum/tidak lagi dibuka." };
    if (!cocokKelas_(v[2], me.kelas)) return { success: false, message: "Tugas ini bukan untuk kelasmu." };
    if (doneMap_(me.username)[id]) return { success: false, message: "Kamu sudah mengumpulkan tugas ini." };
    if (deadlinePassed_(v[8])) return { success: false, message: "Batas waktu pengumpulan sudah lewat." };

    const kode = String(v[13] || "").toUpperCase(), durasi = Number(v[14]) || 0;
    const cache = CacheService.getScriptCache(), mk = "mulai_" + me.username + "_" + id;
    let mulai = Number(cache.get(mk) || 0);
    if (kode && !mulai) {
      const fk = "tokfail_" + me.username + "_" + id, gagal = Number(cache.get(fk) || 0);
      if (gagal >= 5) return { success: false, message: "Terlalu banyak salah memasukkan token. Coba lagi dalam 5 menit." };
      if (String(tokenInput || "").trim().toUpperCase() !== kode) { cache.put(fk, String(gagal + 1), 300); return { success: false, message: "Token salah." }; }
      cache.remove(fk);
    }
    if (!mulai && (kode || durasi)) { mulai = Date.now(); cache.put(mk, String(mulai), 21600); }
    let sisa = 0;
    if (durasi) {
      sisa = Math.floor((mulai + durasi * 60000 - Date.now()) / 1000);
      if (sisa <= 0) return { success: false, message: "Waktu pengerjaan tugas ini sudah habis." };
    }
    // tanpa kunci dan tanpa pembahasan
    const soal = JSON.parse(v[7]).map(s => ({ tipe: s.tipe, pertanyaan: s.pertanyaan, opsi: s.opsi || [], kiri: s.kiri || [], kanan: s.kanan || [] }));
    return ok_({ data: { id: id, judul: String(v[4]), mapel: String(v[3]), instruksi: String(v[6]), deadline: toIso_(v[8]), durasi: durasi, sisaDetik: sisa, soal: soal } });
  } catch (e) { return fail_(e); }
}

// jawaban: array sepanjang soal.
//   pg "A".."D" | pgk "A,C" | jodoh "B,A,C,D" ("-" = belum dipilih) | bs "B"/"S" | isian teks | uraian teks
// Nilai TIDAK dikembalikan ke siswa. Guru melihat nilai di Hasil dan merilisnya lewat setNilaiRilis.
function submitJawaban(token, id, jawaban) {
  try {
    const me = auth_(token, STUDENT);
    const ts = tugasSheet_(), trow = findRow_(ts, 1, id, true);
    if (!trow) return { success: false, message: "Tugas tidak ditemukan." };
    const v = ts.getRange(trow, 1, 1, 15).getValues()[0];
    if (String(v[9]) !== "Aktif") return { success: false, message: "Tugas sudah ditutup." };
    if (!cocokKelas_(v[2], me.kelas)) return { success: false, message: "Tugas ini bukan untuk kelasmu." };
    if (deadlinePassed_(v[8])) return { success: false, message: "Batas waktu pengumpulan sudah lewat." };
    const kode = String(v[13] || ""), durasi = Number(v[14]) || 0;
    if (kode || durasi) {
      const mulai = Number(CacheService.getScriptCache().get("mulai_" + me.username + "_" + id) || 0);
      if (!mulai && kode) return { success: false, message: "Masukkan token dan mulai tugas dari awal." };
      if (mulai && durasi && Date.now() > mulai + durasi * 60000 + 120000) return { success: false, message: "Waktu pengerjaan sudah habis, jawaban tidak diterima." };
    }

    // penilaian di luar lock supaya lock singkat saat banyak siswa mengumpulkan bersamaan
    const soal = JSON.parse(v[7]);
    const ans = (jawaban || []).slice(0, soal.length).map(x => String(x == null ? "" : x).substring(0, 1000));
    const sc = skorTugas_(soal, ans, []);
    const waktu = new Date().toISOString();

    return withLock_(function () {
      const done = doneMap_(me.username);
      if (done[id]) return { success: false, message: "Kamu sudah mengumpulkan tugas ini." };
      const js = jawabSheet_();
      js.appendRow([Utilities.getUuid(), id, me.username, me.nama, me.kelas, waktu, JSON.stringify(ans), sc.skor, sc.poin, sc.total, ""]);
      done[id] = { skor: sc.skor, waktu: waktu, row: js.getLastRow() };
      cachePutJson_("done_" + me.username, done, SESSION_TTL);
      const r = findRow_(ts, 1, id, true);
      if (r) ts.getRange(r, 12).setValue((Number(ts.getRange(r, 12).getValue()) || 0) + 1);
      return ok_({ message: "Tugas berhasil dikumpulkan! Nilai akan diumumkan oleh gurumu." });
    }, 45000);
  } catch (e) { return fail_(e); }
}

// Hasil + pembahasan hanya tersedia setelah guru merilis nilai tugas tersebut.
function getHasilSaya(token, id) {
  try {
    const me = auth_(token, STUDENT);
    const ts = tugasSheet_(), trow = findRow_(ts, 1, id, true);
    if (!trow) return { success: false, message: "Tugas sudah tidak tersedia." };
    const tv = ts.getRange(trow, 1, 1, 13).getValues()[0];
    if (String(tv[12]).toUpperCase() !== "TRUE") return { success: false, message: "Nilai belum dirilis. Tunggu pengumuman dari gurumu." };

    const js = jawabSheet_();
    let d = doneMap_(me.username)[id];
    const baca = function (row) {
      if (!row || row < 2 || row > js.getLastRow()) return null;
      const r = js.getRange(row, 1, 1, 11).getValues()[0];
      return (r[1] === id && r[2] === me.username) ? r : null;
    };
    let r = baca(d && d.row);
    if (!r) { d = doneMap_(me.username, true)[id]; r = baca(d && d.row); }   // nomor baris bisa bergeser setelah penghapusan
    if (!r) return { success: false, message: "Belum ada hasil untuk tugas ini." };
    const soal = JSON.parse(tv[7]), ans = JSON.parse(r[6] || "[]"), kor = korArr_(r[10]);
    const sc = skorTugas_(soal, ans, kor);
    return ok_({ hasil: { skor: sc.skor, benar: sc.poin, totalPG: sc.total, adaUraian: soal.some(s => TIPE_GURU.indexOf(s.tipe) >= 0), review: reviewSoal_(soal, ans, kor) }, judul: String(tv[4]) });
  } catch (e) { return fail_(e); }
}

// ==========================================
// STATISTIK DASHBOARD (ADMIN & SUPERADMIN) - di-cache 2 menit
// ==========================================
function getStatsAdmin(token) {
  try {
    auth_(token, STAFF);
    let s = cacheGetJson_("stats");
    if (!s) {
      const roles = rd_(usersSheet_(), 4, 4);
      const perKelas = {};
      materiMeta_().forEach(m => perKelas[m.kelas] = (perKelas[m.kelas] || 0) + 1);
      s = {
        siswa: roles.filter(r => String(r[0]) === ROLE_USER).length,
        admin: roles.filter(r => String(r[0]) !== ROLE_USER).length,
        tugas: Math.max(0, tugasSheet_().getLastRow() - 1),
        pengumpulan: Math.max(0, jawabSheet_().getLastRow() - 1),
        materi: materiMeta_().length, perKelas: perKelas
      };
      cachePutJson_("stats", s, 120);
    }
    return ok_(Object.assign({}, s));
  } catch (e) { return fail_(e); }
}

// ==========================================
// PEMELIHARAAN DATA (SUPERADMIN)
// ==========================================
function getPemeliharaan(token) {
  try {
    auth_(token, SUPER);
    const ss = ss_();
    let total = 0;
    const sheets = ss.getSheets().map(sh => {
      const c = sh.getMaxRows() * sh.getMaxColumns(); total += c;
      return { nama: sh.getName(), baris: Math.max(0, sh.getLastRow() - 1), sel: c };
    });
    const arsipId = PropertiesService.getScriptProperties().getProperty("ARSIP_ID") || "";
    return ok_({ sheets: sheets, totalSel: total, batas: 10000000, arsipUrl: arsipId ? "https://docs.google.com/spreadsheets/d/" + arsipId : "" });
  } catch (e) { return fail_(e); }
}

// Pindahkan tugas berstatus "Tutup" yang lebih tua dari N hari (beserta jawabannya) ke spreadsheet arsip.
// Aman dijalankan berulang; satu proses memindahkan paling banyak ~3000 jawaban.
function arsipkanTugasLama(token, hari) {
  try { auth_(token, SUPER); return arsipkanCore_(hari); } catch (e) { return fail_(e); }
}

// Dipanggil pemicu harian (hanya jika dipicu oleh Apps Script, bukan dari browser).
function arsipOtomatis(e) {
  if (!e || !e.triggerUid) return;
  try { arsipkanCore_(180); } catch (x) { Logger.log("arsipOtomatis gagal: " + x); }
}

function arsipkanCore_(hari) {
  hari = Math.max(30, int_(hari, 180));
  {
    return withLock_(function () {
      const t0 = Date.now(), batas = t0 - hari * 86400000;
      const ts = tugasSheet_(), js = jawabSheet_(), tl = ts.getLastRow();
      if (tl < 2) return ok_({ message: "Belum ada tugas." });
      const ab = ts.getRange(2, 1, tl - 1, 2).getValues();
      const st = ts.getRange(2, 10, tl - 1, 1).getValues();
      const jb = rd_(js, 2, 2);
      const cnt = {}; jb.forEach(r => cnt[r[0]] = (cnt[r[0]] || 0) + 1);

      const pilih = {}; const tugasRows = []; let jumlahJawab = 0, nTugas = 0;
      for (let i = 0; i < ab.length; i++) {
        if (String(st[i][0]) !== "Tutup") continue;
        const d = new Date(ab[i][1]);
        if (isNaN(d.getTime()) || d.getTime() >= batas) continue;
        const c = cnt[ab[i][0]] || 0;
        if (nTugas > 0 && jumlahJawab + c > 3000) break;
        pilih[ab[i][0]] = true; tugasRows.push(i + 2); jumlahJawab += c; nTugas++;
      }
      if (!nTugas) return ok_({ message: "Tidak ada tugas berstatus Tutup yang lebih tua dari " + hari + " hari." });

      // spreadsheet arsip
      const p = PropertiesService.getScriptProperties();
      let arId = p.getProperty("ARSIP_ID"), ar;
      if (arId) { try { ar = SpreadsheetApp.openById(arId); } catch (e) { ar = null; } }
      if (!ar) { ar = SpreadsheetApp.create("Arsip Siguruku"); p.setProperty("ARSIP_ID", ar.getId()); }
      const arSheet = function (name, headers) {
        let s = ar.getSheetByName(name);
        if (!s) { s = ar.insertSheet(name); s.appendRow(headers); s.setFrozenRows(1); }
        return s;
      };
      const aT = arSheet("Tugas_Arsip", ["ID", "Tanggal", "Kelas", "Mapel", "Judul", "MateriId", "Instruksi", "Soal", "Deadline", "Status", "JumlahSoal", "Terkumpul", "NilaiRilis", "Token", "Durasi"]);
      const aJ = arSheet("Jawaban_Arsip", ["ID", "TugasId", "Username", "Nama", "Kelas", "Waktu", "Jawaban", "Skor", "Benar", "TotalPG", "Koreksi"]);

      // jawaban dahulu
      const jRows = []; jb.forEach((r, i) => { if (pilih[r[0]]) jRows.push(i + 2); });
      const asc = jRows.slice().sort((a, b) => a - b);
      let i = 0;
      while (i < asc.length) {
        let j = i; while (j + 1 < asc.length && asc[j + 1] === asc[j] + 1) j++;
        appendRows_(aJ, js.getRange(asc[i], 1, j - i + 1, 11).getValues());
        i = j + 1;
      }
      // tugas
      tugasRows.forEach(r => appendRows_(aT, ts.getRange(r, 1, 1, 15).getValues()));

      deleteRowsBottomUp_(js, jRows, t0);
      deleteRowsBottomUp_(ts, tugasRows, t0);
      bumpVer_("tugas"); cacheRemove_("stats");
      return ok_({ message: nTugas + " tugas dan " + jumlahJawab + " jawaban dipindah ke arsip. Jalankan lagi jika masih ada yang lama.", arsipUrl: ar.getUrl() });
    }, 60000);
  }
}

// ==========================================
// BUKU DITAMBAH LANGSUNG DI SHEET (karena menu "Tulis Materi" dihapus)
// ==========================================
// Isi ID & Tanggal otomatis untuk baris buku yang kelas/judulnya sudah terisi. Aman dijalankan berulang.
function lengkapiIdMateri_() {
  const sh = initSheet(), last = sh.getLastRow();
  if (last < 2) return 0;
  const v = sh.getRange(2, 1, last - 1, 7).getValues();
  let n = 0;
  v.forEach((r, i) => {
    if (r[2] === "" && r[6] === "") return;
    const id = r[0] || Utilities.getUuid(), tgl = r[1] || new Date().toISOString();
    if (!r[0] || !r[1]) { sh.getRange(i + 2, 1, 1, 2).setValues([[id, tgl]]); n++; }
  });
  return n;
}

// Superadmin: dipanggil dari tombol "Segarkan data" di Pengaturan setelah mengedit sheet langsung.
function segarkanData(token) {
  try {
    auth_(token, SUPER);
    return withLock_(function () {
      const n = lengkapiIdMateri_();
      ["materi", "tugas", "rombel"].forEach(bumpVer_);
      cacheRemove_("stats");
      return ok_({ message: "Data disegarkan." + (n ? " " + n + " baris buku dilengkapi ID-nya." : "") });
    }, 30000);
  } catch (e) { return fail_(e); }
}

// Pemicu: setiap sheet diedit tangan, cache diperbarui otomatis agar buku baru langsung muncul.
function onEditInstalasi(e) {
  if (!e || !e.range) return;
  try {
    const name = e.range.getSheet().getName();
    withLock_(function () {
      if (name === SHEET_NAME) { lengkapiIdMateri_(); bumpVer_("materi"); cacheRemove_("stats"); }
      else if (name === SHEET_TUGAS) bumpVer_("tugas");
      else if (name === SHEET_ROMBEL) bumpVer_("rombel");
    }, 15000);
  } catch (x) { Logger.log("onEditInstalasi: " + x); }
}

// Jalankan sekali dari editor (izinkanAkses sudah memanggilnya). Aman diulang: pemicu lama diganti.
function pasangTriggerOtomatis() {
  ScriptApp.getProjectTriggers().forEach(t => {
    const f = t.getHandlerFunction();
    if (f === "onEditInstalasi" || f === "arsipOtomatis") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("onEditInstalasi").forSpreadsheet(ss_()).onEdit().create();
  ScriptApp.newTrigger("arsipOtomatis").timeBased().everyDays(1).atHour(2).create();
}


// ==========================================
// VERSI 4: TOKEN, DURASI, BAHASA, EKSPOR SISWA
// ==========================================
function buatToken_() {
  const c = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"; let s = "";
  for (let i = 0; i < 6; i++) s += c.charAt(Math.floor(Math.random() * c.length));
  return s;
}

// data: {pakaiToken, token, durasi}. Hasil: {token, durasi} atau {error}
function aksesValid_(data) {
  const durasi = clampInt_(data.durasi, 0, 300);   // menit, 0 = tanpa batas
  let token = "";
  if (data.pakaiToken) {
    token = String(data.token || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
    if (!token) token = buatToken_();
    if (token.length < 4 || token.length > 12) return { error: "Token harus 4 sampai 12 huruf/angka." };
  }
  return { token: token, durasi: durasi };
}

// Ubah token/durasi tugas yang sudah ada (dipakai saat merilis tugas Draft atau mengatur ulang akses).
function setTugasAkses(token, id, data) {
  try {
    auth_(token, STAFF);
    const ak = aksesValid_(data || {});
    if (ak.error) return { success: false, message: ak.error };
    return withLock_(function () {
      const sh = tugasSheet_(), row = findRow_(sh, 1, id, true);
      if (!row) return { success: false, message: "Tugas tidak ditemukan." };
      sh.getRange(row, 14, 1, 2).setValues([[ak.token, ak.durasi]]); bumpVer_("tugas");
      return ok_({ message: "Pengaturan rilis disimpan.", token: ak.token });
    });
  } catch (e) { return fail_(e); }
}

function bahasaInstruksi_(b) {
  const kunci = " Field 'kunci' untuk soal pg/pgk tetap huruf Latin A, B, C, D, E dan untuk soal bs tetap kata Benar atau Salah. Field 'pembahasan' ditulis dalam bahasa Indonesia.";
  switch (String(b || "id")) {
    case "ar": return "Tulis pertanyaan, opsi, pasangan, dan jawaban dalam bahasa Arab fusha yang sederhana sesuai usia MI, beri harakat (syakl) pada kata penting." + kunci;
    case "en": return "Write questions, options, pairs, and answers in simple English suitable for elementary students." + kunci;
    case "dwi-ar": return "Tulis pertanyaan dan opsi dalam bahasa Arab fusha sederhana berharakat, lalu tambahkan terjemah Indonesia singkat dalam tanda kurung di akhir pertanyaan." + kunci;
    case "dwi-en": return "Tulis pertanyaan dan opsi dalam bahasa Inggris sederhana, lalu tambahkan terjemah Indonesia singkat dalam tanda kurung di akhir pertanyaan." + kunci;
    default: return "Gunakan bahasa Indonesia yang sederhana dan sesuai usia anak SD/MI.";
  }
}
function bahasaBuku_(b) {
  switch (String(b || "id")) {
    case "ar": case "dwi-ar": return "Tulis isi buku dalam bahasa Arab fusha sederhana berharakat; setiap istilah glosarium diberi arti dalam bahasa Indonesia.";
    case "en": case "dwi-en": return "Tulis isi buku dalam bahasa Inggris sederhana; setiap istilah glosarium diberi arti dalam bahasa Indonesia.";
    default: return "Gunakan bahasa Indonesia yang sederhana, jelas, dan ramah anak sesuai kelasnya.";
  }
}

// Seluruh siswa (tanpa password: password disimpan ter-hash). kelas: 'all' | '5' | '5A'
function exportSiswa(token, kelas) {
  try {
    auth_(token, SUPER);
    const f = String(kelas || "all");
    const list = rd_(usersSheet_(), 1, 5).filter(r => String(r[3]) === ROLE_USER && cocokKelasSiswa_(r[4], f))
      .map(r => ({ nama: String(r[2]), username: String(r[0]), kelas: String(r[4] || "") }));
    list.sort((a, b) => a.kelas.localeCompare(b.kelas, "id", { numeric: true }) || a.nama.localeCompare(b.nama, "id"));
    return ok_({ data: list });
  } catch (e) { return fail_(e, { data: [] }); }
}


// ==========================================
// VERSI 5: BRANDING (LOGO APLIKASI)
// ==========================================
// Logo disimpan sebagai data URL gambar (sudah diperkecil di browser) dalam potongan 8000 karakter di Script Properties.
const LOGO_MAX = 200000;

// Publik (dipanggil sebelum login agar logo tampil di halaman masuk). Tidak memuat data sensitif.
function getBranding() {
  try {
    let hit = cacheGetJson_("logo");
    if (!hit) {
      const p = PropertiesService.getScriptProperties();
      const n = Number(p.getProperty("LOGO_N") || 0);
      let d = "";
      for (let i = 0; i < n; i++) d += p.getProperty("LOGO_" + i) || "";
      hit = { d: d };
      cachePutJson_("logo", hit, 21600);
    }
    return { success: true, nama: "Siguruku", logo: hit.d || "" };
  } catch (e) { return { success: false, nama: "Siguruku", logo: "" }; }
}

function hapusLogo_() {
  const p = PropertiesService.getScriptProperties();
  const n = Number(p.getProperty("LOGO_N") || 0);
  for (let i = 0; i < n; i++) p.deleteProperty("LOGO_" + i);
  p.deleteProperty("LOGO_N");
  cacheRemove_("logo");
}

function saveLogo(token, dataUrl) {
  try {
    auth_(token, SUPER);
    dataUrl = String(dataUrl || "");
    if (!/^data:image\/(png|jpeg|webp|gif);base64,[A-Za-z0-9+\/=]+$/.test(dataUrl)) return { success: false, message: "Format logo tidak valid. Gunakan PNG, JPG, atau WebP." };
    if (dataUrl.length > LOGO_MAX) return { success: false, message: "Logo terlalu besar. Pilih gambar yang lebih kecil." };
    return withLock_(function () {
      hapusLogo_();
      const p = PropertiesService.getScriptProperties(), SZ = 8000, n = Math.ceil(dataUrl.length / SZ), put = {};
      for (let i = 0; i < n; i++) put["LOGO_" + i] = dataUrl.substr(i * SZ, SZ);
      put["LOGO_N"] = String(n);
      p.setProperties(put, false);
      return ok_({ message: "Logo disimpan." });
    }, 30000);
  } catch (e) { return fail_(e); }
}

function deleteLogo(token) {
  try {
    auth_(token, SUPER);
    return withLock_(function () { hapusLogo_(); return ok_({ message: "Logo dihapus." }); }, 30000);
  } catch (e) { return fail_(e); }
}

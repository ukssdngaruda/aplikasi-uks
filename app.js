var APP_CONFIG = null;
var CURRENT_MODULE = 'dashboard';
var CURRENT_DATA = [];
var EDIT_ID = null;
var SESSION_TOKEN = null;
var CURRENT_USER = null;
var SHOW_WELCOME = false; // true hanya tepat setelah login baru (bukan saat sesi dipulihkan) -> tampilkan popup sambutan
var SISWA_LIST = []; // cache daftar siswa (nisn, nama, kelas, jk) untuk autocomplete di modul lain
var DOCIL_LIST = []; // cache daftar anggota Dokter Cilik (idDocil, nama, kelas, jabatan) untuk autocomplete field "Nama Petugas Docil" di modul Log Kegiatan Docil
var KADER_LIST = []; // (BARU) cache daftar Tim/Kader UKS (idTim, nama, jabatan) untuk autocomplete field pj/petugas/pembina di modul lain
var DELETE_CONFIRM_ACTION = null; // (BARU) fungsi yang dijalankan kalau tombol "Ya, Hapus" di modal konfirmasi hapus diklik
var SELECTED_IDS = []; // (BARU) daftar ID baris yang sedang diceklis untuk fitur "Hapus Terpilih"

// ==================== PENGELOMPOKAN MODUL PER TUGAS/ROLE (BARU) ====================
// Dipakai untuk mengelompokkan menu di sidebar seperti "folder" per tugas, supaya
// Admin (yang aksesnya ke SEMUA modul lewat ROLE_ACCESS.Admin = 'ALL' di Code.gs)
// lebih mudah menemukan modul mana milik tugas siapa. Urutan & isi array 'modules'
// di sini SENGAJA disamakan persis dengan ROLE_ACCESS di Code.gs supaya konsisten —
// kalau suatu saat ROLE_ACCESS diubah/ditambah modul baru, sesuaikan juga di sini.
var MODULE_GROUPS = [
  { role: 'Ketua', label: 'Ketua', icon: '🗂️', modules: ['AgendaKegiatan', 'JadwalKegiatan', 'KaderLog'] },
  { role: 'Sekretaris', label: 'Sekretaris', icon: '🗃️', modules: ['Siswa', 'KaderAnggota', 'InventarisSarana', 'InventarisObat'] },
  { role: 'Bendahara', label: 'Bendahara', icon: '💰', modules: ['RKAS', 'BukuKas'] },
  { role: 'DokterPembina', label: 'Dokter Pembina', icon: '🩺', modules: ['PemeriksaanBerkala', 'BBTB', 'GigiMulut', 'Imunisasi'] },
  { role: 'PetugasUKS', label: 'Petugas UKS', icon: '🏥', modules: ['Kunjungan', 'SuratRujukanSakit', 'MutasiObat', 'Kecelakaan', 'Rujukan'] },
  { role: 'KaderPMR', label: 'Kader PMR / Docil', icon: '🧑\u200d⚕️', modules: ['DocilAnggota', 'DocilPiket', 'DocilLog'] }
];

// ==================== IMPORT EXCEL/CSV (BARU) ====================
// Modul mana saja yang menampilkan tombol "Import Excel" di topbar -> cukup
// tambahkan `NamaModul: true` di sini (harus sama persis dengan key modul di
// SHEET_CONFIG, Code.gs). Pemetaan header kolom file Excel/CSV ke key kolom
// modul dilakukan OTOMATIS berdasarkan label kolom yang sudah ada di
// SHEET_CONFIG (lihat matchHeadersToColumns_), jadi TIDAK perlu menulis
// pemetaan manual per kolom setiap kali menambah modul baru ke sini.
var IMPORT_CONFIG = {
  Siswa: true,
  // Label kolom DocilPiket sengaja dibuat sama persis dengan header Excel
  // "Jadwal Piket Mingguan" supaya import di sini bisa langsung mencocokkan
  // kolom otomatis. CATATAN: karena file Excel aslinya punya 2 baris judul
  // ("JADWAL PIKET RUTIN..." dan "UPTD SPF...") sebelum baris header, baris
  // judul itu harus dihapus dulu di Excel (baris pertama file harus persis
  // baris header: Hari, Jam Operasional, dst) sebelum diimpor lewat sini —
  // proses import selalu membaca baris PERTAMA file sebagai baris header.
  DocilPiket: true
};

// Menyimpan sementara hasil parsing file (untuk modul apa, hasil pemetaan
// header, dan data yang sudah siap dikirim ke importRecords) selama modal
// Import Excel terbuka.
var IMPORT_STATE = { moduleKey: null, mappedData: [], headerMapping: null };

document.addEventListener('DOMContentLoaded', function () {
  // Sidebar versi HP: tombol ☰ buka/tutup, ada tombol ✕ di dalam sidebar,
  // klik area gelap di luar sidebar & tombol Esc juga menutupnya.
  setupSidebarMobile();
  document.getElementById('modalClose').onclick = closeModal;
  document.getElementById('cancelBtn').onclick = closeModal;
  document.getElementById('saveBtn').onclick = saveRecord;
  document.getElementById('publicDashboardBtn').onclick = openPublicDashboard;
  document.getElementById('logoutBtn').onclick = handleLogout;
  document.getElementById('loginForm').addEventListener('submit', handleLogin);

  // (BARU) Klik di area gelap DI LUAR kotak modal langsung menutup/membatalkan
  // modal yang sedang terbuka — berlaku untuk modal Import Excel dan modal
  // konfirmasi Hapus. Dicek `e.target === this` supaya klik DI DALAM kotak
  // modal (mis. pada input/tombol) tidak ikut menutup.
  //
  // CATATAN: modal Tambah/Edit Data (#modalOverlay) SENGAJA TIDAK diberi
  // perilaku ini — klik di luar area form (area gelap) TIDAK akan menutup
  // form, supaya isian yang sedang diketik tidak hilang/kebatal secara tidak
  // sengaja hanya karena klik meleset ke luar kotak modal. Modal ini hanya
  // ditutup lewat tombol "✕", tombol "Batal", atau setelah berhasil Simpan.
  document.getElementById('importModalOverlay').addEventListener('click', function (e) {
    if (e.target === this) closeImportModal();
  });
  document.getElementById('deleteConfirmOverlay').addEventListener('click', function (e) {
    if (e.target === this) closeDeleteConfirm();
  });

  // (BARU) Wiring modal konfirmasi hapus
  document.getElementById('deleteConfirmClose').onclick = closeDeleteConfirm;
  document.getElementById('deleteConfirmCancel').onclick = closeDeleteConfirm;
  document.getElementById('deleteConfirmOk').onclick = function () {
    const aksi = DELETE_CONFIRM_ACTION;
    closeDeleteConfirm();
    if (aksi) aksi();
  };

  // Tombol "Cetak LPJ" (Laporan Pertanggungjawaban) — dibuat lewat JS (bukan diedit
  // langsung di Index.html) supaya file ini tetap satu-satunya file yang perlu diganti.
  // Disisipkan tepat di sebelah kiri tombol "+ Tambah Data" dan hanya tampil untuk
  // Admin/Bendahara saat sedang membuka modul RKAS atau Buku Kas (lihat navigate()).
  const addBtnEl = document.getElementById('addBtn');
  if (addBtnEl && addBtnEl.parentNode) {
    const lpjBtn = document.createElement('button');
    lpjBtn.id = 'lpjBtn';
    lpjBtn.type = 'button';
    lpjBtn.className = 'btn secondary';
    lpjBtn.style.marginRight = '8px';
    lpjBtn.style.display = 'none';
    lpjBtn.textContent = '🖨️ Cetak LPJ';
    lpjBtn.onclick = cetakLPJ;
    addBtnEl.parentNode.insertBefore(lpjBtn, addBtnEl);

    // Tombol cetak Laporan/Matriks generik (BARU) — satu tombol yang dipakai
    // bersama oleh SEMUA modul yang terdaftar di PRINT_CONFIG dengan type
    // 'laporan' atau 'matrix'. Teks tombol & fungsi yang dipanggil menyesuaikan
    // modul yang sedang dibuka (lihat navigate() dan cetakLaporanAtauMatrix()).
    const reportBtn = document.createElement('button');
    reportBtn.id = 'reportBtn';
    reportBtn.type = 'button';
    reportBtn.className = 'btn secondary';
    reportBtn.style.marginRight = '8px';
    reportBtn.style.display = 'none';
    reportBtn.onclick = cetakLaporanAtauMatrix;
    addBtnEl.parentNode.insertBefore(reportBtn, addBtnEl);

    // Tombol "Import Excel" (BARU) — hanya tampil untuk modul yang terdaftar
    // di IMPORT_CONFIG (lihat navigate()). Membuka modal import yang sudah ada
    // markupnya di Index.html (#importModalOverlay).
    const importBtn = document.createElement('button');
    importBtn.id = 'importBtn';
    importBtn.type = 'button';
    importBtn.className = 'btn secondary';
    importBtn.style.marginRight = '8px';
    importBtn.style.display = 'none';
    importBtn.textContent = '📥 Import Excel';
    importBtn.onclick = openImportModal;
    addBtnEl.parentNode.insertBefore(importBtn, addBtnEl);
  }

  // Wiring modal Import Excel (markup-nya sudah ada di Index.html).
  document.getElementById('importModalClose').onclick = closeImportModal;
  document.getElementById('importCancelBtn').onclick = closeImportModal;
  document.getElementById('importFileInput').onchange = handleImportFileSelected;
  document.getElementById('importConfirmBtn').onclick = confirmImport;

  const togglePwBtn = document.getElementById('togglePassword');
  if (togglePwBtn) {
    togglePwBtn.onclick = function () {
      const pwInput = document.getElementById('loginPassword');
      const showing = pwInput.type === 'text';
      pwInput.type = showing ? 'password' : 'text';
      togglePwBtn.textContent = showing ? '👁️' : '🙈';
      togglePwBtn.setAttribute('aria-label', showing ? 'Lihat password' : 'Sembunyikan password');
    };
  }

  const savedToken = sessionStorage.getItem('uksToken');
  if (savedToken) {
    SESSION_TOKEN = savedToken;
    google.script.run.withSuccessHandler(onConfigLoaded).withFailureHandler(function (e) {
      console.error('Sesi lama tidak valid:', e);
      sessionStorage.removeItem('uksToken');
      SESSION_TOKEN = null;
      showLogin();
    }).getAppConfig(SESSION_TOKEN);
  } else {
    showLogin();
  }
});

// ==================== LOADING OVERLAY GLOBAL (BARU) ====================
// Ditampilkan di tengah layar (menutupi seluruh halaman) selama proses yang
// memanggil server sedang berjalan (muat data, simpan, hapus, dsb), supaya
// pengguna tahu aplikasi sedang bekerja dan tidak mengklik berkali-kali.
function showLoading(msg) {
  document.getElementById('loadingText').textContent = msg || 'Memproses...';
  document.getElementById('loadingOverlay').classList.add('show');
}
function hideLoading() {
  document.getElementById('loadingOverlay').classList.remove('show');
}

// ==================== MODAL KONFIRMASI HAPUS (BARU) ====================
function openDeleteConfirm(judul, detailHtml, onConfirm) {
  document.getElementById('deleteConfirmDetail').innerHTML =
    '<div style="font-weight:700;margin-bottom:6px;">' + judul + '</div>' + detailHtml;
  DELETE_CONFIRM_ACTION = onConfirm;
  document.getElementById('deleteConfirmOverlay').classList.add('show');
}
function closeDeleteConfirm() {
  document.getElementById('deleteConfirmOverlay').classList.remove('show');
  DELETE_CONFIRM_ACTION = null;
}

// ==================== LOGIN (tanpa pindah halaman / document.write) ====================
function showLogin() {
  document.getElementById('app').style.display = 'none';
  document.getElementById('loginOverlay').style.display = 'flex';
}

function handleLogin(ev) {
  ev.preventDefault();
  const username = document.getElementById('loginUsername').value.trim();
  const password = document.getElementById('loginPassword').value;
  // Peran yang dipilih di layar login (radio "loginRole" di Login.html)
  const roleEl = document.querySelector('input[name="loginRole"]:checked');
  const peranDipilih = roleEl ? roleEl.value : '';
  const btn = document.getElementById('loginBtn');
  const errBox = document.getElementById('loginError');
  errBox.style.display = 'none';
  btn.disabled = true;
  btn.innerHTML = '<span class="spinner"></span>';

  google.script.run.withSuccessHandler(function (result) {
    btn.disabled = false;
    btn.innerHTML = 'Masuk';
    SESSION_TOKEN = result.token;
    SHOW_WELCOME = true; // popup sambutan sesuai peran tampil sekali setelah login
    sessionStorage.setItem('uksToken', SESSION_TOKEN);
    // Langsung minta konfigurasi & tampilkan app di halaman yang SAMA — tidak ada
    // document.write / navigasi ulang, sehingga google.script.run tetap tersambung.
    google.script.run.withSuccessHandler(onConfigLoaded).withFailureHandler(function (e) {
      onError(e, 'dashboard');
    }).getAppConfig(SESSION_TOKEN);
  }).withFailureHandler(function (e) {
    btn.disabled = false;
    btn.innerHTML = 'Masuk';
    errBox.textContent = e.message || e;
    errBox.style.display = 'block';
  }).login(username, password, peranDipilih);
}

// FIX: sebelumnya fungsi ini memanggil location.reload(), yang menyebabkan LAYAR PUTIH
// kosong saat tombol "Keluar" ditekan. Penyebabnya: Web App Apps Script berjalan di
// dalam iframe ber-sandbox (googleusercontent.com) yang tidak selalu diberi izin
// 'allow-same-origin', sehingga location.reload() gagal memuat ulang kontennya.
// Solusinya: TIDAK reload/navigasi sama sekali — cukup reset semua state di
// JavaScript lalu tampilkan kembali layar login di halaman yang sama.
function handleLogout() {
  if (SESSION_TOKEN) {
    google.script.run.withFailureHandler(function (e) {
      console.error('Gagal logout di server (diabaikan, sesi tetap dihapus di client):', e);
    }).logout(SESSION_TOKEN);
  }

  sessionStorage.removeItem('uksToken');
  SESSION_TOKEN = null;
  CURRENT_USER = null;
  APP_CONFIG = null;
  CURRENT_DATA = [];
  EDIT_ID = null;
  SISWA_LIST = [];

  // Reset form login (kalau ada isian tersisa) & sembunyikan pesan error lama
  const loginForm = document.getElementById('loginForm');
  if (loginForm) loginForm.reset();
  const errBox = document.getElementById('loginError');
  if (errBox) errBox.style.display = 'none';

  // Sembunyikan app, tampilkan login — TANPA reload/navigasi halaman
  closeWelcome();
  closeUserModal_();
  closeSidebar();
  document.getElementById('app').style.display = 'none';
  showLogin();
}

function onConfigLoaded(config) {
  APP_CONFIG = config;
  CURRENT_USER = config.user;
  document.getElementById('loginOverlay').style.display = 'none';
  document.getElementById('app').style.display = 'flex';
  document.getElementById('brandSchool').textContent = config.sekolah.namaSekolah + ' — ' + config.sekolah.tahunAjaran;
  // FIX: baris yang dulu ada di sini — document.getElementById('loginSchoolName').textContent = ...
  // — dihapus. Elemen #loginSchoolName ada di layar LOGIN (Login.html) dan sudah
  // di-hardcode "Aplikasi UKS" di sana sebagai branding aplikasi yang tetap/konsisten.
  // Ia TIDAK BOLEH ditimpa dengan nama sekolah di sini, karena itu menyebabkan judul
  // layar login berubah-ubah / tidak konsisten. Nama sekolah cukup tampil di sidebar
  // lewat #brandSchool di baris atas.
  document.getElementById('userInfo').textContent = config.user.nama + ' (' + config.user.role + ')';
  buildSidebar();
  appendPengaturanNav_(); // menu Pengaturan (Akun Saya untuk semua peran; Manajemen Login khusus Admin)
  loadSiswaList(); // preload daftar siswa untuk autocomplete NISN di form modul lain
  loadDocilList(); // preload daftar anggota Docil untuk autocomplete "Nama Petugas Docil" di form Log Kegiatan Docil
  loadKaderList(); // (BARU) preload daftar Tim/Kader UKS untuk autocomplete pj/petugas/pembina di modul lain
  navigate('dashboard');

  if (SHOW_WELCOME) {
    SHOW_WELCOME = false;
    showWelcome(config);
  }
}

// ==================== POPUP SAMBUTAN SESUAI PERAN (BARU) ====================
// Tampil sekali setelah login. Isi sambutan per peran ada di WELCOME_ROLES;
// daftar modul diambil langsung dari hak akses akun (config.modules), jadi
// otomatis mengikuti ROLE_ACCESS di Code.gs.
var WELCOME_ROLES = {
  Admin: {
    icon: '🛠️', title: 'Administrator',
    desc: 'Anda memiliki akses ke seluruh modul Aplikasi UKS. Pantau kelengkapan data dari semua petugas dan pastikan pencatatan berjalan rapi.'
  },
  Pembina: {
    icon: '🎓', title: 'Pembina UKS',
    desc: 'Pantau perkembangan kegiatan dan kesehatan warga sekolah lewat Dashboard. Anda juga dapat membagikan Dashboard Publik kepada pihak yang memerlukan.'
  },
  Ketua: {
    icon: '⭐', title: 'Ketua UKS',
    desc: 'Atur agenda dan jadwal kegiatan UKS, serta pantau log kegiatan kader agar program berjalan sesuai rencana.'
  },
  Sekretaris: {
    icon: '📝', title: 'Sekretaris UKS',
    desc: 'Kelola data siswa, keanggotaan kader, serta inventaris sarana dan obat agar administrasi UKS selalu terbarui.'
  },
  Bendahara: {
    icon: '💰', title: 'Bendahara UKS',
    desc: 'Catat rencana anggaran (RKAS), kas masuk dan keluar, lalu cetak Laporan Pertanggungjawaban (LPJ) bila diperlukan.'
  },
  DokterPembina: {
    icon: '👨‍⚕️', title: 'Dokter Pembina',
    desc: 'Catat hasil pemeriksaan berkala, pengukuran BB/TB, kesehatan gigi dan mulut, serta imunisasi siswa.'
  },
  PetugasUKS: {
    icon: '🩺', title: 'Petugas UKS',
    desc: 'Catat kunjungan ruang UKS, pemakaian obat, kejadian kecelakaan, serta surat rujukan bagi siswa yang memerlukan.'
  },
  KaderPMR: {
    icon: '➕', title: 'Kader PMR',
    desc: 'Catat keanggotaan, jadwal piket, dan log kegiatan Dokter Cilik. Terima kasih sudah ikut menjaga kesehatan teman-teman.'
  }
};

function welcomeEsc_(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function welcomeGreeting_() {
  const h = new Date().getHours();
  if (h < 11) return 'Selamat pagi';
  if (h < 15) return 'Selamat siang';
  if (h < 18) return 'Selamat sore';
  return 'Selamat malam';
}

function showWelcome(config) {
  const role = config.user.role;
  const info = WELCOME_ROLES[role] || {
    icon: '👋', title: role, desc: 'Silakan pilih modul di menu untuk mulai bekerja.'
  };

  let modulHtml;
  if (role === 'Admin') {
    modulHtml = '<div class="welcome-note">Akses penuh ke semua modul (' +
      Object.keys(config.modules || {}).length + ' modul).</div>';
  } else {
    const keys = Object.keys(config.modules || {});
    if (keys.length === 0) {
      modulHtml = '<div class="welcome-note">Gunakan menu <b>Dashboard</b> dan <b>Dashboard Publik</b> untuk memantau data.</div>';
    } else {
      modulHtml = '<div class="welcome-label">Modul yang dapat Anda kelola</div><div class="welcome-chips">' +
        keys.map(function (k) {
          const m = config.modules[k];
          return '<span class="welcome-chip">' + welcomeEsc_(m.icon || '') + ' ' + welcomeEsc_(m.label) + '</span>';
        }).join('') + '</div>';
    }
  }

  let ov = document.getElementById('welcomeOverlay');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'welcomeOverlay';
    ov.className = 'modal-overlay welcome-overlay';
    ov.addEventListener('click', function (e) { if (e.target === ov) closeWelcome(); });
    document.body.appendChild(ov);
  }
  ov.innerHTML =
    '<div class="modal welcome-modal" role="dialog" aria-modal="true" aria-labelledby="welcomeTitle">' +
      '<div class="welcome-head">' +
        '<div class="welcome-icon">' + info.icon + '</div>' +
        '<div class="welcome-role">' + welcomeEsc_(info.title) + '</div>' +
      '</div>' +
      '<div class="welcome-body">' +
        '<h3 id="welcomeTitle">' + welcomeGreeting_() + ', ' + welcomeEsc_(config.user.nama) + '!</h3>' +
        '<p>' + welcomeEsc_(info.desc) + '</p>' +
        modulHtml +
      '</div>' +
      '<div class="welcome-actions"><button type="button" class="btn" id="welcomeOk">Mulai</button></div>' +
    '</div>';
  document.getElementById('welcomeOk').onclick = closeWelcome;
  ov.classList.add('show');
  document.getElementById('welcomeOk').focus();
}

function closeWelcome() {
  const ov = document.getElementById('welcomeOverlay');
  if (ov) ov.classList.remove('show');
}

// Ambil/segarkan daftar siswa (dipakai untuk autocomplete NISN -> auto-isi Nama/Kelas/JK
// di form modul lain). Dipanggil sekali saat login, dan dipanggil ulang setiap kali ada
// perubahan data di modul Siswa sendiri (tambah/edit/hapus) supaya cache selalu terbaru.
function loadSiswaList() {
  google.script.run.withSuccessHandler(function (list) {
    SISWA_LIST = list || [];
  }).withFailureHandler(function (e) {
    console.error('Gagal memuat daftar siswa untuk autocomplete:', e);
  }).getSiswaList(SESSION_TOKEN);
}

// Ambil/segarkan daftar anggota Docil (dipakai untuk autocomplete "Nama Petugas Docil"
// di form modul Log Kegiatan Docil). Dipanggil sekali saat login, dan dipanggil ulang
// setiap kali ada perubahan data di modul "Anggota Dokter Cilik" sendiri supaya cache
// selalu terbaru.
function loadDocilList() {
  google.script.run.withSuccessHandler(function (list) {
    DOCIL_LIST = list || [];
  }).withFailureHandler(function (e) {
    console.error('Gagal memuat daftar anggota Docil untuk autocomplete:', e);
  }).getDocilList(SESSION_TOKEN);
}

// (BARU) Ambil/segarkan daftar Tim/Kader UKS (dipakai untuk autocomplete field
// Penanggung Jawab/Petugas/Pembina di modul lain — lihat KADER_SYNC_FIELDS_).
// Dipanggil sekali saat login, dan dipanggil ulang setiap kali ada perubahan data
// di modul KaderAnggota sendiri supaya cache selalu terbaru.
function loadKaderList() {
  google.script.run.withSuccessHandler(function (list) {
    KADER_LIST = list || [];
  }).withFailureHandler(function (e) {
    console.error('Gagal memuat daftar Kader UKS untuk autocomplete:', e);
  }).getKaderList(SESSION_TOKEN);
}

// ==================== SIDEBAR: DIKELOMPOKKAN PER TUGAS/ROLE (BARU) ====================
// Sebelumnya semua modul yang boleh diakses ditampilkan sebagai daftar datar (flat list)
// di bawah judul "Modul Data". Untuk Admin (yang bisa melihat SEMUA modul dari SEMUA
// role sekaligus), daftar datar ini jadi panjang dan sulit dicari. Sekarang modul
// dikelompokkan seperti "folder" sesuai tugas/role pemiliknya (lihat MODULE_GROUPS di
// atas), dan tiap folder bisa dibuka/ditutup (collapsible) dengan mengklik judulnya.
//
// Untuk role SELAIN Admin, tampilannya tetap sederhana (tanpa folder) karena mereka
// memang cuma punya modul dari satu tugas saja — pengelompokan hanya benar-benar
// dipakai kalau jumlah modul yang tampil cukup banyak (indikasi kuat sedang login
// sebagai Admin/akses 'ALL').
function buildSidebar() {
  const nav = document.getElementById('navList');
  let html = '<li><a class="nav-link active" data-mod="dashboard" onclick="navigate(\'dashboard\')">📊 Dashboard</a></li>';
  const allKeys = Object.keys(APP_CONFIG.modules);

  if (!allKeys.length) {
    html += '<li class="nav-group-title">Tidak ada modul input untuk role Anda</li>';
    nav.innerHTML = html;
    return;
  }

  // Heuristik sederhana: kalau modul yang tampil lebih dari 6 (lebih banyak dari
  // kelompok terbesar manapun di MODULE_GROUPS), berarti ini akun dengan akses 'ALL'
  // (Admin) -> tampilkan berkelompok. Kalau tidak, tampilkan datar seperti biasa.
  const isMultiGroup = allKeys.length > 6;

  if (!isMultiGroup) {
    html += '<li class="nav-group-title">Modul Data</li>';
    allKeys.forEach(function (key) {
      const m = APP_CONFIG.modules[key];
      html += '<li><a class="nav-link" data-mod="' + key + '" onclick="navigate(\'' + key + '\')">' + m.icon + ' ' + m.label + '</a></li>';
    });
    nav.innerHTML = html;
    return;
  }

  const usedKeys = [];
  MODULE_GROUPS.forEach(function (group) {
    const groupKeys = group.modules.filter(function (k) { return APP_CONFIG.modules.hasOwnProperty(k); });
    if (!groupKeys.length) return;
    usedKeys.push.apply(usedKeys, groupKeys);

    const groupId = 'grp-' + group.role;
    html += '<li class="nav-group">' +
      '<div class="nav-group-header" onclick="toggleGroup(\'' + groupId + '\')">' +
        '<span>' + group.icon + ' ' + group.label + '</span>' +
        '<span class="grp-arrow" id="arrow-' + groupId + '">▾</span>' +
      '</div>' +
      '<ul class="nav-group-list" id="' + groupId + '">';
    groupKeys.forEach(function (key) {
      const m = APP_CONFIG.modules[key];
      html += '<li><a class="nav-link" data-mod="' + key + '" onclick="navigate(\'' + key + '\')">' + m.icon + ' ' + m.label + '</a></li>';
    });
    html += '</ul></li>';
  });

  // Jaring pengaman: kalau ada modul di SHEET_CONFIG yang belum dipetakan ke
  // MODULE_GROUPS manapun (mis. modul baru ditambahkan di Code.gs tapi lupa
  // ditambahkan ke MODULE_GROUPS di atas), modul itu tetap ditampilkan di sini
  // supaya tidak "hilang" dari sidebar tanpa disadari.
  const sisaKeys = allKeys.filter(function (k) { return usedKeys.indexOf(k) === -1; });
  if (sisaKeys.length) {
    html += '<li class="nav-group-title">Lainnya</li>';
    sisaKeys.forEach(function (key) {
      const m = APP_CONFIG.modules[key];
      html += '<li><a class="nav-link" data-mod="' + key + '" onclick="navigate(\'' + key + '\')">' + m.icon + ' ' + m.label + '</a></li>';
    });
  }

  nav.innerHTML = html;
}

// ==================== SIDEBAR HP: BUKA / TUTUP ====================
function openSidebar() {
  document.getElementById('sidebar').classList.add('open');
  const bd = document.getElementById('sidebarBackdrop');
  if (bd) bd.classList.add('show');
}

function closeSidebar() {
  const sb = document.getElementById('sidebar');
  if (sb) sb.classList.remove('open');
  const bd = document.getElementById('sidebarBackdrop');
  if (bd) bd.classList.remove('show');
}

function toggleSidebar() {
  if (document.getElementById('sidebar').classList.contains('open')) closeSidebar();
  else openSidebar();
}

function setupSidebarMobile() {
  const app = document.getElementById('app');
  const sidebar = document.getElementById('sidebar');

  // Latar gelap di belakang sidebar (dibuat lewat JS, diletakkan di dalam #app
  // supaya otomatis ikut tersembunyi saat layar login tampil).
  const backdrop = document.createElement('div');
  backdrop.id = 'sidebarBackdrop';
  backdrop.onclick = closeSidebar;
  app.appendChild(backdrop);

  // Tombol ✕ di pojok kanan atas sidebar (hanya tampil di HP, diatur di Stylesheet).
  const brand = sidebar.querySelector('.brand');
  if (brand) {
    const closeBtn = document.createElement('button');
    closeBtn.type = 'button';
    closeBtn.id = 'sidebarClose';
    closeBtn.className = 'sidebar-close';
    closeBtn.setAttribute('aria-label', 'Tutup menu');
    closeBtn.textContent = '✕';
    closeBtn.onclick = closeSidebar;
    brand.appendChild(closeBtn);
  }

  document.getElementById('menuToggle').onclick = toggleSidebar;

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { closeSidebar(); closeWelcome(); }
  });
  // Kalau layar dilebarkan (mis. HP diputar), pastikan latar gelap tidak tertinggal.
  window.addEventListener('resize', function () {
    if (window.innerWidth > 820) closeSidebar();
  });
}

// Buka/tutup satu folder kelompok modul di sidebar. Dipanggil saat judul folder diklik.
function toggleGroup(groupId) {
  const list = document.getElementById(groupId);
  const arrow = document.getElementById('arrow-' + groupId);
  if (!list) return;
  const collapsed = list.classList.toggle('collapsed');
  if (arrow) arrow.textContent = collapsed ? '▸' : '▾';
}

function navigate(moduleKey) {
  CURRENT_MODULE = moduleKey;
  document.querySelectorAll('.nav-link').forEach(function (a) {
    a.classList.toggle('active', a.dataset.mod === moduleKey);
  });
  closeSidebar();

  // Tombol "Cetak LPJ" hanya tampil untuk Admin/Bendahara saat membuka modul RKAS
  // atau Buku Kas — di luar itu selalu disembunyikan.
  const lpjBtn = document.getElementById('lpjBtn');
  if (lpjBtn) {
    const modulSiapDicetak = (moduleKey === 'RKAS' || moduleKey === 'BukuKas');
    const rolePunyaAkses = CURRENT_USER && (CURRENT_USER.role === 'Admin' || CURRENT_USER.role === 'Bendahara');
    lpjBtn.style.display = (modulSiapDicetak && rolePunyaAkses) ? 'inline-flex' : 'none';
  }

  // Tombol cetak Laporan/Matriks generik — tampil untuk modul apa pun yang
  // terdaftar di PRINT_CONFIG dengan type 'laporan' atau 'matrix' (lihat
  // definisi PRINT_CONFIG di bagian bawah file). Tidak ada pembatasan role
  // khusus (beda dari LPJ) karena laporan ini bukan data keuangan.
  const reportBtn = document.getElementById('reportBtn');
  if (reportBtn) {
    const printCfg = PRINT_CONFIG[moduleKey];
    if (printCfg && (printCfg.type === 'laporan' || printCfg.type === 'matrix')) {
      reportBtn.textContent = printCfg.label;
      reportBtn.style.display = 'inline-flex';
    } else {
      reportBtn.style.display = 'none';
    }
  }

  // Tombol "Import Excel" — tampil untuk modul apa pun yang terdaftar di
  // IMPORT_CONFIG (lihat definisi di bagian atas file).
  const importBtn = document.getElementById('importBtn');
  if (importBtn) {
    importBtn.style.display = IMPORT_CONFIG[moduleKey] ? 'inline-flex' : 'none';
  }

  if (moduleKey === 'dashboard') {
    document.getElementById('pageTitle').textContent = 'Dashboard';
    document.getElementById('addBtn').style.display = 'none';
    renderDashboard();
  } else if (moduleKey === 'pengaturan') {
    document.getElementById('pageTitle').textContent = 'Pengaturan';
    document.getElementById('addBtn').style.display = 'none';
    renderPengaturan();
  } else {
    const m = APP_CONFIG.modules[moduleKey];
    if (!m) { navigate('dashboard'); return; }
    document.getElementById('pageTitle').textContent = m.label;
    document.getElementById('addBtn').style.display = 'inline-flex';
    document.getElementById('addBtn').onclick = function () { openModal(); };
    loadModuleData(moduleKey);
  }
}

// ==================== PENGATURAN: AKUN SAYA & MANAJEMEN LOGIN (BARU) ====================
// - "Akun Saya"       : semua peran (ubah kata sandi sendiri).
// - "Manajemen Login" : hanya Admin (tambah/ubah/nonaktifkan/reset sandi/hapus akun).
// Server-nya ada di Code.gs bagian "PENGATURAN: MANAJEMEN AKUN LOGIN"; pengecekan hak
// akses yang sebenarnya dilakukan di server, tampilan di sini hanya menyesuaikan.
var PENGATURAN_TAB = 'akun';
var PENGATURAN_USERS = [];
var PENGATURAN_ROLES = [];
var ROLE_LABELS = {
  Admin: 'Admin', Pembina: 'Pembina', Ketua: 'Ketua', Sekretaris: 'Sekretaris', Bendahara: 'Bendahara',
  DokterPembina: 'Dokter Pembina', PetugasUKS: 'Petugas UKS', KaderPMR: 'Kader PMR'
};
var MIN_PASSWORD_LEN = 6;

function roleLabel_(r) { return ROLE_LABELS[r] || r; }

function appendPengaturanNav_() {
  const nav = document.getElementById('navList');
  if (!nav) return;
  nav.insertAdjacentHTML('beforeend',
    '<li class="nav-group-title">Sistem</li>' +
    '<li><a class="nav-link" data-mod="pengaturan" onclick="navigate(\'pengaturan\')">⚙️ Pengaturan</a></li>');
}

function renderPengaturan(tab) {
  const isAdmin = !!(CURRENT_USER && CURRENT_USER.role === 'Admin');
  if (tab) PENGATURAN_TAB = tab;
  if (PENGATURAN_TAB === 'login' && !isAdmin) PENGATURAN_TAB = 'akun';

  function tabBtn(key, label) {
    return '<button type="button" class="tab' + (PENGATURAN_TAB === key ? ' active' : '') +
      '" onclick="renderPengaturan(\'' + key + '\')">' + label + '</button>';
  }
  document.getElementById('content').innerHTML =
    '<div class="tabs">' + tabBtn('akun', '👤 Akun Saya') + (isAdmin ? tabBtn('login', '🔐 Manajemen Login') : '') + '</div>' +
    '<div id="pengaturanBody"></div>';

  if (PENGATURAN_TAB === 'login') loadPenggunaList_();
  else renderAkunSaya_();
}

// ---------- Akun Saya ----------
function renderAkunSaya_() {
  const u = CURRENT_USER || {};
  document.getElementById('pengaturanBody').innerHTML =
    '<div class="panel" style="max-width:520px;">' +
      '<h2>Akun Saya</h2>' +
      '<div class="pg-info"><span>Nama</span><b>' + welcomeEsc_(u.nama) + '</b></div>' +
      '<div class="pg-info"><span>Nama akun</span><b>' + welcomeEsc_(u.username || '-') + '</b></div>' +
      '<div class="pg-info"><span>Peran</span><b>' + welcomeEsc_(roleLabel_(u.role)) + '</b></div>' +
    '</div>' +
    '<div class="panel" style="max-width:520px;">' +
      '<h2>Ubah Kata Sandi</h2>' +
      '<div class="pg-stack">' +
        '<div class="field"><label for="pgOldPw">Kata sandi lama</label><input type="password" id="pgOldPw" autocomplete="current-password"></div>' +
        '<div class="field"><label for="pgNewPw">Kata sandi baru (minimal ' + MIN_PASSWORD_LEN + ' karakter)</label><input type="password" id="pgNewPw" autocomplete="new-password"></div>' +
        '<div class="field"><label for="pgNewPw2">Ulangi kata sandi baru</label><input type="password" id="pgNewPw2" autocomplete="new-password"></div>' +
      '</div>' +
      '<label class="pg-check"><input type="checkbox" onchange="pgShowPw_(this, [\'pgOldPw\',\'pgNewPw\',\'pgNewPw2\'])"> Tampilkan kata sandi</label>' +
      '<div class="login-error" id="pgPwErr" style="margin:10px 0 0;"></div>' +
      '<div style="margin-top:14px;"><button type="button" class="btn" id="pgPwBtn" onclick="submitUbahPasswordSaya()">Simpan Kata Sandi</button></div>' +
    '</div>';
}

function pgShowPw_(cb, ids) {
  ids.forEach(function (id) {
    const el = document.getElementById(id);
    if (el) el.type = cb.checked ? 'text' : 'password';
  });
}

function pgShowErr_(id, msg) {
  const el = document.getElementById(id);
  if (!el) return;
  el.textContent = msg || '';
  el.style.display = msg ? 'block' : 'none';
}

// Kesalahan dari server: kalau sesi habis -> keluar; selain itu tampilkan di kotak error.
function pgFail_(e, errBoxId) {
  hideLoading();
  const msg = (e && e.message) ? e.message : String(e);
  if (msg.indexOf('Sesi') > -1) { onError(e); return; }
  if (errBoxId && document.getElementById(errBoxId)) pgShowErr_(errBoxId, msg);
  else showToast(msg, 'error');
}

function submitUbahPasswordSaya() {
  const lama = document.getElementById('pgOldPw').value;
  const baru = document.getElementById('pgNewPw').value;
  const ulang = document.getElementById('pgNewPw2').value;
  pgShowErr_('pgPwErr', '');
  if (!lama) { pgShowErr_('pgPwErr', 'Kata sandi lama wajib diisi.'); return; }
  if (baru.length < MIN_PASSWORD_LEN) { pgShowErr_('pgPwErr', 'Kata sandi baru minimal ' + MIN_PASSWORD_LEN + ' karakter.'); return; }
  if (baru !== ulang) { pgShowErr_('pgPwErr', 'Ulangi kata sandi baru dengan benar (kedua isian tidak sama).'); return; }

  showLoading('Menyimpan kata sandi...');
  google.script.run.withSuccessHandler(function () {
    hideLoading();
    showToast('Kata sandi berhasil diubah', 'success');
    renderAkunSaya_();
  }).withFailureHandler(function (e) { pgFail_(e, 'pgPwErr'); }).ubahPasswordSaya(SESSION_TOKEN, lama, baru);
}

// ---------- Manajemen Login (Admin) ----------
function loadPenggunaList_() {
  const body = document.getElementById('pengaturanBody');
  if (body) body.innerHTML = '<div class="empty-state">Memuat daftar akun...</div>';
  showLoading('Memuat daftar akun...');
  google.script.run.withSuccessHandler(function (res) {
    hideLoading();
    PENGATURAN_USERS = res.users || [];
    PENGATURAN_ROLES = res.roles || [];
    renderPenggunaTable_();
  }).withFailureHandler(function (e) { hideLoading(); onError(e, 'pengaturan'); }).getPenggunaList(SESSION_TOKEN);
}

function renderPenggunaTable_() {
  const body = document.getElementById('pengaturanBody');
  if (!body) return;
  const myId = CURRENT_USER ? CURRENT_USER.userId : null;

  let html = '<div class="panel">' +
    '<div class="pg-head"><h2 style="margin:0;">Daftar Akun Login (' + PENGATURAN_USERS.length + ')</h2>' +
    '<button type="button" class="btn" onclick="openUserForm_(-1)">+ Tambah Akun</button></div>' +
    '<p class="pg-note">Akun <b>Nonaktif</b> tidak bisa masuk. Kalau peran, status, atau kata sandi seseorang diubah, ' +
    'ia akan diminta login ulang.</p>';

  if (!PENGATURAN_USERS.length) {
    html += '<div class="empty-state">Belum ada akun.</div>';
  } else {
    html += '<div class="table-wrap"><table><thead><tr><th>Nama</th><th>Nama Akun</th><th>Peran</th><th>Status</th><th>Aksi</th></tr></thead><tbody>';
    PENGATURAN_USERS.forEach(function (u, i) {
      const saya = (u.id === myId);
      html += '<tr>' +
        '<td>' + welcomeEsc_(u.nama) + (saya ? ' <span class="badge me">Anda</span>' : '') + '</td>' +
        '<td>' + welcomeEsc_(u.username) + '</td>' +
        '<td>' + welcomeEsc_(roleLabel_(u.role)) + '</td>' +
        '<td><span class="badge ' + (u.status === 'Aktif' ? 'ok' : 'off') + '">' + welcomeEsc_(u.status) + '</span></td>' +
        '<td><div class="actions-cell">' +
          '<button type="button" class="btn secondary small" onclick="openUserForm_(' + i + ')">✏️ Ubah</button>' +
          '<button type="button" class="btn secondary small" onclick="openResetPw_(' + i + ')">🔑 Sandi</button>' +
          (saya ? '' : '<button type="button" class="btn danger small" onclick="confirmHapusUser_(' + i + ')">🗑️ Hapus</button>') +
        '</div></td></tr>';
    });
    html += '</tbody></table></div>';
  }
  html += '</div>';
  body.innerHTML = html;
}

// Modal khusus (dibuat lewat JS). Sengaja TIDAK menutup saat area gelap diklik,
// supaya isian yang sedang diketik tidak hilang tanpa sengaja.
function openUserModal_(title, innerHtml) {
  let ov = document.getElementById('userModalOverlay');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'userModalOverlay';
    ov.className = 'modal-overlay';
    document.body.appendChild(ov);
  }
  ov.innerHTML =
    '<div class="modal" style="max-width:480px;">' +
      '<div class="modal-header"><h3>' + welcomeEsc_(title) + '</h3>' +
      '<button type="button" class="btn secondary small" onclick="closeUserModal_()">✕</button></div>' +
      innerHtml +
    '</div>';
  ov.classList.add('show');
}

function closeUserModal_() {
  const ov = document.getElementById('userModalOverlay');
  if (ov) ov.classList.remove('show');
}

// idx = -1 -> tambah akun baru; idx >= 0 -> ubah akun PENGATURAN_USERS[idx]
function openUserForm_(idx) {
  const isNew = idx < 0;
  const u = isNew ? { nama: '', username: '', role: 'PetugasUKS', status: 'Aktif' } : PENGATURAN_USERS[idx];
  const myId = CURRENT_USER ? CURRENT_USER.userId : null;
  const saya = !isNew && u.id === myId;

  const roleOpts = PENGATURAN_ROLES.map(function (r) {
    return '<option value="' + welcomeEsc_(r) + '"' + (r === u.role ? ' selected' : '') + '>' + welcomeEsc_(roleLabel_(r)) + '</option>';
  }).join('');
  const statusOpts = ['Aktif', 'Nonaktif'].map(function (s) {
    return '<option value="' + s + '"' + (s === u.status ? ' selected' : '') + '>' + s + '</option>';
  }).join('');

  const lockAttr = saya ? ' disabled' : '';
  openUserModal_(isNew ? 'Tambah Akun' : 'Ubah Akun',
    '<div class="pg-stack">' +
      '<div class="field"><label for="pgNama">Nama lengkap</label><input type="text" id="pgNama" value="' + welcomeEsc_(u.nama) + '" placeholder="Contoh: Siti Aisyah, S.Pd"></div>' +
      '<div class="field"><label for="pgUsername">Nama akun (untuk login)</label><input type="text" id="pgUsername" value="' + welcomeEsc_(u.username) + '" autocomplete="off" placeholder="Contoh: siti.aisyah"></div>' +
      '<div class="field"><label for="pgRole">Peran</label><select id="pgRole"' + lockAttr + '>' + roleOpts + '</select></div>' +
      '<div class="field"><label for="pgStatus">Status</label><select id="pgStatus"' + lockAttr + '>' + statusOpts + '</select></div>' +
      (isNew ? '<div class="field"><label for="pgPass">Kata sandi (minimal ' + MIN_PASSWORD_LEN + ' karakter)</label><input type="password" id="pgPass" autocomplete="new-password"></div>' +
        '<label class="pg-check"><input type="checkbox" onchange="pgShowPw_(this, [\'pgPass\'])"> Tampilkan kata sandi</label>' : '') +
    '</div>' +
    (saya ? '<p class="pg-note">Peran dan status akun Anda sendiri tidak dapat diubah.</p>' : '') +
    '<div class="login-error" id="pgFormErr" style="margin:10px 0 0;"></div>' +
    '<div class="modal-actions">' +
      '<button type="button" class="btn secondary" onclick="closeUserModal_()">Batal</button>' +
      '<button type="button" class="btn" id="pgSaveBtn" onclick="submitUserForm_(' + idx + ')">Simpan</button>' +
    '</div>');
  document.getElementById('pgNama').focus();
}

function submitUserForm_(idx) {
  const isNew = idx < 0;
  const data = {
    nama: document.getElementById('pgNama').value.trim(),
    username: document.getElementById('pgUsername').value.trim(),
    role: document.getElementById('pgRole').value,
    status: document.getElementById('pgStatus').value
  };
  pgShowErr_('pgFormErr', '');
  if (!data.nama) { pgShowErr_('pgFormErr', 'Nama lengkap wajib diisi.'); return; }
  if (!/^[A-Za-z0-9._ -]{3,40}$/.test(data.username)) {
    pgShowErr_('pgFormErr', 'Nama akun 3-40 karakter, hanya huruf, angka, spasi, titik, garis bawah, atau strip.');
    return;
  }
  if (isNew) {
    data.password = document.getElementById('pgPass').value;
    if (data.password.length < MIN_PASSWORD_LEN) { pgShowErr_('pgFormErr', 'Kata sandi minimal ' + MIN_PASSWORD_LEN + ' karakter.'); return; }
  }

  showLoading('Menyimpan akun...');
  const ok = function () {
    hideLoading();
    closeUserModal_();
    showToast(isNew ? 'Akun berhasil ditambahkan' : 'Akun berhasil diperbarui', 'success');
    loadPenggunaList_();
  };
  const fail = function (e) { pgFail_(e, 'pgFormErr'); };
  if (isNew) {
    google.script.run.withSuccessHandler(ok).withFailureHandler(fail).tambahPengguna(SESSION_TOKEN, data);
  } else {
    // Untuk akun sendiri, kolom peran/status terkunci (disabled) tapi nilainya tetap terbaca.
    google.script.run.withSuccessHandler(ok).withFailureHandler(fail).ubahPengguna(SESSION_TOKEN, PENGATURAN_USERS[idx].id, data);
  }
}

function openResetPw_(idx) {
  const u = PENGATURAN_USERS[idx];
  openUserModal_('Atur Ulang Kata Sandi',
    '<p class="pg-note" style="margin-top:0;">Akun: <b>' + welcomeEsc_(u.nama) + '</b> (' + welcomeEsc_(u.username) + ')</p>' +
    '<div class="pg-stack">' +
      '<div class="field"><label for="pgResetPw">Kata sandi baru (minimal ' + MIN_PASSWORD_LEN + ' karakter)</label><input type="password" id="pgResetPw" autocomplete="new-password"></div>' +
    '</div>' +
    '<label class="pg-check"><input type="checkbox" onchange="pgShowPw_(this, [\'pgResetPw\'])"> Tampilkan kata sandi</label>' +
    '<div class="login-error" id="pgResetErr" style="margin:10px 0 0;"></div>' +
    '<div class="modal-actions">' +
      '<button type="button" class="btn secondary" onclick="closeUserModal_()">Batal</button>' +
      '<button type="button" class="btn" onclick="submitResetPw_(' + idx + ')">Simpan</button>' +
    '</div>');
  document.getElementById('pgResetPw').focus();
}

function submitResetPw_(idx) {
  const pw = document.getElementById('pgResetPw').value;
  pgShowErr_('pgResetErr', '');
  if (pw.length < MIN_PASSWORD_LEN) { pgShowErr_('pgResetErr', 'Kata sandi minimal ' + MIN_PASSWORD_LEN + ' karakter.'); return; }
  showLoading('Menyimpan kata sandi...');
  google.script.run.withSuccessHandler(function () {
    hideLoading();
    closeUserModal_();
    showToast('Kata sandi berhasil diatur ulang', 'success');
  }).withFailureHandler(function (e) { pgFail_(e, 'pgResetErr'); })
    .resetPasswordPengguna(SESSION_TOKEN, PENGATURAN_USERS[idx].id, pw);
}

function confirmHapusUser_(idx) {
  const u = PENGATURAN_USERS[idx];
  openDeleteConfirm('Hapus akun ini?',
    '<div>' + welcomeEsc_(u.nama) + '</div><div style="color:#5f6368;font-size:12.5px;">' +
      welcomeEsc_(u.username) + ' &bull; ' + welcomeEsc_(roleLabel_(u.role)) + '</div>' +
    '<div style="color:#5f6368;font-size:12.5px;margin-top:6px;">Kalau hanya ingin melarang masuk sementara, lebih aman ubah statusnya menjadi Nonaktif.</div>',
    function () {
      showLoading('Menghapus akun...');
      google.script.run.withSuccessHandler(function () {
        hideLoading();
        showToast('Akun berhasil dihapus', 'success');
        loadPenggunaList_();
      }).withFailureHandler(function (e) { pgFail_(e); }).hapusPengguna(SESSION_TOKEN, u.id);
    });
}

// ==================== HELPER FORMAT RUPIAH ====================
function formatRupiah(n) {
  return 'Rp ' + Number(n || 0).toLocaleString('id-ID');
}

// ==================== HELPER: KELOMPOKKAN KOLOM CHECKBOX BERURUTAN (BARU) ====================
// Modul yang punya beberapa kolom bertipe 'checkbox' berurutan di SHEET_CONFIG (mis. 12
// kolom bulan Jan-Des pada modul Jadwal & Program Kerja) ditampilkan sebagai SATU
// kelompok, baik di form Tambah/Edit Data (checkbox berjejer dalam satu field) maupun
// di tabel data (satu kolom "Bulan Pelaksanaan" berisi daftar bulan yang dicentang) —
// supaya form/tabel tidak kepanjangan berisi 12 field/kolom terpisah. Ini generik,
// jadi berlaku otomatis untuk modul APAPUN yang punya kolom bertipe 'checkbox' —
// tidak di-hardcode khusus untuk modul Jadwal & Program Kerja saja.
function groupDisplayColumns_(columns) {
  const groups = [];
  let i = 0;
  while (i < columns.length) {
    const c = columns[i];
    if (c.type === 'checkbox') {
      const group = [c];
      let j = i + 1;
      while (j < columns.length && columns[j].type === 'checkbox') {
        group.push(columns[j]);
        j++;
      }
      groups.push({ type: 'checkboxGroup', columns: group });
      i = j;
    } else {
      groups.push({ type: 'single', column: c });
      i++;
    }
  }
  return groups;
}

function isChecked_(val) {
  return val === true || val === 'TRUE' || val === 'Ya';
}

// ==================== DASHBOARD ====================
function renderDashboard() {
  document.getElementById('content').innerHTML = '<div class="empty-state">Memuat statistik...</div>';
  showLoading('Memuat dashboard...');
  google.script.run.withSuccessHandler(function (stats) {
    hideLoading();
    const cards = [
      { label: 'Total Siswa Terdata', num: stats.totalSiswa, cls: '' },
      { label: 'Kunjungan UKS Hari Ini', num: stats.kunjunganHariIni, cls: '' },
      { label: 'Total Catatan Kecelakaan', num: stats.kecelakaanBulanIni, cls: 'warn' },
      { label: 'Obat Stok Menipis (≤5)', num: stats.obatMenipis, cls: 'danger' },
      { label: 'Obat Mendekati Kadaluarsa', num: stats.obatMendekatiED, cls: 'danger' },
      { label: 'Total Log Kegiatan Tim UKS', num: stats.totalKegiatan, cls: '' }
    ];
    let html = '<div class="stat-grid">';
    cards.forEach(function (c) {
      html += '<div class="stat-card ' + c.cls + '"><div class="num">' + c.num + '</div><div class="label">' + c.label + '</div></div>';
    });
    html += '</div>';

    // Kartu ringkasan RKAS (rencana vs realisasi) — hanya ditampilkan kalau ada
    // datanya, supaya Dashboard utama juga langsung berguna buat Bendahara.
    if (stats.rkasTotalRencana || stats.rkasTotalRealisasi) {
      const sisa = stats.rkasTotalRencana - stats.rkasTotalRealisasi;
      html += '<div class="panel" style="margin-top:16px;"><h2>Ringkasan Anggaran (RKAS)</h2>';
      html += '<div class="stat-grid">';
      html += '<div class="stat-card"><div class="num" style="font-size:18px;">' + formatRupiah(stats.rkasTotalRencana) + '</div><div class="label">Total Rencana Anggaran</div></div>';
      html += '<div class="stat-card"><div class="num" style="font-size:18px;">' + formatRupiah(stats.rkasTotalRealisasi) + '</div><div class="label">Total Realisasi</div></div>';
      html += '<div class="stat-card ' + (sisa < 0 ? 'danger' : '') + '"><div class="num" style="font-size:18px;">' + formatRupiah(sisa) + '</div><div class="label">Sisa Anggaran</div></div>';
      html += '</div></div>';
    }

    html += '<div class="panel" style="margin-top:16px;"><h2>Selamat datang, ' + (CURRENT_USER ? CURRENT_USER.nama : '') + '</h2>' +
      '<p style="color:#5f6368;font-size:13.5px;">Anda masuk sebagai <b>' + (CURRENT_USER ? CURRENT_USER.role : '') +
      '</b>. Gunakan menu di samping untuk mencatat data sesuai tugas Anda. Semua data tersimpan otomatis ke satu Google Spreadsheet yang sama.</p></div>';
    document.getElementById('content').innerHTML = html;
  }).withFailureHandler(function (e) { hideLoading(); onError(e, 'dashboard'); }).getDashboardStats(SESSION_TOKEN);
}

// ==================== LIST DATA MODUL ====================
function loadModuleData(moduleKey) {
  document.getElementById('content').innerHTML = '<div class="empty-state">Memuat data...</div>';
  showLoading('Memuat data...');
  google.script.run.withSuccessHandler(function (data) {
    hideLoading();
    data = data || [];
    // Modul Buku Kas: hitung saldo berjalan di sisi client (tanpa panggilan server
    // tambahan) supaya tabelnya bisa menampilkan kolom "Saldo" per transaksi.
    if (moduleKey === 'BukuKas') data = computeSaldoBukuKas(data);
    CURRENT_DATA = data;
    SELECTED_IDS = []; // reset ceklis setiap kali pindah modul / data dimuat ulang
    renderTable(moduleKey, CURRENT_DATA);
  }).withFailureHandler(function (e) {
    hideLoading();
    onError(e, moduleKey);
  }).getSheetData(SESSION_TOKEN, moduleKey);
}

// Data dari server sudah berurutan terbaru-di-atas. Untuk menghitung saldo berjalan
// yang benar, urutkan dulu naik berdasarkan tanggal, jumlahkan/kurangkan berjalan,
// lalu kembalikan ke urutan aslinya (terbaru di atas) dengan field 'saldo' terisi.
function computeSaldoBukuKas(data) {
  const asc = data.slice().sort(function (a, b) { return String(a.tanggal).localeCompare(String(b.tanggal)); });
  let saldo = 0;
  const saldoById = {};
  asc.forEach(function (r) {
    const jumlah = Number(r.jumlah) || 0;
    if (r.jenisTransaksi === 'Kas Masuk') saldo += jumlah; else saldo -= jumlah;
    saldoById[r.id] = saldo;
  });
  return data.map(function (r) {
    const copy = {};
    Object.keys(r).forEach(function (k) { copy[k] = r[k]; });
    copy.saldo = saldoById[r.id];
    return copy;
  });
}

function renderTable(moduleKey, data) {
  const m = APP_CONFIG.modules[moduleKey];
  data = data || [];
  const isBukuKas = moduleKey === 'BukuKas';
  const displayGroups = groupDisplayColumns_(m.columns);
  let html = '<div class="panel">';

  // (BARU) Toolbar aksi modul — semua tombol aksi dikumpulkan dalam satu baris
  // TEPAT DI ATAS TABEL (bukan lagi tersebar di topbar kanan atas), supaya
  // susunannya seragam di seluruh modul: Tambah Data → Import → Cetak → Export
  // PDF → Hapus Terpilih. Isinya diisi oleh buildModuleToolbar_() setelah HTML
  // ini dipasang, karena tombol yang tampil berbeda-beda per modul & per role.
  html += '<div id="moduleToolbar" style="display:flex;flex-wrap:wrap;align-items:center;gap:8px;margin-bottom:14px;"></div>';
  html += '<div class="search-bar"><input id="searchInput" placeholder="Cari data..." oninput="filterTable()"></div>';

  if (!data.length) {
    html += '<div class="empty-state">Belum ada data. Klik "+ Tambah Data" untuk mulai mengisi.</div>';
  } else {
    html += '<div class="table-wrap"><table id="dataTable"><thead><tr>';
    // (BARU) Kolom ceklis "pilih semua". Hanya mencentang baris yang SEDANG TAMPIL,
    // jadi kalau kotak pencarian sedang aktif, yang terpilih hanya hasil pencarian.
    html += '<th style="width:34px;text-align:center;"><input type="checkbox" id="checkAll" onclick="toggleSelectAll(this)" title="Pilih semua data yang tampil" style="width:16px;height:16px;cursor:pointer;vertical-align:middle;"></th>';
    displayGroups.forEach(function (g) {
      html += '<th>' + (g.type === 'checkboxGroup' ? 'Bulan Pelaksanaan' : g.column.label) + '</th>';
    });
    if (isBukuKas) html += '<th>Saldo (Rp)</th>'; // kolom hasil hitung, bukan bagian form input
    html += '<th>Diinput Oleh</th>'; // kolom otomatis dari sistem, bukan bagian form input
    html += '<th>Aksi</th></tr></thead><tbody>';
    data.forEach(function (row) {
      html += '<tr>';
      html += '<td style="text-align:center;"><input type="checkbox" class="row-check" value="' + row.id + '" onclick="toggleRowSelect(this)" style="width:16px;height:16px;cursor:pointer;vertical-align:middle;"></td>';
      displayGroups.forEach(function (g) {
        if (g.type === 'checkboxGroup') {
          const checkedLabels = g.columns.filter(function (gc) { return isChecked_(row[gc.key]); }).map(function (gc) { return gc.label; });
          html += '<td>' + (checkedLabels.length ? checkedLabels.join(', ') : '-') + '</td>';
          return;
        }
        const c = g.column;
        let val = row[c.key] !== undefined && row[c.key] !== null ? row[c.key] : '';
        if ((c.key === 'jumlah' || c.key === 'totalAnggaran' || c.key === 'hargaSatuan' || c.key === 'realisasi') && val !== '') {
          val = formatRupiah(val);
        }
        html += '<td>' + val + '</td>';
      });
      if (isBukuKas) html += '<td><b>' + formatRupiah(row.saldo) + '</b></td>';
      html += '<td style="color:#5f6368;">' + (row.diinputOleh || '-') + '</td>';
      // Tombol cetak per-baris (🖨️) hanya muncul untuk modul yang terdaftar di
      // PRINT_CONFIG dengan type 'surat' (surat/laporan per satu data, mis. per
      // siswa/per kejadian) — lihat definisi PRINT_CONFIG & cetakSurat().
      const printCfg = PRINT_CONFIG[moduleKey];
      const printBtnHtml = (printCfg && printCfg.type === 'surat')
        ? '<button class="btn secondary small" onclick="cetakSurat(\'' + moduleKey + '\',\'' + row.id + '\')" title="' + printCfg.label + '">🖨️</button>'
        : '';
      html += '<td class="actions-cell">' + printBtnHtml +
        '<button class="btn secondary small" onclick="openModal(\'' + row.id + '\')">Edit</button>' +
        '<button class="btn danger small" onclick="confirmDelete(\'' + row.id + '\')">Hapus</button>' +
        '</td></tr>';
    });
    html += '</tbody></table></div>';
  }
  html += '</div>';
  document.getElementById('content').innerHTML = html;
  buildModuleToolbar_(moduleKey);
  updateBulkBar();
}

// ==================== TOOLBAR AKSI PER MODUL (BARU) ====================
// Membangun deretan tombol aksi di atas tabel. Tombol yang muncul menyesuaikan
// modul & hak akses — logikanya SAMA PERSIS dengan yang dipakai navigate() untuk
// tombol topbar, hanya lokasinya dipindah ke dalam panel. Tombol kembar di topbar
// disembunyikan supaya tidak ada dua tombol dengan fungsi sama di satu layar.
function buildModuleToolbar_(moduleKey) {
  const bar = document.getElementById('moduleToolbar');
  if (!bar) return;

  let h = '';

  // Tambah Data — mengikuti hak akses yang sudah ditentukan navigate() lewat #addBtn.
  const topAdd = document.getElementById('addBtn');
  if (topAdd && topAdd.style.display !== 'none') {
    h += '<button type="button" class="btn" onclick="openModal()">+ Tambah Data</button>';
  }

  // Import Excel — hanya modul yang terdaftar di IMPORT_CONFIG.
  if (IMPORT_CONFIG[moduleKey]) {
    h += '<button type="button" class="btn secondary" onclick="openImportModal()">📥 Import Excel</button>';
  }

  // Cetak Laporan/Matriks — hanya modul yang terdaftar di PRINT_CONFIG.
  const printCfg = PRINT_CONFIG[moduleKey];
  if (printCfg && (printCfg.type === 'laporan' || printCfg.type === 'matrix')) {
    h += '<button type="button" class="btn secondary" onclick="cetakLaporanAtauMatrix()">' + printCfg.label + '</button>';
  }

  // Cetak LPJ — khusus modul keuangan & hanya untuk Admin/Bendahara.
  const rolePunyaAksesLPJ = CURRENT_USER && (CURRENT_USER.role === 'Admin' || CURRENT_USER.role === 'Bendahara');
  if ((moduleKey === 'RKAS' || moduleKey === 'BukuKas') && rolePunyaAksesLPJ) {
    h += '<button type="button" class="btn secondary" onclick="cetakLPJ()">🖨️ Cetak LPJ</button>';
  }

  // Export PDF — generik, tersedia di SEMUA modul.
  h += '<button type="button" class="btn secondary" onclick="exportTabelPDF()">📄 Export PDF</button>';

  // Hapus Terpilih — selalu tampil supaya susunan tombol konsisten antar modul,
  // tapi nonaktif (abu-abu) selama belum ada baris yang diceklis.
  h += '<button type="button" class="btn danger" id="bulkDeleteBtn" onclick="confirmDeleteSelected()" disabled style="opacity:.45;cursor:not-allowed;">🗑️ Hapus Terpilih</button>';
  h += '<span id="bulkCount" style="font-size:13px;font-weight:700;color:#b3261e;"></span>';
  h += '<button type="button" class="btn secondary small" id="bulkClearBtn" onclick="clearSelection()" style="display:none;">Batal Pilih</button>';

  bar.innerHTML = h;

  // Sembunyikan tombol kembar di topbar — fungsinya sudah pindah ke toolbar ini.
  ['addBtn', 'importBtn', 'reportBtn', 'lpjBtn'].forEach(function (id) {
    const el = document.getElementById(id);
    if (el) el.style.display = 'none';
  });
}

// ==================== EXPORT PDF TABEL (BARU) ====================
// Mencetak isi tabel modul yang sedang dibuka apa adanya (tanpa kolom ceklis &
// kolom Aksi) memakai jendela cetak yang sama dengan fitur cetak lain, jadi
// hasilnya berkop sekolah. Aturan isi:
// - Kalau ADA baris yang diceklis  -> yang dicetak hanya baris terpilih.
// - Kalau TIDAK ada yang diceklis  -> yang dicetak semua baris yang sedang
//   tampil, jadi hasil kotak "Cari data" ikut terbawa ke PDF.
function exportTabelPDF() {
  const table = document.getElementById('dataTable');
  if (!table) { showToast('Belum ada data untuk diexport', 'error'); return; }

  const m = APP_CONFIG.modules[CURRENT_MODULE];
  const sekolah = APP_CONFIG.sekolah;

  // Kolom pertama (ceklis) dan kolom terakhir (Aksi) tidak ikut dicetak.
  const allHeads = [];
  table.querySelectorAll('thead th').forEach(function (th) { allHeads.push(th.textContent.trim()); });
  const kolomAwal = 1;
  const kolomAkhir = allHeads.length - 1; // eksklusif
  const heads = allHeads.slice(kolomAwal, kolomAkhir);

  const barisTerpilihSaja = SELECTED_IDS.length > 0;
  const rows = [];
  table.querySelectorAll('tbody tr').forEach(function (tr) {
    const cb = tr.querySelector('.row-check');
    if (barisTerpilihSaja) {
      if (!cb || !cb.checked) return;
    } else if (tr.style.display === 'none') {
      return;
    }
    const cells = [];
    const tds = tr.querySelectorAll('td');
    for (let i = kolomAwal; i < kolomAkhir && i < tds.length; i++) {
      cells.push(tds[i].textContent.trim());
    }
    rows.push(cells);
  });

  if (!rows.length) { showToast('Tidak ada baris yang bisa diexport', 'error'); return; }

  // Tabel modul umumnya lebar (banyak kolom), jadi dipaksa orientasi landscape.
  let body = '<style>@page{size:A4 landscape;margin:12mm;}</style>';
  body += '<div class="kop"><h1>' + sekolah.namaSekolah + '</h1>' +
    '<div class="kop-sub">Usaha Kesehatan Sekolah (UKS) — ' + sekolah.kotaKabupaten + '</div></div>' +
    '<hr class="kop-line"><h2 class="judul-surat">' + m.label + '</h2>';
  body += '<div style="font-size:11.5px;color:#555;margin-bottom:8px;">Tahun Ajaran ' + sekolah.tahunAjaran +
    ' — Jumlah data: ' + rows.length + (barisTerpilihSaja ? ' (data terpilih)' : '') + '</div>';

  body += '<table class="rekap-table"><thead><tr><th class="center" style="width:32px;">No</th>';
  heads.forEach(function (hd) { body += '<th>' + hd + '</th>'; });
  body += '</tr></thead><tbody>';
  rows.forEach(function (cells, i) {
    body += '<tr><td class="center">' + (i + 1) + '</td>';
    cells.forEach(function (c) { body += '<td>' + (c || '-') + '</td>'; });
    body += '</tr>';
  });
  body += '</tbody></table>';

  const pencetak = CURRENT_USER ? CURRENT_USER.nama + ' (' + CURRENT_USER.role + ')' : '-';
  body += '<div style="margin-top:14px;font-size:11px;color:#666;">Dicetak oleh: ' + pencetak +
    ' pada ' + new Date().toLocaleString('id-ID') + '</div>';

  openPrintWindow_('Export ' + m.label, body);
}

function filterTable() {
  const q = document.getElementById('searchInput').value.toLowerCase();
  document.querySelectorAll('#dataTable tbody tr').forEach(function (tr) {
    tr.style.display = tr.textContent.toLowerCase().indexOf(q) > -1 ? '' : 'none';
  });
  // Status ceklis "pilih semua" ikut menyesuaikan hasil pencarian yang sedang tampil.
  updateBulkBar();
}

// ==================== PILIH BANYAK DATA / HAPUS MASSAL (BARU) ====================
// SELECTED_IDS menyimpan ID baris yang sedang diceklis pada modul yang sedang dibuka.
// Isinya dikosongkan setiap kali pindah modul atau data dimuat ulang (lihat
// loadModuleData), supaya tidak ada ID "nyasar" dari modul sebelumnya.

// Ambil hanya checkbox dari baris yang benar-benar terlihat (tidak disembunyikan
// oleh filter pencarian). Ini yang membuat "pilih semua" aman dipakai bersama
// kotak Cari data: yang terpilih hanya data hasil pencarian, bukan seluruh sheet.
function visibleRowChecks_() {
  const out = [];
  document.querySelectorAll('#dataTable tbody tr').forEach(function (tr) {
    if (tr.style.display === 'none') return;
    const cb = tr.querySelector('.row-check');
    if (cb) out.push(cb);
  });
  return out;
}

function toggleRowSelect(cb) {
  const i = SELECTED_IDS.indexOf(cb.value);
  if (cb.checked && i === -1) SELECTED_IDS.push(cb.value);
  if (!cb.checked && i > -1) SELECTED_IDS.splice(i, 1);
  updateBulkBar();
}

function toggleSelectAll(cb) {
  visibleRowChecks_().forEach(function (rc) {
    rc.checked = cb.checked;
    const i = SELECTED_IDS.indexOf(rc.value);
    if (cb.checked && i === -1) SELECTED_IDS.push(rc.value);
    if (!cb.checked && i > -1) SELECTED_IDS.splice(i, 1);
  });
  updateBulkBar();
}

function clearSelection() {
  SELECTED_IDS = [];
  document.querySelectorAll('.row-check').forEach(function (c) { c.checked = false; });
  const ca = document.getElementById('checkAll');
  if (ca) { ca.checked = false; ca.indeterminate = false; }
  updateBulkBar();
}

function updateBulkBar() {
  const ada = SELECTED_IDS.length > 0;

  // Tombol "Hapus Terpilih" tetap di tempatnya, hanya berubah aktif/nonaktif —
  // jadi posisi tombol lain tidak bergeser saat pengguna menceklis data.
  const delBtn = document.getElementById('bulkDeleteBtn');
  if (delBtn) {
    delBtn.disabled = !ada;
    delBtn.style.opacity = ada ? '1' : '.45';
    delBtn.style.cursor = ada ? 'pointer' : 'not-allowed';
  }
  const cnt = document.getElementById('bulkCount');
  if (cnt) cnt.textContent = ada ? SELECTED_IDS.length + ' data dipilih' : '';
  const clearBtn = document.getElementById('bulkClearBtn');
  if (clearBtn) clearBtn.style.display = ada ? 'inline-flex' : 'none';

  // Kotak header: tercentang penuh kalau semua baris tampil terpilih, dan
  // "indeterminate" (garis strip) kalau baru sebagian — memberi isyarat visual
  // bahwa masih ada baris tampil yang belum ikut terpilih.
  const ca = document.getElementById('checkAll');
  if (ca) {
    const vis = visibleRowChecks_();
    const checked = vis.filter(function (c) { return c.checked; }).length;
    ca.checked = vis.length > 0 && checked === vis.length;
    ca.indeterminate = checked > 0 && checked < vis.length;
  }
}

function confirmDeleteSelected() {
  if (!SELECTED_IDS.length) { showToast('Belum ada data yang dipilih', 'error'); return; }
  const m = APP_CONFIG.modules[CURRENT_MODULE];
  const identifierKeys = ['namaSiswa', 'nama', 'namaKegiatan', 'namaBarang', 'namaObat', 'uraianKegiatan', 'uraian', 'hari', 'tanggal'];

  const ids = SELECTED_IDS.slice();
  const daftarNama = ids.map(function (id) {
    const r = CURRENT_DATA.find(function (x) { return x.id === id; }) || {};
    for (let i = 0; i < identifierKeys.length; i++) {
      if (r[identifierKeys[i]]) return r[identifierKeys[i]];
    }
    return id;
  });

  // Tampilkan maksimal 10 nama supaya popup tidak kepanjangan, sisanya diringkas.
  let detail = '<div style="text-align:left;font-size:13px;line-height:1.7;max-height:220px;overflow:auto;">';
  daftarNama.slice(0, 10).forEach(function (n, i) { detail += '<div>' + (i + 1) + '. ' + n + '</div>'; });
  if (daftarNama.length > 10) {
    detail += '<div style="color:#5f6368;margin-top:4px;">…dan ' + (daftarNama.length - 10) + ' data lainnya</div>';
  }
  detail += '</div><div style="margin-top:10px;color:#b3261e;font-size:12.5px;font-weight:600;">Data yang sudah dihapus tidak bisa dikembalikan.</div>';

  openDeleteConfirm('Hapus ' + ids.length + ' data ' + m.label + '?', detail, function () {
    showLoading('Menghapus ' + ids.length + ' data...');
    google.script.run.withSuccessHandler(function (res) {
      hideLoading();
      const jml = (res && res.count) ? res.count : ids.length;
      showToast(jml + ' data berhasil dihapus', 'success');
      SELECTED_IDS = [];
      loadModuleData(CURRENT_MODULE);
      if (CURRENT_MODULE === 'Siswa') loadSiswaList();
      if (CURRENT_MODULE === 'DocilAnggota') loadDocilList();
      if (CURRENT_MODULE === 'KaderAnggota') loadKaderList();
    }).withFailureHandler(function (e) { hideLoading(); onError(e); }).deleteRecords(SESSION_TOKEN, CURRENT_MODULE, ids);
  });
}

// ==================== FORM MODAL ====================
// SINKRONISASI DATA SISWA:
// Kolom bertipe 'text' dengan key 'nisn' pada modul APAPUN selain modul Siswa sendiri
// dirender sebagai input dengan <datalist> berisi seluruh siswa terdaftar (format:
// "NISN — Nama (Kelas)"). Begitu petugas mengetik/memilih NISN yang cocok dengan data
// di modul Siswa, field Nama/Kelas/JK pada form yang sama otomatis terisi lewat
// onNisnPicked() — sehingga tidak perlu diketik manual dan tidak mungkin beda ejaan
// dengan data master di modul Siswa.
//
// (BARU) SINKRONISASI DATA KADER UKS: dipetakan eksplisit per modul+key (bukan
// dicocokkan lewat nama key saja seperti nisn/namaPetugas) karena key 'pj' dan
// 'petugas' dipakai juga di modul lain (RKAS, PemeriksaanBerkala, dst) yang
// SENGAJA tidak ikut disinkronkan — supaya field "Penanggung Jawab" di RKAS
// misalnya, tetap bebas diisi nama pihak luar (bendahara BOS, dsb), bukan hanya
// anggota Tim/Kader UKS.
var KADER_SYNC_FIELDS_ = {
  KaderLog: 'pj',
  JadwalKegiatan: 'pj',
  AgendaKegiatan: 'pj',
  DocilPiket: 'pembina',
  Kunjungan: 'petugas',
  Kecelakaan: 'petugas'
};

function openModal(id) {
  EDIT_ID = id || null;
  const m = APP_CONFIG.modules[CURRENT_MODULE];
  document.getElementById('modalTitle').textContent = id ? 'Edit Data' : 'Tambah Data';

  const record = id ? CURRENT_DATA.find(function (r) { return r.id === id; }) : {};
  const showSiswaAutocomplete = CURRENT_MODULE !== 'Siswa' && SISWA_LIST.length > 0;
  // Khusus modul Log Kegiatan Docil: field "Nama Petugas Docil" dibuatkan datalist
  // dari daftar anggota Docil yang sudah terdaftar di modul Anggota Dokter Cilik,
  // supaya tidak perlu diketik ulang manual dan tidak mungkin typo/beda ejaan.
  const showDocilAutocomplete = CURRENT_MODULE === 'DocilLog' && DOCIL_LIST.length > 0;
  // (BARU) Key field yang disinkronkan ke daftar Kader UKS pada modul yang sedang
  // dibuka (undefined kalau modul ini tidak ada di KADER_SYNC_FIELDS_).
  const kaderSyncKey = KADER_SYNC_FIELDS_[CURRENT_MODULE];
  const showKaderAutocomplete = !!kaderSyncKey && KADER_LIST.length > 0;

  let html = '';
  const displayGroups = groupDisplayColumns_(m.columns);
  displayGroups.forEach(function (g) {
    // Kelompok kolom checkbox berurutan (mis. 12 bulan Jan-Des) dirender sebagai
    // SATU field berisi checkbox berjejer, bukan satu field terpisah per bulan —
    // supaya modal Tambah/Edit Data tidak kepanjangan.
    if (g.type === 'checkboxGroup') {
      html += '<div class="field full"><label>Bulan Pelaksanaan</label>' +
        '<div style="display:flex;flex-wrap:wrap;gap:12px;padding:4px 0;">';
      g.columns.forEach(function (gc) {
        const checked = isChecked_(record[gc.key]) ? 'checked' : '';
        html += '<label style="display:flex;align-items:center;gap:5px;font-weight:400;font-size:12.5px;color:var(--text);cursor:pointer;">' +
          '<input type="checkbox" data-key="' + gc.key + '" ' + checked + ' style="width:16px;height:16px;">' + gc.label + '</label>';
      });
      html += '</div></div>';
      return;
    }

    const c = g.column;
    const val = record[c.key] !== undefined ? record[c.key] : '';
    const req = c.required ? 'required' : '';
    const fullClass = (c.type === 'text' && c.label.length > 20) ? ' full' : '';
    html += '<div class="field' + fullClass + '"><label>' + c.label + (c.required ? ' *' : '') + '</label>';
    if (c.type === 'select') {
      html += '<select data-key="' + c.key + '" ' + req + '><option value="">-- Pilih --</option>';
      c.options.forEach(function (opt) {
        html += '<option value="' + opt + '"' + (val === opt ? ' selected' : '') + '>' + opt + '</option>';
      });
      html += '</select>';
    } else if (c.key === 'nisn' && showSiswaAutocomplete) {
      html += '<input type="text" list="siswaDatalist" data-key="nisn" value="' + (val + '').replace(/"/g, '&quot;') +
        '" placeholder="Ketik/pilih NISN siswa..." oninput="onNisnPicked(this.value)" autocomplete="off" ' + req + '>' +
        '<small style="color:#5f6368;font-size:11px;margin-top:2px;">Pilih dari daftar siswa terdaftar — Nama/Kelas/JK akan terisi otomatis.</small>';
    } else if (c.key === 'namaPetugas' && showDocilAutocomplete) {
      html += '<input type="text" list="docilDatalist" data-key="namaPetugas" value="' + (val + '').replace(/"/g, '&quot;') +
        '" placeholder="Ketik/pilih nama anggota Docil..." autocomplete="off" ' + req + '>' +
        '<small style="color:#5f6368;font-size:11px;margin-top:2px;">Pilih dari daftar anggota Dokter Cilik yang sudah terdaftar.</small>';
    } else if (c.key === kaderSyncKey && showKaderAutocomplete) {
      html += '<input type="text" list="kaderDatalist" data-key="' + c.key + '" value="' + (val + '').replace(/"/g, '&quot;') +
        '" placeholder="Ketik/pilih nama anggota Tim/Kader UKS..." autocomplete="off" ' + req + '>' +
        '<small style="color:#5f6368;font-size:11px;margin-top:2px;">Pilih dari daftar Tim/Kader UKS yang sudah terdaftar.</small>';
    } else if (c.key === 'buktiUrl') {
      html += '<input type="url" data-key="buktiUrl" value="' + (val + '').replace(/"/g, '&quot;') +
        '" placeholder="Tempel link foto nota/kwitansi (Google Drive, dsb)..." ' + req + '>' +
        '<small style="color:#5f6368;font-size:11px;margin-top:2px;">Upload foto ke Google Drive, klik "Bagikan" -> "Siapa saja yang memiliki link", lalu tempel link-nya di sini.</small>';
    } else if (c.type === 'number') {
      html += '<input type="number" step="any" data-key="' + c.key + '" value="' + val + '" ' + req + '>';
    } else if (c.type === 'date') {
      html += '<input type="date" data-key="' + c.key + '" value="' + val + '" ' + req + '>';
    } else if (c.type === 'time') {
      html += '<input type="time" data-key="' + c.key + '" value="' + val + '" ' + req + '>';
    } else if (c.type === 'datetime') {
      html += '<input type="datetime-local" data-key="' + c.key + '" value="' + val + '" ' + req + '>';
    } else {
      html += '<input type="text" data-key="' + c.key + '" value="' + (val + '').replace(/"/g, '&quot;') + '" ' + req + '>';
    }
    html += '</div>';
  });

  // Datalist bersama untuk semua field NISN di modul ini (kalau ada).
  if (showSiswaAutocomplete && m.columns.some(function (c) { return c.key === 'nisn'; })) {
    html += '<datalist id="siswaDatalist">';
    SISWA_LIST.forEach(function (s) {
      html += '<option value="' + String(s.nisn).replace(/"/g, '&quot;') + '">' + s.nama + ' — Kelas ' + s.kelas + '</option>';
    });
    html += '</datalist>';
  }

  // Datalist untuk field "Nama Petugas Docil" (khusus modul DocilLog).
  if (showDocilAutocomplete && m.columns.some(function (c) { return c.key === 'namaPetugas'; })) {
    html += '<datalist id="docilDatalist">';
    DOCIL_LIST.forEach(function (d) {
      const ket = d.jabatan ? (d.jabatan + ' — Kelas ' + d.kelas) : ('Kelas ' + d.kelas);
      html += '<option value="' + String(d.nama).replace(/"/g, '&quot;') + '">' + ket + '</option>';
    });
    html += '</datalist>';
  }

  // (BARU) Datalist untuk field Penanggung Jawab/Petugas/Pembina yang disinkronkan
  // ke daftar Tim/Kader UKS (lihat KADER_SYNC_FIELDS_ di atas).
  if (showKaderAutocomplete && m.columns.some(function (c) { return c.key === kaderSyncKey; })) {
    html += '<datalist id="kaderDatalist">';
    KADER_LIST.forEach(function (k) {
      html += '<option value="' + String(k.nama).replace(/"/g, '&quot;') + '">' + (k.jabatan || 'Anggota Tim UKS') + '</option>';
    });
    html += '</datalist>';
  }

  document.getElementById('modalForm').innerHTML = html;
  document.getElementById('modalOverlay').classList.add('show');
}

// Dipanggil setiap kali field NISN diketik/dipilih. Kalau nilainya persis cocok dengan
// salah satu NISN di daftar siswa, field Nama/Kelas/JK pada form yang sedang terbuka
// otomatis diisi dari data siswa tersebut (menimpa isian manual sebelumnya).
function onNisnPicked(nisnValue) {
  const siswa = SISWA_LIST.find(function (s) { return String(s.nisn) === String(nisnValue).trim(); });
  if (!siswa) return;

  const namaField = document.querySelector('#modalForm [data-key="namaSiswa"]') ||
    document.querySelector('#modalForm [data-key="nama"]');
  if (namaField) namaField.value = siswa.nama;

  const kelasField = document.querySelector('#modalForm [data-key="kelas"]');
  if (kelasField) kelasField.value = siswa.kelas;

  const jkField = document.querySelector('#modalForm [data-key="jk"]');
  if (jkField) jkField.value = siswa.jk;
}

function closeModal() {
  document.getElementById('modalOverlay').classList.remove('show');
  EDIT_ID = null;
}

function saveRecord() {
  const data = {};
  let valid = true;
  document.querySelectorAll('#modalForm [data-key]').forEach(function (el) {
    if (el.type === 'checkbox') {
      // Checkbox tidak punya konsep "wajib diisi" seperti field lain — nilainya
      // selalu true/false (dicentang/tidak), bukan berdasarkan el.value.
      data[el.dataset.key] = el.checked;
    } else {
      data[el.dataset.key] = el.value;
      if (el.required && !el.value) valid = false;
    }
  });
  if (!valid) { showToast('Mohon lengkapi field wajib (*)', 'error'); return; }

  const saveBtn = document.getElementById('saveBtn');
  saveBtn.innerHTML = '<span class="spinner"></span>';
  saveBtn.disabled = true;
  showLoading('Menyimpan data...'); // (BARU) popup loading tengah layar saat proses simpan

  function done() {
    saveBtn.innerHTML = 'Simpan';
    saveBtn.disabled = false;
    hideLoading(); // (BARU)
    closeModal();
    loadModuleData(CURRENT_MODULE);
    // Kalau yang barusan diubah adalah data Siswa, segarkan cache SISWA_LIST supaya
    // modul lain langsung melihat NISN/nama terbaru di autocomplete-nya.
    if (CURRENT_MODULE === 'Siswa') loadSiswaList();
    // Sama halnya kalau yang diubah adalah data Anggota Docil, segarkan cache
    // DOCIL_LIST supaya modul Log Kegiatan Docil langsung melihat nama terbaru.
    if (CURRENT_MODULE === 'DocilAnggota') loadDocilList();
    // (BARU) Sama halnya kalau yang diubah adalah data Tim/Kader UKS, segarkan cache
    // KADER_LIST supaya modul KaderLog/JadwalKegiatan/AgendaKegiatan/DocilPiket/
    // Kunjungan/Kecelakaan langsung melihat nama terbaru di autocomplete-nya.
    if (CURRENT_MODULE === 'KaderAnggota') loadKaderList();
  }

  if (EDIT_ID) {
    google.script.run.withSuccessHandler(function () {
      showToast('Data berhasil diperbarui', 'success'); done();
    }).withFailureHandler(function (e) { onError(e); done(); }).updateRecord(SESSION_TOKEN, CURRENT_MODULE, EDIT_ID, data);
  } else {
    google.script.run.withSuccessHandler(function () {
      showToast('Data berhasil ditambahkan', 'success'); done();
    }).withFailureHandler(function (e) { onError(e); done(); }).addRecord(SESSION_TOKEN, CURRENT_MODULE, data);
  }
}

// ==================== HAPUS DATA (DIPERBARUI) ====================
// Sebelumnya memakai confirm() bawaan browser yang hanya menampilkan teks generik
// "Hapus data ini?" — mudah membuat salah hapus kalau daftar datanya panjang.
// Sekarang menampilkan modal konfirmasi khusus (#deleteConfirmOverlay) berisi
// ringkasan field penting dari data yang akan dihapus (mis. Nama Siswa, Nama
// Kegiatan, Uraian, dsb — field pertama yang cocok & field lain yang terisi),
// supaya pengguna bisa memastikan data yang benar sebelum menekan "Ya, Hapus".
function confirmDelete(id) {
  const record = CURRENT_DATA.find(function (r) { return r.id === id; });
  if (!record) { showToast('Data tidak ditemukan', 'error'); return; }
  const m = APP_CONFIG.modules[CURRENT_MODULE];

  // Cari field yang paling mewakili "identitas" data ini untuk judul popup
  const identifierKeys = ['namaSiswa', 'nama', 'namaKegiatan', 'namaBarang', 'namaObat', 'uraianKegiatan', 'uraian'];
  let judul = record.id;
  for (let i = 0; i < identifierKeys.length; i++) {
    if (record[identifierKeys[i]]) { judul = record[identifierKeys[i]]; break; }
  }

  // Tampilkan beberapa field pertama modul ini sebagai ringkasan, supaya jelas
  // data mana yang akan terhapus (menghindari salah hapus).
  let detailHtml = '<div style="text-align:left;font-size:13px;line-height:1.7;">';
  m.columns.slice(0, 4).forEach(function (c) {
    if (c.type === 'checkbox') return;
    const val = record[c.key];
    if (val !== undefined && val !== '' && val !== null) {
      detailHtml += '<div><b>' + c.label + ':</b> ' + val + '</div>';
    }
  });
  detailHtml += '</div>';

  openDeleteConfirm(judul, detailHtml, function () {
    showLoading('Menghapus data...');
    google.script.run.withSuccessHandler(function () {
      hideLoading();
      showToast('Data berhasil dihapus', 'success');
      loadModuleData(CURRENT_MODULE);
      if (CURRENT_MODULE === 'Siswa') loadSiswaList();
      if (CURRENT_MODULE === 'DocilAnggota') loadDocilList();
      if (CURRENT_MODULE === 'KaderAnggota') loadKaderList();
    }).withFailureHandler(function (e) { hideLoading(); onError(e); }).deleteRecord(SESSION_TOKEN, CURRENT_MODULE, id);
  });
}

// ==================== LPJ (LAPORAN PERTANGGUNGJAWABAN) ====================
// Memanggil getLPJData(token) di server (RKAS rencana vs realisasi + Buku Kas
// dengan saldo berjalan), lalu membuka jendela baru berisi halaman siap-cetak.
// Pengguna tinggal klik tombol "Cetak / Simpan sebagai PDF" di jendela tsb
// (memakai window.print() bawaan browser -> pilih "Simpan sebagai PDF").
function cetakLPJ() {
  showToast('Menyiapkan Laporan Pertanggungjawaban...', 'success');
  showLoading('Menyiapkan LPJ...');
  google.script.run.withSuccessHandler(function (data) {
    hideLoading();
    openLPJPrintWindow(data);
  }).withFailureHandler(function (e) { hideLoading(); onError(e); }).getLPJData(SESSION_TOKEN);
}

function openLPJPrintWindow(data) {
  const w = window.open('', '_blank');
  if (!w) {
    showToast('Popup diblokir browser. Izinkan popup untuk situs ini agar bisa mencetak LPJ.', 'error');
    return;
  }

  let html = '<html><head><title>LPJ Bendahara UKS - ' + data.sekolah.namaSekolah + '</title>';
  html += '<style>' +
    'body{font-family:Arial,Helvetica,sans-serif;font-size:12px;color:#222;padding:24px;}' +
    'h1{font-size:16px;margin:0 0 2px;}' +
    'h2{font-size:14px;margin-top:26px;margin-bottom:6px;border-bottom:2px solid #1a73e8;padding-bottom:4px;}' +
    'table{width:100%;border-collapse:collapse;margin-top:6px;}' +
    'th,td{border:1px solid #ccc;padding:5px 7px;font-size:11px;text-align:left;}' +
    'th{background:#1a73e8;color:#fff;}' +
    '.right{text-align:right;}' +
    '.meta{color:#666;margin-bottom:12px;}' +
    '.summary{margin-top:10px;font-size:12.5px;}' +
    '.summary div{margin-bottom:3px;}' +
    '.summary b{display:inline-block;min-width:200px;}' +
    '.ttd{margin-top:48px;display:flex;justify-content:space-between;max-width:420px;}' +
    '.printBtn{margin-top:24px;padding:9px 18px;font-size:13px;cursor:pointer;background:#1a73e8;color:#fff;border:none;border-radius:6px;}' +
    '@media print{.printBtn{display:none;}}' +
    '</style></head><body>';

  html += '<h1>Laporan Pertanggungjawaban (LPJ) Bendahara UKS</h1>';
  html += '<div class="meta">' + data.sekolah.namaSekolah + ' — Tahun Ajaran ' + data.sekolah.tahunAjaran +
    '<br>Dicetak oleh: ' + data.dicetakOleh + ' pada ' + data.tanggalCetak + '</div>';

  html += '<h2>Ringkasan Rencana Anggaran (RKAS) per Pilar Trias UKS</h2>';
  html += '<table><thead><tr><th>Pilar Trias UKS</th><th class="right">Rencana (Rp)</th><th class="right">Realisasi (Rp)</th><th class="right">Sisa (Rp)</th></tr></thead><tbody>';
  data.rkas.rekapPilar.forEach(function (p) {
    html += '<tr><td>' + p.pilar + '</td><td class="right">' + formatRupiah(p.rencana) + '</td><td class="right">' + formatRupiah(p.realisasi) + '</td><td class="right">' + formatRupiah(p.rencana - p.realisasi) + '</td></tr>';
  });
  html += '</tbody></table>';
  html += '<div class="summary">' +
    '<div><b>Total Rencana Anggaran</b> ' + formatRupiah(data.rkas.totalRencana) + '</div>' +
    '<div><b>Total Realisasi</b> ' + formatRupiah(data.rkas.totalRealisasi) + '</div>' +
    '<div><b>Sisa Anggaran</b> ' + formatRupiah(data.rkas.sisa) + '</div>' +
    '</div>';

  html += '<h2>Detail Kegiatan RKAS</h2>';
  html += '<table><thead><tr><th>Uraian Kegiatan</th><th>Pilar</th><th class="right">Rencana (Rp)</th><th class="right">Realisasi (Rp)</th><th>Status Realisasi</th></tr></thead><tbody>';
  if (data.rkas.rows.length) {
    data.rkas.rows.forEach(function (r) {
      html += '<tr><td>' + (r.uraianKegiatan || '') + '</td><td>' + (r.pilarTrias || '') + '</td><td class="right">' +
        formatRupiah(r.totalAnggaran) + '</td><td class="right">' + formatRupiah(r.realisasi) + '</td><td>' + (r.statusRealisasi || '-') + '</td></tr>';
    });
  } else {
    html += '<tr><td colspan="5" style="text-align:center;color:#888;">Belum ada data RKAS</td></tr>';
  }
  html += '</tbody></table>';

  html += '<h2>Buku Kas</h2>';
  html += '<table><thead><tr><th>Tanggal</th><th>Uraian</th><th>Jenis</th><th class="right">Jumlah (Rp)</th><th class="right">Saldo (Rp)</th></tr></thead><tbody>';
  if (data.bukuKas.rows.length) {
    data.bukuKas.rows.forEach(function (r) {
      html += '<tr><td>' + (r.tanggal || '') + '</td><td>' + (r.uraian || '') + '</td><td>' + (r.jenisTransaksi || '') +
        '</td><td class="right">' + formatRupiah(r.jumlah) + '</td><td class="right">' + formatRupiah(r.saldo) + '</td></tr>';
    });
  } else {
    html += '<tr><td colspan="5" style="text-align:center;color:#888;">Belum ada data Buku Kas</td></tr>';
  }
  html += '</tbody></table>';
  html += '<div class="summary">' +
    '<div><b>Total Kas Masuk</b> ' + formatRupiah(data.bukuKas.totalMasuk) + '</div>' +
    '<div><b>Total Kas Keluar</b> ' + formatRupiah(data.bukuKas.totalKeluar) + '</div>' +
    '<div><b>Saldo Akhir</b> ' + formatRupiah(data.bukuKas.saldoAkhir) + '</div>' +
    '</div>';

  html += '<div class="ttd">' +
    '<div>Mengetahui,<br>Kepala Sekolah<br><br><br><br>____________________</div>' +
    '<div>Bendahara UKS<br><br><br><br>____________________</div>' +
    '</div>';

  html += '<button class="printBtn" onclick="window.print()">🖨️ Cetak / Simpan sebagai PDF</button>';
  html += '</body></html>';

  w.document.open();
  w.document.write(html);
  w.document.close();
}

// ==================== CETAK SURAT & LAPORAN MODUL LAIN (BARU) ====================
// PRINT_CONFIG menentukan cara cetak per modul, tanpa perlu menulis ulang kode
// CRUD untuk tiap modul (sama semangatnya dengan SHEET_CONFIG di Code.gs):
// - type 'surat'   -> tombol cetak (🖨️) muncul di SETIAP BARIS tabel; mencetak
//                     SATU data sebagai surat/laporan resmi (mis. per siswa,
//                     per kejadian). Perlu template khusus di SURAT_BUILDERS.
// - type 'laporan' -> tombol cetak muncul di ATAS tabel (seperti "Cetak LPJ");
//                     mencetak SELURUH data modul yang sedang dibuka sebagai
//                     satu tabel rekap. TIDAK perlu template khusus — otomatis
//                     memakai kolom modul itu sendiri (lihat cetakLaporanGenerik).
// - type 'matrix'  -> khusus JadwalKegiatan: dicetak sebagai matriks bulanan
//                     Jan-Des per kategori Program Kerja UKS, meniru Excel asli.
var PRINT_CONFIG = {
  Rujukan: { type: 'surat', label: 'Cetak Surat Rujukan' },
  SuratRujukanSakit: { type: 'surat', label: 'Cetak Surat Izin/Rujukan' },
  Kecelakaan: { type: 'surat', label: 'Cetak Laporan Kejadian' },
  PemeriksaanBerkala: { type: 'surat', label: 'Cetak Hasil Pemeriksaan' },
  BBTB: { type: 'laporan', label: '🖨️ Cetak Laporan Status Gizi' },
  Imunisasi: { type: 'laporan', label: '🖨️ Cetak Laporan Imunisasi' },
  Kunjungan: { type: 'laporan', label: '🖨️ Cetak Rekap Kunjungan' },
  GigiMulut: { type: 'laporan', label: '🖨️ Cetak Laporan Gigi & Mulut' },
  KaderLog: { type: 'laporan', label: '🖨️ Cetak Laporan Kegiatan' },
  DocilLog: { type: 'laporan', label: '🖨️ Cetak Laporan Kegiatan' },
  AgendaKegiatan: { type: 'laporan', label: '🖨️ Cetak Agenda Kegiatan' },
  JadwalKegiatan: { type: 'matrix', label: '🖨️ Cetak Matriks Program Kerja' }
};

// CSS khusus untuk halaman cetak surat/laporan (format kop surat resmi), terpisah
// dari CSS aplikasi utama supaya hasil cetak rapi tanpa terpengaruh gaya UI app.
var SURAT_PRINT_CSS =
  'body{font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#222;padding:30px 40px;}' +
  '.kop{text-align:center;}' +
  '.kop h1{font-size:16px;margin:0;text-transform:uppercase;}' +
  '.kop-sub{font-size:12px;color:#444;margin-top:2px;}' +
  '.kop-line{border:none;border-top:2px solid #222;margin:8px 0 4px;}' +
  '.judul-surat{text-align:center;text-decoration:underline;font-size:14px;margin:18px 0 2px;text-transform:uppercase;}' +
  '.no-surat{text-align:center;font-size:12px;margin-bottom:16px;}' +
  '.info-table{margin:14px 0;border-collapse:collapse;}' +
  '.info-table td{padding:3px 6px;vertical-align:top;font-size:12.5px;}' +
  '.info-label{width:180px;font-weight:600;}' +
  '.info-colon{width:10px;}' +
  'p{font-size:12.5px;line-height:1.6;}' +
  '.tgl-kanan{text-align:right;margin-top:24px;}' +
  '.ttd-row{display:flex;justify-content:flex-end;gap:50px;margin-top:6px;flex-wrap:wrap;}' +
  '.ttd-col{text-align:center;font-size:12.5px;min-width:160px;}' +
  '.ttd-space{height:60px;}' +
  '.ttd-name{text-decoration:underline;font-weight:600;}' +
  '.rekap-table{width:100%;border-collapse:collapse;margin-top:8px;}' +
  '.rekap-table th,.rekap-table td{border:1px solid #ccc;padding:5px 7px;font-size:11px;text-align:left;}' +
  '.rekap-table th{background:#1a73e8;color:#fff;}' +
  '.rekap-table td.center,.rekap-table th.center{text-align:center;}' +
  'h3{font-size:13px;margin:22px 0 4px;border-bottom:1px solid #ccc;padding-bottom:3px;}' +
  '.printBtn{margin-top:24px;padding:9px 18px;font-size:13px;cursor:pointer;background:#1a73e8;color:#fff;border:none;border-radius:6px;}' +
  '@media print{.printBtn{display:none;}}';

function tanggalCetakHariIni_() {
  const d = new Date();
  const bulan = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  return d.getDate() + ' ' + bulan[d.getMonth()] + ' ' + d.getFullYear();
}

// Tabel 2-kolom (label : nilai) untuk isi surat resmi — dipakai semua template
// di SURAT_BUILDERS supaya tampilannya konsisten.
function renderInfoTable_(record, fields) {
  let html = '<table class="info-table">';
  fields.forEach(function (f) {
    const raw = record[f.key];
    const val = (raw !== undefined && raw !== null && raw !== '') ? raw : '-';
    html += '<tr><td class="info-label">' + f.label + '</td><td class="info-colon">:</td><td class="info-value">' + val + '</td></tr>';
  });
  html += '</table>';
  return html;
}

function suratHeaderHtml_(sekolah, judul, noSurat) {
  let html = '<div class="kop"><h1>' + sekolah.namaSekolah + '</h1>' +
    '<div class="kop-sub">Usaha Kesehatan Sekolah (UKS) — ' + sekolah.kotaKabupaten + '</div></div>' +
    '<hr class="kop-line"><h2 class="judul-surat">' + judul + '</h2>';
  if (noSurat) html += '<div class="no-surat">Nomor: ' + noSurat + '</div>';
  return html;
}

// Blok tanda tangan di kanan bawah surat/laporan. cols: [{label, sublabel, nama}]
function ttdBlockHtml_(cols) {
  let html = '<div class="ttd-row">';
  cols.forEach(function (c) {
    html += '<div class="ttd-col"><div>' + c.label + '</div>' +
      (c.sublabel ? '<div style="font-size:11px;color:#666;">' + c.sublabel + '</div>' : '') +
      '<div class="ttd-space"></div><div class="ttd-name">' + (c.nama ? c.nama : '&nbsp;') + '</div></div>';
  });
  html += '</div>';
  return html;
}

function openPrintWindow_(judulTab, bodyHtml) {
  const w = window.open('', '_blank');
  if (!w) {
    showToast('Popup diblokir browser. Izinkan popup untuk situs ini agar bisa mencetak.', 'error');
    return;
  }
  const html = '<html><head><title>' + judulTab + '</title><style>' + SURAT_PRINT_CSS + '</style></head><body>' +
    bodyHtml + '<button class="printBtn" onclick="window.print()">🖨️ Cetak / Simpan sebagai PDF</button></body></html>';
  w.document.open();
  w.document.write(html);
  w.document.close();
}

// ---- Template surat/laporan PER-BARIS (type: 'surat' di PRINT_CONFIG) ----
// Setiap fungsi menerima 1 record data + objek sekolah, mengembalikan HTML isi
// surat (tanpa <html>/<head>/<style> — itu ditambahkan oleh openPrintWindow_).
var SURAT_BUILDERS = {
  Rujukan: function (r, sekolah) {
    let body = suratHeaderHtml_(sekolah, 'Surat Rujukan Kesehatan', r.noSurat);
    body += '<p>Yang bertanda tangan di bawah ini, Petugas UKS ' + sekolah.namaSekolah + ', dengan ini merujuk peserta didik berikut ke fasilitas kesehatan untuk mendapatkan penanganan lebih lanjut:</p>';
    body += renderInfoTable_(r, [
      { key: 'namaSiswa', label: 'Nama Siswa' }, { key: 'nisn', label: 'NISN' },
      { key: 'kelas', label: 'Kelas' }, { key: 'jkUmur', label: 'JK / Umur' },
      { key: 'namaOrtu', label: 'Nama Orang Tua/Wali' }, { key: 'hpOrtu', label: 'No. Telepon Orang Tua' }
    ]);
    body += '<p><b>Dirujuk ke:</b><br>' + (r.namaFaskes || '-') + '<br>' + (r.alamatFaskes || '-') + '</p>';
    body += '<p><b>Hasil Pemeriksaan Awal UKS:</b><br>' + (r.hasilAwal || '-') + '</p>';
    body += '<p>Demikian surat rujukan ini dibuat untuk dipergunakan sebagaimana mestinya.</p>';
    body += '<p class="tgl-kanan">' + sekolah.kotaKabupaten + ', ' + (r.tanggal || tanggalCetakHariIni_()) + '</p>';
    body += ttdBlockHtml_([{ label: 'Petugas UKS' }]);
    return body;
  },
  SuratRujukanSakit: function (r, sekolah) {
    let body = suratHeaderHtml_(sekolah, 'Surat Izin / Rujukan Sakit', r.noSurat);
    body += '<p>Dengan ini menerangkan bahwa peserta didik berikut mengalami sakit dan diizinkan/dirujuk untuk mendapatkan penanganan kesehatan:</p>';
    body += renderInfoTable_(r, [
      { key: 'namaSiswa', label: 'Nama Lengkap Siswa' }, { key: 'nisn', label: 'NISN/NIK' },
      { key: 'kelas', label: 'Kelas/Rombel' }, { key: 'jkUmur', label: 'Jenis Kelamin/Umur' },
      { key: 'namaOrtu', label: 'Nama Orang Tua/Wali' }, { key: 'hpOrtu', label: 'No. Telepon Orang Tua' }
    ]);
    if (r.namaFaskes) body += '<p><b>Dirujuk ke:</b><br>' + r.namaFaskes + '<br>' + (r.alamatFaskes || '-') + '</p>';
    body += '<p>Demikian surat ini dibuat untuk dipergunakan sebagaimana mestinya.</p>';
    body += '<p class="tgl-kanan">' + sekolah.kotaKabupaten + ', ' + (r.tanggalWaktu || tanggalCetakHariIni_()) + '</p>';
    body += ttdBlockHtml_([{ label: 'Petugas UKS' }]);
    return body;
  },
  Kecelakaan: function (r, sekolah) {
    let body = suratHeaderHtml_(sekolah, 'Laporan Kejadian Kecelakaan / Cedera', null);
    body += renderInfoTable_(r, [
      { key: 'tanggalWaktu', label: 'Tanggal & Waktu' }, { key: 'namaSiswa', label: 'Nama Siswa' },
      { key: 'nisn', label: 'NISN' }, { key: 'kelas', label: 'Kelas' },
      { key: 'lokasi', label: 'Lokasi Kejadian' }, { key: 'jenisCedera', label: 'Jenis Cedera' },
      { key: 'keparahan', label: 'Tingkat Keparahan' }, { key: 'ortuDiberitahu', label: 'Orang Tua Diberitahu' }
    ]);
    body += '<p><b>Kronologi/Penyebab:</b><br>' + (r.kronologi || '-') + '</p>';
    body += '<p><b>Pertolongan Pertama (P3K):</b><br>' + (r.pertolongan || '-') + '</p>';
    body += '<p><b>Tindak Lanjut:</b><br>' + (r.tindakLanjut || '-') + '</p>';
    body += '<p class="tgl-kanan">' + sekolah.kotaKabupaten + ', ' + tanggalCetakHariIni_() + '</p>';
    body += ttdBlockHtml_([{ label: 'Petugas/Guru Penanggung Jawab', nama: r.petugas }]);
    return body;
  },
  PemeriksaanBerkala: function (r, sekolah) {
    let body = suratHeaderHtml_(sekolah, 'Hasil Pemeriksaan Kesehatan Berkala', null);
    body += renderInfoTable_(r, [
      { key: 'namaSiswa', label: 'Nama Siswa' }, { key: 'nisn', label: 'NISN' },
      { key: 'kelas', label: 'Kelas' }, { key: 'tanggal', label: 'Tanggal Periksa' },
      { key: 'tekananDarah', label: 'Tekanan Darah' }, { key: 'kebersihanTelinga', label: 'Kebersihan Telinga' },
      { key: 'kesehatanGigi', label: 'Kesehatan Gigi & Mulut' }, { key: 'kesehatanMata', label: 'Kesehatan Mata' },
      { key: 'kebersihanKulit', label: 'Kebersihan Kulit & Kuku' }, { key: 'hbAnemia', label: 'HB / Anemia' },
      { key: 'kesimpulan', label: 'Kesimpulan Kesehatan' }
    ]);
    body += '<p><b>Rekomendasi/Tindak Lanjut:</b><br>' + (r.rekomendasi || '-') + '</p>';
    body += '<p class="tgl-kanan">' + sekolah.kotaKabupaten + ', ' + (r.tanggal || tanggalCetakHariIni_()) + '</p>';
    body += ttdBlockHtml_([{ label: 'Petugas Pemeriksa' }]);
    return body;
  }
};

function cetakSurat(moduleKey, id) {
  const record = CURRENT_DATA.find(function (r) { return r.id === id; });
  if (!record) { showToast('Data tidak ditemukan', 'error'); return; }
  const builder = SURAT_BUILDERS[moduleKey];
  if (!builder) { showToast('Format cetak untuk modul ini belum tersedia', 'error'); return; }
  const bodyHtml = builder(record, APP_CONFIG.sekolah);
  const judulTab = (PRINT_CONFIG[moduleKey] && PRINT_CONFIG[moduleKey].label) || 'Cetak Data';
  openPrintWindow_(judulTab + ' - ' + (record.namaSiswa || record.id), bodyHtml);
}

// ---- Laporan rekap GENERIK (type: 'laporan') ----
// Otomatis memakai kolom modul yang sedang dibuka — termasuk pengelompokan
// kolom checkbox lewat groupDisplayColumns_ (sama seperti tabel di aplikasi) —
// jadi modul baru dengan type 'laporan' di PRINT_CONFIG langsung bisa dicetak
// TANPA perlu menulis template khusus.
function cetakLaporanGenerik(moduleKey) {
  const m = APP_CONFIG.modules[moduleKey];
  const data = CURRENT_DATA || [];
  const displayGroups = groupDisplayColumns_(m.columns);
  const sekolah = APP_CONFIG.sekolah;

  let body = suratHeaderHtml_(sekolah, m.label, null);
  body += '<p style="text-align:center;color:#555;">Tahun Ajaran ' + sekolah.tahunAjaran + ' — Dicetak ' + tanggalCetakHariIni_() +
    ' oleh ' + (CURRENT_USER ? CURRENT_USER.nama + ' (' + CURRENT_USER.role + ')' : '-') + '</p>';

  if (!data.length) {
    body += '<p style="text-align:center;color:#888;">Belum ada data untuk dicetak.</p>';
  } else {
    body += '<table class="rekap-table"><thead><tr>';
    displayGroups.forEach(function (g) {
      body += '<th>' + (g.type === 'checkboxGroup' ? 'Bulan Pelaksanaan' : g.column.label) + '</th>';
    });
    body += '</tr></thead><tbody>';
    data.forEach(function (row) {
      body += '<tr>';
      displayGroups.forEach(function (g) {
        if (g.type === 'checkboxGroup') {
          const labels = g.columns.filter(function (gc) { return isChecked_(row[gc.key]); }).map(function (gc) { return gc.label; });
          body += '<td>' + (labels.length ? labels.join(', ') : '-') + '</td>';
          return;
        }
        const val = row[g.column.key] !== undefined && row[g.column.key] !== null ? row[g.column.key] : '';
        body += '<td>' + val + '</td>';
      });
      body += '</tr>';
    });
    body += '</tbody></table>';
  }

  body += '<div style="margin-top:36px;">' + ttdBlockHtml_([
    { label: 'Mengetahui,', sublabel: 'Kepala Sekolah' },
    { label: 'Dibuat oleh,', sublabel: CURRENT_USER ? CURRENT_USER.role : '' }
  ]) + '</div>';

  openPrintWindow_('Laporan - ' + m.label, body);
}

// ---- Matriks Program Kerja UKS (khusus JadwalKegiatan, type: 'matrix') ----
// Meniru format matriks bulanan Jan-Des di template Excel asli sekolah:
// dikelompokkan per kategori "Program Kerja UKS", dengan kolom centang per bulan.
function cetakMatrixJadwal() {
  const sekolah = APP_CONFIG.sekolah;
  const data = CURRENT_DATA || [];
  const bulanCols = [
    { key: 'bulanJan', label: 'J' }, { key: 'bulanFeb', label: 'F' }, { key: 'bulanMar', label: 'M' },
    { key: 'bulanApr', label: 'A' }, { key: 'bulanMei', label: 'M' }, { key: 'bulanJun', label: 'J' },
    { key: 'bulanJul', label: 'J' }, { key: 'bulanAgs', label: 'A' }, { key: 'bulanSep', label: 'S' },
    { key: 'bulanOkt', label: 'O' }, { key: 'bulanNov', label: 'N' }, { key: 'bulanDes', label: 'D' }
  ];
  const programList = [
    '1. Pendidikan Kesehatan', '2. Pelayanan Kesehatan', '3. Pengadaan Sarana UKS',
    '4. Pembinaan dan Pengawasan Kebersihan', '5. Kegiatan Fisik (Olahraga)',
    '6. Kerjasama dengan Orangtua dan Masyarakat'
  ];

  let body = suratHeaderHtml_(sekolah, 'Matriks Program Kerja UKS', null);
  body += '<p style="text-align:center;color:#555;">Tahun Ajaran ' + sekolah.tahunAjaran + ' — Dicetak ' + tanggalCetakHariIni_() + '</p>';

  if (!data.length) {
    body += '<p style="text-align:center;color:#888;">Belum ada data untuk dicetak.</p>';
  } else {
    programList.forEach(function (prog) {
      const rows = data.filter(function (r) { return r.programKerja === prog; });
      if (!rows.length) return;
      body += '<h3>' + prog + '</h3>';
      body += '<table class="rekap-table"><thead><tr><th style="width:26px;">No</th><th>Nama Kegiatan</th><th>Sasaran</th><th>Pelaksana</th><th>Frekuensi</th>';
      bulanCols.forEach(function (b) { body += '<th class="center" style="width:20px;">' + b.label + '</th>'; });
      body += '<th>Keterangan</th></tr></thead><tbody>';
      rows.forEach(function (r, idx) {
        body += '<tr><td class="center">' + (idx + 1) + '</td><td>' + (r.namaKegiatan || '') + '</td><td>' + (r.sasaran || '') +
          '</td><td>' + (r.pj || '') + '</td><td>' + (r.frekuensi || '') + '</td>';
        bulanCols.forEach(function (b) { body += '<td class="center">' + (isChecked_(r[b.key]) ? '✓' : '') + '</td>'; });
        body += '<td>' + (r.keterangan || '') + '</td></tr>';
      });
      body += '</tbody></table>';
    });
  }

  body += '<div style="margin-top:36px;">' + ttdBlockHtml_([
    { label: 'Mengetahui,', sublabel: 'Kepala Sekolah' },
    { label: 'Ketua Pelaksana UKS', sublabel: '' }
  ]) + '</div>';

  openPrintWindow_('Matriks Program Kerja UKS', body);
}

// Dipanggil oleh tombol generik #reportBtn — memilih fungsi cetak yang sesuai
// berdasarkan type modul yang sedang dibuka di PRINT_CONFIG.
function cetakLaporanAtauMatrix() {
  const cfg = PRINT_CONFIG[CURRENT_MODULE];
  if (!cfg) return;
  if (cfg.type === 'matrix') cetakMatrixJadwal();
  else cetakLaporanGenerik(CURRENT_MODULE);
}

// ==================== IMPORT EXCEL/CSV — IMPLEMENTASI (BARU) ====================
// Membaca file Excel (.xlsx/.xls) atau CSV yang dipilih user (memakai SheetJS,
// sudah dimuat di Index.html), mencocokkan baris header file ke kolom modul yang
// sedang dibuka (lihat IMPORT_CONFIG di atas), menampilkan pratinjau, lalu
// mengirim semua baris sekaligus ke importRecords(token, sheetName, dataArray)
// di Code.gs.

// Samakan sebuah teks header supaya perbandingan tidak peduli spasi, tanda baca,
// dan besar/kecil huruf — mis. "TB (cm)" akan cocok dengan "tb(cm)" atau "TB cm".
function normalizeHeaderLabel_(s) {
  return String(s || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '');
}

// Cocokkan setiap kolom modul (berdasarkan LABEL atau KEY-nya di SHEET_CONFIG)
// dengan salah satu header di baris pertama file. Urutan kolom di file BEBAS —
// yang penting namanya cocok/mirip dengan label kolom modul.
function matchHeadersToColumns_(headerRow, columns) {
  const mapping = {}; // { columnKey: headerIndex }
  const usedHeaderIdx = {};
  columns.forEach(function (col) {
    const normLabel = normalizeHeaderLabel_(col.label);
    const normKey = normalizeHeaderLabel_(col.key);
    for (let i = 0; i < headerRow.length; i++) {
      if (usedHeaderIdx[i]) continue;
      const normHeader = normalizeHeaderLabel_(headerRow[i]);
      if (normHeader && (normHeader === normLabel || normHeader === normKey)) {
        mapping[col.key] = i;
        usedHeaderIdx[i] = true;
        break;
      }
    }
  });
  const missingRequired = columns.filter(function (col) {
    return col.required && mapping[col.key] === undefined;
  }).map(function (col) { return col.label; });
  const unmatchedHeaders = headerRow.filter(function (h, i) {
    return String(h).trim() !== '' && !usedHeaderIdx[i];
  });
  return { mapping: mapping, missingRequired: missingRequired, unmatchedHeaders: unmatchedHeaders };
}

function openImportModal() {
  IMPORT_STATE = { moduleKey: CURRENT_MODULE, mappedData: [], headerMapping: null };
  document.getElementById('importFileInput').value = '';
  document.getElementById('importPreviewArea').innerHTML = '';
  document.getElementById('importConfirmBtn').style.display = 'none';
  const m = APP_CONFIG.modules[CURRENT_MODULE];
  document.getElementById('importModalTitle').textContent = 'Import Data dari Excel — ' + (m ? m.label : '');
  document.getElementById('importModalOverlay').classList.add('show');
}

function closeImportModal() {
  document.getElementById('importModalOverlay').classList.remove('show');
  document.getElementById('importFileInput').value = '';
  IMPORT_STATE = { moduleKey: null, mappedData: [], headerMapping: null };
}

function handleImportFileSelected(ev) {
  const file = ev.target.files && ev.target.files[0];
  const previewArea = document.getElementById('importPreviewArea');
  const confirmBtn = document.getElementById('importConfirmBtn');
  confirmBtn.style.display = 'none';
  IMPORT_STATE.mappedData = [];
  if (!file) { previewArea.innerHTML = ''; return; }

  previewArea.innerHTML = '<div class="empty-state">Membaca file...</div>';
  const m = APP_CONFIG.modules[IMPORT_STATE.moduleKey];

  const reader = new FileReader();
  reader.onload = function (e) {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      // header:1 -> ambil sebagai array baris mentah (bukan objek), defval agar
      // sel kosong tidak hilang dari array (supaya index kolom tetap sejajar).
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false });

      const filledRows = rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ''; }); });
      if (!filledRows.length) throw new Error('File kosong atau tidak terbaca.');

      const headerRow = filledRows[0].map(function (h) { return String(h).trim(); });
      const dataRows = filledRows.slice(1);

      const matched = matchHeadersToColumns_(headerRow, m.columns);
      if (matched.missingRequired.length) {
        throw new Error('Kolom wajib tidak ditemukan di file: ' + matched.missingRequired.join(', ') +
          '. Pastikan baris pertama file berisi nama kolom yang sesuai (mis. "' + matched.missingRequired[0] + '").');
      }

      const mappedData = dataRows
        .filter(function (row) { return row.some(function (v) { return String(v).trim() !== ''; }); })
        .map(function (row) {
          const obj = {};
          m.columns.forEach(function (col) {
            const idx = matched.mapping[col.key];
            obj[col.key] = (idx !== undefined && row[idx] !== undefined) ? row[idx] : '';
          });
          return obj;
        });

      if (!mappedData.length) throw new Error('Tidak ada baris data untuk diimpor (file hanya berisi baris header).');

      IMPORT_STATE.mappedData = mappedData;
      IMPORT_STATE.headerMapping = matched;
      renderImportPreview_(m, matched, mappedData);
      confirmBtn.style.display = 'inline-flex';
      confirmBtn.disabled = false;
      confirmBtn.textContent = 'Impor Sekarang (' + mappedData.length + ' baris)';
    } catch (err) {
      previewArea.innerHTML = '<div class="empty-state" style="color:#d93025;">Gagal membaca file: ' + (err.message || err) + '</div>';
      IMPORT_STATE.mappedData = [];
    }
  };
  reader.onerror = function () {
    previewArea.innerHTML = '<div class="empty-state" style="color:#d93025;">Gagal membaca file.</div>';
  };
  reader.readAsArrayBuffer(file);
}

function renderImportPreview_(m, matched, mappedData) {
  const previewArea = document.getElementById('importPreviewArea');
  let html = '';

  if (matched.unmatchedHeaders.length) {
    html += '<div class="login-error" style="display:block;margin-bottom:12px;">' +
      'Kolom berikut di file tidak dikenali dan akan DIABAIKAN: ' + matched.unmatchedHeaders.join(', ') + '</div>';
  }

  html += '<p style="font-size:12.5px;color:#188038;font-weight:600;margin:0 0 8px;">' +
    mappedData.length + ' baris data siap diimpor. Periksa pratinjau di bawah sebelum melanjutkan.</p>';

  html += '<div class="table-wrap"><table><thead><tr>';
  m.columns.forEach(function (col) { html += '<th>' + col.label + '</th>'; });
  html += '</tr></thead><tbody>';
  mappedData.slice(0, 10).forEach(function (row) {
    html += '<tr>';
    m.columns.forEach(function (col) { html += '<td>' + (row[col.key] !== undefined ? row[col.key] : '') + '</td>'; });
    html += '</tr>';
  });
  html += '</tbody></table></div>';
  if (mappedData.length > 10) {
    html += '<p style="font-size:12px;color:#5f6368;margin-top:6px;">Menampilkan 10 dari ' + mappedData.length + ' baris. Sisanya tetap akan diimpor.</p>';
  }
  previewArea.innerHTML = html;
}

function confirmImport() {
  if (!IMPORT_STATE.mappedData.length) { showToast('Tidak ada data untuk diimpor', 'error'); return; }
  const confirmBtn = document.getElementById('importConfirmBtn');
  const originalText = confirmBtn.textContent;
  confirmBtn.disabled = true;
  confirmBtn.innerHTML = '<span class="spinner"></span>';
  showLoading('Mengimpor data...'); // (BARU)

  const moduleKey = IMPORT_STATE.moduleKey;
  const jumlahBaris = IMPORT_STATE.mappedData.length;

  google.script.run.withSuccessHandler(function (result) {
    confirmBtn.disabled = false;
    confirmBtn.textContent = originalText;
    hideLoading(); // (BARU)
    showToast((result && result.count ? result.count : jumlahBaris) + ' baris data berhasil diimpor', 'success');
    closeImportModal();
    if (CURRENT_MODULE === moduleKey) loadModuleData(moduleKey);
    // Segarkan cache autocomplete kalau yang diimpor adalah data Siswa/Docil.
    if (moduleKey === 'Siswa') loadSiswaList();
    if (moduleKey === 'DocilAnggota') loadDocilList();
    if (moduleKey === 'KaderAnggota') loadKaderList();
  }).withFailureHandler(function (e) {
    confirmBtn.disabled = false;
    confirmBtn.textContent = originalText;
    hideLoading(); // (BARU)
    onError(e);
  }).importRecords(SESSION_TOKEN, moduleKey, IMPORT_STATE.mappedData);
}

// ==================== HELPERS ====================
function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast show' + (type ? ' ' + type : '');
  // Pesan panjang (mis. alasan penolakan dari server) ditampilkan lebih lama supaya sempat terbaca.
  const durasi = Math.min(7000, Math.max(2500, String(msg).length * 55));
  setTimeout(function () { t.className = 'toast'; }, durasi);
}

function onError(e, moduleKeyForRetry) {
  console.error(e);
  const msg = (e && e.message) ? e.message : String(e);
  showToast('Terjadi kesalahan: ' + msg, 'error');

  if (moduleKeyForRetry) {
    const content = document.getElementById('content');
    if (content) {
      content.innerHTML =
        '<div class="empty-state" style="color:#c00;">' +
        'Gagal memuat data.<br><span style="font-size:12px;color:#888;">' + msg + '</span><br><br>' +
        '<button class="btn secondary small" onclick="navigate(\'' + moduleKeyForRetry + '\')">Coba Lagi</button>' +
        '</div>';
    }
  }

  if (typeof msg === 'string' && msg.indexOf('Sesi') > -1) {
    sessionStorage.removeItem('uksToken');
    SESSION_TOKEN = null;
    showLogin();
  }
}

function openPublicDashboard() {
  // Versi GitHub Pages: dashboard publik adalah halaman statis public.html di folder yang sama.
  const publicUrl = new URL('public.html', window.location.href).href;
  window.open(publicUrl, '_blank');
  showToast('Link Dashboard Publik dibuka — bisa dibagikan ke atasan', 'success');
}

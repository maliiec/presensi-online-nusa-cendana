<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Presensi Online - Nusa Cendana</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Font: Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
    <!-- SheetJS for Excel export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    
    <style>
        body { font-family: 'Inter', sans-serif; }
        .custom-scrollbar::-webkit-scrollbar { width: 6px; height: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; border-radius: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .custom-scrollbar::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    </style>
</head>
<body class="bg-slate-100 text-slate-800 min-h-screen flex flex-col antialiased">

    <header class="bg-[#0c2340] text-white shadow-lg sticky top-0 z-30 border-b border-amber-500/30">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3 flex justify-between items-center">
            <div class="flex items-center space-x-3">
                <!-- Logo Perusahaan -->
                <div class="bg-white p-1.5 rounded-xl shadow-md border border-amber-400/40 flex items-center justify-center min-w-[42px] h-11">
                    <img id="company-logo" src="Logo%20NC.jpeg" alt="Nusa Cendana" class="h-8 w-auto object-contain" onerror="handleLogoError(this)">
                </div>
                <div>
                    <h1 class="text-lg sm:text-xl font-black tracking-tight text-white flex items-center gap-1.5" id="company-title">
                        NUSA CENDANA
                    </h1>
                    <p class="text-[11px] text-amber-300/90 font-medium tracking-wide" id="header-subtitle">Portal Presensi Karyawan</p>
                </div>
            </div>
            
            <div class="flex items-center gap-3">
                <!-- Clock Display -->
                <div class="hidden md:flex flex-col text-right text-xs text-slate-200 pr-2 border-r border-slate-700">
                    <span id="live-time" class="font-bold text-sm text-amber-300">00:00:00</span>
                    <span id="live-date" class="text-[11px] text-slate-300">-</span>
                </div>

                <!-- Admin Mode Toggle Button -->
                <button id="btn-mode-toggle" onclick="toggleAdminMode()" class="flex items-center gap-2 bg-[#173860] hover:bg-[#1f487e] text-white px-3.5 py-2 rounded-xl text-xs font-bold transition border border-amber-500/30 shadow-sm active:scale-95">
                    <i class="fa-solid fa-user-gear text-amber-400" id="icon-mode"></i>
                    <span id="label-mode">Akses Admin</span>
                </button>
            </div>
        </div>
    </header>

    <main class="flex-grow max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">

        <!-- ========================================== -->
        <!-- VIEW KARYAWAN: Form Pengisian Absen        -->
        <!-- ========================================== -->
        <div id="view-karyawan" class="max-w-xl mx-auto space-y-5">
            
            <!-- Information Banner Nusa Cendana -->
            <div class="bg-gradient-to-r from-[#0c2340] to-[#1a3d6d] text-white p-5 rounded-2xl shadow-md flex items-center justify-between border-l-4 border-amber-400">
                <div>
                    <h2 class="text-base font-bold text-amber-300">Form Kehadiran Harian</h2>
                    <p class="text-xs text-slate-200 mt-1">Silakan isi data kehadiran karyawan Nusa Cendana hari ini.</p>
                </div>
                <div class="text-3xl text-amber-400 opacity-90">
                    <i class="fa-solid fa-calendar-check"></i>
                </div>
            </div>

            <!-- Card Form Karyawan -->
            <div class="bg-white p-6 sm:p-7 rounded-2xl shadow-sm border border-slate-200">
                <form id="attendance-form" onsubmit="submitAttendance(event)" class="space-y-4">
                    
                    <input type="hidden" id="entry-id" value="">

                    <!-- Nama Karyawan -->
                    <div>
                        <label for="nama" class="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wider">
                            Nama Lengkap <span class="text-rose-500">*</span>
                        </label>
                        <div class="relative">
                            <span class="absolute inset-y-0 left-0 flex items-center pl-3.5 text-slate-400">
                                <i class="fa-regular fa-user"></i>
                            </span>
                            <input type="text" id="nama" required placeholder="Contoh: Budi Santoso" class="w-full pl-10 pr-3.5 py-2.5 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition">
                        </div>
                    </div>

                    <!-- Jabatan & Departemen -->
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <!-- Jabatan -->
                        <div>
                            <label for="jabatan" class="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wider">
                                Jabatan <span class="text-rose-500">*</span>
                            </label>
                            <select id="jabatan" required class="w-full px-3.5 py-2.5 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition bg-white">
                                <option value="" disabled selected>-- Pilih Jabatan --</option>
                                <option value="Staff">Staff</option>
                                <option value="Manajer">Manajer</option>
                            </select>
                        </div>

                        <!-- Departemen -->
                        <div>
                            <label for="departemen" class="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wider">
                                Departemen <span class="text-rose-500">*</span>
                            </label>
                            <select id="departemen" required class="w-full px-3.5 py-2.5 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition bg-white">
                                <option value="" disabled selected>-- Pilih Departemen --</option>
                                <option value="Keuangan">Keuangan</option>
                                <option value="Administrasi">Administrasi</option>
                                <option value="Persediaan">Persediaan</option>
                                <option value="Lain-lain">Lain-lain</option>
                            </select>
                        </div>
                    </div>

                    <!-- Tanggal Presensi -->
                    <div>
                        <label for="tanggal" class="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wider">
                            Tanggal Kehadiran <span class="text-rose-500">*</span>
                        </label>
                        <input type="date" id="tanggal" required class="w-full px-3.5 py-2.5 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 focus:border-indigo-500 outline-none transition">
                    </div>

                    <!-- Waktu Presensi (Otomatis Tercatat Realtime) -->
                    <div class="bg-indigo-50/70 border border-indigo-100 rounded-xl p-3.5 flex items-center justify-between" id="time-inputs-container">
                        <div class="flex items-center gap-3">
                            <div class="w-10 h-10 bg-[#0c2340] text-amber-400 rounded-xl flex items-center justify-center text-base shadow-sm">
                                <i class="fa-regular fa-clock"></i>
                            </div>
                            <div>
                                <p class="text-xs font-bold text-slate-800 uppercase tracking-wider">Waktu Submit Presensi</p>
                                <p class="text-[11px] text-slate-500">Tercatat otomatis secara presisi saat tombol diklik.</p>
                            </div>
                        </div>
                        <div class="text-right">
                            <span id="auto-submit-time" class="text-base font-black text-indigo-900 font-mono">00:00:00</span>
                            <span class="block text-[10px] text-emerald-600 font-bold"><i class="fa-solid fa-lock mr-0.5"></i>Terkunci Sistem</span>
                        </div>
                    </div>

                    <!-- Status Kehadiran (Interactive Cards) -->
                    <div>
                        <label class="block text-xs font-bold text-slate-700 mb-2 uppercase tracking-wider">
                            Status Kehadiran <span class="text-rose-500">*</span>
                        </label>
                        <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
                            <label class="relative flex flex-col items-center justify-center p-3 rounded-xl border border-slate-200 cursor-pointer hover:bg-emerald-50 has-[:checked]:bg-emerald-100 has-[:checked]:border-emerald-500 transition">
                                <input type="radio" name="status" value="Hadir" checked class="sr-only" onchange="handleStatusChange()">
                                <i class="fa-solid fa-circle-check text-emerald-600 text-base mb-1"></i>
                                <span class="text-xs font-bold text-emerald-900">Hadir</span>
                            </label>

                            <label class="relative flex flex-col items-center justify-center p-3 rounded-xl border border-slate-200 cursor-pointer hover:bg-amber-50 has-[:checked]:bg-amber-100 has-[:checked]:border-amber-500 transition">
                                <input type="radio" name="status" value="Izin" class="sr-only" onchange="handleStatusChange()">
                                <i class="fa-solid fa-file-signature text-amber-600 text-base mb-1"></i>
                                <span class="text-xs font-bold text-amber-900">Izin</span>
                            </label>

                            <label class="relative flex flex-col items-center justify-center p-3 rounded-xl border border-slate-200 cursor-pointer hover:bg-blue-50 has-[:checked]:bg-blue-100 has-[:checked]:border-blue-500 transition">
                                <input type="radio" name="status" value="Sakit" class="sr-only" onchange="handleStatusChange()">
                                <i class="fa-solid fa-notes-medical text-blue-600 text-base mb-1"></i>
                                <span class="text-xs font-bold text-blue-900">Sakit</span>
                            </label>

                            <label class="relative flex flex-col items-center justify-center p-3 rounded-xl border border-slate-200 cursor-pointer hover:bg-rose-50 has-[:checked]:bg-rose-100 has-[:checked]:border-rose-500 transition">
                                <input type="radio" name="status" value="Alpa" class="sr-only" onchange="handleStatusChange()">
                                <i class="fa-solid fa-user-slash text-rose-600 text-base mb-1"></i>
                                <span class="text-xs font-bold text-rose-900">Alpa</span>
                            </label>
                        </div>
                    </div>

                    <!-- Catatan Tambahan -->
                    <div>
                        <label for="catatan" class="block text-xs font-bold text-slate-700 mb-1.5 uppercase tracking-wider">Catatan / Keterangan (Opsional)</label>
                        <textarea id="catatan" rows="2" placeholder="Tulis alasan jika Izin/Sakit atau catatan tambahan..." class="w-full px-3.5 py-2 text-sm border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none transition resize-none"></textarea>
                    </div>

                    <!-- Submit Button -->
                    <div class="pt-2">
                        <button type="submit" id="btn-submit" class="w-full bg-[#0c2340] hover:bg-[#163761] text-white py-3.5 px-4 rounded-xl font-bold text-sm transition shadow-lg flex justify-center items-center gap-2 active:scale-[0.99]">
                            <i class="fa-solid fa-paper-plane text-amber-400"></i>
                            <span id="btn-submit-text">Kirim Kehadiran Sekarang</span>
                        </button>
                    </div>
                </form>
            </div>
        </div>

        <!-- ========================================== -->
        <!-- VIEW ADMIN: Dashboard Pemantauan Absensi   -->
        <!-- ========================================== -->
        <div id="view-admin" class="hidden space-y-6">
            
            <!-- Admin Alert & Config Banner -->
            <div class="bg-slate-900 text-white rounded-2xl p-4 sm:p-5 shadow-md flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-l-4 border-amber-400">
                <div>
                    <div class="flex items-center gap-2">
                        <span class="px-2.5 py-0.5 bg-amber-500 text-slate-950 rounded-md text-[10px] font-black uppercase tracking-wider">Mode Pengawasan Admin</span>
                        <h2 class="text-base font-bold">Dashboard Pemantauan Absensi Karyawan</h2>
                    </div>
                    <p class="text-xs text-slate-300 mt-1">Pantau seluruh data kehadiran karyawan Nusa Cendana secara terpusat.</p>
                </div>

                <div class="flex flex-wrap items-center gap-2 w-full md:w-auto">
                    <button onclick="openLogoModal()" class="flex-1 md:flex-initial flex items-center justify-center gap-1.5 bg-amber-500 hover:bg-amber-600 text-slate-950 px-3.5 py-2 rounded-xl text-xs font-bold transition shadow-sm">
                        <i class="fa-solid fa-image"></i>
                        <span>Atur Logo Perusahaan</span>
                    </button>
                    <button onclick="openGoogleSheetModal()" class="flex-1 md:flex-initial flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-2 rounded-xl text-xs font-semibold transition shadow-sm">
                        <i class="fa-solid fa-file-csv"></i>
                        <span>Google Sheet</span>
                    </button>
                    <button onclick="changeAdminPinModal()" class="flex-1 md:flex-initial flex items-center justify-center gap-1.5 bg-slate-800 hover:bg-slate-700 text-white px-3.5 py-2 rounded-xl text-xs font-semibold transition shadow-sm border border-slate-700">
                        <i class="fa-solid fa-key"></i>
                        <span>Ubah PIN</span>
                    </button>
                </div>
            </div>

            <!-- Ringkasan Statistik Cards -->
            <div class="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
                <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-[11px] font-bold uppercase tracking-wider text-slate-400">Total Hadir</p>
                        <h3 id="stat-hadir" class="text-2xl font-black text-emerald-600 mt-0.5">0</h3>
                    </div>
                    <div class="w-11 h-11 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-user-check"></i>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-[11px] font-bold uppercase tracking-wider text-slate-400">Total Izin</p>
                        <h3 id="stat-izin" class="text-2xl font-black text-amber-500 mt-0.5">0</h3>
                    </div>
                    <div class="w-11 h-11 rounded-xl bg-amber-100 text-amber-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-file-signature"></i>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-[11px] font-bold uppercase tracking-wider text-slate-400">Total Sakit</p>
                        <h3 id="stat-sakit" class="text-2xl font-black text-blue-600 mt-0.5">0</h3>
                    </div>
                    <div class="w-11 h-11 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-notes-medical"></i>
                    </div>
                </div>

                <div class="bg-white p-4 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
                    <div>
                        <p class="text-[11px] font-bold uppercase tracking-wider text-slate-400">Total Alpa</p>
                        <h3 id="stat-alpa" class="text-2xl font-black text-rose-600 mt-0.5">0</h3>
                    </div>
                    <div class="w-11 h-11 rounded-xl bg-rose-100 text-rose-600 flex items-center justify-center text-xl">
                        <i class="fa-solid fa-user-xmark"></i>
                    </div>
                </div>
            </div>

            <!-- Tabel Monitoring Absensi -->
            <div class="bg-white p-5 rounded-2xl shadow-sm border border-slate-200 space-y-4">
                
                <!-- Table Header Actions -->
                <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 pb-3 border-b border-slate-100">
                    <div>
                        <h3 class="text-base font-bold text-slate-800 flex items-center gap-2">
                            <i class="fa-solid fa-list-check text-indigo-600"></i>
                            Daftar Siapa Saja Yang Telah Absen
                        </h3>
                        <p class="text-xs text-slate-500">Data berikut tercatat secara otomatis saat karyawan mengisi form.</p>
                    </div>

                    <div class="flex items-center gap-2 w-full sm:w-auto">
                        <button onclick="exportToExcel()" class="flex-1 sm:flex-initial inline-flex items-center justify-center gap-1.5 px-3.5 py-2 bg-emerald-50 text-emerald-700 border border-emerald-200 hover:bg-emerald-100 rounded-xl text-xs font-bold transition">
                            <i class="fa-solid fa-file-excel text-emerald-600"></i>
                            <span>Download Excel</span>
                        </button>
                        <button onclick="syncAllToSheet()" class="flex-1 sm:flex-initial inline-flex items-center justify-center gap-1.5 px-3.5 py-2 bg-indigo-50 text-indigo-700 border border-indigo-200 hover:bg-indigo-100 rounded-xl text-xs font-bold transition">
                            <i class="fa-solid fa-cloud-arrow-up text-indigo-600"></i>
                            <span>Kirim ke Sheet</span>
                        </button>
                    </div>
                </div>

                <!-- Filter Controls -->
                <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-2.5">
                    <!-- Search Input -->
                    <div class="relative">
                        <span class="absolute inset-y-0 left-0 flex items-center pl-3 text-slate-400 text-xs">
                            <i class="fa-solid fa-magnifying-glass"></i>
                        </span>
                        <input type="text" id="filter-search" oninput="applyFilters()" placeholder="Cari nama karyawan..." class="w-full pl-8 pr-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-1 focus:ring-indigo-500 outline-none">
                    </div>

                    <!-- Filter Departemen -->
                    <select id="filter-dept" onchange="applyFilters()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-1 focus:ring-indigo-500 outline-none bg-white">
                        <option value="ALL">Semua Departemen</option>
                        <option value="Keuangan">Keuangan</option>
                        <option value="Administrasi">Administrasi</option>
                        <option value="Persediaan">Persediaan</option>
                        <option value="Lain-lain">Lain-lain</option>
                    </select>

                    <!-- Filter Status -->
                    <select id="filter-status" onchange="applyFilters()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-1 focus:ring-indigo-500 outline-none bg-white">
                        <option value="ALL">Semua Status</option>
                        <option value="Hadir">Hadir</option>
                        <option value="Izin">Izin</option>
                        <option value="Sakit">Sakit</option>
                        <option value="Alpa">Alpa</option>
                    </select>

                    <!-- Filter Tanggal -->
                    <input type="date" id="filter-tanggal" onchange="applyFilters()" class="w-full px-3 py-2 text-xs border border-slate-300 rounded-xl focus:ring-1 focus:ring-indigo-500 outline-none">
                </div>

                <!-- Data Table -->
                <div class="overflow-x-auto custom-scrollbar border border-slate-200 rounded-xl">
                    <table class="w-full text-left text-xs">
                        <thead class="bg-slate-50 text-slate-600 font-bold uppercase tracking-wider border-b border-slate-200">
                            <tr>
                                <th class="p-3.5">Tanggal / Waktu</th>
                                <th class="p-3.5">Nama Karyawan</th>
                                <th class="p-3.5">Jabatan & Dept</th>
                                <th class="p-3.5 text-center">Status</th>
                                <th class="p-3.5">Catatan</th>
                                <th class="p-3.5 text-center">Aksi Admin</th>
                            </tr>
                        </thead>
                        <tbody id="attendance-tbody" class="divide-y divide-slate-100">
                            <!-- Dynamic Content Rendered Here -->
                        </tbody>
                    </table>
                </div>

                <!-- Empty State Message -->
                <div id="empty-state" class="hidden text-center py-10">
                    <div class="w-16 h-16 bg-slate-100 text-slate-400 rounded-full flex items-center justify-center mx-auto mb-3 text-2xl">
                        <i class="fa-regular fa-folder-open"></i>
                    </div>
                    <p class="text-sm font-bold text-slate-600">Belum Ada Data Presensi</p>
                    <p class="text-xs text-slate-400 mt-1">Data akan otomatis muncul di sini begitu karyawan mengisi form absen.</p>
                </div>

                <div class="flex justify-between items-center text-xs text-slate-500 pt-1">
                    <span>Total Ditampilkan: <strong id="filtered-count" class="text-slate-800">0</strong> data</span>
                    <button onclick="clearAllData()" class="text-rose-600 hover:underline text-[11px] font-medium">Hapus Semua Data</button>
                </div>
            </div>
        </div>

    </main>

    <!-- MODAL 1: Login PIN Admin -->
    <div id="modal-pin" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-sm w-full p-6 shadow-2xl border border-slate-100 text-center space-y-4">
            <div class="w-12 h-12 bg-indigo-100 text-indigo-600 rounded-2xl flex items-center justify-center mx-auto text-xl">
                <i class="fa-solid fa-lock"></i>
            </div>
            <div>
                <h3 class="text-lg font-bold text-slate-800">Verifikasi Akses Admin</h3>
                <p class="text-xs text-slate-500 mt-1">Masukkan PIN keamanan untuk membuka portal pemantauan admin.</p>
                <div class="mt-2 inline-block px-2.5 py-1 bg-amber-50 border border-amber-200 text-amber-800 rounded-lg text-[11px] font-mono font-bold">
                    PIN Bawaan: 1234
                </div>
            </div>

            <div>
                <input type="password" id="input-pin" maxlength="6" placeholder="Masukkan PIN" class="w-full text-center tracking-widest text-xl font-bold py-3 border border-slate-300 rounded-xl focus:ring-2 focus:ring-indigo-500 outline-none">
                <p id="pin-error-msg" class="text-xs text-rose-600 font-semibold mt-1.5 hidden"><i class="fa-solid fa-circle-exclamation mr-1"></i>PIN salah! Coba lagi.</p>
            </div>

            <div class="flex gap-2">
                <button onclick="closePinModal()" class="flex-1 py-2.5 border border-slate-300 rounded-xl text-slate-600 font-bold text-xs hover:bg-slate-50 transition">
                    Batal
                </button>
                <button onclick="verifyAdminPin()" class="flex-1 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl font-bold text-xs transition shadow-md">
                    Masuk Admin
                </button>
            </div>
        </div>
    </div>

    <!-- MODAL 2: Sukses Kirim Karyawan -->
    <div id="modal-success" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-100 text-center space-y-4">
            <div class="w-16 h-16 bg-emerald-100 text-emerald-600 rounded-full flex items-center justify-center mx-auto text-3xl">
                <i class="fa-solid fa-circle-check"></i>
            </div>
            <div>
                <h3 class="text-xl font-bold text-slate-800">Presensi Berhasil Dicatat!</h3>
                <p class="text-xs text-slate-600 mt-1" id="success-modal-desc">Terima kasih, data kehadiran Anda telah tersimpan ke sistem.</p>
            </div>
            <button onclick="closeSuccessModal()" class="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl font-bold text-xs transition shadow-md">
                Selesai / Tutup
            </button>
        </div>
    </div>

    <!-- MODAL 3: Pengaturan Google Sheet -->
    <div id="modal-sheet" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-lg w-full p-6 shadow-2xl border border-slate-100 space-y-4 max-h-[90vh] overflow-y-auto custom-scrollbar">
            <div class="flex justify-between items-center border-b pb-3">
                <div class="flex items-center gap-2">
                    <i class="fa-solid fa-file-csv text-emerald-600 text-lg"></i>
                    <h3 class="text-base font-bold text-slate-800">Integrasi Google Sheet</h3>
                </div>
                <button onclick="closeGoogleSheetModal()" class="text-slate-400 hover:text-slate-600">
                    <i class="fa-solid fa-xmark text-lg"></i>
                </button>
            </div>

            <div class="space-y-3 text-xs">
                <div>
                    <label class="block font-bold text-slate-700 mb-1">1. Salin Kode Google Apps Script Berikut:</label>
                    <div class="relative">
                        <textarea id="apps-script-code" readonly rows="5" class="w-full font-mono text-[11px] bg-slate-900 text-emerald-400 p-3 rounded-xl border border-slate-800 outline-none resize-none">
function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  if (data.test) return ContentService.createTextOutput("OK");
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["ID", "Tanggal", "Nama", "Jabatan", "Departemen", "Jam Masuk", "Jam Keluar", "Status", "Catatan"]);
  }
  sheet.appendRow([data.id, data.tanggal, data.nama, data.jabatan, data.departemen, data.jamMasuk, data.jamKeluar, data.status, data.catatan]);
  return ContentService.createTextOutput("Success");
}
                        </textarea>
                        <button onclick="copyAppsScript()" class="absolute top-2 right-2 bg-slate-800 hover:bg-slate-700 text-white px-2.5 py-1 rounded-lg text-[10px] font-semibold border border-slate-700">
                            <i class="fa-regular fa-copy mr-1"></i> Salin Skrip
                        </button>
                    </div>
                    <p class="text-[10px] text-slate-500 mt-1">Buka Google Sheet Anda > Extensions > Apps Script > Paste skrip ini > Deploy as Web App (Anyone).</p>
                </div>

                <div>
                    <label class="block font-bold text-slate-700 mb-1">2. Tempelkan Webhook URL Google Apps Script:</label>
                    <input type="url" id="webhook-url" placeholder="https://script.google.com/macros/s/.../exec" class="w-full px-3 py-2 border border-slate-300 rounded-xl outline-none focus:ring-2 focus:ring-emerald-500 text-xs">
                </div>

                <div id="connection-status" class="hidden p-2.5 rounded-xl text-xs font-medium"></div>
            </div>

            <div class="flex gap-2 pt-2 border-t">
                <button onclick="testConnection()" class="px-3.5 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-xs font-bold transition">
                    Uji Koneksi
                </button>
                <button onclick="saveWebhookUrl()" class="flex-1 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold shadow-md transition">
                    Simpan Webhook
                </button>
            </div>
        </div>
    </div>

    <!-- MODAL 4: Pengaturan Logo & Nama Perusahaan -->
    <div id="modal-logo" class="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-slate-100 space-y-4">
            <div class="flex justify-between items-center border-b pb-3">
                <div class="flex items-center gap-2">
                    <i class="fa-solid fa-image text-amber-500 text-lg"></i>
                    <h3 class="text-base font-bold text-slate-800">Pengaturan Logo & Brand</h3>
                </div>
                <button onclick="closeLogoModal()" class="text-slate-400 hover:text-slate-600">
                    <i class="fa-solid fa-xmark text-lg"></i>
                </button>
            </div>

            <div class="space-y-3 text-xs">
                <div>
                    <label class="block font-bold text-slate-700 mb-1">Nama Perusahaan / Organisasi:</label>
                    <input type="text" id="input-company-name" placeholder="NUSA CENDANA" class="w-full px-3 py-2 border border-slate-300 rounded-xl outline-none focus:ring-2 focus:ring-amber-500 text-xs font-semibold">
                </div>

                <div>
                    <label class="block font-bold text-slate-700 mb-1">Upload Gambar Logo Baru:</label>
                    <input type="file" id="input-logo-file" accept="image/*" onchange="previewUploadedLogo(event)" class="w-full text-xs text-slate-500 file:mr-3 file:py-2 file:px-3 file:rounded-xl file:border-0 file:text-xs file:font-semibold file:bg-amber-50 file:text-amber-700 hover:file:bg-amber-100">
                    <p class="text-[10px] text-slate-400 mt-1">Atau letakkan file bernama <strong>Logo NC.jpeg</strong> di folder GitHub yang sama.</p>
                </div>

                <!-- Preview Box -->
                <div class="bg-slate-50 p-3 rounded-xl border border-slate-200 text-center">
                    <p class="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-2">Pratinjau Logo Saat Ini:</p>
                    <img id="logo-preview" src="Logo NC.jpeg" alt="Preview Logo" class="h-12 mx-auto object-contain" onerror="handleLogoError(this)">
                </div>
            </div>

            <div class="flex gap-2 pt-2 border-t">
                <button onclick="resetLogoToDefault()" class="px-3 py-2 border border-slate-300 text-slate-600 rounded-xl text-xs font-bold hover:bg-slate-50">
                    Reset
                </button>
                <button onclick="saveLogoSettings()" class="flex-1 py-2 bg-amber-500 hover:bg-amber-600 text-slate-950 rounded-xl text-xs font-bold shadow-md">
                    Simpan Perubahan
                </button>
            </div>
        </div>
    </div>

    <!-- Toast Notification -->
    <div id="toast" class="fixed bottom-5 right-5 z-50 flex items-center gap-3 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl transition-all duration-300 transform translate-y-20 opacity-0 pointer-events-none">
        <i id="toast-icon" class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
        <span id="toast-msg" class="text-xs font-semibold">Notifikasi</span>
    </div>

    <script>
        let isAdminLoggedIn = false;
        let adminPin = localStorage.getItem('presensi_admin_pin') || '1234';
        let webhookUrl = localStorage.getItem('presensi_webhook_url') || '';
        let companyName = localStorage.getItem('presensi_company_name') || 'NUSA CENDANA';
        
        // Logo bawaan SVG Nusa Cendana yang di-encode secara aman agar pasti ke-render 100%
        const defaultLogoSvg = `data:image/svg+xml;charset=utf-8,${encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 320 160"><g transform="translate(10, 0)"><path d="M 40,20 L 85,20 L 170,120 L 170,20 L 210,20 L 210,135 L 165,135 L 80,35 L 80,135 L 40,135 Z" fill="#0c2340"/><path d="M 10,95 C 40,130 110,125 190,80 C 220,65 240,48 255,30 C 240,38 220,48 200,52 C 220,38 235,25 240,12 C 220,25 200,35 180,40 C 110,68 50,82 10,95 Z" fill="#d97706"/><path d="M 10,98 C 40,125 105,120 180,80 C 120,105 60,110 10,98 Z" fill="#f59e0b"/></g><text x="160" y="152" font-family="Inter, sans-serif" font-weight="900" font-size="24" fill="#0c2340" text-anchor="middle" letter-spacing="1">NUSA CENDANA</text></svg>')}`;
        
        let companyLogo = localStorage.getItem('presensi_company_logo') || 'https://i.postimg.cc/8sZ8m4hD/Logo-NC.jpeg';
        
        // Data Presensi Lokal
        let attendanceData = JSON.parse(localStorage.getItem('presensi_data')) || [
            {
                id: 'PRE-1001',
                nama: 'Budi Santoso',
                jabatan: 'Manajer',
                departemen: 'Keuangan',
                tanggal: new Date().toISOString().split('T')[0],
                jamMasuk: '08:00',
                jamKeluar: '-',
                status: 'Hadir',
                catatan: 'Hadir tepat waktu',
                synced: false
            }
        ];

        // Init App On Window Load
        window.addEventListener('DOMContentLoaded', () => {
            applyLogoAndBranding();

            const today = new Date().toISOString().split('T')[0];
            const tanggalInput = document.getElementById('tanggal');
            if (tanggalInput) tanggalInput.value = today;
            
            const webhookInput = document.getElementById('webhook-url');
            if (webhookInput && webhookUrl) {
                webhookInput.value = webhookUrl;
            }

            startClock();
            updateViewMode();
        });

        function formatLogoUrl(url) {
            if (!url) return 'https://i.postimg.cc/8sZ8m4hD/Logo-NC.jpeg';
            if (url.includes('postimg.cc/8sZ8m4hD')) {
                return 'https://i.postimg.cc/8sZ8m4hD/Logo-NC.jpeg';
            }
            if (url.includes('postimg.cc/') && !url.includes('i.postimg.cc/')) {
                const match = url.match(/postimg\.cc\/([a-zA-Z0-9]+)/);
                if (match && match[1]) {
                    return `https://i.postimg.cc/${match[1]}/Logo-NC.jpeg`;
                }
            }
            return url;
        }

        function applyLogoAndBranding() {
            const formattedLogo = formatLogoUrl(companyLogo);
            const compTitle = document.getElementById('company-title');
            const compLogo = document.getElementById('company-logo');
            const logoPrev = document.getElementById('logo-preview');
            const inputCompName = document.getElementById('input-company-name');

            if (compTitle) compTitle.textContent = companyName;
            if (compLogo) compLogo.src = formattedLogo;
            if (logoPrev) logoPrev.src = formattedLogo;
            if (inputCompName) inputCompName.value = companyName;
        }

        function handleLogoError(img) {
            img.onerror = null;
            img.src = defaultLogoSvg;
        }

        function openLogoModal() {
            document.getElementById('modal-logo').classList.remove('hidden');
        }

        function closeLogoModal() {
            document.getElementById('modal-logo').classList.add('hidden');
        }

        function previewUploadedLogo(event) {
            const file = event.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    document.getElementById('logo-preview').src = e.target.result;
                    companyLogo = e.target.result;
                };
                reader.readAsDataURL(file);
            }
        }

        function saveLogoSettings() {
            const nameInput = document.getElementById('input-company-name').value.trim();
            if (nameInput) {
                companyName = nameInput;
            }
            
            localStorage.setItem('presensi_company_name', companyName);
            localStorage.setItem('presensi_company_logo', companyLogo);
            
            applyLogoAndBranding();
            closeLogoModal();
            showToast("Pengaturan Logo & Nama Perusahaan berhasil disimpan!");
        }

        function resetLogoToDefault() {
            companyName = 'NUSA CENDANA';
            companyLogo = 'https://i.postimg.cc/8sZ8m4hD/Logo-NC.jpeg';
            localStorage.removeItem('presensi_company_name');
            localStorage.removeItem('presensi_company_logo');
            applyLogoAndBranding();
            closeLogoModal();
            showToast("Logo & Nama berhasil dikembalikan ke bawaan Nusa Cendana.");
        }

        function startClock() {
            function update() {
                const now = new Date();
                const timeStr = now.toLocaleTimeString('id-ID', { hour12: false });
                const dateStr = now.toLocaleDateString('id-ID', { day: 'numeric', month: 'short', year: 'numeric' });
                
                const liveTime = document.getElementById('live-time');
                const liveDate = document.getElementById('live-date');
                const autoTime = document.getElementById('auto-submit-time');

                if (liveTime) liveTime.textContent = timeStr;
                if (liveDate) liveDate.textContent = dateStr;
                if (autoTime) autoTime.textContent = timeStr;
            }
            update();
            setInterval(update, 1000);
        }

        function handleStatusChange() {
            // Waktu presensi dikunci secara otomatis oleh sistem saat tombol kirim diklik
        }

        function toggleAdminMode() {
            if (isAdminLoggedIn) {
                isAdminLoggedIn = false;
                updateViewMode();
                showToast("Anda telah keluar dari Mode Admin.", "info");
            } else {
                document.getElementById('input-pin').value = '';
                document.getElementById('pin-error-msg').classList.add('hidden');
                document.getElementById('modal-pin').classList.remove('hidden');
                setTimeout(() => document.getElementById('input-pin').focus(), 100);
            }
        }

        function closePinModal() {
            document.getElementById('modal-pin').classList.add('hidden');
        }

        function verifyAdminPin() {
            const inputPin = document.getElementById('input-pin').value.trim();
            if (inputPin === adminPin) {
                isAdminLoggedIn = true;
                closePinModal();
                updateViewMode();
                showToast("Berhasil masuk sebagai Admin!");
            } else {
                document.getElementById('pin-error-msg').classList.remove('hidden');
            }
        }

        function updateViewMode() {
            const viewKaryawan = document.getElementById('view-karyawan');
            const viewAdmin = document.getElementById('view-admin');
            const btnModeToggle = document.getElementById('btn-mode-toggle');
            const iconMode = document.getElementById('icon-mode');
            const labelMode = document.getElementById('label-mode');
            const headerSubtitle = document.getElementById('header-subtitle');

            if (isAdminLoggedIn) {
                viewKaryawan.classList.add('hidden');
                viewAdmin.classList.remove('hidden');
                
                btnModeToggle.className = "flex items-center gap-2 bg-rose-600 hover:bg-rose-700 text-white px-3.5 py-2 rounded-xl text-xs font-bold transition border border-rose-500/40 shadow-sm";
                iconMode.className = "fa-solid fa-right-from-bracket";
                labelMode.textContent = "Keluar Admin";
                headerSubtitle.textContent = "Sistem Pengawasan Admin";

                renderTable();
                updateStats();
            } else {
                viewAdmin.classList.add('hidden');
                viewKaryawan.classList.remove('hidden');

                btnModeToggle.className = "flex items-center gap-2 bg-[#173860] hover:bg-[#1f487e] text-white px-3.5 py-2 rounded-xl text-xs font-bold transition border border-amber-500/30 shadow-sm";
                iconMode.className = "fa-solid fa-user-gear text-amber-400";
                labelMode.textContent = "Akses Admin";
                headerSubtitle.textContent = "Portal Presensi Karyawan";
            }
        }

        async function submitAttendance(event) {
            event.preventDefault();

            const entryId = document.getElementById('entry-id').value;
            const nama = document.getElementById('nama').value.trim();
            const jabatan = document.getElementById('jabatan').value;
            const departemen = document.getElementById('departemen').value;
            const tanggal = document.getElementById('tanggal').value;
            
            // Otomatis mengambil jam & menit saat ini dari perangkat saat tombol submit diklik
            const now = new Date();
            const hours = String(now.getHours()).padStart(2, '0');
            const minutes = String(now.getMinutes()).padStart(2, '0');
            const exactCurrentTime = `${hours}:${minutes}`;

            const status = document.querySelector('input[name="status"]:checked').value;
            const catatan = document.getElementById('catatan').value.trim();

            if (!nama || !jabatan || !departemen || !tanggal) {
                showToast("Mohon isi semua bidang formulir yang wajib!", "error");
                return;
            }

            const newEntry = {
                id: entryId || 'PRE-' + Date.now().toString().slice(-6),
                nama,
                jabatan,
                departemen,
                tanggal,
                jamMasuk: status === 'Hadir' ? exactCurrentTime : '-',
                jamKeluar: '-',
                status,
                catatan: catatan || '-',
                synced: false
            };

            if (entryId) {
                const idx = attendanceData.findIndex(item => item.id === entryId);
                if (idx !== -1) attendanceData[idx] = newEntry;
            } else {
                attendanceData.unshift(newEntry);
            }

            saveToLocalStorage();

            if (webhookUrl) {
                sendToGoogleSheet(newEntry);
            }

            document.getElementById('success-modal-desc').textContent = `Terima kasih ${nama}, presensi (${status}) Anda pada tanggal ${tanggal} telah tersimpan.`;
            document.getElementById('modal-success').classList.remove('hidden');

            resetForm();

            if (isAdminLoggedIn) {
                renderTable();
                updateStats();
            }
        }

        function closeSuccessModal() {
            document.getElementById('modal-success').classList.add('hidden');
        }

        function resetForm() {
            document.getElementById('attendance-form').reset();
            document.getElementById('entry-id').value = '';
            document.getElementById('btn-submit-text').textContent = "Kirim Kehadiran Sekarang";
            document.getElementById('tanggal').value = new Date().toISOString().split('T')[0];
            handleStatusChange();
        }

        function renderTable(dataToRender = null) {
            const tbody = document.getElementById('attendance-tbody');
            const emptyState = document.getElementById('empty-state');
            const countSpan = document.getElementById('filtered-count');

            const list = dataToRender || attendanceData;
            tbody.innerHTML = '';

            if (list.length === 0) {
                emptyState.classList.remove('hidden');
                countSpan.textContent = '0';
                return;
            } else {
                emptyState.classList.add('hidden');
            }

            countSpan.textContent = list.length;

            list.forEach(item => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50 transition border-b border-slate-100";

                let statusBadge = '';
                if (item.status === 'Hadir') {
                    statusBadge = `<span class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800">Hadir</span>`;
                } else if (item.status === 'Izin') {
                    statusBadge = `<span class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800">Izin</span>`;
                } else if (item.status === 'Sakit') {
                    statusBadge = `<span class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-blue-100 text-blue-800">Sakit</span>`;
                } else {
                    statusBadge = `<span class="px-2.5 py-1 rounded-full text-[11px] font-bold bg-rose-100 text-rose-800">Alpa</span>`;
                }

                tr.innerHTML = `
                    <td class="p-3.5">
                        <div class="font-bold text-slate-800">${item.tanggal}</div>
                        <div class="text-[11px] text-slate-500">
                            ${item.status === 'Hadir' ? `<i class="fa-regular fa-clock mr-1"></i>${item.jamMasuk} - ${item.jamKeluar}` : '-'}
                        </div>
                    </td>
                    <td class="p-3.5">
                        <div class="font-extrabold text-slate-900">${escapeHtml(item.nama)}</div>
                        <div class="text-[10px] text-slate-400 font-mono">${item.id}</div>
                    </td>
                    <td class="p-3.5">
                        <div class="font-semibold text-slate-700">${escapeHtml(item.jabatan)}</div>
                        <div class="text-[11px] text-indigo-600 font-bold">${escapeHtml(item.departemen)}</div>
                    </td>
                    <td class="p-3.5 text-center">
                        ${statusBadge}
                    </td>
                    <td class="p-3.5 text-slate-600 max-w-[140px] truncate" title="${escapeHtml(item.catatan)}">
                        ${escapeHtml(item.catatan)}
                    </td>
                    <td class="p-3.5 text-center">
                        <div class="flex items-center justify-center gap-1.5">
                            <button onclick="editEntry('${item.id}')" class="p-1.5 text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 rounded-lg transition" title="Edit">
                                <i class="fa-regular fa-pen-to-square"></i>
                            </button>
                            <button onclick="deleteEntry('${item.id}')" class="p-1.5 text-slate-500 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition" title="Hapus">
                                <i class="fa-regular fa-trash-can"></i>
                            </button>
                        </div>
                    </td>
                `;

                tbody.appendChild(tr);
            });
        }

        function applyFilters() {
            const searchVal = document.getElementById('filter-search').value.toLowerCase();
            const deptVal = document.getElementById('filter-dept').value;
            const statusVal = document.getElementById('filter-status').value;
            const dateVal = document.getElementById('filter-tanggal').value;

            const filtered = attendanceData.filter(item => {
                const matchSearch = item.nama.toLowerCase().includes(searchVal);
                const matchDept = (deptVal === 'ALL') || (item.departemen === deptVal);
                const matchStatus = (statusVal === 'ALL') || (item.status === statusVal);
                const matchDate = (!dateVal) || (item.tanggal === dateVal);

                return matchSearch && matchDept && matchStatus && matchDate;
            });

            renderTable(filtered);
        }

        function updateStats() {
            const hadir = attendanceData.filter(i => i.status === 'Hadir').length;
            const izin = attendanceData.filter(i => i.status === 'Izin').length;
            const sakit = attendanceData.filter(i => i.status === 'Sakit').length;
            const alpa = attendanceData.filter(i => i.status === 'Alpa').length;

            document.getElementById('stat-hadir').textContent = hadir;
            document.getElementById('stat-izin').textContent = izin;
            document.getElementById('stat-sakit').textContent = sakit;
            document.getElementById('stat-alpa').textContent = alpa;
        }

        function editEntry(id) {
            const item = attendanceData.find(i => i.id === id);
            if (!item) return;

            isAdminLoggedIn = false;
            updateViewMode();

            document.getElementById('entry-id').value = item.id;
            document.getElementById('nama').value = item.nama;
            document.getElementById('jabatan').value = item.jabatan;
            document.getElementById('departemen').value = item.departemen;
            document.getElementById('tanggal').value = item.tanggal;
            document.getElementById('jam-masuk').value = item.jamMasuk !== '-' ? item.jamMasuk : '';
            document.getElementById('jam-keluar').value = item.jamKeluar !== '-' ? item.jamKeluar : '';
            document.getElementById('catatan').value = item.catatan !== '-' ? item.catatan : '';

            const radios = document.querySelectorAll('input[name="status"]');
            radios.forEach(r => { if (r.value === item.status) r.checked = true; });

            handleStatusChange();
            document.getElementById('btn-submit-text').textContent = "Perbarui Data Presensi";
            window.scrollTo({ top: 0, behavior: 'smooth' });
        }

        function deleteEntry(id) {
            if (confirm("Apakah Anda yakin ingin menghapus data presensi ini?")) {
                attendanceData = attendanceData.filter(i => i.id !== id);
                saveToLocalStorage();
                applyFilters();
                updateStats();
                showToast("Data presensi berhasil dihapus.");
            }
        }

        function clearAllData() {
            if (confirm("Apakah Anda YAKIN ingin menghapus seluruh data presensi karyawan?")) {
                attendanceData = [];
                saveToLocalStorage();
                renderTable();
                updateStats();
                showToast("Seluruh data presensi telah dibersihkan.");
            }
        }

        function saveToLocalStorage() {
            localStorage.setItem('presensi_data', JSON.stringify(attendanceData));
        }

        async function sendToGoogleSheet(entry) {
            if (!webhookUrl) return;
            try {
                await fetch(webhookUrl, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(entry)
                });
                entry.synced = true;
                saveToLocalStorage();
            } catch (err) {
                console.error("GSheet Error:", err);
            }
        }

        async function syncAllToSheet() {
            if (!webhookUrl) {
                openGoogleSheetModal();
                showToast("Atur Webhook URL Google Sheet terlebih dahulu!", "error");
                return;
            }

            showToast("Mengirimkan seluruh data ke Google Sheet...", "info");
            let count = 0;
            for (let item of attendanceData) {
                try {
                    await fetch(webhookUrl, {
                        method: 'POST',
                        mode: 'no-cors',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify(item)
                    });
                    item.synced = true;
                    count++;
                } catch (e) { console.error(e); }
            }
            saveToLocalStorage();
            showToast(`Berhasil mengirimkan ${count} data ke Google Sheet!`);
        }

        function exportToExcel() {
            if (attendanceData.length === 0) {
                showToast("Belum ada data untuk diexport!", "error");
                return;
            }

            const exportList = attendanceData.map(i => ({
                "ID Presensi": i.id,
                "Tanggal": i.tanggal,
                "Nama Karyawan": i.nama,
                "Jabatan": i.jabatan,
                "Departemen": i.departemen,
                "Jam Masuk": i.jamMasuk,
                "Jam Keluar": i.jamKeluar,
                "Status Kehadiran": i.status,
                "Catatan": i.catatan
            }));

            const ws = XLSX.utils.json_to_sheet(exportList);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Rekap Presensi");

            const filename = `Presensi_Karyawan_${new Date().toISOString().split('T')[0]}.xlsx`;
            XLSX.writeFile(wb, filename);
            showToast("Laporan Excel berhasil diunduh!");
        }

        function openGoogleSheetModal() {
            document.getElementById('modal-sheet').classList.remove('hidden');
        }

        function closeGoogleSheetModal() {
            document.getElementById('modal-sheet').classList.add('hidden');
        }

        function saveWebhookUrl() {
            const url = document.getElementById('webhook-url').value.trim();
            webhookUrl = url;
            localStorage.setItem('presensi_webhook_url', webhookUrl);
            closeGoogleSheetModal();
            showToast("Webhook URL Google Sheet berhasil disimpan!");
        }

        function copyAppsScript() {
            const text = document.getElementById('apps-script-code').value;
            navigator.clipboard.writeText(text).then(() => {
                showToast("Kode skrip disalin ke clipboard!");
            }).catch(() => {
                document.getElementById('apps-script-code').select();
                document.execCommand('copy');
                showToast("Kode disalin ke clipboard!");
            });
        }

        async function testConnection() {
            const url = document.getElementById('webhook-url').value.trim();
            const statusDiv = document.getElementById('connection-status');

            if (!url) {
                statusDiv.className = "p-2.5 rounded-xl text-xs bg-rose-100 text-rose-800 block";
                statusDiv.textContent = "Masukkan URL Webhook Google Script terlebih dahulu!";
                return;
            }

            statusDiv.className = "p-2.5 rounded-xl text-xs bg-blue-100 text-blue-800 block";
            statusDiv.textContent = "Menguji sinyal koneksi ke Google Sheet...";

            try {
                await fetch(url, {
                    method: 'POST',
                    mode: 'no-cors',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ test: true })
                });

                statusDiv.className = "p-2.5 rounded-xl text-xs bg-emerald-100 text-emerald-800 block";
                statusDiv.textContent = "✓ Sinyal terkirim! Silakan periksa Google Sheet Anda.";
            } catch (err) {
                statusDiv.className = "p-2.5 rounded-xl text-xs bg-rose-100 text-rose-800 block font-semibold";
                statusDiv.textContent = "Gagal terhubung: " + err.message;
            }
        }

        function changeAdminPinModal() {
            const newPin = prompt("Masukkan PIN Baru Admin (min 4 angka):", adminPin);
            if (newPin && newPin.trim().length >= 4) {
                adminPin = newPin.trim();
                localStorage.setItem('presensi_admin_pin', adminPin);
                showToast("PIN Admin berhasil diperbarui!");
            } else if (newPin !== null) {
                showToast("PIN harus minimal 4 karakter!", "error");
            }
        }

        function showToast(message, type = "success") {
            const toast = document.getElementById('toast');
            const toastMsg = document.getElementById('toast-msg');
            const toastIcon = document.getElementById('toast-icon');

            if (!toast || !toastMsg || !toastIcon) return;

            toastMsg.textContent = message;

            if (type === "error") {
                toastIcon.className = "fa-solid fa-circle-exclamation text-rose-400 text-base";
            } else if (type === "info") {
                toastIcon.className = "fa-solid fa-circle-info text-blue-400 text-base";
            } else {
                toastIcon.className = "fa-solid fa-circle-check text-emerald-400 text-base";
            }

            toast.classList.remove('translate-y-20', 'opacity-0');
            toast.classList.add('translate-y-0', 'opacity-100');

            setTimeout(() => {
                toast.classList.remove('translate-y-0', 'opacity-100');
                toast.classList.add('translate-y-20', 'opacity-0');
            }, 3000);
        }

        function escapeHtml(str) {
            if (!str) return '';
            return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
        }
    </script>
</body>
</html>

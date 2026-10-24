<!DOCTYPE html>
<html lang="vi" class="h-full bg-slate-50">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hệ Thống Quản Lý & Theo Dõi Mức Độ Học Tập Giữa HK1 (GVBM & GVCN)</title>
    <!-- Tailwind CSS -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Google Fonts: Be Vietnam Pro -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Be+Vietnam+Pro:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <!-- FontAwesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Chart.js -->
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <!-- SheetJS for Excel Export/Import -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

    <script>
        tailwind.config = {
            darkMode: 'class',
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Be Vietnam Pro', 'sans-serif'],
                    },
                    colors: {
                        brand: {
                            50: '#f0f9ff',
                            100: '#e0f2fe',
                            500: '#0284c7',
                            600: '#0369a1',
                            700: '#0369a1',
                            800: '#075985',
                            900: '#0c4a6e',
                        }
                    }
                }
            }
        }
    </script>
    <style>
        body { font-family: 'Be Vietnam Pro', sans-serif; }
        .tab-btn.active {
            color: #0369a1;
            border-bottom-color: #0369a1;
            background-color: #f0f9ff;
        }
        .dark .tab-btn.active {
            color: #38bdf8;
            border-bottom-color: #38bdf8;
            background-color: #0f172a;
        }
        @media print {
            .no-print { display: none !important; }
            .print-only { display: block !important; }
            body { background: white !important; color: black !important; }
            .card { border: none !important; box-shadow: none !important; }
            table { width: 100% !important; border-collapse: collapse !important; }
            th, td { border: 1px solid #cbd5e1 !important; padding: 4px 6px !important; font-size: 10px !important; }
        }
        .custom-scrollbar::-webkit-scrollbar { height: 6px; width: 6px; }
        .custom-scrollbar::-webkit-scrollbar-track { background: #f1f5f9; }
        .custom-scrollbar::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        .dark .custom-scrollbar::-webkit-scrollbar-track { background: #1e293b; }
        .dark .custom-scrollbar::-webkit-scrollbar-thumb { background: #475569; }
    </style>
</head>
<body class="h-full flex flex-col text-slate-800 dark:text-slate-100 dark:bg-slate-900 antialiased selection:bg-brand-500 selection:text-white transition-colors duration-200">

    <!-- TOP NAVBAR & ROLE SWITCHER -->
    <header class="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 sticky top-0 z-30 shadow-sm no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex flex-col md:flex-row justify-between md:h-16 items-center py-2 md:py-0 gap-3">
                <!-- App Title -->
                <div class="flex items-center space-x-3 w-full md:w-auto justify-between md:justify-start">
                    <div class="flex items-center space-x-3">
                        <div id="role-logo-bg" class="p-2.5 bg-gradient-to-tr from-brand-600 to-teal-500 text-white rounded-xl shadow-md transition-all">
                            <i id="role-icon" class="fa-solid fa-graduation-cap text-xl"></i>
                        </div>
                        <div>
                            <div class="flex items-center gap-2">
                                <h1 class="text-base font-bold text-slate-900 dark:text-white leading-tight">EduTrack Giữa HK1</h1>
                                <span id="role-badge" class="px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-brand-100 text-brand-800 dark:bg-brand-900 dark:text-brand-200">GV Bộ Môn</span>
                            </div>
                            <p id="role-subtitle" class="text-xs text-slate-500 dark:text-slate-400 font-medium">Sổ điểm & Theo dõi Bộ môn</p>
                        </div>
                    </div>

                    <!-- Mobile Dark Toggle -->
                    <button onclick="toggleDarkMode()" class="md:hidden p-2 text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-white">
                        <i class="fa-solid fa-moon dark:hidden"></i>
                        <i class="fa-solid fa-sun hidden dark:block"></i>
                    </button>
                </div>

                <!-- ROLE SWITCHER & CONTROLS -->
                <div class="flex flex-wrap items-center gap-2 sm:gap-3 w-full md:w-auto justify-center md:justify-end">
                    <!-- MODE TOGGLE BUTTONS -->
                    <div class="bg-slate-100 dark:bg-slate-700 p-1 rounded-xl flex items-center shadow-inner">
                        <button id="btn-mode-gvbm" onclick="switchRoleMode('GVBM')" class="px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 bg-white dark:bg-slate-800 text-brand-700 dark:text-brand-400 shadow-sm">
                            <i class="fa-solid fa-book-bookmark"></i> Giáo Viên Bộ Môn
                        </button>
                        <button id="btn-mode-gvcn" onclick="switchRoleMode('GVCN')" class="px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 text-slate-600 dark:text-slate-300 hover:text-slate-900">
                            <i class="fa-solid fa-people-roof"></i> Giáo Viên Chủ Nhiệm
                        </button>
                    </div>

                    <div class="h-6 w-px bg-slate-200 dark:bg-slate-700 hidden sm:block"></div>

                    <!-- Actions -->
                    <button onclick="loadDemoData()" class="px-2.5 py-1.5 text-xs font-semibold text-slate-600 dark:text-slate-300 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:hover:bg-slate-600 rounded-lg transition flex items-center gap-1" title="Tải dữ liệu thử nghiệm">
                        <i class="fa-solid fa-flask text-amber-500"></i>
                        <span class="hidden lg:inline">Dữ liệu mẫu</span>
                    </button>
                    <button onclick="clearCurrentRoleData()" class="px-2.5 py-1.5 text-xs font-semibold text-rose-600 bg-rose-50 hover:bg-rose-100 dark:bg-rose-950/40 dark:text-rose-300 rounded-lg transition flex items-center gap-1" title="Xóa dữ liệu">
                        <i class="fa-solid fa-trash-can"></i>
                    </button>
                    <button onclick="toggleDarkMode()" class="hidden md:flex p-2 text-slate-500 hover:text-slate-700 dark:text-slate-400 dark:hover:text-white rounded-lg bg-slate-100 dark:bg-slate-700">
                        <i class="fa-solid fa-moon dark:hidden"></i>
                        <i class="fa-solid fa-sun hidden dark:block"></i>
                    </button>
                    <button onclick="window.print()" class="px-3 py-1.5 text-xs font-semibold text-white bg-slate-800 hover:bg-slate-900 dark:bg-slate-700 dark:hover:bg-slate-600 rounded-lg transition shadow-sm flex items-center gap-1.5">
                        <i class="fa-solid fa-print"></i>
                        <span class="hidden sm:inline">In báo cáo</span>
                    </button>
                </div>
            </div>
        </div>
    </header>

    <!-- GVBM NAVIGATION TABS -->
    <nav id="nav-gvbm" class="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 no-print">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex space-x-1 overflow-x-auto custom-scrollbar">
                <button onclick="switchTab('gvbm-dashboard')" id="btn-gvbm-dashboard" class="tab-btn active px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-chart-pie"></i> Thống Kê Bộ Môn
                </button>
                <button onclick="switchTab('gvbm-gradebook')" id="btn-gvbm-gradebook" class="tab-btn px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-table"></i> Sổ Điểm & Nhận Xét
                </button>
                <button onclick="switchTab('gvbm-attention')" id="btn-gvbm-attention" class="tab-btn px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap relative">
                    <i class="fa-solid fa-triangle-exclamation text-amber-500"></i> Báo GVCN / Phụ Huynh
                    <span id="gvbm-attention-badge" class="ml-1 bg-rose-500 text-white text-xs px-1.5 py-0.5 rounded-full hidden">0</span>
                </button>
                <button onclick="switchTab('gvbm-settings')" id="btn-gvbm-settings" class="tab-btn px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-sliders"></i> Cấu Hình Cột Điểm
                </button>
            </div>
        </div>
    </nav>

    <!-- GVCN NAVIGATION TABS -->
    <nav id="nav-gvcn" class="bg-white dark:bg-slate-800 border-b border-slate-200 dark:border-slate-700 no-print hidden">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
            <div class="flex space-x-1 overflow-x-auto custom-scrollbar">
                <button onclick="switchTab('gvcn-dashboard')" id="btn-gvcn-dashboard" class="tab-btn active px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-chart-line"></i> Bức Tranh Toàn Cảnh Lớp
                </button>
                <button onclick="switchTab('gvcn-master')" id="btn-gvcn-master" class="tab-btn px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-table-cells"></i> Bảng Điểm Tổng Hợp Các Môn
                </button>
                <button onclick="switchTab('gvcn-reports')" id="btn-gvcn-reports" class="tab-btn px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap">
                    <i class="fa-solid fa-id-card"></i> Sổ Liên Lạc / Phiếu Báo Điểm
                </button>
                <button onclick="switchTab('gvcn-atrisk')" id="btn-gvcn-atrisk" class="tab-btn px-4 py-3 text-sm font-semibold text-slate-600 dark:text-slate-300 border-b-2 border-transparent hover:text-brand-600 flex items-center gap-2 whitespace-nowrap relative">
                    <i class="fa-solid fa-shield-cat text-rose-500"></i> Danh Sách Cần Giúp Đỡ (>2 môn yếu)
                    <span id="gvcn-atrisk-badge" class="ml-1 bg-rose-500 text-white text-xs px-1.5 py-0.5 rounded-full hidden">0</span>
                </button>
            </div>
        </div>
    </nav>

    <!-- MAIN CONTAINER -->
    <main class="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6">

        <!-- ==================== GVBM PANELS ==================== -->

        <!-- GVBM TAB 1: DASHBOARD -->
        <section id="gvbm-dashboard" class="tab-content space-y-6">
            <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4 no-print">
                <div class="flex items-center gap-3 w-full sm:w-auto">
                    <label class="text-xs font-bold uppercase text-slate-500 dark:text-slate-400 whitespace-nowrap"><i class="fa-solid fa-filter mr-1"></i> Lọc Lớp Môn Học:</label>
                    <select id="gvbm-filter-class" onchange="renderGvbmDashboard()" class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm rounded-lg p-2 font-medium">
                        <option value="ALL">Tất cả các lớp</option>
                    </select>
                </div>
                <div class="text-xs text-slate-500 dark:text-slate-400 font-medium">
                    Môn: <b id="gvbm-subject-display-name" class="text-brand-600 dark:text-brand-400">Toán</b> | Chuẩn Đánh Giá Giữa HK1 (TT 22)
                </div>
            </div>

            <!-- STAT CARDS -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-indigo-50 dark:bg-indigo-950 text-indigo-600 dark:text-indigo-400 rounded-lg"><i class="fa-solid fa-users text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Sĩ Số Môn Học</p>
                        <h3 id="gvbm-stat-total" class="text-2xl font-bold text-slate-800 dark:text-white">0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-brand-50 dark:bg-brand-950 text-brand-600 dark:text-brand-400 rounded-lg"><i class="fa-solid fa-calculator text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">ĐTB Môn Chung</p>
                        <h3 id="gvbm-stat-avg" class="text-2xl font-bold text-brand-600 dark:text-brand-400">0.0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-emerald-50 dark:bg-emerald-950 text-emerald-600 dark:text-emerald-400 rounded-lg"><i class="fa-solid fa-trophy text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Điểm Cao Nhất</p>
                        <h3 id="gvbm-stat-max" class="text-2xl font-bold text-emerald-600 dark:text-emerald-400">0.0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-rose-50 dark:bg-rose-950 text-rose-600 dark:text-rose-400 rounded-lg"><i class="fa-solid fa-arrow-trend-down text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Điểm Thấp Nhất</p>
                        <h3 id="gvbm-stat-min" class="text-2xl font-bold text-rose-600 dark:text-rose-400">0.0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-teal-50 dark:bg-teal-950 text-teal-600 dark:text-teal-400 rounded-lg"><i class="fa-solid fa-chart-line text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Tỷ Lệ Đạt (≥5.0)</p>
                        <h3 id="gvbm-stat-pass" class="text-2xl font-bold text-teal-600 dark:text-teal-400">0%</h3>
                    </div>
                </div>
            </div>

            <!-- CHARTS -->
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="bg-white dark:bg-slate-800 p-5 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col items-center">
                    <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 self-start mb-4"><i class="fa-solid fa-chart-pie text-brand-500 mr-2"></i>Tỷ Lệ Phân Loại Mức Độ</h3>
                    <div class="w-full h-64 relative"><canvas id="gvbmChartDoughnut"></canvas></div>
                </div>
                <div class="bg-white dark:bg-slate-800 p-5 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col items-center lg:col-span-2">
                    <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 self-start mb-4"><i class="fa-solid fa-chart-column text-brand-500 mr-2"></i>Thống Kê Chi Tiết Số Lượng</h3>
                    <div class="w-full h-64 relative"><canvas id="gvbmChartBar"></canvas></div>
                </div>
            </div>
        </section>

        <!-- GVBM TAB 2: GRADEBOOK -->
        <section id="gvbm-gradebook" class="tab-content hidden space-y-4">
            <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm space-y-4 no-print">
                <div class="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
                    <div class="flex flex-wrap items-center gap-3">
                        <div class="relative w-full sm:w-64">
                            <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
                            <input type="text" id="gvbm-search" oninput="renderGvbmGradebook()" placeholder="Tìm tên học sinh..." class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-xs rounded-lg pl-8 pr-3 py-2 w-full">
                        </div>
                        <select id="gvbm-gradebook-filter-class" onchange="renderGvbmGradebook()" class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-xs rounded-lg p-2 font-medium">
                            <option value="ALL">Tất cả các lớp</option>
                        </select>
                    </div>

                    <div class="flex flex-wrap items-center gap-2">
                        <button onclick="openAddStudentModal()" class="px-3 py-2 text-xs font-semibold text-white bg-brand-600 hover:bg-brand-700 rounded-lg transition shadow-sm flex items-center gap-1.5"><i class="fa-solid fa-user-plus"></i> Thêm HS</button>
                        <button onclick="openBatchAddModal()" class="px-3 py-2 text-xs font-semibold text-slate-700 dark:text-slate-200 bg-slate-100 hover:bg-slate-200 dark:bg-slate-700 dark:hover:bg-slate-600 rounded-lg transition flex items-center gap-1.5"><i class="fa-solid fa-paste"></i> Nhập hàng loạt</button>
                        <button onclick="exportGvbmToExcel()" class="px-3 py-2 text-xs font-semibold text-emerald-700 dark:text-emerald-300 bg-emerald-50 dark:bg-emerald-950 hover:bg-emerald-100 rounded-lg transition flex items-center gap-1.5"><i class="fa-solid fa-file-excel"></i> Xuất File gửi GVCN</button>
                        <label class="px-3 py-2 text-xs font-semibold text-sky-700 dark:text-sky-300 bg-sky-50 dark:bg-sky-950 hover:bg-sky-100 rounded-lg transition cursor-pointer flex items-center gap-1.5">
                            <i class="fa-solid fa-file-import"></i> Nhập Excel/CSV
                            <input type="file" accept=".xlsx, .xls, .csv" onchange="importGvbmExcel(event)" class="hidden">
                        </label>
                    </div>
                </div>
            </div>

            <!-- GVBM GRADEBOOK TABLE -->
            <div class="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm overflow-hidden">
                <div class="overflow-x-auto custom-scrollbar">
                    <table class="w-full text-xs text-left border-collapse" id="gvbm-table">
                        <thead class="text-xs uppercase bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-bold border-b border-slate-200 dark:border-slate-600">
                            <tr id="gvbm-header-row">
                                <!-- Dynamic JS Injection -->
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-200 dark:divide-slate-700 font-medium" id="gvbm-tbody">
                            <!-- Dynamic JS Injection -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- GVBM TAB 3: ATTENTION & ZALO MSGS -->
        <section id="gvbm-attention" class="tab-content hidden space-y-4">
            <div class="bg-amber-50 dark:bg-amber-950/40 border border-amber-200 dark:border-amber-800/60 rounded-xl p-4 flex items-start gap-3 text-amber-900 dark:text-amber-200 no-print">
                <i class="fa-solid fa-triangle-exclamation text-amber-600 text-xl mt-0.5"></i>
                <div class="space-y-1">
                    <h4 class="text-sm font-bold">Học Sinh Cần Báo Cho GVCN & Phụ Huynh (ĐTB Môn < 5.0)</h4>
                    <p class="text-xs text-amber-800 dark:text-amber-300">Tự động lọc các học sinh sa sút môn bộ môn. GVBM chỉ cần bấm chọn tin nhắn để gửi Zalo cho GVCN hoặc PHHS.</p>
                </div>
            </div>

            <div class="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm overflow-hidden">
                <div class="overflow-x-auto">
                    <table class="w-full text-xs text-left text-slate-700 dark:text-slate-200">
                        <thead class="text-xs uppercase bg-slate-100 dark:bg-slate-700 font-bold border-b border-slate-200 dark:border-slate-600">
                            <tr>
                                <th class="p-3 w-10 text-center">STT</th>
                                <th class="p-3">Họ và Tên</th>
                                <th class="p-3 text-center">Lớp</th>
                                <th class="p-3 text-center">ĐGK Môn</th>
                                <th class="p-3 text-center">ĐTB Môn Giữa HK1</th>
                                <th class="p-3">Ghi Chú Môn Học</th>
                                <th class="p-3 text-center no-print">Tạo Tin Nhắn Zalo/SMS</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-200 dark:divide-slate-700" id="gvbm-attention-tbody">
                            <!-- JS Injection -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- GVBM TAB 4: SETTINGS -->
        <section id="gvbm-settings" class="tab-content hidden space-y-6">
            <div class="bg-white dark:bg-slate-800 p-6 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm max-w-2xl mx-auto space-y-6">
                <h3 class="text-base font-bold text-slate-900 dark:text-white border-b border-slate-200 dark:border-slate-700 pb-3 flex items-center gap-2">
                    <i class="fa-solid fa-sliders text-brand-600"></i> Cấu Hình Bộ Môn Của GVBM
                </h3>
                <div class="space-y-4">
                    <div>
                        <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Tên Môn Học Đảm Nhận:</label>
                        <input type="text" id="gvbm-cfg-subject" value="Toán" class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm rounded-lg p-2.5">
                    </div>
                    <div>
                        <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Số Cột Đánh Giá Thường Xuyên (ĐGTX):</label>
                        <input type="number" id="gvbm-cfg-dgtx" min="1" max="8" value="4" class="w-24 bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm rounded-lg p-2.5 font-bold text-center">
                    </div>
                    <button onclick="saveGvbmConfig()" class="px-5 py-2.5 bg-brand-600 hover:bg-brand-700 text-white font-semibold text-xs rounded-lg shadow-sm transition flex items-center gap-2">
                        <i class="fa-solid fa-floppy-disk"></i> Lưu cấu hình bộ môn
                    </button>
                </div>
            </div>
        </section>


        <!-- ==================== GVCN PANELS ==================== -->

        <!-- GVCN TAB 1: OVERVIEW DASHBOARD -->
        <section id="gvcn-dashboard" class="tab-content hidden space-y-6">
            <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4 no-print">
                <div class="flex items-center gap-3">
                    <label class="text-xs font-bold uppercase text-slate-500 dark:text-slate-400 whitespace-nowrap"><i class="fa-solid fa-house-user mr-1"></i> Lớp Chủ Nhiệm:</label>
                    <input type="text" id="gvcn-class-name" value="6A1" onchange="updateGvcnClassName(this.value)" class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm font-bold rounded-lg p-2 w-28 text-center uppercase">
                </div>
                <div class="text-xs text-slate-500 dark:text-slate-400 font-medium">
                    <i class="fa-solid fa-certificate text-amber-500 mr-1"></i> Xếp loại Học lực Tổng hợp Lớp theo Thông tư 22/58
                </div>
            </div>

            <!-- GVCN STAT CARDS -->
            <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-4">
                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-indigo-50 dark:bg-indigo-950 text-indigo-600 dark:text-indigo-400 rounded-lg"><i class="fa-solid fa-users text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Sĩ Số Lớp CN</p>
                        <h3 id="gvcn-stat-total" class="text-2xl font-bold text-slate-800 dark:text-white">0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-emerald-50 dark:bg-emerald-950 text-emerald-600 dark:text-emerald-400 rounded-lg"><i class="fa-solid fa-award text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Học Sinh Tốt / Giỏi</p>
                        <h3 id="gvcn-stat-good" class="text-2xl font-bold text-emerald-600 dark:text-emerald-400">0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-sky-50 dark:bg-sky-950 text-sky-600 dark:text-sky-400 rounded-lg"><i class="fa-solid fa-user-check text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Học Sinh Khá</p>
                        <h3 id="gvcn-stat-kha" class="text-2xl font-bold text-sky-600 dark:text-sky-400">0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-amber-50 dark:bg-amber-950 text-amber-600 dark:text-amber-400 rounded-lg"><i class="fa-solid fa-user-minus text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Học Sinh Đạt (TB)</p>
                        <h3 id="gvcn-stat-dat" class="text-2xl font-bold text-amber-600 dark:text-amber-400">0</h3>
                    </div>
                </div>

                <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex items-center space-x-4">
                    <div class="p-3 bg-rose-50 dark:bg-rose-950 text-rose-600 dark:text-rose-400 rounded-lg"><i class="fa-solid fa-triangle-exclamation text-xl"></i></div>
                    <div>
                        <p class="text-xs font-medium text-slate-500 dark:text-slate-400">Chưa Đạt / Yếu</p>
                        <h3 id="gvcn-stat-weak" class="text-2xl font-bold text-rose-600 dark:text-rose-400">0</h3>
                    </div>
                </div>
            </div>

            <!-- CHARTS & SUMMARY -->
            <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
                <div class="bg-white dark:bg-slate-800 p-5 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col items-center">
                    <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 self-start mb-4"><i class="fa-solid fa-chart-pie text-brand-500 mr-2"></i>Tỷ Lệ Học Lực Tổng hợp Lớp</h3>
                    <div class="w-full h-64 relative"><canvas id="gvcnChartDoughnut"></canvas></div>
                </div>
                <div class="bg-white dark:bg-slate-800 p-5 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col items-center lg:col-span-2">
                    <h3 class="text-sm font-bold text-slate-800 dark:text-slate-200 self-start mb-4"><i class="fa-solid fa-chart-column text-brand-500 mr-2"></i>Thống Kê Điểm Trung Bình Các Môn</h3>
                    <div class="w-full h-64 relative"><canvas id="gvcnChartBar"></canvas></div>
                </div>
            </div>
        </section>

        <!-- GVCN TAB 2: MASTER SYNTHESIS TABLE -->
        <section id="gvcn-master" class="tab-content hidden space-y-4">
            <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4 no-print">
                <div class="relative w-full sm:w-64">
                    <i class="fa-solid fa-magnifying-glass absolute left-3 top-2.5 text-slate-400 text-xs"></i>
                    <input type="text" id="gvcn-search" oninput="renderGvcnMasterTable()" placeholder="Tìm tên học sinh..." class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-xs rounded-lg pl-8 pr-3 py-2 w-full">
                </div>

                <div class="flex flex-wrap items-center gap-2">
                    <button onclick="openAddStudentModal()" class="px-3 py-2 text-xs font-semibold text-white bg-brand-600 hover:bg-brand-700 rounded-lg shadow-sm flex items-center gap-1.5"><i class="fa-solid fa-user-plus"></i> Thêm HS Lớp</button>
                    <button onclick="exportGvcnMasterToExcel()" class="px-3 py-2 text-xs font-semibold text-emerald-700 dark:text-emerald-300 bg-emerald-50 dark:bg-emerald-950 hover:bg-emerald-100 rounded-lg transition flex items-center gap-1.5"><i class="fa-solid fa-file-excel"></i> Xuất Bảng Điểm Tổng Hợp</button>
                    <label class="px-3 py-2 text-xs font-semibold text-sky-700 dark:text-sky-300 bg-sky-50 dark:bg-sky-950 hover:bg-sky-100 rounded-lg cursor-pointer flex items-center gap-1.5">
                        <i class="fa-solid fa-file-import"></i> Ghép Bảng Điểm Môn Từ GVBM
                        <input type="file" accept=".xlsx, .xls, .csv" onchange="importGvcnSubjectFile(event)" class="hidden">
                    </label>
                </div>
            </div>

            <!-- GVCN MATRIX TABLE -->
            <div class="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm overflow-hidden">
                <div class="overflow-x-auto custom-scrollbar">
                    <table class="w-full text-xs text-left border-collapse" id="gvcn-master-table">
                        <thead class="text-xs uppercase bg-slate-100 dark:bg-slate-700 text-slate-700 dark:text-slate-200 font-bold border-b border-slate-200 dark:border-slate-600">
                            <tr id="gvcn-master-header">
                                <th class="p-2.5 w-10 text-center border-r">STT</th>
                                <th class="p-2.5 min-w-[160px] border-r">Họ và Tên</th>
                                <!-- Subject columns injected dynamically -->
                                <th class="p-2.5 w-20 text-center border-r bg-brand-50 dark:bg-brand-950 text-brand-900 dark:text-brand-200">ĐTB Tổng</th>
                                <th class="p-2.5 w-24 text-center border-r">Xếp Loại</th>
                                <th class="p-2.5 w-16 text-center no-print">Xóa</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-200 dark:divide-slate-700 font-medium" id="gvcn-master-tbody">
                            <!-- JS Injection -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- GVCN TAB 3: INDIVIDUAL STUDENT REPORT CARDS / PARENT SLIPS -->
        <section id="gvcn-reports" class="tab-content hidden space-y-6">
            <div class="bg-white dark:bg-slate-800 p-4 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm flex flex-col sm:flex-row items-center justify-between gap-4 no-print">
                <div class="flex items-center gap-3 w-full sm:w-auto">
                    <label class="text-xs font-bold uppercase text-slate-500 dark:text-slate-400 whitespace-nowrap"><i class="fa-solid fa-user-check mr-1"></i> Chọn Học Sinh:</label>
                    <select id="gvcn-report-student-select" onchange="renderStudentReportCard()" class="bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm font-semibold rounded-lg p-2.5 w-full sm:w-64">
                        <!-- Dynamic Student Options -->
                    </select>
                </div>
                <div class="flex items-center gap-2">
                    <button onclick="copyParentZaloSlip()" class="px-3 py-2 text-xs font-semibold text-sky-700 bg-sky-50 dark:bg-sky-950 dark:text-sky-300 rounded-lg hover:bg-sky-100 flex items-center gap-1.5"><i class="fa-solid fa-copy"></i> Copy Phiếu gửi Zalo PHHS</button>
                    <button onclick="window.print()" class="px-3 py-2 text-xs font-semibold text-white bg-slate-800 dark:bg-slate-700 rounded-lg hover:bg-slate-900 flex items-center gap-1.5"><i class="fa-solid fa-print"></i> In Phiếu Báo Điểm</button>
                </div>
            </div>

            <!-- PRINTABLE REPORT CARD / PHIẾU BÁO ĐIỂM -->
            <div id="student-report-card" class="bg-white text-slate-900 p-8 rounded-xl border border-slate-200 shadow-md max-w-3xl mx-auto space-y-6 print:shadow-none print:border-none print:p-0">
                <!-- Report Header -->
                <div class="text-center space-y-1 border-b pb-4">
                    <h2 class="text-lg font-bold uppercase tracking-wide">TRƯỜNG THCS / THPT .....................................</h2>
                    <h3 class="text-base font-extrabold text-brand-700">PHIẾU BÁO KẾT QUẢ HỌC TẬP GIỮA HỌC KỲ 1</h3>
                    <p class="text-xs text-slate-500 italic">Năm học: 2026 - 2027</p>
                </div>

                <!-- Student Profile Info -->
                <div class="grid grid-cols-2 gap-4 text-xs font-semibold bg-slate-50 p-4 rounded-lg border border-slate-200">
                    <div>Họ và tên học sinh: <span id="rpt-student-name" class="font-bold text-brand-800 uppercase text-sm">...</span></div>
                    <div>Lớp: <span id="rpt-student-class" class="font-bold">...</span></div>
                    <div>Điểm TB Tổng hợp Các môn: <span id="rpt-student-gpa" class="font-bold text-brand-700 text-sm">...</span></div>
                    <div>Xếp loại Học lực: <span id="rpt-student-rank" class="font-bold text-emerald-700">...</span></div>
                </div>

                <!-- Grades Breakdown Table -->
                <div>
                    <h4 class="text-xs font-bold uppercase text-slate-700 mb-2">Chi Tiết Điểm Trung Bình Các Môn Học:</h4>
                    <table class="w-full text-xs text-left border-collapse border border-slate-300">
                        <thead>
                            <tr class="bg-slate-100 text-slate-800 font-bold border-b border-slate-300">
                                <th class="p-2 border-r border-slate-300 text-center w-12">STT</th>
                                <th class="p-2 border-r border-slate-300">Môn Học</th>
                                <th class="p-2 border-r border-slate-300 text-center w-28">ĐTB Giữa HK1</th>
                                <th class="p-2 border-r border-slate-300 text-center w-28">Đánh giá</th>
                            </tr>
                        </thead>
                        <tbody id="rpt-grades-tbody" class="divide-y divide-slate-200">
                            <!-- Dynamic Injection -->
                        </tbody>
                    </table>
                </div>

                <!-- Homeroom Teacher Comments -->
                <div class="space-y-2">
                    <h4 class="text-xs font-bold uppercase text-slate-700">Nhận xét của Giáo viên Chủ nhiệm:</h4>
                    <textarea id="rpt-teacher-comment" onchange="saveGvcnComment(this.value)" rows="3" placeholder="Nhập nhận xét chung về tình hình nếp, chuyên cần và học tập..." class="w-full text-xs p-3 bg-slate-50 border border-slate-300 rounded-lg no-print"></textarea>
                    <p id="rpt-teacher-comment-print" class="hidden print:block text-xs italic p-2 border border-slate-200 rounded min-h-[60px]"></p>
                </div>

                <!-- Signatures Section -->
                <div class="grid grid-cols-2 text-center text-xs font-semibold pt-6">
                    <div>
                        <p class="uppercase font-bold">XÁC NHẬN CỦA PHỤ HUYNH</p>
                        <p class="text-[10px] text-slate-400 italic">(Ký và ghi rõ họ tên)</p>
                    </div>
                    <div>
                        <p class="italic text-slate-500">..., Ngày ... tháng ... năm 2026</p>
                        <p class="uppercase font-bold mt-1">GIÁO VIÊN CHỦ NHIỆM</p>
                        <p class="text-[10px] text-slate-400 italic mb-12">(Ký và ghi rõ họ tên)</p>
                    </div>
                </div>
            </div>
        </section>

        <!-- GVCN TAB 4: AT RISK STUDENTS -->
        <section id="gvcn-atrisk" class="tab-content hidden space-y-4">
            <div class="bg-rose-50 dark:bg-rose-950/40 border border-rose-200 dark:border-rose-800/60 rounded-xl p-4 flex items-start gap-3 text-rose-900 dark:text-rose-200 no-print">
                <i class="fa-solid fa-shield-cat text-rose-600 text-xl mt-0.5"></i>
                <div class="space-y-1">
                    <h4 class="text-sm font-bold">Danh Sách Học Sinh Có Nguy Cơ / Cần Giúp Đỡ Đặc Biệt</h4>
                    <p class="text-xs text-rose-800 dark:text-rose-300">Tự động phát hiện các học sinh có <b>từ 2 môn học trở lên bị điểm Chưa đạt (<5.0)</b>. GVCN cần lên kế hoạch phối hợp với các GVBM và gia đình.</p>
                </div>
            </div>

            <div class="bg-white dark:bg-slate-800 rounded-xl border border-slate-200 dark:border-slate-700 shadow-sm overflow-hidden">
                <div class="overflow-x-auto">
                    <table class="w-full text-xs text-left text-slate-700 dark:text-slate-200">
                        <thead class="text-xs uppercase bg-slate-100 dark:bg-slate-700 font-bold border-b border-slate-200 dark:border-slate-600">
                            <tr>
                                <th class="p-3 w-10 text-center">STT</th>
                                <th class="p-3">Họ và Tên</th>
                                <th class="p-3 text-center">Số Môn Dưới 5.0</th>
                                <th class="p-3">Các Môn Bị Yếu/Chưa Đạt</th>
                                <th class="p-3 text-center">ĐTB Tổng Hợp</th>
                                <th class="p-3 text-center no-print">Hành Động</th>
                            </tr>
                        </thead>
                        <tbody class="divide-y divide-slate-200 dark:divide-slate-700" id="gvcn-atrisk-tbody">
                            <!-- JS Injection -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

    </main>

    <!-- FOOTER -->
    <footer class="bg-white dark:bg-slate-800 border-t border-slate-200 dark:border-slate-700 mt-auto py-4 no-print">
        <div class="max-w-7xl mx-auto px-4 text-center text-xs text-slate-500 dark:text-slate-400 flex flex-col sm:flex-row justify-between items-center gap-2">
            <div>
                <span>Hệ thống Quản lý Mức độ Học tập Giữa HK1 (TT 22/58) &copy; 2026</span>
            </div>
            <div class="flex items-center gap-4">
                <span class="inline-flex items-center gap-1 text-emerald-600 dark:text-emerald-400 font-medium"><i class="fa-solid fa-shield-halved"></i> Tự động lưu trữ LocalStorage</span>
            </div>
        </div>
    </footer>

    <!-- MODAL 1: ADD SINGLE STUDENT -->
    <div id="modal-add-student" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center hidden p-4">
        <div class="bg-white dark:bg-slate-800 rounded-xl shadow-xl border border-slate-200 dark:border-slate-700 w-full max-w-md overflow-hidden">
            <div class="px-5 py-4 border-b border-slate-100 dark:border-slate-700 flex justify-between items-center bg-slate-50 dark:bg-slate-900">
                <h3 class="font-bold text-slate-800 dark:text-slate-200 text-sm flex items-center gap-2"><i class="fa-solid fa-user-plus text-brand-600"></i> Thêm Học Sinh Mới</h3>
                <button onclick="closeAddStudentModal()" class="text-slate-400 hover:text-slate-600"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            <form onsubmit="handleAddStudentSubmit(event)" class="p-5 space-y-4">
                <div>
                    <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Họ và Tên Học Sinh <span class="text-rose-500">*</span></label>
                    <input type="text" id="modal-add-name" required placeholder="Nguyễn Văn An" class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm rounded-lg p-2.5">
                </div>
                <div>
                    <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Lớp Học <span class="text-rose-500">*</span></label>
                    <input type="text" id="modal-add-class" required placeholder="6A1" class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm rounded-lg p-2.5">
                </div>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeAddStudentModal()" class="px-4 py-2 text-xs font-semibold text-slate-600 dark:text-slate-300 bg-slate-100 dark:bg-slate-700 rounded-lg">Hủy</button>
                    <button type="submit" class="px-4 py-2 text-xs font-semibold text-white bg-brand-600 hover:bg-brand-700 rounded-lg shadow-sm">Lưu Học Sinh</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL 2: BATCH ADD STUDENTS -->
    <div id="modal-batch-add" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center hidden p-4">
        <div class="bg-white dark:bg-slate-800 rounded-xl shadow-xl border border-slate-200 dark:border-slate-700 w-full max-w-lg overflow-hidden">
            <div class="px-5 py-4 border-b border-slate-100 dark:border-slate-700 flex justify-between items-center bg-slate-50 dark:bg-slate-900">
                <h3 class="font-bold text-slate-800 dark:text-slate-200 text-sm flex items-center gap-2"><i class="fa-solid fa-paste text-brand-600"></i> Nhập Hàng Loạt Từ Excel/Word</h3>
                <button onclick="closeBatchAddModal()" class="text-slate-400 hover:text-slate-600"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            <form onsubmit="handleBatchAddSubmit(event)" class="p-5 space-y-4">
                <div>
                    <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Lớp Áp Dụng <span class="text-rose-500">*</span></label>
                    <input type="text" id="modal-batch-class" required placeholder="6A1" class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-sm rounded-lg p-2.5">
                </div>
                <div>
                    <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Dán Danh Sách Họ Tên (Mỗi tên 1 dòng):</label>
                    <textarea id="modal-batch-text" rows="7" required placeholder="Nguyễn Văn A&#10;Trần Thị B&#10;Lê Văn C" class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-xs rounded-lg p-2.5 font-mono"></textarea>
                </div>
                <div class="flex justify-end gap-2 pt-2">
                    <button type="button" onclick="closeBatchAddModal()" class="px-4 py-2 text-xs font-semibold text-slate-600 dark:text-slate-300 bg-slate-100 dark:bg-slate-700 rounded-lg">Hủy</button>
                    <button type="submit" class="px-4 py-2 text-xs font-semibold text-white bg-brand-600 hover:bg-brand-700 rounded-lg shadow-sm">Thêm Vào Sổ Điểm</button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL 3: ZALO/SMS GENERATOR -->
    <div id="modal-sms" class="fixed inset-0 bg-slate-900/50 backdrop-blur-sm z-50 flex items-center justify-center hidden p-4">
        <div class="bg-white dark:bg-slate-800 rounded-xl shadow-xl border border-slate-200 dark:border-slate-700 w-full max-w-lg overflow-hidden">
            <div class="px-5 py-4 border-b border-slate-100 dark:border-slate-700 flex justify-between items-center bg-slate-50 dark:bg-slate-900">
                <h3 class="font-bold text-slate-800 dark:text-slate-200 text-sm flex items-center gap-2"><i class="fa-solid fa-comment-dots text-brand-600"></i> Mẫu Tin Nhắn Phối Hợp</h3>
                <button onclick="closeSmsModal()" class="text-slate-400 hover:text-slate-600"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            <div class="p-5 space-y-4">
                <div>
                    <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Đối tượng nhận:</label>
                    <div class="flex gap-4 text-xs font-medium">
                        <label class="flex items-center gap-1.5 cursor-pointer">
                            <input type="radio" name="sms-target" value="GVCN" checked onchange="generateSmsContent()" class="text-brand-600"> Gửi GVCN
                        </label>
                        <label class="flex items-center gap-1.5 cursor-pointer">
                            <input type="radio" name="sms-target" value="PHHS" onchange="generateSmsContent()" class="text-brand-600"> Gửi Phụ huynh HS
                        </label>
                    </div>
                </div>

                <div>
                    <label class="block text-xs font-bold uppercase text-slate-700 dark:text-slate-300 mb-1">Nội dung tin nhắn:</label>
                    <textarea id="sms-content-area" rows="6" readonly class="w-full bg-slate-50 dark:bg-slate-900 border border-slate-300 dark:border-slate-600 text-slate-800 dark:text-slate-200 text-xs rounded-lg p-3 font-mono"></textarea>
                </div>

                <div class="flex justify-between items-center pt-2">
                    <span id="sms-copy-msg" class="text-xs text-emerald-600 dark:text-emerald-400 font-bold opacity-0 transition-opacity"><i class="fa-solid fa-circle-check"></i> Đã sao chép!</span>
                    <div class="flex gap-2">
                        <button type="button" onclick="closeSmsModal()" class="px-4 py-2 text-xs font-semibold text-slate-600 dark:text-slate-300 bg-slate-100 dark:bg-slate-700 rounded-lg">Đóng</button>
                        <button type="button" onclick="copySmsContent()" class="px-4 py-2 text-xs font-semibold text-white bg-brand-600 hover:bg-brand-700 rounded-lg shadow-sm flex items-center gap-1.5"><i class="fa-solid fa-copy"></i> Sao Chép Tin Nhắn</button>
                    </div>
                </div>
            </div>
        </div>
    </div>


    <!-- JAVASCRIPT LOGIC -->
    <script>
        // STORAGE KEYS
        const KEY_GVBM = 'EDUTRACK_GVBM_DATA_V2';
        const KEY_GVCN = 'EDUTRACK_GVCN_DATA_V2';

        // STATE VARIABLES
        let currentRole = 'GVBM'; // 'GVBM' or 'GVCN'
        
        let gvbmData = {
            config: { subjectName: "Toán", dgtxCount: 4 },
            students: []
        };

        let gvcnData = {
            className: "6A1",
            subjects: ["Toán", "Ngữ Văn", "Tiếng Anh", "KHTN", "KHXH", "Tin Học", "GDCD"],
            students: []
        };

        let activeSmsStudent = null;
        let chartInstances = {};

        // INITIALIZATION
        window.onload = function() {
            loadAllDataFromStorage();
            renderRoleView();
        };

        // DARK MODE TOGGLE
        function toggleDarkMode() {
            document.documentElement.classList.toggle('dark');
        }

        // ROLE SWITCHER
        function switchRoleMode(role) {
            currentRole = role;

            const btnGvbm = document.getElementById('btn-mode-gvbm');
            const btnGvcn = document.getElementById('btn-mode-gvcn');
            const navGvbm = document.getElementById('nav-gvbm');
            const navGvcn = document.getElementById('nav-gvcn');
            const badge = document.getElementById('role-badge');
            const subtitle = document.getElementById('role-subtitle');
            const logoBg = document.getElementById('role-logo-bg');
            const roleIcon = document.getElementById('role-icon');

            if (role === 'GVBM') {
                btnGvbm.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 bg-white dark:bg-slate-800 text-brand-700 dark:text-brand-400 shadow-sm";
                btnGvcn.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 text-slate-600 dark:text-slate-300 hover:text-slate-900";
                navGvbm.classList.remove('hidden');
                navGvcn.classList.add('hidden');
                badge.innerText = "GV Bộ Môn";
                badge.className = "px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-brand-100 text-brand-800 dark:bg-brand-900 dark:text-brand-200";
                subtitle.innerText = "Sổ điểm & Theo dõi Bộ môn";
                logoBg.className = "p-2.5 bg-gradient-to-tr from-brand-600 to-teal-500 text-white rounded-xl shadow-md";
                roleIcon.className = "fa-solid fa-graduation-cap text-xl";
                switchTab('gvbm-dashboard');
            } else {
                btnGvcn.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 bg-white dark:bg-slate-800 text-indigo-700 dark:text-indigo-400 shadow-sm";
                btnGvbm.className = "px-3 py-1.5 rounded-lg text-xs font-bold transition-all flex items-center gap-1.5 text-slate-600 dark:text-slate-300 hover:text-slate-900";
                navGvcn.classList.remove('hidden');
                navGvbm.classList.add('hidden');
                badge.innerText = "GV Chủ Nhiệm";
                badge.className = "px-2 py-0.5 rounded text-[10px] font-bold uppercase bg-indigo-100 text-indigo-800 dark:bg-indigo-900 dark:text-indigo-200";
                subtitle.innerText = "Quản lý & Tổng hợp Lớp Chủ Nhiệm";
                logoBg.className = "p-2.5 bg-gradient-to-tr from-indigo-600 to-purple-600 text-white rounded-xl shadow-md";
                roleIcon.className = "fa-solid fa-people-roof text-xl";
                switchTab('gvcn-dashboard');
            }

            renderRoleView();
        }

        function switchTab(tabId) {
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));

            document.getElementById(tabId).classList.remove('hidden');
            const activeBtn = document.getElementById(`btn-${tabId}`);
            if (activeBtn) activeBtn.classList.add('active');

            if (tabId === 'gvbm-dashboard') renderGvbmDashboard();
            if (tabId === 'gvbm-gradebook') renderGvbmGradebook();
            if (tabId === 'gvbm-attention') renderGvbmAttentionList();
            if (tabId === 'gvcn-dashboard') renderGvcnDashboard();
            if (tabId === 'gvcn-master') renderGvcnMasterTable();
            if (tabId === 'gvcn-reports') renderGvcnReportCardSelect();
            if (tabId === 'gvcn-atrisk') renderGvcnAtriskList();
        }

        // STORAGE MANAGEMENT
        function loadAllDataFromStorage() {
            const rawBm = localStorage.getItem(KEY_GVBM);
            if (rawBm) {
                try { gvbmData = JSON.parse(rawBm); } catch(e) {}
            } else {
                loadDemoGvbmData(false);
            }

            const rawCn = localStorage.getItem(KEY_GVCN);
            if (rawCn) {
                try { gvcnData = JSON.parse(rawCn); } catch(e) {}
            } else {
                loadDemoGvcnData(false);
            }

            // Sync Config Form
            document.getElementById('gvbm-cfg-subject').value = gvbmData.config.subjectName || "Toán";
            document.getElementById('gvbm-cfg-dgtx').value = gvbmData.config.dgtxCount || 4;
            document.getElementById('gvcn-class-name').value = gvcnData.className || "6A1";
        }

        function saveGvbmData() { localStorage.setItem(KEY_GVBM, JSON.stringify(gvbmData)); }
        function saveGvcnData() { localStorage.setItem(KEY_GVCN, JSON.stringify(gvcnData)); }

        function renderRoleView() {
            if (currentRole === 'GVBM') {
                recalculateGvbmGrades();
                populateGvbmClassDropdowns();
                renderGvbmDashboard();
            } else {
                recalculateGvcnGrades();
                renderGvcnDashboard();
            }
        }

        // ==================== GVBM CORE CALCULATIONS & RENDER ====================

        function calculateGvbmStudent(s) {
            const cols = gvbmData.config.dgtxCount;
            let sumTx = 0;
            let countTx = 0;

            for (let i = 1; i <= cols; i++) {
                const val = s.dgtx[`tx${i}`];
                if (val !== undefined && val !== null && val !== '') {
                    sumTx += parseFloat(val);
                    countTx++;
                }
            }

            const dgk = (s.dgk !== undefined && s.dgk !== null && s.dgk !== '') ? parseFloat(s.dgk) : null;

            if (countTx > 0 || dgk !== null) {
                const totalPoints = sumTx + (dgk !== null ? dgk * 2 : 0);
                const totalWeight = countTx + (dgk !== null ? 2 : 0);
                s.dtb = totalWeight > 0 ? Math.round((totalPoints / totalWeight) * 10) / 10 : null;
            } else {
                s.dtb = null;
            }

            if (s.dtb === null) s.level = "Chưa có";
            else if (s.dtb >= 8.0) s.level = "Tốt";
            else if (s.dtb >= 6.5) s.level = "Khá";
            else if (s.dtb >= 5.0) s.level = "Đạt";
            else s.level = "Chưa đạt";
        }

        function recalculateGvbmGrades() {
            gvbmData.students.forEach(s => calculateGvbmStudent(s));
            saveGvbmData();
        }

        function populateGvbmClassDropdowns() {
            const classes = [...new Set(gvbmData.students.map(s => s.className))].filter(Boolean).sort();
            const dashSel = document.getElementById('gvbm-filter-class');
            const gbSel = document.getElementById('gvbm-gradebook-filter-class');

            const dashVal = dashSel.value;
            const gbVal = gbSel.value;

            dashSel.innerHTML = '<option value="ALL">Tất cả các lớp</option>';
            gbSel.innerHTML = '<option value="ALL">Tất cả các lớp</option>';

            classes.forEach(c => {
                dashSel.innerHTML += `<option value="${c}">Lớp ${c}</option>`;
                gbSel.innerHTML += `<option value="${c}">Lớp ${c}</option>`;
            });

            if (classes.includes(dashVal)) dashSel.value = dashVal;
            if (classes.includes(gbVal)) gbSel.value = gbVal;
        }

        function renderGvbmDashboard() {
            document.getElementById('gvbm-subject-display-name').innerText = gvbmData.config.subjectName || "Toán";
            const filterClass = document.getElementById('gvbm-filter-class').value;

            let list = gvbmData.students;
            if (filterClass !== 'ALL') list = list.filter(s => s.className === filterClass);

            const total = list.length;
            const graded = list.filter(s => s.dtb !== null);

            let max = 0, min = 10, sum = 0, passCount = 0;
            let counts = { "Tốt": 0, "Khá": 0, "Đạt": 0, "Chưa đạt": 0 };

            if (graded.length > 0) {
                graded.forEach(s => {
                    if (s.dtb > max) max = s.dtb;
                    if (s.dtb < min) min = s.dtb;
                    sum += s.dtb;
                    if (s.dtb >= 5.0) passCount++;
                    if (counts[s.level] !== undefined) counts[s.level]++;
                });
            } else min = 0;

            document.getElementById('gvbm-stat-total').innerText = total;
            document.getElementById('gvbm-stat-avg').innerText = graded.length > 0 ? (sum / graded.length).toFixed(1) : "0.0";
            document.getElementById('gvbm-stat-max').innerText = graded.length > 0 ? max.toFixed(1) : "0.0";
            document.getElementById('gvbm-stat-min').innerText = graded.length > 0 ? min.toFixed(1) : "0.0";
            document.getElementById('gvbm-stat-pass').innerText = graded.length > 0 ? ((passCount / graded.length) * 100).toFixed(1) + "%" : "0%";

            renderGvbmCharts(counts);
        }

        function renderGvbmCharts(counts) {
            const labels = ['Tốt (≥8.0)', 'Khá (6.5-7.9)', 'Đạt (5.0-6.4)', 'Chưa đạt (<5.0)'];
            const values = [counts["Tốt"], counts["Khá"], counts["Đạt"], counts["Chưa đạt"]];
            const colors = ['#10b981', '#0284c7', '#f59e0b', '#ef4444'];

            if (chartInstances.gvbmDoughnut) chartInstances.gvbmDoughnut.destroy();
            chartInstances.gvbmDoughnut = new Chart(document.getElementById('gvbmChartDoughnut'), {
                type: 'doughnut',
                data: { labels: labels, datasets: [{ data: values, backgroundColor: colors }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
            });

            if (chartInstances.gvbmBar) chartInstances.gvbmBar.destroy();
            chartInstances.gvbmBar = new Chart(document.getElementById('gvbmChartBar'), {
                type: 'bar',
                data: { labels: labels, datasets: [{ label: 'Số lượng HS', data: values, backgroundColor: colors, borderRadius: 6 }] },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
            });
        }

        function renderGvbmGradebook() {
            const cols = gvbmData.config.dgtxCount;
            const header = document.getElementById('gvbm-header-row');

            let headerHtml = `
                <th class="p-2.5 w-10 text-center border-r">STT</th>
                <th class="p-2.5 min-w-[160px] border-r">Họ và Tên</th>
                <th class="p-2.5 w-20 text-center border-r">Lớp</th>
            `;
            for (let i = 1; i <= cols; i++) {
                headerHtml += `<th class="p-2.5 w-20 text-center border-r bg-slate-50 dark:bg-slate-800">ĐGTX ${i}</th>`;
            }
            headerHtml += `
                <th class="p-2.5 w-24 text-center border-r bg-indigo-50/70 dark:bg-indigo-950/70 text-indigo-900 dark:text-indigo-200">ĐGK (x2)</th>
                <th class="p-2.5 w-24 text-center border-r bg-brand-50 dark:bg-brand-950 text-brand-900 dark:text-brand-200">ĐTB Môn</th>
                <th class="p-2.5 w-28 text-center border-r">Mức Độ</th>
                <th class="p-2.5 min-w-[180px] border-r">Nhận Xét Bộ Môn</th>
                <th class="p-2.5 w-14 text-center no-print">Xóa</th>
            `;
            header.innerHTML = headerHtml;

            const search = document.getElementById('gvbm-search').value.toLowerCase().trim();
            const filterClass = document.getElementById('gvbm-gradebook-filter-class').value;

            let filtered = gvbmData.students.filter(s => {
                const matchSearch = s.name.toLowerCase().includes(search);
                const matchClass = (filterClass === 'ALL' || s.className === filterClass);
                return matchSearch && matchClass;
            });

            const tbody = document.getElementById('gvbm-tbody');
            tbody.innerHTML = '';

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="${cols + 7}" class="p-8 text-center text-slate-400">Không có dữ liệu học sinh.</td></tr>`;
                return;
            }

            filtered.forEach((s, idx) => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50/80 dark:hover:bg-slate-700/50";

                let html = `
                    <td class="p-2 text-center border-r text-slate-400">${idx + 1}</td>
                    <td class="p-2 border-r font-bold text-slate-800 dark:text-slate-100">${escapeHtml(s.name)}</td>
                    <td class="p-2 text-center border-r font-bold text-slate-600 dark:text-slate-300">${escapeHtml(s.className)}</td>
                `;

                for (let i = 1; i <= cols; i++) {
                    const val = (s.dgtx && s.dgtx[`tx${i}`] !== undefined) ? s.dgtx[`tx${i}`] : '';
                    html += `
                        <td class="p-1 border-r text-center">
                            <input type="number" step="0.1" min="0" max="10" value="${val}"
                                data-id="${s.id}" data-type="tx${i}" onchange="handleGvbmInputChange(this)"
                                class="w-full text-center py-1 bg-transparent hover:bg-white dark:hover:bg-slate-900 focus:bg-white dark:focus:bg-slate-900 focus:ring-2 focus:ring-brand-500 rounded font-semibold text-slate-800 dark:text-slate-100 outline-none">
                        </td>
                    `;
                }

                const dgk = s.dgk !== undefined ? s.dgk : '';
                html += `
                    <td class="p-1 border-r text-center bg-indigo-50/30 dark:bg-indigo-950/30">
                        <input type="number" step="0.1" min="0" max="10" value="${dgk}"
                            data-id="${s.id}" data-type="dgk" onchange="handleGvbmInputChange(this)"
                            class="w-full text-center py-1 bg-transparent hover:bg-white dark:hover:bg-slate-900 focus:bg-white dark:focus:bg-slate-900 focus:ring-2 focus:ring-indigo-500 rounded font-bold text-indigo-700 dark:text-indigo-300 outline-none">
                    </td>
                    <td class="p-2 text-center border-r font-bold text-sm ${s.dtb !== null && s.dtb < 5.0 ? 'text-rose-600 bg-rose-50/50 dark:bg-rose-950/40' : 'text-slate-800 dark:text-white'}">
                        ${s.dtb !== null ? s.dtb.toFixed(1) : '-'}
                    </td>
                    <td class="p-2 text-center border-r">
                        <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${getBadgeClass(s.level)}">${s.level}</span>
                    </td>
                    <td class="p-1 border-r">
                        <input type="text" value="${escapeHtml(s.note || '')}"
                            data-id="${s.id}" data-type="note" onchange="handleGvbmInputChange(this)" placeholder="Nhận xét môn..."
                            class="w-full py-1 px-2 text-xs bg-transparent hover:bg-white dark:hover:bg-slate-900 border border-transparent focus:ring-1 focus:ring-brand-500 rounded">
                    </td>
                    <td class="p-2 text-center no-print">
                        <button onclick="deleteGvbmStudent('${s.id}')" class="text-slate-300 hover:text-rose-600 p-1"><i class="fa-solid fa-trash-can"></i></button>
                    </td>
                `;

                tr.innerHTML = html;
                tbody.appendChild(tr);
            });
        }

        function handleGvbmInputChange(input) {
            const id = input.getAttribute('data-id');
            const type = input.getAttribute('data-type');
            let val = input.value.trim();

            const student = gvbmData.students.find(s => s.id === id);
            if (!student) return;

            if (type.startsWith('tx')) {
                if (val !== '') val = Math.min(10, Math.max(0, parseFloat(val)));
                student.dgtx[type] = val;
            } else if (type === 'dgk') {
                if (val !== '') val = Math.min(10, Math.max(0, parseFloat(val)));
                student.dgk = val;
            } else if (type === 'note') {
                student.note = val;
            }

            calculateGvbmStudent(student);
            saveGvbmData();
            renderGvbmGradebook();
            renderGvbmAttentionList();
        }

        function renderGvbmAttentionList() {
            const attention = gvbmData.students.filter(s => s.dtb !== null && s.dtb < 5.0);
            const tbody = document.getElementById('gvbm-attention-tbody');
            const badge = document.getElementById('gvbm-attention-badge');

            if (attention.length > 0) {
                badge.innerText = attention.length;
                badge.classList.remove('hidden');
            } else badge.classList.add('hidden');

            tbody.innerHTML = '';
            if (attention.length === 0) {
                tbody.innerHTML = `<tr><td colspan="7" class="p-8 text-center text-emerald-600 font-bold"><i class="fa-solid fa-circle-check text-xl mb-1 block"></i>Không có học sinh nào bị sa sút môn ${gvbmData.config.subjectName}!</td></tr>`;
                return;
            }

            attention.forEach((s, idx) => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-rose-50/30 dark:hover:bg-rose-950/20";
                tr.innerHTML = `
                    <td class="p-3 text-center text-slate-400 font-medium">${idx + 1}</td>
                    <td class="p-3 font-bold text-slate-800 dark:text-slate-100">${escapeHtml(s.name)}</td>
                    <td class="p-3 text-center font-bold text-slate-600 dark:text-slate-300">${escapeHtml(s.className)}</td>
                    <td class="p-3 text-center font-semibold text-rose-600">${s.dgk || '-'}</td>
                    <td class="p-3 text-center font-bold text-rose-600 text-sm">${s.dtb.toFixed(1)}</td>
                    <td class="p-3 text-xs text-slate-600 dark:text-slate-300">${escapeHtml(s.note || 'ĐTB môn dưới 5.0')}</td>
                    <td class="p-3 text-center no-print">
                        <button onclick="openSmsModal('${s.id}')" class="px-2.5 py-1 text-xs font-bold text-brand-700 dark:text-brand-300 bg-brand-50 dark:bg-brand-950 rounded-lg hover:bg-brand-100"><i class="fa-solid fa-paper-plane mr-1"></i> Soạn Zalo/SMS</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function saveGvbmConfig() {
            const subject = document.getElementById('gvbm-cfg-subject').value.trim();
            const cols = parseInt(document.getElementById('gvbm-cfg-dgtx').value);

            if (!subject || isNaN(cols) || cols < 1 || cols > 8) {
                alert("Vui lòng nhập tên môn và số cột ĐGTX hợp lệ (1-8).");
                return;
            }

            gvbmData.config.subjectName = subject;
            gvbmData.config.dgtxCount = cols;
            saveGvbmData();
            renderRoleView();
            alert("Đã lưu cấu hình bộ môn!");
        }

        // ==================== GVCN CORE CALCULATIONS & RENDER ====================

        function calculateGvcnStudent(s) {
            const subjects = gvcnData.subjects;
            let sum = 0;
            let count = 0;
            let weakSubjects = [];

            subjects.forEach(subj => {
                const val = s.grades[subj];
                if (val !== undefined && val !== null && val !== '') {
                    const num = parseFloat(val);
                    sum += num;
                    count++;
                    if (num < 5.0) weakSubjects.push(`${subj} (${num})`);
                }
            });

            s.gpa = count > 0 ? Math.round((sum / count) * 10) / 10 : null;
            s.weakSubjects = weakSubjects;

            // Classification TT 22
            if (s.gpa === null) s.rank = "Chưa xếp loại";
            else if (s.gpa >= 8.0 && weakSubjects.length === 0) s.rank = "Tốt";
            else if (s.gpa >= 6.5 && weakSubjects.length <= 1) s.rank = "Khá";
            else if (s.gpa >= 5.0 && weakSubjects.length <= 2) s.rank = "Đạt";
            else s.rank = "Chưa đạt";
        }

        function recalculateGvcnGrades() {
            gvcnData.students.forEach(s => calculateGvcnStudent(s));
            saveGvcnData();
        }

        function updateGvcnClassName(val) {
            gvcnData.className = val.trim().toUpperCase();
            saveGvcnData();
        }

        function renderGvcnDashboard() {
            const list = gvcnData.students;
            const total = list.length;

            let good = 0, kha = 0, dat = 0, weak = 0;
            let subjectSums = {};
            let subjectCounts = {};

            gvcnData.subjects.forEach(sb => { subjectSums[sb] = 0; subjectCounts[sb] = 0; });

            list.forEach(s => {
                if (s.rank === "Tốt") good++;
                else if (s.rank === "Khá") kha++;
                else if (s.rank === "Đạt") dat++;
                else if (s.rank === "Chưa đạt") weak++;

                gvcnData.subjects.forEach(sb => {
                    if (s.grades[sb] !== undefined && s.grades[sb] !== '') {
                        subjectSums[sb] += parseFloat(s.grades[sb]);
                        subjectCounts[sb]++;
                    }
                });
            });

            document.getElementById('gvcn-stat-total').innerText = total;
            document.getElementById('gvcn-stat-good').innerText = good;
            document.getElementById('gvcn-stat-kha').innerText = kha;
            document.getElementById('gvcn-stat-dat').innerText = dat;
            document.getElementById('gvcn-stat-weak').innerText = weak;

            renderGvcnCharts({ "Tốt": good, "Khá": kha, "Đạt": dat, "Chưa đạt": weak }, subjectSums, subjectCounts);
        }

        function renderGvcnCharts(rankCounts, subjectSums, subjectCounts) {
            // Doughnut Chart
            if (chartInstances.gvcnDoughnut) chartInstances.gvcnDoughnut.destroy();
            chartInstances.gvcnDoughnut = new Chart(document.getElementById('gvcnChartDoughnut'), {
                type: 'doughnut',
                data: {
                    labels: ['Tốt / Giỏi', 'Khá', 'Đạt / TB', 'Chưa đạt / Yếu'],
                    datasets: [{ data: [rankCounts["Tốt"], rankCounts["Khá"], rankCounts["Đạt"], rankCounts["Chưa đạt"]], backgroundColor: ['#10b981', '#0284c7', '#f59e0b', '#ef4444'] }]
                },
                options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom' } } }
            });

            // Bar Chart Subject Averages
            const subjLabels = gvcnData.subjects;
            const subjAvgs = subjLabels.map(sb => subjectCounts[sb] > 0 ? (subjectSums[sb] / subjectCounts[sb]).toFixed(1) : 0);

            if (chartInstances.gvcnBar) chartInstances.gvcnBar.destroy();
            chartInstances.gvcnBar = new Chart(document.getElementById('gvcnChartBar'), {
                type: 'bar',
                data: {
                    labels: subjLabels,
                    datasets: [{ label: 'ĐTB Môn Chung', data: subjAvgs, backgroundColor: '#0369a1', borderRadius: 4 }]
                },
                options: { responsive: true, maintainAspectRatio: false, scales: { y: { min: 0, max: 10 } } }
            });
        }

        function renderGvcnMasterTable() {
            const header = document.getElementById('gvcn-master-header');
            const subjects = gvcnData.subjects;

            let headerHtml = `
                <th class="p-2.5 w-10 text-center border-r">STT</th>
                <th class="p-2.5 min-w-[160px] border-r">Họ và Tên</th>
            `;
            subjects.forEach(sb => {
                headerHtml += `<th class="p-2.5 w-20 text-center border-r bg-slate-50 dark:bg-slate-800">${escapeHtml(sb)}</th>`;
            });
            headerHtml += `
                <th class="p-2.5 w-20 text-center border-r bg-brand-50 dark:bg-brand-950 text-brand-900 dark:text-brand-200">ĐTB Tổng</th>
                <th class="p-2.5 w-24 text-center border-r">Xếp Loại</th>
                <th class="p-2.5 w-14 text-center no-print">Xóa</th>
            `;
            header.innerHTML = headerHtml;

            const search = document.getElementById('gvcn-search').value.toLowerCase().trim();
            const filtered = gvcnData.students.filter(s => s.name.toLowerCase().includes(search));

            const tbody = document.getElementById('gvcn-master-tbody');
            tbody.innerHTML = '';

            if (filtered.length === 0) {
                tbody.innerHTML = `<tr><td colspan="${subjects.length + 5}" class="p-8 text-center text-slate-400">Chưa có dữ liệu học sinh lớp chủ nhiệm.</td></tr>`;
                return;
            }

            filtered.forEach((s, idx) => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-50/80 dark:hover:bg-slate-700/50";

                let html = `
                    <td class="p-2 text-center border-r text-slate-400">${idx + 1}</td>
                    <td class="p-2 border-r font-bold text-slate-800 dark:text-slate-100">${escapeHtml(s.name)}</td>
                `;

                subjects.forEach(sb => {
                    const val = s.grades[sb] !== undefined ? s.grades[sb] : '';
                    html += `
                        <td class="p-1 border-r text-center">
                            <input type="number" step="0.1" min="0" max="10" value="${val}"
                                data-id="${s.id}" data-subject="${sb}" onchange="handleGvcnGradeChange(this)"
                                class="w-full text-center py-1 bg-transparent hover:bg-white dark:hover:bg-slate-900 focus:bg-white dark:focus:bg-slate-900 focus:ring-1 focus:ring-brand-500 font-semibold text-slate-800 dark:text-slate-100 rounded">
                        </td>
                    `;
                });

                html += `
                    <td class="p-2 text-center border-r font-bold text-sm text-brand-700 dark:text-brand-300 bg-brand-50/40 dark:bg-brand-950/40">
                        ${s.gpa !== null ? s.gpa.toFixed(1) : '-'}
                    </td>
                    <td class="p-2 text-center border-r">
                        <span class="px-2 py-0.5 rounded-full text-[11px] font-bold ${getBadgeClass(s.rank)}">${s.rank}</span>
                    </td>
                    <td class="p-2 text-center no-print">
                        <button onclick="deleteGvcnStudent('${s.id}')" class="text-slate-300 hover:text-rose-600 p-1"><i class="fa-solid fa-trash-can"></i></button>
                    </td>
                `;

                tr.innerHTML = html;
                tbody.appendChild(tr);
            });
        }

        function handleGvcnGradeChange(input) {
            const id = input.getAttribute('data-id');
            const sb = input.getAttribute('data-subject');
            let val = input.value.trim();

            const student = gvcnData.students.find(s => s.id === id);
            if (!student) return;

            if (val !== '') val = Math.min(10, Math.max(0, parseFloat(val)));
            student.grades[sb] = val;

            calculateGvcnStudent(student);
            saveGvcnData();
            renderGvcnMasterTable();
            renderGvcnAtriskList();
        }

        function renderGvcnAtriskList() {
            const list = gvcnData.students.filter(s => s.weakSubjects.length >= 2);
            const tbody = document.getElementById('gvcn-atrisk-tbody');
            const badge = document.getElementById('gvcn-atrisk-badge');

            if (list.length > 0) {
                badge.innerText = list.length;
                badge.classList.remove('hidden');
            } else badge.classList.add('hidden');

            tbody.innerHTML = '';
            if (list.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="p-8 text-center text-emerald-600 font-bold"><i class="fa-solid fa-circle-check text-xl mb-1 block"></i>Không có học sinh nào bị yếu/chưa đạt trên 2 môn học!</td></tr>`;
                return;
            }

            list.forEach((s, idx) => {
                const tr = document.createElement('tr');
                tr.className = "hover:bg-rose-50/30 dark:hover:bg-rose-950/20";
                tr.innerHTML = `
                    <td class="p-3 text-center text-slate-400 font-medium">${idx + 1}</td>
                    <td class="p-3 font-bold text-slate-800 dark:text-slate-100">${escapeHtml(s.name)}</td>
                    <td class="p-3 text-center font-bold text-rose-600">${s.weakSubjects.length} Môn</td>
                    <td class="p-3 font-medium text-rose-600 text-xs">${escapeHtml(s.weakSubjects.join(', '))}</td>
                    <td class="p-3 text-center font-bold text-sm">${s.gpa ? s.gpa.toFixed(1) : '-'}</td>
                    <td class="p-3 text-center no-print">
                        <button onclick="switchTab('gvcn-reports'); document.getElementById('gvcn-report-student-select').value='${s.id}'; renderStudentReportCard();" class="px-2.5 py-1 text-xs font-bold text-brand-700 dark:text-brand-300 bg-brand-50 dark:bg-brand-950 rounded-lg hover:bg-brand-100"><i class="fa-solid fa-id-card mr-1"></i> Xem Phiếu Điểm</button>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        // REPORT CARD SLIP GENERATOR
        function renderGvcnReportCardSelect() {
            const select = document.getElementById('gvcn-report-student-select');
            select.innerHTML = '';

            if (gvcnData.students.length === 0) {
                select.innerHTML = '<option value="">(Chưa có học sinh)</option>';
                return;
            }

            gvcnData.students.forEach(s => {
                select.innerHTML += `<option value="${s.id}">${s.name} (Lớp ${gvcnData.className})</option>`;
            });

            renderStudentReportCard();
        }

        function renderStudentReportCard() {
            const id = document.getElementById('gvcn-report-student-select').value;
            const student = gvcnData.students.find(s => s.id === id);

            if (!student) return;

            document.getElementById('rpt-student-name').innerText = student.name;
            document.getElementById('rpt-student-class').innerText = gvcnData.className;
            document.getElementById('rpt-student-gpa').innerText = student.gpa !== null ? student.gpa.toFixed(1) : '-';
            document.getElementById('rpt-student-rank').innerText = student.rank;

            const tbody = document.getElementById('rpt-grades-tbody');
            tbody.innerHTML = '';

            gvcnData.subjects.forEach((sb, idx) => {
                const val = student.grades[sb] !== undefined && student.grades[sb] !== '' ? parseFloat(student.grades[sb]).toFixed(1) : '-';
                let evalText = "Đạt";
                if (val !== '-' && parseFloat(val) < 5.0) evalText = "Chưa đạt";
                else if (val !== '-' && parseFloat(val) >= 8.0) evalText = "Tốt";

                tbody.innerHTML += `
                    <tr>
                        <td class="p-2 text-center border-r border-slate-300">${idx + 1}</td>
                        <td class="p-2 border-r border-slate-300 font-semibold">${escapeHtml(sb)}</td>
                        <td class="p-2 text-center border-r border-slate-300 font-bold ${val !== '-' && parseFloat(val) < 5.0 ? 'text-rose-600' : ''}">${val}</td>
                        <td class="p-2 text-center border-r border-slate-300">${evalText}</td>
                    </tr>
                `;
            });

            const commentInput = document.getElementById('rpt-teacher-comment');
            const commentPrint = document.getElementById('rpt-teacher-comment-print');
            commentInput.value = student.teacherComment || '';
            commentPrint.innerText = student.teacherComment || 'Em chấp hành tốt nội quy trường lớp, cần phát huy kết quả học tập trong thời gian tới.';
        }

        function saveGvcnComment(val) {
            const id = document.getElementById('gvcn-report-student-select').value;
            const student = gvcnData.students.find(s => s.id === id);
            if (student) {
                student.teacherComment = val;
                saveGvcnData();
                document.getElementById('rpt-teacher-comment-print').innerText = val;
            }
        }

        function copyParentZaloSlip() {
            const id = document.getElementById('gvcn-report-student-select').value;
            const student = gvcnData.students.find(s => s.id === id);
            if (!student) return;

            let msg = `Kính gửi Phụ huynh em ${student.name} (Lớp ${gvcnData.className}),\nGVCN xin thông báo Phiếu điểm Giữa HK1 các môn của em:\n`;
            gvcnData.subjects.forEach(sb => {
                const val = student.grades[sb] !== undefined && student.grades[sb] !== '' ? student.grades[sb] : '-';
                msg += `- Môn ${sb}: ${val}\n`;
            });
            msg += `=> ĐTB TỔNG HỢP: ${student.gpa ? student.gpa.toFixed(1) : '-'}\n`;
            msg += `=> XẾP LOẠI HỌC LỰC: ${student.rank}\n`;
            if (student.teacherComment) msg += `* Nhận xét GVCN: ${student.teacherComment}\n`;
            msg += `Trân trọng cảm ơn sự phối hợp của Quý Phụ huynh!`;

            copyToClipboard(msg);
            alert("Đã sao chép Phiếu báo điểm chi tiết để dán vào Zalo cho Phụ huynh!");
        }

        // ==================== COMMON UTILITIES & MODALS ====================

        function openAddStudentModal() { document.getElementById('modal-add-student').classList.remove('hidden'); }
        function closeAddStudentModal() { document.getElementById('modal-add-student').classList.add('hidden'); }
        function handleAddStudentSubmit(e) {
            e.preventDefault();
            const name = document.getElementById('modal-add-name').value.trim();
            const className = document.getElementById('modal-add-class').value.trim().toUpperCase();

            if (currentRole === 'GVBM') {
                gvbmData.students.push({
                    id: 'bm_' + Date.now(),
                    name: name,
                    className: className,
                    dgtx: {}, dgk: '', dtb: null, level: 'Chưa có', note: ''
                });
                saveGvbmData();
            } else {
                gvcnData.students.push({
                    id: 'cn_' + Date.now(),
                    name: name,
                    grades: {}, gpa: null, rank: 'Chưa xếp loại', weakSubjects: []
                });
                saveGvcnData();
            }

            closeAddStudentModal();
            renderRoleView();
            document.getElementById('modal-add-name').value = '';
        }

        function openBatchAddModal() { document.getElementById('modal-batch-add').classList.remove('hidden'); }
        function closeBatchAddModal() { document.getElementById('modal-batch-add').classList.add('hidden'); }
        function handleBatchAddSubmit(e) {
            e.preventDefault();
            const className = document.getElementById('modal-batch-class').value.trim().toUpperCase();
            const rawText = document.getElementById('modal-batch-text').value;
            const names = rawText.split('\n').map(n => n.trim().replace(/^[0-9]+\s*[\.\-\)]\s*/, '')).filter(n => n.length > 0);

            if (currentRole === 'GVBM') {
                names.forEach(name => {
                    gvbmData.students.push({
                        id: 'bm_' + Date.now() + Math.random().toString(36).substr(2,3),
                        name: name, className: className, dgtx: {}, dgk: '', dtb: null, level: 'Chưa có', note: ''
                    });
                });
                saveGvbmData();
            } else {
                names.forEach(name => {
                    gvcnData.students.push({
                        id: 'cn_' + Date.now() + Math.random().toString(36).substr(2,3),
                        name: name, grades: {}, gpa: null, rank: 'Chưa xếp loại', weakSubjects: []
                    });
                });
                saveGvcnData();
            }

            closeBatchAddModal();
            renderRoleView();
            document.getElementById('modal-batch-text').value = '';
        }

        function deleteGvbmStudent(id) {
            if (confirm("Xóa học sinh này khỏi danh sách bộ môn?")) {
                gvbmData.students = gvbmData.students.filter(s => s.id !== id);
                saveGvbmData();
                renderRoleView();
            }
        }

        function deleteGvcnStudent(id) {
            if (confirm("Xóa học sinh này khỏi lớp chủ nhiệm?")) {
                gvcnData.students = gvcnData.students.filter(s => s.id !== id);
                saveGvcnData();
                renderRoleView();
            }
        }

        // SMS ZALO MODAL
        function openSmsModal(id) {
            activeSmsStudent = gvbmData.students.find(s => s.id === id);
            if (!activeSmsStudent) return;

            document.getElementById('modal-sms').classList.remove('hidden');
            generateSmsContent();
        }
        function closeSmsModal() { document.getElementById('modal-sms').classList.add('hidden'); }
        function generateSmsContent() {
            if (!activeSmsStudent) return;
            const target = document.querySelector('input[name="sms-target"]:checked').value;
            const s = activeSmsStudent;
            const subj = gvbmData.config.subjectName || "Môn học";

            let msg = "";
            if (target === 'GVCN') {
                msg = `Thầy/Cô GVCN Lớp ${s.className} ơi, em gửi kết quả Giữa HK1 môn ${subj} của em ${s.name}: ĐTB Môn đạt ${s.dtb ? s.dtb.toFixed(1) : '-'}. Nhờ Thầy/Cô nhắc nhở em củng cố bài vở giúp bộ môn ạ. Cảm ơn Thầy/Cô!`;
            } else {
                msg = `Kính gửi Phụ huynh em ${s.name} (Lớp ${s.className}), Giáo viên bộ môn ${subj} xin thông báo ĐTB Giữa HK1 của em là ${s.dtb ? s.dtb.toFixed(1) : '-'}. Kính mong Gia đình đôn đốc em ôn tập thêm tại nhà. Trân trọng!`;
            }

            document.getElementById('sms-content-area').value = msg;
        }

        function copySmsContent() {
            copyToClipboard(document.getElementById('sms-content-area').value);
            const copyMsg = document.getElementById('sms-copy-msg');
            copyMsg.classList.remove('opacity-0');
            setTimeout(() => copyMsg.classList.add('opacity-0'), 2000);
        }

        function copyToClipboard(text) {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            document.body.appendChild(textarea);
            textarea.select();
            try { document.execCommand('copy'); } catch(e) {}
            document.body.removeChild(textarea);
        }

        // EXPORT / IMPORT EXCEL
        function exportGvbmToExcel() {
            const exportData = gvbmData.students.map((s, idx) => {
                let row = { "STT": idx + 1, "Họ và Tên": s.name, "Lớp": s.className };
                for (let i = 1; i <= gvbmData.config.dgtxCount; i++) row[`ĐGTX ${i}`] = s.dgtx[`tx${i}`] || '';
                row["ĐGK (x2)"] = s.dgk || '';
                row["ĐTB Môn Giữa HK1"] = s.dtb !== null ? s.dtb.toFixed(1) : '';
                row["Mức Độ"] = s.level;
                row["Ghi Chú"] = s.note || '';
                return row;
            });

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "BangDiemMon");
            XLSX.writeFile(wb, `BaoCao_Mon_${gvbmData.config.subjectName}_GiuHK1.xlsx`);
        }

        function importGvbmExcel(evt) {
            const file = evt.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(sheet);

                    rows.forEach(r => {
                        const name = r["Họ và Tên"] || r["Họ tên"] || "";
                        const className = r["Lớp"] || "6A1";
                        if (name) {
                            gvbmData.students.push({
                                id: 'bm_' + Date.now() + Math.random().toString(36).substr(2,3),
                                name: String(name).trim(),
                                className: String(className).trim().toUpperCase(),
                                dgtx: {}, dgk: '', dtb: null, level: 'Chưa có', note: ''
                            });
                        }
                    });
                    saveGvbmData();
                    renderRoleView();
                    alert("Đã nhập thành công danh sách từ Excel!");
                } catch(err) { alert("Lỗi định dạng file Excel!"); }
            };
            reader.readAsArrayBuffer(file);
        }

        function exportGvcnMasterToExcel() {
            const exportData = gvcnData.students.map((s, idx) => {
                let row = { "STT": idx + 1, "Họ và Tên": s.name };
                gvcnData.subjects.forEach(sb => { row[sb] = s.grades[sb] || ''; });
                row["ĐTB Tổng Hợp"] = s.gpa !== null ? s.gpa.toFixed(1) : '';
                row["Xếp Loại Học Lực"] = s.rank;
                return row;
            });

            const ws = XLSX.utils.json_to_sheet(exportData);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "TongHopCacMon");
            XLSX.writeFile(wb, `BangDiemTongHop_Lop_${gvcnData.className}_GiuHK1.xlsx`);
        }

        function importGvcnSubjectFile(evt) {
            const file = evt.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const wb = XLSX.read(data, { type: 'array' });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    const rows = XLSX.utils.sheet_to_json(sheet);

                    let subjectName = prompt("Nhập tên môn học muốn nhập điểm (VD: Toán, Ngữ Văn, Tiếng Anh):", "Toán");
                    if (!subjectName) return;

                    if (!gvcnData.subjects.includes(subjectName)) {
                        gvcnData.subjects.push(subjectName);
                    }

                    rows.forEach(r => {
                        const name = r["Họ và Tên"] || r["Họ tên"] || "";
                        const score = r["ĐTB Môn Giữa HK1"] || r["ĐTB Môn"] || r[subjectName] || "";

                        if (name) {
                            let student = gvcnData.students.find(s => s.name.toLowerCase() === String(name).trim().toLowerCase());
                            if (!student) {
                                student = {
                                    id: 'cn_' + Date.now() + Math.random().toString(36).substr(2,3),
                                    name: String(name).trim(),
                                    grades: {}, gpa: null, rank: 'Chưa xếp loại', weakSubjects: []
                                };
                                gvcnData.students.push(student);
                            }
                            if (score !== '') student.grades[subjectName] = score;
                        }
                    });

                    saveGvcnData();
                    renderRoleView();
                    alert(`Đã ghép thành công điểm môn ${subjectName} vào bảng điểm tổng hợp lớp!`);
                } catch(err) { alert("Lỗi đọc file ghép điểm!"); }
            };
            reader.readAsArrayBuffer(file);
        }

        // DEMO DATA PRESETS
        function loadDemoGvbmData(confirmNeeded = true) {
            if (confirmNeeded && !confirm("Nạp dữ liệu mẫu cho Giáo viên Bộ môn?")) return;

            gvbmData = {
                config: { subjectName: "Toán", dgtxCount: 4 },
                students: [
                    { id: 'bm_1', name: 'Nguyễn Trần Minh Anh', className: '6A1', dgtx: { tx1: 8.5, tx2: 9.0, tx3: 8.0, tx4: 8.5 }, dgk: 8.5, dtb: null, level: '', note: 'Học tiếp thu nhanh' },
                    { id: 'bm_2', name: 'Lê Hoàng Quốc Bảo', className: '6A1', dgtx: { tx1: 7.0, tx2: 6.5, tx3: 7.5, tx4: 7.0 }, dgk: 7.0, dtb: null, level: '', note: '' },
                    { id: 'bm_3', name: 'Phạm Vũ Bảo Châu', className: '6A1', dgtx: { tx1: 5.0, tx2: 5.5, tx3: 6.0, tx4: 5.0 }, dgk: 5.5, dtb: null, level: '', note: '' },
                    { id: 'bm_4', name: 'Trần Đình Quốc Dũng', className: '6A1', dgtx: { tx1: 4.0, tx2: 3.5, tx3: 4.5, tx4: 4.0 }, dgk: 3.5, dtb: null, level: '', note: 'Tính toán còn sai sót nhiều' },
                    { id: 'bm_5', name: 'Ngô Thanh Hương Giang', className: '6A1', dgtx: { tx1: 9.0, tx2: 9.5, tx3: 9.0, tx4: 10 }, dgk: 9.5, dtb: null, level: '', note: 'Xuất sắc' }
                ]
            };
            saveGvbmData();
            if (confirmNeeded) renderRoleView();
        }

        function loadDemoGvcnData(confirmNeeded = true) {
            if (confirmNeeded && !confirm("Nạp dữ liệu mẫu cho Giáo viên Chủ nhiệm?")) return;

            gvcnData = {
                className: "6A1",
                subjects: ["Toán", "Ngữ Văn", "Tiếng Anh", "KHTN", "KHXH", "Tin Học"],
                students: [
                    { id: 'cn_1', name: 'Nguyễn Trần Minh Anh', grades: { "Toán": 8.6, "Ngữ Văn": 8.0, "Tiếng Anh": 9.0, "KHTN": 8.5, "KHXH": 8.8, "Tin Học": 9.5 }, gpa: null, rank: '', weakSubjects: [], teacherComment: 'Em chăm ngoan, có ý thức học tập xuất sắc.' },
                    { id: 'cn_2', name: 'Lê Hoàng Quốc Bảo', grades: { "Toán": 7.0, "Ngữ Văn": 6.8, "Tiếng Anh": 7.5, "KHTN": 7.2, "KHXH": 7.0, "Tin Học": 8.0 }, gpa: null, rank: '', weakSubjects: [], teacherComment: 'Cần tích cực phát biểu xây dựng bài hơn.' },
                    { id: 'cn_3', name: 'Phạm Vũ Bảo Châu', grades: { "Toán": 5.4, "Ngữ Văn": 6.0, "Tiếng Anh": 5.5, "KHTN": 5.8, "KHXH": 6.2, "Tin Học": 6.0 }, gpa: null, rank: '', weakSubjects: [], teacherComment: 'Chú ý tập trung hơn trong giờ học.' },
                    { id: 'cn_4', name: 'Trần Đình Quốc Dũng', grades: { "Toán": 3.8, "Ngữ Văn": 4.5, "Tiếng Anh": 4.0, "KHTN": 4.2, "KHXH": 5.5, "Tin Học": 6.0 }, gpa: null, rank: '', weakSubjects: [], teacherComment: 'Em còn yếu Toán, Văn, Anh. Cần phụ đạo gấp.' }
                ]
            };
            saveGvcnData();
            if (confirmNeeded) renderRoleView();
        }

        function loadDemoData() {
            if (currentRole === 'GVBM') loadDemoGvbmData(true);
            else loadDemoGvcnData(true);
        }

        function clearCurrentRoleData() {
            if (confirm(`Thầy/Cô muốn xóa toàn bộ dữ liệu ở vai trò ${currentRole === 'GVBM' ? 'Giáo viên Bộ môn' : 'Giáo viên Chủ nhiệm'}?`)) {
                if (currentRole === 'GVBM') { gvbmData.students = []; saveGvbmData(); }
                else { gvcnData.students = []; saveGvcnData(); }
                renderRoleView();
            }
        }

        function getBadgeClass(level) {
            switch(level) {
                case 'Tốt': return 'bg-emerald-100 text-emerald-800 dark:bg-emerald-950 dark:text-emerald-300';
                case 'Khá': return 'bg-sky-100 text-sky-800 dark:bg-sky-950 dark:text-sky-300';
                case 'Đạt': return 'bg-amber-100 text-amber-800 dark:bg-amber-950 dark:text-amber-300';
                case 'Chưa đạt': return 'bg-rose-100 text-rose-800 dark:bg-rose-950 dark:text-rose-300';
                default: return 'bg-slate-100 text-slate-600 dark:bg-slate-700 dark:text-slate-300';
            }
        }

        function escapeHtml(str) {
            if (!str) return '';
            return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
        }
    </script>
</body>
</html># 7-1

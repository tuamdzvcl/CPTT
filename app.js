// app.js
const excelFileInput = document.getElementById('excelFile');
const expenseFileInput = document.getElementById('expenseFile');
const idSuffixInput = document.getElementById('idSuffixInput');
const tableContainer = document.getElementById('tableContainer');
const emptyState = document.getElementById('emptyState');
const tableBody = document.getElementById('tableBody');
const globalThresholdInput = document.getElementById('globalThreshold');
const applyGlobalBtn = document.getElementById('applyGlobalBtn');
const sheetNameLabel = document.getElementById('sheetNameLabel');

const btnTuanAnh = document.getElementById('btnTuanAnh');
const btnKiet = document.getElementById('btnKiet');
const btnExport = document.getElementById('btnExport');
const calc10PercentBtn = document.getElementById('calc10PercentBtn');
const filterBtns = document.querySelectorAll('#statusFilters > div');

const dateRangeInput = document.getElementById('dateRange');
const summaryCards = document.getElementById('summaryCards');
const sumRevenueEl = document.getElementById('sumRevenue');
const sumOrdersEl = document.getElementById('sumOrders');
const sumSpendEl = document.getElementById('sumSpend');
const sumCpaEl = document.getElementById('sumCpa');
const sumCostPercentEl = document.getElementById('sumCostPercent');

// Target & KPI Elements
const targetMonthRevInput = document.getElementById('targetMonthRevenue');
const targetMonthSpendInput = document.getElementById('targetMonthSpend');
const targetDaysInMonthInput = document.getElementById('targetDaysInMonth');
const targetRevPreview = document.getElementById('targetRevPreview');
const targetSpendPreview = document.getElementById('targetSpendPreview');
const dailyTargetRevEl = document.getElementById('dailyTargetRev');
const dailyTargetSpendEl = document.getElementById('dailyTargetSpend');
const monthDetectedLabel = document.getElementById('monthDetectedLabel');
const targetCostRatioEl = document.getElementById('targetCostRatio');
const targetSelectedDaysText = document.getElementById('targetSelectedDaysText');
const compareRangeLabel = document.getElementById('compareRangeLabel');
const compareDaysCountBadge = document.getElementById('compareDaysCountBadge');
const periodTargetRevText = document.getElementById('periodTargetRevText');
const periodActualRevText = document.getElementById('periodActualRevText');
const revProgressPercent = document.getElementById('revProgressPercent');
const revProgressBar = document.getElementById('revProgressBar');
const revDiffAlert = document.getElementById('revDiffAlert');
const revDiffLabel = document.getElementById('revDiffLabel');
const revDiffValue = document.getElementById('revDiffValue');
const periodTargetSpendText = document.getElementById('periodTargetSpendText');
const periodActualSpendText = document.getElementById('periodActualSpendText');
const spendProgressPercent = document.getElementById('spendProgressPercent');
const spendProgressBar = document.getElementById('spendProgressBar');
const spendDiffAlert = document.getElementById('spendDiffAlert');
const spendDiffLabel = document.getElementById('spendDiffLabel');
const spendDiffValue = document.getElementById('spendDiffValue');
const targetActionBanner = document.getElementById('targetActionBanner');
const actualRevSubtext = document.getElementById('actualRevSubtext');
const actualSpendSubtext = document.getElementById('actualSpendSubtext');

let adsData = [];
let rawAdsData = [];
let rawSalesData = [];
let rawExpenseData = [];
let adsDataMap = { 'tuananh': [], 'kiet': [] };
let activeTab = 'tuananh';
let sheetNameMap = { 'tuananh': '', 'kiet': '' };
let currentThresholds = {};
let sortColumn = 'orders';
let sortDirection = 'desc';
let is10PercentMode = false;
let currentFilter = 'all';

// Target calculation state
let currentGrandTotalRevenue = 0;
let currentFinalSpendTotal = 0;
let currentAllSalesRevenue = 0;
let currentAllSalesOrders = 0;
let currentSelectedDays = 0;
let currentDateRangeText = 'Chưa chọn ngày';
let isCustomDaysInMonth = false;
let minAvailableDate = null;
let maxAvailableDate = null;

// Format currency VND
const formatCurrency = (value) => {
    if (isNaN(value) || value === null) return '0 đ';
    return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(value);
};

// Convert cell Date or Serial to local JS Date (midnight)
const extractDateFromCellDate = (val) => {
    if (!val) return null;
    if (typeof val === 'number') {
        // Excel serial to UTC date
        const utcMs = Math.round((val - 25569) * 86400 * 1000);
        const d = new Date(utcMs);
        return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
    }
    if (typeof val === 'string') {
        const str = val.trim();
        // Hỗ trợ DD/MM/YYYY hoặc DD-MM-YYYY phổ biến ở VN
        const dmyMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
        if (dmyMatch) {
            const day = parseInt(dmyMatch[1], 10);
            const month = parseInt(dmyMatch[2], 10) - 1;
            const year = parseInt(dmyMatch[3], 10);
            return new Date(year, month, day);
        }
        const parsed = new Date(str);
        if (!isNaN(parsed.getTime())) {
            return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
        }
    }
    if (val instanceof Date && !isNaN(val.getTime())) {
        return new Date(val.getUTCFullYear(), val.getUTCMonth(), val.getUTCDate());
    }
    return null;
};

// Initialize Flatpickr
let datePicker = flatpickr("#dateRange", {
    mode: "range",
    dateFormat: "Y-m-d",
    locale: "vn",
    theme: "dark",
    onChange: function (selectedDates) {
        if (selectedDates.length === 2 || selectedDates.length === 0) {
            processData();
        }
    },
    onClose: function (selectedDates) {
        if (selectedDates.length === 1) {
            processData();
        }
    }
});

// Format hiển thị mục tiêu dễ nhìn (kèm chữ triệu / tỷ)
const formatTargetDisplay = (val) => {
    if (!val || val <= 0) return '0 đ';
    if (val >= 1000000000) {
        const ty = (val / 1000000000).toFixed(val % 1000000000 === 0 ? 0 : 2);
        return `${formatCurrency(val)} (${ty} tỷ)`;
    } else if (val >= 1000000) {
        const tr = (val / 1000000).toFixed(val % 1000000 === 0 ? 0 : 1);
        return `${formatCurrency(val)} (${tr} triệu)`;
    }
    return formatCurrency(val);
};

// Parse numeric value (handles currency symbol ₫/đ/$, spaces, VN dots, commas)
const parseNumeric = (val) => {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (typeof val === 'string') {
        let str = val.trim().replace(/[₫đĐ$\s\u00a0]|vnd/gi, '');
        if (!str) return 0;

        // If it has both dot and comma
        if (str.includes('.') && str.includes(',')) {
            if (str.lastIndexOf(',') > str.lastIndexOf('.')) {
                // VN format: 1.000.000,50
                str = str.replace(/\./g, '').replace(',', '.');
            } else {
                // US format: 1,000,000.50
                str = str.replace(/,/g, '');
            }
        } else if ((str.match(/\./g) || []).length > 1) {
            // Multiple dots: e.g. "500.000.000"
            str = str.replace(/\./g, '');
        } else if ((str.match(/,/g) || []).length > 1) {
            // Multiple commas: e.g. "500,000,000"
            str = str.replace(/,/g, '');
        } else if (str.includes('.')) {
            const parts = str.split('.');
            // Single dot followed by 3 digits -> thousands separator (VN) e.g. 500.000
            if (parts[1] && parts[1].length === 3) {
                str = str.replace(/\./g, '');
            }
        } else if (str.includes(',')) {
            const parts = str.split(',');
            // Single comma followed by 3 digits -> thousands separator e.g. 500,000
            if (parts[1] && parts[1].length === 3) {
                str = str.replace(/,/g, '');
            } else {
                str = str.replace(',', '.');
            }
        }

        const parsed = parseFloat(str.replace(/[^\d.-]/g, ''));
        return isNaN(parsed) ? 0 : parsed;
    }
    return 0;
};

// Helper: lấy giá trị trong row theo danh sách tên cột ứng viên (bỏ qua hoa/thường, khoảng trắng)
const getRowVal = (row, candidates) => {
    if (!row) return null;
    for (const key of Object.keys(row)) {
        const cleanKey = key.trim().toLowerCase();
        for (const c of candidates) {
            if (cleanKey === c.toLowerCase()) {
                if (row[key] !== null && row[key] !== undefined && row[key] !== '') {
                    return row[key];
                }
            }
        }
    }
    return null;
};

// Nhận diện danh sách nhân viên Sale phụ trách của sheet Ads (dựa trên các cột "Sale - ..." hoặc "Sale ...")
const getAssignedSaleStaff = (adsRows) => {
    if (!adsRows || adsRows.length === 0) return [];
    const sample = adsRows[0];
    const staffList = [];
    Object.keys(sample).forEach(key => {
        const clean = key.trim();
        // Không nhận các cột tổng tiền, chi phí, cp
        if (/tổng|tiền|cp\/|chi phí/i.test(clean)) return;
        if (/^sale\s*[-_:]?\s*/i.test(clean)) {
            const name = clean.replace(/^sale\s*[-_:]?\s*/i, '').trim();
            if (name && !staffList.includes(name)) staffList.push(name);
        }
    });
    return staffList;
};

// Kiểm tra xem nhân viên bán hàng có thuộc danh sách phụ trách của tab không
const isStaffMatched = (saleName, staffList) => {
    if (!staffList || staffList.length === 0) return true;
    const lower = (saleName || '').toLowerCase().trim();
    return staffList.some(s => {
        const staffLower = s.toLowerCase().trim();
        // Đặc thù: Thảo lớn là Thu Thảo, không phải Phương Thảo
        if (staffLower.includes('thảo lớn')) {
            return lower.includes('thu thảo') || lower.includes('thảo lớn') || (lower.includes('thảo') && !lower.includes('phương'));
        }
        if (staffLower.includes('phương thảo')) {
            return lower.includes('phương thảo');
        }
        return lower.includes(staffLower);
    });
};

// Parse target inputs: hỗ trợ số đầy đủ, viết tắt 500tr, 500k, hoặc gõ nhanh 500 (= 500 triệu)
const parseTargetInput = (inputEl) => {
    if (!inputEl) return 0;
    let str = (typeof inputEl === 'string' ? inputEl : (inputEl.value || '')).toString().trim().toLowerCase();
    if (!str) return 0;

    // Hỗ trợ "tỷ", "ty", "b"
    if (str.includes('tỷ') || str.includes('ty') || str.includes('b')) {
        const num = parseFloat(str.replace(/[^\d.-]/g, ''));
        return isNaN(num) ? 0 : Math.round(num * 1000000000);
    }
    // Hỗ trợ "tr", "triệu", "trieu", "m"
    if (str.includes('tr') || str.includes('m') || str.includes('triệu') || str.includes('trieu')) {
        const num = parseFloat(str.replace(/[^\d.-]/g, ''));
        return isNaN(num) ? 0 : Math.round(num * 1000000);
    }
    // Hỗ trợ "k", "nghìn", "ngàn"
    if (str.includes('k') || str.includes('nghìn') || str.includes('ngàn')) {
        const num = parseFloat(str.replace(/[^\d.-]/g, ''));
        return isNaN(num) ? 0 : Math.round(num * 1000);
    }

    const cleanNum = parseNumeric(str);
    if (cleanNum > 0) {
        // Nếu nhập số nhỏ dưới 10.000 (ví dụ 500 hoặc 50): tự động hiểu là TRIỆU đồng
        if (cleanNum < 10000) {
            return Math.round(cleanNum * 1000000);
        }
        return Math.round(cleanNum);
    }
    return 0;
};

// Handle file upload
excelFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });

        adsDataMap = { 'tuananh': [], 'kiet': [] };
        sheetNameMap = { 'tuananh': '', 'kiet': '' };
        rawSalesData = [];
        minAvailableDate = null;
        maxAvailableDate = null;

        const candidateSaleSheets = [];

        workbook.SheetNames.forEach(sheetName => {
            const lowerName = sheetName.toLowerCase();
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            // Nhận diện sheet Ads trước
            const isAdsSheet = lowerName.includes('t.anh') || lowerName.includes('tuấn anh') || lowerName.includes('kiệt') || lowerName.includes('kiet');

            // Sheet hoàn/hủy/loại trừ (tránh nhận nhầm tên nhân viên như Thu Huyền, Huyền Trang)
            const isExcludedSheet = /(đơn\s*hoàn|đơn\s*hủy|hoàn\s*hàng|trả\s*hàng|refund|return)/i.test(lowerName) ||
                (lowerName.includes('hoàn') && !lowerName.includes('sale')) ||
                (lowerName.includes('hủy') && !lowerName.includes('sale'));

            // Detect SALE / Doanh Thu / Bán Hàng sheets
            const isSaleSheet = !isAdsSheet && !isExcludedSheet && (
                lowerName.includes('sale') || 
                lowerName.includes('doanh') || 
                lowerName.includes('bán') || 
                lowerName.includes('đơn') || 
                lowerName.includes('order') || 
                lowerName.includes('revenue')
            );

            if (isSaleSheet) {
                candidateSaleSheets.push({ sheetName, data });
            }

            // Detect ADS sheet Tuấn Anh
            if (lowerName.includes('t.anh') || lowerName.includes('tuấn anh')) {
                sheetNameMap['tuananh'] = sheetName;
                adsDataMap['tuananh'] = data;
            }
            
            // Detect ADS sheet Kiệt
            if (lowerName.includes('kiệt') || lowerName.includes('kiet')) {
                sheetNameMap['kiet'] = sheetName;
                adsDataMap['kiet'] = data;
            }
        });

        // Nếu không tìm thấy sheet tên "sale/doanh/bán...", thử tìm sheet có cột tiền
        if (candidateSaleSheets.length === 0) {
            workbook.SheetNames.forEach(sheetName => {
                const lowerName = sheetName.toLowerCase();
                const isAdsSheet = lowerName.includes('t.anh') || lowerName.includes('tuấn anh') || lowerName.includes('kiệt') || lowerName.includes('kiet');
                const isExcludedSheet = /(đơn\s*hoàn|đơn\s*hủy|hoàn\s*hàng|trả\s*hàng|refund|return)/i.test(lowerName) ||
                    (lowerName.includes('hoàn') && !lowerName.includes('sale')) ||
                    (lowerName.includes('hủy') && !lowerName.includes('sale'));
                if (!isAdsSheet && !isExcludedSheet) {
                    const worksheet = workbook.Sheets[sheetName];
                    const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });
                    if (data.length > 0) {
                        const sampleRow = data[0];
                        const hasMoneyCol = getRowVal(sampleRow, ['Thành Tiền (Trước VAT)', 'Thành tiền', 'Doanh thu', 'Doanh số', 'Tổng tiền', 'Đơn Giá (gồm VAT)']);
                        if (hasMoneyCol !== null) {
                            candidateSaleSheets.push({ sheetName, data });
                        }
                    }
                }
            });
        }

        // Đọc dữ liệu bán hàng từ các sheet Sale
        candidateSaleSheets.forEach(({ sheetName, data }) => {
            const defaultSaleName = sheetName.replace(/^(sale|nv|bán\s*hàng)\s*[-_:]\s*/i, '').trim();

            data.forEach(row => {
                const adId = getRowVal(row, ['ID Quảng Cáo', 'ID Ads', 'ID QC', 'Mã quảng cáo', 'Mã QC', 'Mã Ads', 'ID quảng cáo', 'ID']);
                const dateVal = getRowVal(row, ['Ngày đặt', 'Ngày Đặt', 'Ngày', 'Ngày lên đơn', 'Ngày bán', 'Thời gian', 'Ngày tạo', 'Date']);
                const amountVal = getRowVal(row, [
                    'Thành Tiền (Trước VAT)', 'Thành tiền (trước VAT)', 'Thành tiền trước VAT', 'Thành Tiền (TRƯỚC VAT)', 
                    'Thành tiền', 'Thành Tiền', 'Doanh thu', 'Doanh Thu', 'Doanh số', 'Doanh Số', 
                    'Tổng tiền', 'Tổng Tiền', 'Tiền', 'Giá trị', 'Đơn Giá (gồm VAT)'
                ]);
                const amount = parseNumeric(amountVal);
                const qty = parseNumeric(getRowVal(row, ['Số Lượng', 'Số lượng', 'SL', 'Quantity']) || 1);
                const slDonRaw = getRowVal(row, ['SL đơn tháng này', 'SL đơn', 'Số đơn']);
                const slDon = (slDonRaw !== null && slDonRaw !== undefined && slDonRaw !== '') ? parseInt(slDonRaw, 10) : null;
                const khCode = (getRowVal(row, ['SĐT KH', 'SDT KH', 'KH Số', 'Mã KH']) || '').toString().trim();

                const jsDate = extractDateFromCellDate(dateVal);
                if (jsDate) {
                    if (!minAvailableDate || jsDate < minAvailableDate) minAvailableDate = jsDate;
                    if (!maxAvailableDate || jsDate > maxAvailableDate) maxAvailableDate = jsDate;
                }

                const safeAdId = (adId && adId.toString().trim() !== '') ? adId.toString().trim() : null;
                const saleName = getRowVal(row, ['Sale', 'Tên Sale', 'Người bán', 'Nhân viên', 'NVKD', 'Tên NV', 'Nhân viên sale']) || defaultSaleName || 'Không rõ';

                rawSalesData.push({
                    adId: safeAdId,
                    saleName: saleName.toString().trim(),
                    date: jsDate,
                    revenue: amount,
                    qty: qty > 0 ? qty : 1,
                    slDon: slDon,
                    khCode: khCode
                });
            });
        });

        // Set default rawAdsData to current active tab
        rawAdsData = adsDataMap[activeTab];
        sheetNameLabel.textContent = sheetNameMap[activeTab] ? `(Dữ liệu từ ${sheetNameMap[activeTab]})` : '(Không tìm thấy sheet)';

        // Đặt khoảng ngày mặc định vào THÁNG MỚI NHẤT của dữ liệu (tránh gộp các tháng cũ trong file)
        if (maxAvailableDate) {
            const defaultStart = new Date(maxAvailableDate.getFullYear(), maxAvailableDate.getMonth(), 1);
            const fmtStart = `${defaultStart.getFullYear()}-${(defaultStart.getMonth() + 1).toString().padStart(2, '0')}-01`;
            const fmtEnd = `${maxAvailableDate.getFullYear()}-${(maxAvailableDate.getMonth() + 1).toString().padStart(2, '0')}-${maxAvailableDate.getDate().toString().padStart(2, '0')}`;
            datePicker.setDate([fmtStart, fmtEnd], true);
        } else if (minAvailableDate) {
            const fmtMin = `${minAvailableDate.getFullYear()}-${(minAvailableDate.getMonth() + 1).toString().padStart(2, '0')}-${minAvailableDate.getDate().toString().padStart(2, '0')}`;
            datePicker.setDate([fmtMin, fmtMin], true);
        }

        processData();
    };
    reader.readAsArrayBuffer(file);
});

// Handle expense file upload
expenseFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        rawExpenseData = XLSX.utils.sheet_to_json(worksheet, { defval: null });
        
        processData();
    };
    reader.readAsArrayBuffer(file);
});

idSuffixInput.addEventListener('input', () => {
    if (rawExpenseData.length > 0) {
        processData();
    }
});

// Tab switching
const updateTabUI = () => {
    if (activeTab === 'tuananh') {
        btnTuanAnh.className = 'flex-1 bg-brand-600 hover:bg-brand-500 text-white font-semibold py-3 px-6 rounded-2xl transition-all shadow-lg shadow-brand-500/30 border border-brand-400';
        btnKiet.className = 'flex-1 bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold py-3 px-6 rounded-2xl transition-all border border-slate-700';
    } else {
        btnKiet.className = 'flex-1 bg-brand-600 hover:bg-brand-500 text-white font-semibold py-3 px-6 rounded-2xl transition-all shadow-lg shadow-brand-500/30 border border-brand-400';
        btnTuanAnh.className = 'flex-1 bg-slate-800 hover:bg-slate-700 text-slate-300 font-semibold py-3 px-6 rounded-2xl transition-all border border-slate-700';
    }
};

btnTuanAnh.addEventListener('click', () => {
    if (activeTab === 'tuananh') return;
    activeTab = 'tuananh';
    updateTabUI();
    rawAdsData = adsDataMap['tuananh'] || [];
    sheetNameLabel.textContent = sheetNameMap['tuananh'] ? `(Dữ liệu từ ${sheetNameMap['tuananh']})` : '(Không tìm thấy sheet)';
    processData();
});

btnKiet.addEventListener('click', () => {
    if (activeTab === 'kiet') return;
    activeTab = 'kiet';
    updateTabUI();
    rawAdsData = adsDataMap['kiet'] || [];
    sheetNameLabel.textContent = sheetNameMap['kiet'] ? `(Dữ liệu từ ${sheetNameMap['kiet']})` : '(Không tìm thấy sheet)';
    processData();
});

// Export Excel
btnExport.addEventListener('click', () => {
    if (adsData.length === 0) {
        alert('Không có dữ liệu để xuất!');
        return;
    }

    let fromFilterStr = 'Tất cả';
    let toFilterStr = 'Tất cả';

    if (datePicker.selectedDates.length > 0) {
        fromFilterStr = datePicker.selectedDates[0].toLocaleDateString('vi-VN');
        if (datePicker.selectedDates.length === 2) {
            toFilterStr = datePicker.selectedDates[1].toLocaleDateString('vi-VN');
        } else {
            toFilterStr = fromFilterStr;
        }
    }

    const exportData = adsData.map(ad => ({
        'Từ Ngày': fromFilterStr,
        'Đến Ngày': toFilterStr,
        'Tên Chiến Dịch': ad.campaign,
        'Tên Quảng Cáo': ad.adName,
        'ID Quảng Cáo': ad.id,
        'Tổng Đơn': ad.orders,
        'Doanh Thu (VND)': ad.revenue,
        'Chi Phí Báo Cáo (VND)': ad.spend,
        'Chi Phí Thực Tế (VND)': ad.realSpend,
        'Chi Phí / Đơn (VND)': ad.cpa
    }));

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    const sheetName = activeTab === 'tuananh' ? 'BaoCao_TuanAnh' : 'BaoCao_Kiet';
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    // Xuất thêm sheet Mục Tiêu & Kế Hoạch nếu người dùng có nhập
    const targetRev = parseTargetInput(targetMonthRevInput);
    const targetSpend = parseTargetInput(targetMonthSpendInput);
    if (targetRev > 0 || targetSpend > 0) {
        const daysInMonth = parseInt(targetDaysInMonthInput ? targetDaysInMonthInput.value : 30, 10) || 30;
        const dailyRev = targetRev / daysInMonth;
        const dailySpend = targetSpend / daysInMonth;
        const pTargetRev = currentSelectedDays * dailyRev;
        const pTargetSpend = currentSelectedDays * dailySpend;
        const diffRev = currentGrandTotalRevenue - pTargetRev;
        const diffSpend = currentFinalSpendTotal - pTargetSpend;
        const isOver = diffSpend > 0 && pTargetSpend > 0;

        const targetExportData = [
            { 'Hạng Mục': 'Mục tiêu Doanh Thu Tháng', 'Số Liệu': targetRev, 'Đơn Vị': 'VND' },
            { 'Hạng Mục': 'Mục tiêu Chi Phí Ads Tháng', 'Số Liệu': targetSpend, 'Đơn Vị': 'VND' },
            { 'Hạng Mục': 'Số ngày trong tháng', 'Số Liệu': daysInMonth, 'Đơn Vị': 'Ngày' },
            { 'Hạng Mục': 'Mục tiêu Doanh Thu / Ngày', 'Số Liệu': Math.round(dailyRev), 'Đơn Vị': 'VND/ngày' },
            { 'Hạng Mục': 'Mục tiêu Chi Phí Ads / Ngày', 'Số Liệu': Math.round(dailySpend), 'Đơn Vị': 'VND/ngày' },
            { 'Hạng Mục': 'Khoảng thời gian chọn', 'Số Liệu': currentDateRangeText, 'Đơn Vị': '' },
            { 'Hạng Mục': 'Số ngày được chọn', 'Số Liệu': currentSelectedDays, 'Đơn Vị': 'Ngày' },
            { 'Hạng Mục': `Mục tiêu Doanh Thu (${currentSelectedDays} ngày)`, 'Số Liệu': Math.round(pTargetRev), 'Đơn Vị': 'VND' },
            { 'Hạng Mục': 'Doanh Thu thực tế thu được', 'Số Liệu': currentGrandTotalRevenue, 'Đơn Vị': 'VND' },
            { 'Hạng Mục': 'Chênh lệch Doanh Thu', 'Số Liệu': diffRev, 'Đơn Vị': diffRev < 0 ? 'Đang thiếu' : 'Đạt / Vượt' },
            { 'Hạng Mục': `Mục tiêu Chi Phí Ads (${currentSelectedDays} ngày)`, 'Số Liệu': Math.round(pTargetSpend), 'Đơn Vị': 'VND' },
            { 'Hạng Mục': 'Chi Phí Ads thực tế đã chi', 'Số Liệu': currentFinalSpendTotal, 'Đơn Vị': 'VND' },
            { 'Hạng Mục': 'Chênh lệch Chi Phí', 'Số Liệu': diffSpend, 'Đơn Vị': isOver ? 'VƯỢT NGÂN SÁCH (CẦN CẮT GIẢM)' : 'Trong định mức / Tiết kiệm' }
        ];
        const wsTarget = XLSX.utils.json_to_sheet(targetExportData);
        XLSX.utils.book_append_sheet(wb, wsTarget, 'KeHoach_MucTieu');
    }
    
    // Generate filename
    const now = new Date();
    const dateStr = `${now.getFullYear()}${(now.getMonth()+1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
    XLSX.writeFile(wb, `BaoCao_DoanhSo_${activeTab}_${dateStr}.xlsx`);
});

// Target Management Logic
const calculateTargets = () => {
    const targetRev = parseTargetInput(targetMonthRevInput);
    const targetSpend = parseTargetInput(targetMonthSpendInput);
    let daysInMonth = parseInt(targetDaysInMonthInput ? targetDaysInMonthInput.value : 30, 10);
    if (isNaN(daysInMonth) || daysInMonth <= 0) daysInMonth = 30;

    // Save to localStorage
    if (targetRev > 0) localStorage.setItem('cptt_target_month_rev', targetRev);
    if (targetSpend > 0) localStorage.setItem('cptt_target_month_spend', targetSpend);
    localStorage.setItem('cptt_target_days_month', daysInMonth);

    // 1. Calculate Daily Breakdown
    const dailyRev = daysInMonth > 0 ? (targetRev / daysInMonth) : 0;
    const dailySpend = daysInMonth > 0 ? (targetSpend / daysInMonth) : 0;

    if (dailyTargetRevEl) dailyTargetRevEl.textContent = formatCurrency(Math.round(dailyRev)) + '/ngày';
    if (dailyTargetSpendEl) dailyTargetSpendEl.textContent = formatCurrency(Math.round(dailySpend)) + '/ngày';

    const targetRatio = targetRev > 0 ? (targetSpend / targetRev) * 100 : 0;
    if (targetCostRatioEl) targetCostRatioEl.textContent = targetRatio.toFixed(1) + '%';

    // 2. Selected Period Calculation
    const days = currentSelectedDays > 0 ? currentSelectedDays : 0;
    if (targetSelectedDaysText) targetSelectedDaysText.textContent = `${days} ngày`;
    if (compareDaysCountBadge) compareDaysCountBadge.textContent = `${days} ngày`;
    if (compareRangeLabel) compareRangeLabel.textContent = currentDateRangeText;

    document.querySelectorAll('.revDaysLabel').forEach(el => el.textContent = days);
    document.querySelectorAll('.spendDaysLabel').forEach(el => el.textContent = days);

    const periodTargetRev = days * dailyRev;
    const periodTargetSpend = days * dailySpend;

    if (periodTargetRevText) periodTargetRevText.textContent = formatCurrency(Math.round(periodTargetRev));
    if (periodTargetSpendText) periodTargetSpendText.textContent = formatCurrency(Math.round(periodTargetSpend));

    const actualRev = currentGrandTotalRevenue;
    const actualSpend = currentFinalSpendTotal;

    if (periodActualRevText) periodActualRevText.textContent = formatCurrency(actualRev);
    if (periodActualSpendText) periodActualSpendText.textContent = formatCurrency(actualSpend);

    if (actualRevSubtext) {
        if (currentAllSalesOrders > 0) {
            actualRevSubtext.textContent = `(Từ các mã ID Ads: ${currentAllSalesOrders} đơn)`;
        } else {
            actualRevSubtext.textContent = `(0 đơn có ID Ads)`;
        }
    }

    if (actualSpendSubtext) {
        if (rawExpenseData.length > 0) {
            actualSpendSubtext.textContent = `Nguồn: Chi tiêu thực tế (Đã đổi đuôi ID)`;
        } else {
            actualSpendSubtext.textContent = `Nguồn: Báo cáo Ads trong file`;
        }
    }

    // 3. Revenue Comparison
    let revPct = periodTargetRev > 0 ? (actualRev / periodTargetRev) * 100 : 0;
    if (revProgressPercent) revProgressPercent.textContent = revPct.toFixed(1) + '%';
    if (revProgressBar) revProgressBar.style.width = Math.min(revPct, 100) + '%';

    const diffRev = actualRev - periodTargetRev;
    if (periodTargetRev === 0) {
        if (revDiffAlert) {
            revDiffAlert.className = 'p-3 rounded-lg text-xs flex items-center justify-between font-medium bg-slate-900/60 border border-slate-700';
            revDiffLabel.textContent = 'Chưa đặt mục tiêu doanh thu';
            revDiffValue.textContent = '-';
            revDiffValue.className = 'font-bold text-slate-400';
        }
    } else if (diffRev < 0) {
        if (revDiffAlert) {
            revDiffAlert.className = 'p-3 rounded-lg text-xs flex items-center justify-between font-medium bg-amber-950/40 border border-amber-500/50';
            revDiffLabel.innerHTML = '<span class="flex items-center gap-1.5 text-amber-400 font-semibold"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg> Đang thiếu doanh số:</span>';
            revDiffValue.innerHTML = `<span class="text-amber-300 font-extrabold text-sm">-${formatCurrency(Math.abs(diffRev))}</span> <span class="text-[11px] text-amber-400/80">(Đạt ${revPct.toFixed(1)}%)</span>`;
        }
    } else {
        if (revDiffAlert) {
            revDiffAlert.className = 'p-3 rounded-lg text-xs flex items-center justify-between font-medium bg-emerald-950/40 border border-emerald-500/50';
            revDiffLabel.innerHTML = '<span class="flex items-center gap-1.5 text-emerald-400 font-semibold"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg> Vượt mục tiêu doanh thu:</span>';
            revDiffValue.innerHTML = `<span class="text-emerald-300 font-extrabold text-sm">+${formatCurrency(diffRev)}</span> <span class="text-[11px] text-emerald-400/80">(Đạt ${revPct.toFixed(1)}%)</span>`;
        }
    }

    // 4. Spend Comparison & Warning
    let spendPct = periodTargetSpend > 0 ? (actualSpend / periodTargetSpend) * 100 : 0;
    if (spendProgressPercent) spendProgressPercent.textContent = spendPct.toFixed(1) + '%';
    if (spendProgressBar) spendProgressBar.style.width = Math.min(spendPct, 100) + '%';

    const diffSpend = actualSpend - periodTargetSpend;
    const isOverBudget = diffSpend > 0 && periodTargetSpend > 0;

    if (periodTargetSpend === 0) {
        if (spendDiffAlert) {
            spendDiffAlert.className = 'p-3 rounded-lg text-xs flex items-center justify-between font-medium bg-slate-900/60 border border-slate-700';
            spendDiffLabel.textContent = 'Chưa đặt mục tiêu chi phí';
            spendDiffValue.textContent = '-';
            spendDiffValue.className = 'font-bold text-slate-400';
        }
    } else if (isOverBudget) {
        if (spendDiffAlert) {
            spendDiffAlert.className = 'p-3 rounded-lg text-xs flex items-center justify-between font-medium bg-rose-950/50 border border-rose-500/60 shadow-lg shadow-rose-900/30 animate-pulse';
            spendDiffLabel.innerHTML = '<span class="flex items-center gap-1.5 text-rose-400 font-bold"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"></path></svg> CHI PHÍ VƯỢT QUÁ:</span>';
            spendDiffValue.innerHTML = `<span class="text-rose-300 font-black text-sm">+${formatCurrency(diffSpend)}</span> <span class="text-[11px] text-rose-400/80">(Vượt ${(spendPct - 100).toFixed(1)}%)</span>`;
        }
    } else {
        if (spendDiffAlert) {
            spendDiffAlert.className = 'p-3 rounded-lg text-xs flex items-center justify-between font-medium bg-emerald-950/40 border border-emerald-500/50';
            spendDiffLabel.innerHTML = '<span class="flex items-center gap-1.5 text-emerald-400 font-semibold"><svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg> Tiết kiệm / Dư ngân sách:</span>';
            spendDiffValue.innerHTML = `<span class="text-emerald-300 font-bold text-sm">-${formatCurrency(Math.abs(diffSpend))}</span> <span class="text-[11px] text-emerald-400/80">(Chi ${spendPct.toFixed(1)}%)</span>`;
        }
    }

    // 5. Actionable Recommendation Banner
    renderActionBanner(periodTargetRev, actualRev, diffRev, periodTargetSpend, actualSpend, diffSpend, isOverBudget, days, daysInMonth);
};

const renderActionBanner = (pTargetRev, actualRev, diffRev, pTargetSpend, actualSpend, diffSpend, isOverBudget, selDays, daysInMonth) => {
    if (!targetActionBanner) return;

    if (actualRev === 0 && actualSpend === 0 && pTargetRev === 0 && pTargetSpend === 0) {
        targetActionBanner.classList.add('hidden');
        return;
    }

    targetActionBanner.classList.remove('hidden');

    const dangerCount = adsData.filter(ad => {
        const threshold = ad.threshold !== null ? ad.threshold : (parseFloat(globalThresholdInput.value) || 0);
        const defaultZero = 350000;
        const spend = ad.spendToEvaluate;
        if (ad.orders === 0) {
            const t = threshold > 0 ? threshold : defaultZero;
            return spend > t;
        }
        return threshold > 0 && ad.cpa > threshold;
    }).length;

    const remainingDays = Math.max(0, daysInMonth - selDays);

    if (isOverBudget) {
        targetActionBanner.className = 'mt-4 p-4 rounded-xl border border-rose-500/50 bg-gradient-to-r from-rose-950/60 via-slate-900/80 to-slate-900 transition-all';
        targetActionBanner.innerHTML = `
            <div class="flex flex-col md:flex-row items-start md:items-center justify-between gap-4">
                <div class="flex items-start gap-3">
                    <span class="text-3xl mt-0.5">⚠️</span>
                    <div>
                        <h4 class="text-sm font-bold text-rose-300 flex items-center gap-2">
                            CẢNH BÁO CẮT GIẢM: CHI PHÍ QUẢNG CÁO VƯỢT QUÁ NGÂN SÁCH!
                        </h4>
                        <p class="text-xs text-slate-300 mt-1 leading-relaxed">
                            Trong <strong class="text-white">${selDays} ngày</strong> qua, bạn đã chi <strong class="text-rose-400">${formatCurrency(actualSpend)}</strong>, vượt ngân sách định mức <strong class="text-rose-400">${formatCurrency(diffSpend)}</strong>. 
                            ${diffRev < 0 ? `Đồng thời doanh thu đang thiếu <strong class="text-amber-300">${formatCurrency(Math.abs(diffRev))}</strong>.` : ''}
                        </p>
                        <div class="mt-2 text-xs text-rose-200 font-medium flex flex-wrap items-center gap-2">
                            <span class="bg-rose-500/20 px-2.5 py-1 rounded-md border border-rose-500/40 text-rose-300">
                                ✂️ Đề xuất cắt giảm: <strong>${formatCurrency(diffSpend)}</strong>
                            </span>
                            ${remainingDays > 0 ? `
                                <span class="text-slate-400">
                                    (Còn ${remainingDays} ngày trong tháng, ngân sách còn lại nên siết chặt trung bình <strong>${formatCurrency(Math.max(0, (parseNumeric(targetMonthSpendInput.value) - actualSpend) / remainingDays))}/ngày</strong>)
                                </span>
                            ` : ''}
                        </div>
                    </div>
                </div>
                ${dangerCount > 0 ? `
                    <button id="quickFilterDangerBtn" class="flex-none bg-rose-600 hover:bg-rose-500 text-white font-semibold py-2.5 px-4 rounded-xl text-xs transition-all shadow-lg shadow-rose-500/30 whitespace-nowrap flex items-center gap-1.5 cursor-pointer">
                        <span>🔍 Lọc ${dangerCount} QC Vượt Ngưỡng</span>
                    </button>
                ` : ''}
            </div>
        `;

        const filterBtn = document.getElementById('quickFilterDangerBtn');
        if (filterBtn) {
            filterBtn.addEventListener('click', () => {
                const dangerFilterEl = document.getElementById('filter-danger');
                if (dangerFilterEl) dangerFilterEl.click();
                tableContainer.scrollIntoView({ behavior: 'smooth' });
            });
        }
    } else if (diffRev < 0 && pTargetRev > 0) {
        targetActionBanner.className = 'mt-4 p-4 rounded-xl border border-amber-500/50 bg-gradient-to-r from-amber-950/50 via-slate-900/80 to-slate-900 transition-all';
        targetActionBanner.innerHTML = `
            <div class="flex items-start gap-3">
                <span class="text-3xl mt-0.5">📉</span>
                <div>
                    <h4 class="text-sm font-bold text-amber-300">
                        CẦN TĂNG TỐC DOANH THU ĐỂ KỊP TIẾN ĐỘ THÁNG
                    </h4>
                    <p class="text-xs text-slate-300 mt-1 leading-relaxed">
                        Doanh thu đạt <strong class="text-emerald-400">${formatCurrency(actualRev)}</strong> trong ${selDays} ngày, đang hụt <strong class="text-amber-300">${formatCurrency(Math.abs(diffRev))}</strong> so với kế hoạch (${((actualRev/pTargetRev)*100).toFixed(1)}%). 
                        Chi phí Ads vẫn trong tầm kiểm soát (tiết kiệm <strong class="text-emerald-400">${formatCurrency(Math.abs(diffSpend))}</strong>).
                    </p>
                    ${remainingDays > 0 ? `
                        <div class="mt-2 text-xs text-amber-200 font-medium">
                            🎯 Mục tiêu còn lại trong ${remainingDays} ngày cuối tháng: cần đạt trung bình <strong>${formatCurrency(Math.max(0, (parseNumeric(targetMonthRevInput.value) - actualRev) / remainingDays))}/ngày</strong>.
                        </div>
                    ` : ''}
                </div>
            </div>
        `;
    } else if (pTargetRev > 0 && pTargetSpend > 0) {
        targetActionBanner.className = 'mt-4 p-4 rounded-xl border border-emerald-500/50 bg-gradient-to-r from-emerald-950/50 via-slate-900/80 to-slate-900 transition-all';
        targetActionBanner.innerHTML = `
            <div class="flex items-start gap-3">
                <span class="text-3xl mt-0.5">🎉</span>
                <div>
                    <h4 class="text-sm font-bold text-emerald-300">
                        TIẾN ĐỘ RẤT TỐT: DOANH SỐ ĐẠT MỤC TIÊU & CHI PHÍ TRONG TẦM KIỂM SOÁT
                    </h4>
                    <p class="text-xs text-slate-300 mt-1 leading-relaxed">
                        Doanh thu thực tế vượt mục tiêu <strong class="text-emerald-400">+${formatCurrency(diffRev)}</strong> và chi phí Ads tiết kiệm <strong class="text-emerald-400">${formatCurrency(Math.abs(diffSpend))}</strong> so với định mức ngân sách. Hãy tiếp tục duy trì và nhân bản các chiến dịch hiệu quả!
                    </p>
                </div>
            </div>
        `;
    } else {
        targetActionBanner.classList.add('hidden');
    }
};

// Process data 
const processData = () => {
    adsData = [];

    // Parse filters from Flatpickr
    let fromFilter = null;
    let toFilter = null;
    currentSelectedDays = 0;
    currentDateRangeText = 'Tất cả thời gian';

    if (datePicker.selectedDates.length > 0) {
        fromFilter = new Date(datePicker.selectedDates[0]);
        fromFilter.setHours(0, 0, 0, 0);

        // Auto update days in month if user hasn't explicitly customized
        const filterYear = fromFilter.getFullYear();
        const filterMonth = fromFilter.getMonth() + 1;
        const daysInMonthDetected = new Date(filterYear, filterMonth, 0).getDate();
        if (!isCustomDaysInMonth && targetDaysInMonthInput) {
            targetDaysInMonthInput.value = daysInMonthDetected;
            if (monthDetectedLabel) monthDetectedLabel.textContent = `(Tháng ${filterMonth}: ${daysInMonthDetected} ngày)`;
        }

        if (datePicker.selectedDates.length === 2) {
            toFilter = new Date(datePicker.selectedDates[1]);
            toFilter.setHours(23, 59, 59, 999);
            const toMidnight = new Date(datePicker.selectedDates[1]);
            toMidnight.setHours(0, 0, 0, 0);
            const diffMs = Math.abs(toMidnight.getTime() - fromFilter.getTime());
            currentSelectedDays = Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1;
            currentDateRangeText = `${fromFilter.toLocaleDateString('vi-VN')} - ${toMidnight.toLocaleDateString('vi-VN')}`;
        } else {
            toFilter = new Date(datePicker.selectedDates[0]);
            toFilter.setHours(23, 59, 59, 999);
            currentSelectedDays = 1;
            currentDateRangeText = fromFilter.toLocaleDateString('vi-VN');
        }
    } else if (minAvailableDate && maxAvailableDate) {
        const diffMs = Math.abs(maxAvailableDate.getTime() - minAvailableDate.getTime());
        currentSelectedDays = Math.round(diffMs / (1000 * 60 * 60 * 24)) + 1;
        currentDateRangeText = `${minAvailableDate.toLocaleDateString('vi-VN')} - ${maxAvailableDate.toLocaleDateString('vi-VN')}`;
    } else {
        currentSelectedDays = 0;
        currentDateRangeText = 'Chưa chọn khoảng ngày';
    }

    // 1. Aggregate Sales by Ad ID & Total Sales
    const salesAgg = {};
    let totalSalesRevenueAll = 0;
    let totalSalesOrdersAll = 0;
    const staffList = getAssignedSaleStaff(rawAdsData);

    rawSalesData.forEach(sale => {
        let isMatch = true;

        if (sale.date) {
            if (fromFilter && sale.date < fromFilter) isMatch = false;
            if (toFilter && sale.date > toFilter) isMatch = false;
        }

        if (isMatch) {
            // Xác định số đơn: nếu có slDon thì dùng slDon, ngược lại nếu revenue > 0 tính 1
            let orderInc = 0;
            if (sale.slDon !== undefined && sale.slDon !== null && !isNaN(sale.slDon)) {
                if (sale.slDon > 0) {
                    orderInc = sale.slDon;
                } else if (maxAvailableDate && sale.date && (sale.date.getMonth() !== maxAvailableDate.getMonth() || sale.date.getFullYear() !== maxAvailableDate.getFullYear())) {
                    // Dữ liệu tháng trước trong file Excel có slDonThangNay = 0, nên nếu có doanh thu thì tính 1 đơn
                    orderInc = sale.revenue > 0 ? 1 : 0;
                }
            } else {
                orderInc = sale.revenue > 0 ? 1 : 0;
            }

            totalSalesRevenueAll += sale.revenue;
            totalSalesOrdersAll += orderInc;

            // Chỉ cộng dồn theo các mã ID Quảng Cáo
            if (sale.adId) {
                if (!salesAgg[sale.adId]) {
                    salesAgg[sale.adId] = { orders: 0, revenue: 0 };
                }
                salesAgg[sale.adId].orders += orderInc;
                salesAgg[sale.adId].revenue += sale.revenue;
            }
        }
    });

    let grandTotalOrders = 0;
    let grandTotalRevenue = 0;
    let grandTotalSpend = 0;
    let grandTotalRealSpend = 0;
    const hasExpenseFile = rawExpenseData.length > 0;
    const suffix = idSuffixInput.value.trim() || '501';

    // Build realSpendMap
    const realSpendMap = {};
    rawExpenseData.forEach(row => {
        let expenseId = (row['Mã quảng cáo'] || row['ID quảng cáo'] || row['ID QC'] || row['ID chiến dịch'] || '').toString().trim();
        if (expenseId) {
            // Đổi đuôi ID thành giá trị người dùng nhập
            const len = suffix.length;
            if (expenseId.length >= len) {
                expenseId = expenseId.slice(0, -len) + suffix;
            } else {
                expenseId = expenseId + suffix;
            }
            const spend = parseNumeric(row['Số tiền đã chi tiêu (VND)'] || row['Chi tiêu'] || 0);
            if (!realSpendMap[expenseId]) realSpendMap[expenseId] = 0;
            realSpendMap[expenseId] += spend;
        }
    });

    const processedIds = new Set();

    // 2. Map with Ads data
    rawAdsData.forEach(row => {
        const campaignName = row['Tên chiến dịch'] || row['Tên QC'] || row['Chiến dịch'] || 'N/A';
        const adName = row['Tên QC'] || row['Tên nhóm QC'] || '';
        const adId = (row['ID QC'] || row['ID Ads'] || row['ID Quảng Cáo'] || '').toString().trim();

        if (campaignName === 'N/A' && adName === '' && !adId) return;

        let totalOrders = 0;
        let totalRevenue = 0;

        if (adId && salesAgg[adId]) {
            totalOrders = salesAgg[adId].orders;
            totalRevenue = salesAgg[adId].revenue;
        }

        let spend = parseNumeric(getRowVal(row, [
            'Tiền Ads (chưa VAT)', 'Tiền Ads (chưa vat)', 'Tiền Ads (Vat 10%)', 'Tiền Ads (VAT 10%)',
            'Tổng Tiền ADS', 'Tổng tiền ADS', 'Tiền Ads / 1.1', 'Chi phí', 'Tiền ADS', 'Tiền Ads'
        ]));
        const uniqueId = adId || Math.random().toString(36).substr(2, 9);
        
        processedIds.add(uniqueId);
        
        const realSpend = realSpendMap[uniqueId] || 0;
        const spendToEvaluate = hasExpenseFile ? realSpend : spend;

        let cpa = 0;
        if (totalOrders > 0) {
            cpa = spendToEvaluate / totalOrders;
        }

        // Accumulate grand totals
        grandTotalOrders += totalOrders;
        grandTotalRevenue += totalRevenue;
        grandTotalSpend += spend;
        grandTotalRealSpend += realSpend;

        adsData.push({
            id: uniqueId,
            campaign: campaignName,
            adName: adName,
            orders: totalOrders,
            revenue: totalRevenue,
            spend: spend,
            realSpend: realSpend,
            cpa: cpa,
            threshold: currentThresholds[uniqueId] !== undefined ? currentThresholds[uniqueId] : null,
            spendToEvaluate: spendToEvaluate
        });
    });

    // Thêm các ID có trong file chi tiêu nhưng không có trong file sale/ads
    Object.keys(realSpendMap).forEach(expId => {
        if (!processedIds.has(expId)) {
            const realSpend = realSpendMap[expId];
            grandTotalRealSpend += realSpend;
            
            adsData.push({
                id: expId,
                campaign: 'Không có trong Báo Cáo',
                adName: '',
                orders: 0,
                revenue: 0,
                spend: 0,
                realSpend: realSpend,
                cpa: 0,
                threshold: currentThresholds[expId] !== undefined ? currentThresholds[expId] : null,
                spendToEvaluate: realSpend
            });
        }
    });

    // Update Summary UI
    summaryCards.classList.remove('hidden');
    sumRevenueEl.textContent = formatCurrency(grandTotalRevenue);
    sumOrdersEl.textContent = grandTotalOrders;
    
    const finalSpendTotal = hasExpenseFile ? grandTotalRealSpend : grandTotalSpend;
    sumSpendEl.textContent = formatCurrency(finalSpendTotal);
    
    let grandCpa = 0;
    if (grandTotalOrders > 0) {
        grandCpa = finalSpendTotal / grandTotalOrders;
    }
    sumCpaEl.textContent = formatCurrency(grandCpa);

    let costPercent = 0;
    if (grandTotalRevenue > 0) {
        costPercent = (finalSpendTotal / grandTotalRevenue) * 100;
    }
    sumCostPercentEl.textContent = costPercent.toFixed(2) + '%';

    currentAllSalesRevenue = grandTotalRevenue;
    currentAllSalesOrders = grandTotalOrders;
    currentGrandTotalRevenue = grandTotalRevenue;
    currentFinalSpendTotal = finalSpendTotal;
    calculateTargets();

    sortDataAndRender();
    renderSalesPerformance();
};

const updateSortIcons = () => {
    document.querySelectorAll('th[data-sort] .sort-icon').forEach(icon => {
        icon.innerHTML = ''; // Clear all icons
    });
    const activeTh = document.querySelector(`th[data-sort="${sortColumn}"] .sort-icon`);
    if (activeTh) {
        if (sortDirection === 'asc') {
            activeTh.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" class="w-4 h-4"><path fill-rule="evenodd" d="M14.77 12.79a.75.75 0 01-1.06-.02L10 8.832 6.29 12.77a.75.75 0 11-1.08-1.04l4.25-4.5a.75.75 0 011.08 0l4.25 4.5a.75.75 0 01-.02 1.06z" clip-rule="evenodd" /></svg>`;
        } else {
            activeTh.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 20" fill="currentColor" class="w-4 h-4"><path fill-rule="evenodd" d="M5.23 7.21a.75.75 0 011.06.02L10 11.168l3.71-3.938a.75.75 0 111.08 1.04l-4.25 4.5a.75.75 0 01-1.08 0l-4.25-4.5a.75.75 0 01.02-1.06z" clip-rule="evenodd" /></svg>`;
        }
    }
};

const sortDataAndRender = () => {
    adsData.sort((a, b) => {
        let valA = a[sortColumn];
        let valB = b[sortColumn];

        if (sortColumn === 'status') {
            const getSeverity = (ad) => {
                const defaultZeroOrderThreshold = 350000;
                const spendToEvaluate = ad.spendToEvaluate;

                if (is10PercentMode) {
                    if (spendToEvaluate === 0) return 0;
                    if (ad.orders === 0) {
                        if (spendToEvaluate > defaultZeroOrderThreshold) return 3;
                        if (spendToEvaluate >= defaultZeroOrderThreshold * 0.8) return 2;
                        return 1;
                    }
                    const tenPercentRev = ad.revenue * 0.1;
                    if (spendToEvaluate > tenPercentRev) return 3;
                    if (spendToEvaluate >= tenPercentRev * 0.8) return 2;
                    return 1;
                } else {
                    const threshold = ad.threshold !== null ? ad.threshold : (parseFloat(globalThresholdInput.value) || 0);
                    
                    if (ad.orders === 0) {
                        if (spendToEvaluate === 0) return 0;
                        const t = threshold > 0 ? threshold : defaultZeroOrderThreshold;
                        if (spendToEvaluate > t) return 3;
                        if (spendToEvaluate >= t * 0.8) return 2;
                        return 1;
                    }
                    
                    if (threshold > 0) {
                        if (ad.cpa > threshold) return 3;
                        if (ad.cpa >= threshold * 0.8) return 2;
                    }
                    return 1;
                }
            };
            valA = getSeverity(a);
            valB = getSeverity(b);
        }

        if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
        if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
        return 0;
    });

    updateSortIcons();
    renderTable();
};

document.querySelectorAll('th[data-sort]').forEach(th => {
    th.addEventListener('click', () => {
        const col = th.getAttribute('data-sort');
        if (sortColumn === col) {
            sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
            sortColumn = col;
            sortDirection = col === 'campaign' ? 'asc' : 'desc';
        }
        sortDataAndRender();
    });
});

// Render Table
const renderTable = () => {
    if (adsData.length === 0) {
        tableContainer.classList.add('hidden');
        emptyState.classList.remove('hidden');
        return;
    }

    tableContainer.classList.remove('hidden');
    emptyState.classList.add('hidden');
    tableBody.innerHTML = '';

    const globalThreshold = parseFloat(globalThresholdInput.value) || 0;
    const defaultZeroOrderThreshold = 350000;
    const hasExpenseFile = rawExpenseData.length > 0;

    adsData.forEach(ad => {
        let statusClass = 'status-good';
        let statusText = 'Tốt';
        let rowClass = '';
        
        const rowThreshold = ad.threshold !== null ? ad.threshold : globalThreshold;
        const spendToEvaluate = ad.spendToEvaluate;

        if (is10PercentMode) {
            if (spendToEvaluate === 0) {
                statusClass = 'status-inactive';
                statusText = 'Chưa tiêu tiền';
                rowClass = 'row-inactive';
            } else if (ad.orders === 0) {
                if (spendToEvaluate > defaultZeroOrderThreshold) {
                    statusClass = 'status-danger';
                    statusText = 'Vượt 350k (0 đơn)';
                    rowClass = 'row-danger';
                } else if (spendToEvaluate >= defaultZeroOrderThreshold * 0.8) {
                    statusClass = 'status-warning';
                    statusText = 'Cảnh báo 350k';
                    rowClass = 'row-warning';
                } else {
                    statusClass = 'status-good';
                    statusText = 'Tốt (< 350k)';
                }
            } else {
                const tenPercentRev = ad.revenue * 0.1;
                if (spendToEvaluate > tenPercentRev) {
                    statusClass = 'status-danger';
                    statusText = 'Lỗ (>10% DT)';
                    rowClass = 'row-danger';
                } else if (spendToEvaluate >= tenPercentRev * 0.8) {
                    statusClass = 'status-warning';
                    statusText = 'Cảnh báo (>=8% DT)';
                    rowClass = 'row-warning';
                } else {
                    statusClass = 'status-good';
                    statusText = 'Lãi (<8% DT)';
                }
            }
        } else {
            if (ad.orders === 0) {
                if (spendToEvaluate === 0) {
                    statusClass = 'status-inactive';
                    statusText = 'Chưa tiêu tiền';
                    rowClass = 'row-inactive';
                } else {
                    const thresholdToUse = rowThreshold > 0 ? rowThreshold : defaultZeroOrderThreshold;
                    if (spendToEvaluate > thresholdToUse) {
                        statusClass = 'status-danger';
                        statusText = `Vượt ngưỡng ${rowThreshold > 0 ? '(0 đơn)' : '350k'}`;
                        rowClass = 'row-danger';
                    } else if (spendToEvaluate >= thresholdToUse * 0.8) {
                        statusClass = 'status-warning';
                        statusText = `Cảnh báo ${rowThreshold > 0 ? '(0 đơn)' : '350k'}`;
                        rowClass = 'row-warning';
                    } else {
                        statusClass = 'status-good';
                        statusText = 'Tốt (0 đơn)';
                    }
                }
            } else {
                if (rowThreshold > 0) {
                    if (ad.cpa > rowThreshold) {
                        statusClass = 'status-danger';
                        statusText = 'Vượt ngưỡng';
                        rowClass = 'row-danger';
                    } else if (ad.cpa >= rowThreshold * 0.8) {
                        statusClass = 'status-warning';
                        statusText = 'Cảnh báo (>80%)';
                        rowClass = 'row-warning';
                    }
                }
            }
        }

        // Apply filters
        let isDanger = statusClass === 'status-danger';
        let isWarning = statusClass === 'status-warning';
        let isGood = statusClass === 'status-good' || (statusClass === 'status-inactive' && ad.orders > 0);

        if (currentFilter === 'danger' && !isDanger) return;
        if (currentFilter === 'warning' && !isWarning) return;
        if (currentFilter === 'good' && !isGood) return;

        const tr = document.createElement('tr');
        tr.className = rowClass;

        const displayName = ad.adName && ad.adName !== ad.campaign
            ? `<div class="font-medium text-white">${ad.campaign}</div><div class="text-xs text-slate-400 mt-1">${ad.adName}</div><div class="text-xs text-slate-500 mt-1">ID: ${ad.id}</div>`
            : `<div class="font-medium text-white">${ad.campaign}</div><div class="text-xs text-slate-500 mt-1">ID: ${ad.id}</div>`;

        tr.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap">
                ${displayName}
            </td>
            <td class="px-6 py-4 text-center font-bold ${ad.orders > 0 ? 'text-emerald-400' : 'text-slate-500'}">
                ${ad.orders}
            </td>
            <td class="px-6 py-4 text-right font-medium text-emerald-300">
                ${formatCurrency(ad.revenue)}
            </td>
            <td class="px-6 py-4 text-right font-medium text-slate-400">
                ${formatCurrency(ad.spend)}
            </td>
            <td class="px-6 py-4 text-right font-medium text-blue-300">
                ${ad.realSpend > 0 ? formatCurrency(ad.realSpend) : '<span class="text-slate-600 font-normal">-</span>'}
            </td>
            <td class="px-6 py-4 text-right font-bold ${statusClass === 'status-danger' ? 'text-red-400' : (statusClass === 'status-warning' ? 'text-yellow-400' : 'text-slate-300')}">
                ${ad.orders > 0 ? formatCurrency(ad.cpa) : '<span class="text-slate-500 font-normal">N/A</span>'}
            </td>
            <td class="px-6 py-4">
                ${is10PercentMode ? 
                    (ad.orders === 0 ? 
                        `<span class="text-purple-400 font-medium whitespace-nowrap">Mặc định: 350.000 đ</span>` : 
                        `<span class="text-purple-400 font-medium whitespace-nowrap">Ngưỡng: ${formatCurrency(ad.revenue * 0.1)}</span>`) :
                    `<input type="number" 
                           data-id="${ad.id}" 
                           value="${ad.threshold !== null ? ad.threshold : ''}" 
                           placeholder="${globalThreshold > 0 ? 'Mặc: ' + globalThreshold : 'Nhập...'}"
                           class="row-threshold w-full bg-slate-900/50 border border-slate-700/50 text-white text-xs rounded focus:ring-brand-500 focus:border-brand-500 block p-2 outline-none placeholder-slate-600 transition-all">`
                }
            </td>
            <td class="px-6 py-4 text-center">
                <div class="flex flex-col items-center justify-center gap-1">
                    <span class="status-indicator ${statusClass}"></span>
                    <span class="text-[10px] uppercase font-bold tracking-wide">${statusText}</span>
                </div>
            </td>
        `;
        tableBody.appendChild(tr);
    });

    document.querySelectorAll('.row-threshold').forEach(input => {
        input.addEventListener('input', (e) => {
            const id = e.target.getAttribute('data-id');
            const val = parseFloat(e.target.value);

            if (isNaN(val)) {
                delete currentThresholds[id];
            } else {
                currentThresholds[id] = val;
            }

            const adIndex = adsData.findIndex(a => a.id === id);
            if (adIndex > -1) {
                adsData[adIndex].threshold = isNaN(val) ? null : val;
            }
        });

        input.addEventListener('blur', () => {
            renderTable();
        });

        input.addEventListener('keypress', (e) => {
            if (e.key === 'Enter') {
                renderTable();
            }
        });
    });
};

applyGlobalBtn.addEventListener('click', () => {
    is10PercentMode = false;
    currentThresholds = {};
    adsData.forEach(ad => ad.threshold = null);
    renderTable();
});

calc10PercentBtn.addEventListener('click', () => {
    is10PercentMode = true;
    renderTable();
});

globalThresholdInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        is10PercentMode = false;
        currentThresholds = {};
        adsData.forEach(ad => ad.threshold = null);
        renderTable();
    }
});

filterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        const filter = btn.getAttribute('data-filter');
        currentFilter = filter;
        
        // Update UI
        filterBtns.forEach(b => {
            if (b.getAttribute('data-filter') === filter) {
                b.classList.remove('opacity-50');
                b.classList.add('border-white/20', 'shadow-md');
            } else {
                b.classList.add('opacity-50');
                b.classList.remove('border-white/20', 'shadow-md');
            }
        });

        renderTable();
    });
});

// Render Sales Performance
const renderSalesPerformance = () => {
    const salesPerfContainer = document.getElementById('salesPerformanceContainer');
    const salesPerfBody = document.getElementById('salesPerfBody');
    if (!salesPerfContainer || !salesPerfBody) return;

    if (rawSalesData.length === 0) {
        salesPerfContainer.classList.add('hidden');
        return;
    }

    salesPerfContainer.classList.remove('hidden');
    salesPerfBody.innerHTML = '';

    // Aggregate by Sale Name based on date filter
    let fromFilter = null;
    let toFilter = null;
    
    if (datePicker.selectedDates.length > 0) {
        fromFilter = new Date(datePicker.selectedDates[0]);
        fromFilter.setHours(0, 0, 0, 0);
        if (datePicker.selectedDates.length === 2) {
            toFilter = new Date(datePicker.selectedDates[1]);
            toFilter.setHours(23, 59, 59, 999);
        } else {
            toFilter = new Date(datePicker.selectedDates[0]); 
            toFilter.setHours(23, 59, 59, 999);
        }
    }

    // Tập hợp danh sách ID QC của tab hiện tại để chỉ tính hiệu suất Sale cho các đơn có ID QC này
    const currentTabAdIds = new Set(rawAdsData.map(r => (r['ID QC'] || r['ID Ads'] || r['ID Quảng Cáo'] || '').toString().trim()).filter(Boolean));

    const saleAgg = {};
    rawSalesData.forEach(sale => {
        let isMatch = true;
        
        if (sale.date) {
            if (fromFilter && sale.date < fromFilter) isMatch = false;
            if (toFilter && sale.date > toFilter) isMatch = false;
        }

        // Chỉ ghi nhận vào hiệu suất Sale nếu đơn hàng có mã ID QC thuộc tab hiện tại
        if (isMatch && sale.adId && currentTabAdIds.has(sale.adId)) {
            let orderInc = 0;
            if (sale.slDon !== undefined && sale.slDon !== null && !isNaN(sale.slDon)) {
                if (sale.slDon > 0) {
                    orderInc = sale.slDon;
                } else if (maxAvailableDate && sale.date && (sale.date.getMonth() !== maxAvailableDate.getMonth() || sale.date.getFullYear() !== maxAvailableDate.getFullYear())) {
                    orderInc = sale.revenue > 0 ? 1 : 0;
                }
            } else {
                orderInc = sale.revenue > 0 ? 1 : 0;
            }

            if (!saleAgg[sale.saleName]) {
                saleAgg[sale.saleName] = { orders: 0, revenue: 0 };
            }
            saleAgg[sale.saleName].orders += orderInc;
            saleAgg[sale.saleName].revenue += sale.revenue;
        }
    });

    const sortedSales = Object.keys(saleAgg).map(name => ({
        name: name,
        orders: saleAgg[name].orders,
        revenue: saleAgg[name].revenue
    })).sort((a, b) => b.revenue - a.revenue);

    if (sortedSales.length === 0) {
        salesPerfContainer.classList.add('hidden');
        return;
    }

    sortedSales.forEach((s, index) => {
        const tr = document.createElement('tr');
        
        // Highlight top 3
        let rankClass = 'text-slate-400';
        if (index === 0) rankClass = 'text-yellow-400 font-bold text-lg drop-shadow-[0_0_5px_rgba(250,204,21,0.5)]';
        else if (index === 1) rankClass = 'text-slate-300 font-bold text-base';
        else if (index === 2) rankClass = 'text-amber-600 font-bold text-base';

        tr.innerHTML = `
            <td class="px-6 py-4 whitespace-nowrap ${rankClass} text-center w-16">
                #${index + 1}
            </td>
            <td class="px-6 py-4 whitespace-nowrap font-medium text-white">
                <div class="flex items-center gap-3">
                    <div class="w-8 h-8 rounded-full bg-brand-500/20 text-brand-400 flex items-center justify-center font-bold text-xs uppercase">
                        ${s.name.substring(0, 2)}
                    </div>
                    ${s.name}
                </div>
            </td>
            <td class="px-6 py-4 text-center font-bold text-white">
                ${s.orders}
            </td>
            <td class="px-6 py-4 text-right font-medium text-emerald-300">
                ${formatCurrency(s.revenue)}
            </td>
        `;
        salesPerfBody.appendChild(tr);
    });
};

// Target Inputs Event Listeners & Initialization
if (targetMonthRevInput) {
    targetMonthRevInput.addEventListener('input', () => {
        const val = targetMonthRevInput.value.trim();
        if (!val) {
            if (targetRevPreview) targetRevPreview.textContent = '0 đ';
        } else {
            const parsed = parseTargetInput(targetMonthRevInput);
            if (targetRevPreview) targetRevPreview.textContent = formatTargetDisplay(parsed);
        }
        calculateTargets();
    });

    targetMonthRevInput.addEventListener('blur', () => {
        const parsed = parseTargetInput(targetMonthRevInput);
        if (parsed > 0) {
            targetMonthRevInput.value = parsed.toLocaleString('vi-VN');
            if (targetRevPreview) targetRevPreview.textContent = formatTargetDisplay(parsed);
        }
        calculateTargets();
    });
}

if (targetMonthSpendInput) {
    targetMonthSpendInput.addEventListener('input', () => {
        const val = targetMonthSpendInput.value.trim();
        if (!val) {
            if (targetSpendPreview) targetSpendPreview.textContent = '0 đ';
        } else {
            const parsed = parseTargetInput(targetMonthSpendInput);
            if (targetSpendPreview) targetSpendPreview.textContent = formatTargetDisplay(parsed);
        }
        calculateTargets();
    });

    targetMonthSpendInput.addEventListener('blur', () => {
        const parsed = parseTargetInput(targetMonthSpendInput);
        if (parsed > 0) {
            targetMonthSpendInput.value = parsed.toLocaleString('vi-VN');
            if (targetSpendPreview) targetSpendPreview.textContent = formatTargetDisplay(parsed);
        }
        calculateTargets();
    });
}

if (targetDaysInMonthInput) {
    targetDaysInMonthInput.addEventListener('input', () => {
        isCustomDaysInMonth = true;
        const days = parseInt(targetDaysInMonthInput.value, 10);
        if (!isNaN(days) && days > 0 && monthDetectedLabel) {
            monthDetectedLabel.textContent = `(Tùy chỉnh: ${days} ngày)`;
        }
        calculateTargets();
    });

    targetDaysInMonthInput.addEventListener('blur', () => {
        let days = parseInt(targetDaysInMonthInput.value, 10);
        if (isNaN(days) || days < 1) {
            days = 30;
            targetDaysInMonthInput.value = 30;
        }
        calculateTargets();
    });
}

// Restore saved targets from localStorage
const initTargets = () => {
    try {
        const savedRev = localStorage.getItem('cptt_target_month_rev');
        const savedSpend = localStorage.getItem('cptt_target_month_spend');
        const savedDays = localStorage.getItem('cptt_target_days_month');

        if (savedRev) {
            let num = parseNumeric(savedRev);
            if (num > 0 && num < 10000) num *= 1000000;
            if (num > 0 && targetMonthRevInput) {
                targetMonthRevInput.value = num.toLocaleString('vi-VN');
                if (targetRevPreview) targetRevPreview.textContent = formatTargetDisplay(num);
            }
        }
        if (savedSpend) {
            let num = parseNumeric(savedSpend);
            if (num > 0 && num < 10000) num *= 1000000;
            if (num > 0 && targetMonthSpendInput) {
                targetMonthSpendInput.value = num.toLocaleString('vi-VN');
                if (targetSpendPreview) targetSpendPreview.textContent = formatTargetDisplay(num);
            }
        }
        if (savedDays) {
            const days = parseInt(savedDays, 10);
            if (days > 0 && targetDaysInMonthInput) {
                targetDaysInMonthInput.value = days;
                isCustomDaysInMonth = true;
                if (monthDetectedLabel) monthDetectedLabel.textContent = `(Tùy chỉnh: ${days} ngày)`;
            }
        }
    } catch (e) {
        console.warn('LocalStorage error:', e);
    }

    calculateTargets();
};

// Initialize Targets on load
initTargets();

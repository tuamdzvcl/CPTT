// app.js
const excelFileInput = document.getElementById('excelFile');
const tableContainer = document.getElementById('tableContainer');
const emptyState = document.getElementById('emptyState');
const tableBody = document.getElementById('tableBody');
const globalThresholdInput = document.getElementById('globalThreshold');
const applyGlobalBtn = document.getElementById('applyGlobalBtn');
const sheetNameLabel = document.getElementById('sheetNameLabel');

const dateRangeInput = document.getElementById('dateRange');
const summaryCards = document.getElementById('summaryCards');
const sumRevenueEl = document.getElementById('sumRevenue');
const sumOrdersEl = document.getElementById('sumOrders');
const sumSpendEl = document.getElementById('sumSpend');
const sumCpaEl = document.getElementById('sumCpa');
const sumCostPercentEl = document.getElementById('sumCostPercent');

let adsData = []; 
let rawAdsData = []; 
let rawSalesData = [];
let currentThresholds = {}; 
let sortColumn = 'orders';
let sortDirection = 'desc'; 

// Format currency VND
const formatCurrency = (value) => {
    if (isNaN(value) || value === null) return '0 đ';
    return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(value);
};

// Convert cell Date or Serial to local JS Date (midnight)
const extractDateFromCellDate = (val) => {
    if (!val) return null;
    let date = val;
    if (typeof val === 'number') {
        // Excel serial to UTC date
        const utcMs = Math.round((val - 25569) * 86400 * 1000);
        date = new Date(utcMs);
    } else if (typeof val === 'string') {
        date = new Date(val);
    }
    if (!(date instanceof Date) || isNaN(date)) return null;
    
    // cellDates=true creates a Date where UTC matches Excel's local time.
    // e.g. Excel 2026-06-01 -> 2026-06-01T00:00:00Z
    // We extract the UTC components to form a local Date object at midnight.
    return new Date(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate());
};

// Initialize Flatpickr
let datePicker = flatpickr("#dateRange", {
    mode: "range",
    dateFormat: "Y-m-d",
    locale: "vn",
    theme: "dark",
    onChange: function(selectedDates) {
        if (selectedDates.length === 2 || selectedDates.length === 0) {
            processData();
        }
    }
});

// Parse numeric value
const parseNumeric = (val) => {
    if (typeof val === 'number') return val;
    if (typeof val === 'string') {
        const parsed = parseFloat(val.replace(/[^\d.-]/g, ''));
        return isNaN(parsed) ? 0 : parsed;
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
        
        rawSalesData = [];
        rawAdsData = [];
        let targetAdsSheetName = '';
        let minDate = null;
        let maxDate = null;

        workbook.SheetNames.forEach(sheetName => {
            const lowerName = sheetName.toLowerCase();
            const worksheet = workbook.Sheets[sheetName];
            const data = XLSX.utils.sheet_to_json(worksheet, { defval: null });

            // Detect SALE sheets
            if (lowerName.includes('sale')) {
                data.forEach(row => {
                    const adId = row['ID Quảng Cáo'] || row['ID Ads'] || row['ID QC'] || null;
                    const dateVal = row['Ngày đặt'] || row['Ngày'] || null;
                    const amount = parseNumeric(row['Thành Tiền (Trước VAT)'] || row['Thành tiền'] || row['Thành Tiền'] || 0);
                    const qty = parseNumeric(row['Số Lượng'] || row['Số lượng'] || 1);

                    const jsDate = extractDateFromCellDate(dateVal);
                    
                    if (jsDate) {
                        if (!minDate || jsDate < minDate) minDate = jsDate;
                        if (!maxDate || jsDate > maxDate) maxDate = jsDate;
                    }

                    if (adId && adId.toString().trim() !== '') {
                        rawSalesData.push({
                            adId: adId.toString().trim(),
                            date: jsDate,
                            revenue: amount,
                            qty: qty > 0 ? qty : 1
                        });
                    }
                });
            }

            // Detect ADS sheet
            if (lowerName.includes('t.anh') || lowerName.includes('tuấn anh')) {
                targetAdsSheetName = sheetName;
                rawAdsData = data;
            }
        });

        // Fallback
        if (rawAdsData.length === 0 && workbook.SheetNames.length > 0) {
            targetAdsSheetName = workbook.SheetNames[0];
            rawAdsData = XLSX.utils.sheet_to_json(workbook.Sheets[targetAdsSheetName], { defval: null });
        }

        sheetNameLabel.textContent = `(Dữ liệu từ ${targetAdsSheetName})`;

        // Set default dates
        if (minDate && maxDate) {
            // Default to showing all available data from the file
            datePicker.setDate([minDate, maxDate]);
        }

        processData();
    };
    reader.readAsArrayBuffer(file);
});

// Process data 
const processData = () => {
    adsData = [];
    
    // Parse filters from Flatpickr
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

    // 1. Aggregate Sales by Ad ID
    const salesAgg = {};
    rawSalesData.forEach(sale => {
        let isMatch = true;
        
        if (sale.date) {
            if (fromFilter && sale.date < fromFilter) isMatch = false;
            if (toFilter && sale.date > toFilter) isMatch = false;
        }

        if (isMatch) {
            if (!salesAgg[sale.adId]) {
                salesAgg[sale.adId] = { orders: 0, revenue: 0 };
            }
            salesAgg[sale.adId].orders += 1; 
            salesAgg[sale.adId].revenue += sale.revenue;
        }
    });

    let grandTotalOrders = 0;
    let grandTotalRevenue = 0;
    let grandTotalSpend = 0;

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

        // Spend from ADS sheet
        // Note: The spend in ADS sheet might be the all-time or monthly total. We use it as is for that row.
        let spend = parseNumeric(row['Tiền Ads (chưa VAT)'] || row['Tiền Ads (Vat 10%)'] || row['Chi phí'] || 0);
        
        // If the user filtered dates and this ad has no orders in that date, should we still count its total spend in the Grand Total?
        // Let's count it. If they ran ads but got 0 orders in that date range, it's still spend. 
        // Wait, what if the spend is from another month entirely?
        // Since we can't filter the spend by date, we'll just include the spend if it's in the rawAdsData.
        // We'll calculate CPA based on the orders they got in this date range vs the total spend shown.
        
        let cpa = 0;
        if (totalOrders > 0) {
            cpa = spend / totalOrders;
        }

        const uniqueId = adId || Math.random().toString(36).substr(2, 9);

        // Accumulate grand totals
        grandTotalOrders += totalOrders;
        grandTotalRevenue += totalRevenue;
        grandTotalSpend += spend;

        adsData.push({
            id: uniqueId,
            campaign: campaignName,
            adName: adName,
            orders: totalOrders,
            revenue: totalRevenue,
            spend: spend,
            cpa: cpa,
            threshold: currentThresholds[uniqueId] !== undefined ? currentThresholds[uniqueId] : null
        });
    });

    // Update Summary UI
    summaryCards.classList.remove('hidden');
    sumRevenueEl.textContent = formatCurrency(grandTotalRevenue);
    sumOrdersEl.textContent = grandTotalOrders;
    sumSpendEl.textContent = formatCurrency(grandTotalSpend);
    
    let grandCpa = 0;
    if (grandTotalOrders > 0) {
        grandCpa = grandTotalSpend / grandTotalOrders;
    }
    sumCpaEl.textContent = formatCurrency(grandCpa);

    let costPercent = 0;
    if (grandTotalRevenue > 0) {
        costPercent = (grandTotalSpend / grandTotalRevenue) * 100;
    }
    sumCostPercentEl.textContent = costPercent.toFixed(2) + '%';

    sortDataAndRender();
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
            // Helper to get status severity: Danger=3, Warning=2, Good=1, Inactive=0
            const getSeverity = (ad) => {
                const threshold = ad.threshold !== null ? ad.threshold : (parseFloat(globalThresholdInput.value) || 0);
                if (ad.orders === 0 && ad.spend === 0) return 0;
                if (ad.orders === 0 && threshold > 0 && ad.spend >= threshold) return 3;
                if (ad.orders === 0 && threshold > 0 && ad.spend >= threshold * 0.8) return 2;
                if (ad.orders === 0) return 0;
                if (threshold > 0) {
                    if (ad.cpa > threshold) return 3;
                    if (ad.cpa >= threshold * 0.8) return 2;
                }
                return 1;
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

    adsData.forEach(ad => {
        const thresholdToUse = ad.threshold !== null ? ad.threshold : globalThreshold;
        
        let statusClass = 'status-good';
        let statusText = 'Tốt';
        let rowClass = '';

        if (ad.orders === 0) {
            if (ad.spend === 0) {
                statusClass = 'status-inactive';
                statusText = 'Chưa tiêu tiền';
                rowClass = 'row-inactive';
            } else {
                if (thresholdToUse > 0 && ad.spend >= thresholdToUse) {
                    statusClass = 'status-danger';
                    statusText = 'Vượt ngưỡng (0 đơn)';
                    rowClass = 'row-danger';
                } else if (thresholdToUse > 0 && ad.spend >= thresholdToUse * 0.8) {
                    statusClass = 'status-warning';
                    statusText = 'Cảnh báo (0 đơn)';
                    rowClass = 'row-warning';
                } else {
                    statusClass = 'status-inactive';
                    statusText = 'Chưa có đơn';
                }
            }
        } else {
            if (thresholdToUse > 0) {
                if (ad.cpa > thresholdToUse) {
                    statusClass = 'status-danger';
                    statusText = 'Vượt ngưỡng';
                    rowClass = 'row-danger';
                } else if (ad.cpa >= thresholdToUse * 0.8) {
                    statusClass = 'status-warning';
                    statusText = 'Cảnh báo (>80%)';
                    rowClass = 'row-warning';
                }
            }
        }

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
            <td class="px-6 py-4 text-right font-medium text-slate-300">
                ${formatCurrency(ad.spend)}
            </td>
            <td class="px-6 py-4 text-right font-bold ${statusClass === 'status-danger' ? 'text-red-400' : (statusClass === 'status-warning' ? 'text-yellow-400' : 'text-slate-300')}">
                ${ad.orders > 0 ? formatCurrency(ad.cpa) : '<span class="text-slate-500 font-normal">N/A</span>'}
            </td>
            <td class="px-6 py-4">
                <input type="number" 
                       data-id="${ad.id}" 
                       value="${ad.threshold !== null ? ad.threshold : ''}" 
                       placeholder="${globalThreshold > 0 ? 'Mặc: ' + globalThreshold : 'Nhập...'}"
                       class="row-threshold w-full bg-slate-900/50 border border-slate-700/50 text-white text-xs rounded focus:ring-brand-500 focus:border-brand-500 block p-2 outline-none placeholder-slate-600 transition-all">
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
            if(e.key === 'Enter') {
                renderTable();
            }
        });
    });
};

applyGlobalBtn.addEventListener('click', () => {
    currentThresholds = {}; 
    adsData.forEach(ad => ad.threshold = null);
    renderTable();
});

globalThresholdInput.addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        currentThresholds = {};
        adsData.forEach(ad => ad.threshold = null);
        renderTable();
    }
});

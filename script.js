/**
 * ==== การเชื่อมต่อ GOOGLE SHEETS ====
 * ให้คุณใส่ URL ของ Web App จาก Google Apps Script ตรงนี้
 */
const GOOGLE_APP_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbzlT9ZOGKz4RgD3Np19SRAVOTNOAd9C-HGbsUFA-F2g-GGU5Th5zpwP2mYnEf4pI7Xe/exec';
// Optional: link to the Google Sheets UI (set this to your spreadsheet URL if you want a direct link)
const GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1N0yeiILt8A468mCQw7WQnIhMeeGEKjBO1jSMiHE-hgQ/edit?gid=608803957#gid=608803957'


// ตัวแปรเก็บข้อมูลทั้งหมดจาก Google Sheets
let allData = [];
let currentFilteredData = [];
// ตัวแปรเก็บกราฟ
let donutChart;
let barChart;
// Creditor dropdown data + selection state
let creditorData = [];
let selectedCreditors = new Set();
let selectedPayDocCreditors = new Set(); // New for Date Summary multi-select
let selectedOverdueRanges = new Set();
let overdueRanges = [
    { label: '1-30', min: 1, max: 30 },
    { label: '31-60', min: 31, max: 60 },
    { label: '61-90', min: 61, max: 90 },
    { label: '91-120', min: 91, max: 120 },
    { label: '121-150', min: 121, max: 150 },
    { label: '151-180', min: 151, max: 180 },
    { label: '181-210', min: 181, max: 210 },
    { label: '211-240', min: 211, max: 240 },
    { label: '241-270', min: 241, max: 270 },
    { label: '271-300', min: 271, max: 300 },
    { label: '301-330', min: 301, max: 330 }
];
let currentModalDate = ''; // Store date for PDF report header

// Thai month mapping for sorting and comparison
const monthMap = {
    'ม.ค.': 1, 'ก.พ.': 2, 'มี.ค.': 3, 'เม.ย.': 4, 'พ.ค.': 5, 'มิ.ย.': 6,
    'ก.ค.': 7, 'ส.ค.': 8, 'ก.ย.': 9, 'ต.ค.': 10, 'พ.ย.': 11, 'ธ.ค.': 12
};

// Short month names for printing (Thai) and a helper to format dates using Buddhist year
const thaiMonthsShort = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];

const formatThaiDateTime = (date) => {
    if (!date || !(date instanceof Date)) return '';
    const day = String(date.getDate()).padStart(2, '0');
    const month = thaiMonthsShort[date.getMonth()] || '';
    const yearBE = date.getFullYear() + 543; // Buddhist Era
    const hh = String(date.getHours()).padStart(2, '0');
    const mm = String(date.getMinutes()).padStart(2, '0');
    const ss = String(date.getSeconds()).padStart(2, '0');
    return `${day} ${month} ${yearBE} ${hh}:${mm}:${ss}`;
};

// Format numbers as Thai Baht currency
const formatCurrency = (number) => {
    return new Intl.NumberFormat('th-TH', {
        style: 'decimal',
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(number);
};

// Number counter animation function
const animateValue = (obj, start, end, duration, isCurrency = false) => {
    let startTimestamp = null;
    const step = (timestamp) => {
        if (!startTimestamp) startTimestamp = timestamp;
        const progress = Math.min((timestamp - startTimestamp) / duration, 1);
        const ease = 1 - Math.pow(1 - progress, 4); // easeOutQuart
        const currentVal = ease * (end - start) + start;

        if (isCurrency) {
            obj.innerText = formatCurrency(currentVal);
        } else {
            obj.innerText = `คิดเป็น ${currentVal.toFixed(2)}%`;
        }

        if (progress < 1) {
            window.requestAnimationFrame(step);
        } else {
            if (isCurrency) obj.innerText = formatCurrency(end);
            else obj.innerText = `คิดเป็น ${end.toFixed(2)}%`;
        }
    };
    window.requestAnimationFrame(step);
};

// DOM Elements
const paymentStatusFilter = document.getElementById('paymentStatusFilter');
const categoryFilter = document.getElementById('categoryFilter');
const monthFilter = document.getElementById('monthFilter');
const dayFilter = document.getElementById('dayFilter');
const yearFilter = document.getElementById('yearFilter');
const refreshBtn = document.getElementById('refreshBtn');
const loading = document.getElementById('loading');

const exportPdfBtn = document.getElementById('exportPdfBtn');

const totalAmountEl = document.getElementById('totalAmount');
const overdueAmountEl = document.getElementById('overdueAmount');
const ontimeAmountEl = document.getElementById('ontimeAmount');
const notdueAmountEl = document.getElementById('notdueAmount');

const totalPercentEl = document.getElementById('totalPercent');
const overduePercentEl = document.getElementById('overduePercent');
const ontimePercentEl = document.getElementById('ontimePercent');
const notduePercentEl = document.getElementById('notduePercent');

const nodateAmountEl = document.getElementById('nodateAmount');
const earlyAmountEl = document.getElementById('earlyAmount');
const nodatePercentEl = document.getElementById('nodatePercent');
const earlyPercentEl = document.getElementById('earlyPercent');

const pendingAmountEl = document.getElementById('pendingAmount');
const pendingPercentEl = document.getElementById('pendingPercent');

// Initialize the dashboard
const init = async () => {
    setupEventListeners();

    // Do not preselect the current year — defaults should start at 'all'

    initCharts();

    // Populate dayFilter 1-31
    if (dayFilter) {
        for (let i = 1; i <= 31; i++) {
            const opt = document.createElement('option');
            opt.value = i.toString().padStart(2, '0');
            opt.textContent = i.toString();
            dayFilter.appendChild(opt);
        }
    }

    // Ensure all filters default to 'all' on initial load
    const setDefaultAll = (id) => {
        const el = document.getElementById(id);
        if (el) el.value = 'all';
    };

    setDefaultAll('paymentStatusFilter');
    setDefaultAll('categoryFilter');
    setDefaultAll('dayFilter');
    setDefaultAll('monthFilter');
    setDefaultAll('yearFilter');
    setDefaultAll('payDocStatusFilter');
    setDefaultAll('payDocMonthFilter');
    setDefaultAll('payDocYearFilter');
    setDefaultAll('payDocCategoryFilter');
    setDefaultAll('payDocUrgencyFilter');

    if (GOOGLE_APP_SCRIPT_URL === 'YOUR_WEB_APP_URL_HERE') {
        document.getElementById('setupModal').classList.remove('hidden');
        renderOverdueDropdown(); // Populate ranges
        loadMockData();
    } else {
        renderOverdueDropdown(); // Populate ranges
        await fetchData();
    }
};

// Keep last opened details items for quick re-render/filtering
let lastDetailsItems = [];
let lastDateDetailItems = []; // New for Date Detail Modal

// Populate grouped summary table (One row per creditor + month + year) - Simplified to be reusable
function populateGroupedDetailsTable(items = [], tbodyId = 'detailsTableBody', tfootId = 'detailsTableFooter', btnViewItemsId = 'btnViewItems') {
    const detailsTableBody = document.getElementById(tbodyId);
    const detailsTableFooter = document.getElementById(tfootId);
    const table = detailsTableBody ? detailsTableBody.closest('table') : null;
    const detailsTableHeader = table ? table.querySelector('thead') : null;

    if (table) {
        table.classList.remove('grouped-mode');
        table.classList.add('grouped-mode-5col');
    }

    // Change Table Header for Summary View — now with Month Due & Year columns
    if (detailsTableHeader) {
        detailsTableHeader.innerHTML = `
            <tr>
                <th style="text-align: left; width: 40%;">ชื่อเจ้าหนี้</th>
                <th style="text-align: center; width: 18%;">Month Due</th>
                <th style="text-align: center; width: 6%;">Year</th>
                <th style="text-align: center !important; width: 12%;">จำนวนรายการ</th>
                <th style="text-align: right; width: 24%;">ยอดเงินรวม (฿)</th>
            </tr>
        `;
    }

    detailsTableBody.innerHTML = '';
    detailsTableFooter.innerHTML = '';

    if (!items || items.length === 0) {
        detailsTableBody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 32px; color: var(--text-muted);">ไม่มีข้อมูล</td></tr>`;
        return;
    }

    // Aggregate by docNo + monthDue + yearDue (creditor name + month + year)
    const groups = {};
    items.forEach(it => {
        const name = it.docNo || 'ไม่ระบุชื่อ';
        const month = it.monthDue || '-';
        const year = it.yearDue || '-';
        const key = `${name}|||${month}|||${year}`;
        const amt = Number(it.amount) || 0;
        if (!groups[key]) groups[key] = { name, month, year, total: 0, count: 0 };
        groups[key].total += amt;
        groups[key].count += 1;
    });

    const sortedGroups = Object.values(groups)
        .sort((a, b) => {
            // Sort by creditor name ascending (group same names together)
            const nameCompare = a.name.localeCompare(b.name, 'th');
            if (nameCompare !== 0) return nameCompare;
            // Then by year ascending
            const yearA = parseInt(a.year) || 9999;
            const yearB = parseInt(b.year) || 9999;
            if (yearA !== yearB) return yearA - yearB;
            // Then by month ascending
            const monthA = monthMap[a.month] || 99;
            const monthB = monthMap[b.month] || 99;
            if (monthA !== monthB) return monthA - monthB;
            // Then by total descending
            return (b.total - a.total) || (b.count - a.count);
        });

    // Count unique creditor names
    const uniqueCreditors = new Set(sortedGroups.map(g => g.name));

    let grandTotal = 0;
    sortedGroups.forEach(g => {
        grandTotal += g.total;
        const tr = document.createElement('tr');
        tr.className = 'summary-row-clickable';
        tr.innerHTML = `
            <td style="text-align: left; font-weight: 700; color: #fff;">${g.name}</td>
            <td style="text-align: center; color: var(--accent-primary); font-weight: 600;">${g.month}</td>
            <td style="text-align: center; color: var(--text-muted);">${g.year}</td>
            <td style="text-align: center; color: var(--text-muted);">${g.count} รายการ</td>
            <td style="text-align: right; color: var(--color-total); font-weight: 800; font-size: 15px; font-variant-numeric: tabular-nums;">${formatCurrency(g.total)}</td>
        `;
        // Drill-down: Click row to see details for this creditor + month + year
        tr.addEventListener('click', () => {
            const creditorItems = items.filter(it => {
                const itName = it.docNo || 'ไม่ระบุชื่อ';
                const itMonth = it.monthDue || '-';
                const itYear = it.yearDue || '-';
                return itName === g.name && itMonth === g.month && itYear === g.year;
            });
            const btnViewItems = document.getElementById(btnViewItemsId);
            if (btnViewItems) btnViewItems.click(); // Switch back to items view
            populateDetailsTable(creditorItems, tbodyId, tfootId);
        });
        detailsTableBody.appendChild(tr);
    });

    // Total footer
    const totalTr = document.createElement('tr');
    totalTr.className = 'total-row-summary';
    totalTr.innerHTML = `
        <td colspan="4" class="total-label">ยอดยกมาทั้งหมด (${uniqueCreditors.size} ราย / ${sortedGroups.length} รายการ):</td>
        <td class="total-amount-val">${formatCurrency(grandTotal)}</td>
    `;
    detailsTableFooter.appendChild(totalTr);
}

// Popluates the newly added Category summary
function populateCategoryDetailsTable(items = []) {
    const detailsTableBody = document.getElementById('detailsTableBody');
    const detailsTableFooter = document.getElementById('detailsTableFooter');
    const table = detailsTableBody ? detailsTableBody.closest('table') : null;
    const detailsTableHeader = table ? table.querySelector('thead') : null;

    if (table) {
        table.classList.add('grouped-mode');
    }

    if (detailsTableHeader) {
        detailsTableHeader.innerHTML = `
            <tr>
                <th style="text-align: left; width: 60%;">หมวดหมู่</th>
                <th style="text-align: center !important; width: 12%;">จำนวนรายการ</th>
                <th style="text-align: right; width: 28%;">ยอดเงินรวม (฿)</th>
            </tr>
        `;
    }

    detailsTableBody.innerHTML = '';
    detailsTableFooter.innerHTML = '';

    if (!items || items.length === 0) {
        detailsTableBody.innerHTML = `<tr><td colspan="3" style="text-align:center; padding: 32px; color: var(--text-muted);">ไม่มีข้อมูลสำหรับตัวกรองนี้</td></tr>`;
        return;
    }

    // Group items by category instead of creditor
    const groupMap = {};
    items.forEach(it => {
        const key = it.category ? it.category : 'ไม่ระบุหมวดหมู่';
        if (!groupMap[key]) {
            groupMap[key] = { count: 0, total: 0, name: key };
        }
        groupMap[key].count += 1;
        groupMap[key].total += (Number(it.amount) || 0);
    });

    const sortedGroups = Object.values(groupMap)
        .sort((a, b) => (b.total - a.total) || (b.count - a.count));

    let grandTotal = 0;
    sortedGroups.forEach(g => {
        grandTotal += g.total;
        const tr = document.createElement('tr');
        tr.className = 'summary-row-clickable';
        tr.innerHTML = `
            <td style="text-align: left; font-weight: 700; color: #fff;">${g.name}</td>
            <td style="text-align: center; color: var(--text-muted);">${g.count} รายการ</td>
            <td style="text-align: right; color: var(--color-total); font-weight: 800; font-size: 15px; font-variant-numeric: tabular-nums;">${formatCurrency(g.total)}</td>
        `;
        // Drill-down: Click row to see details for this category
        tr.addEventListener('click', () => {
            const catItems = items.filter(it => (it.category || 'ไม่ระบุหมวดหมู่') === g.name);
            const btnViewItems = document.getElementById('btnViewItems');
            if (btnViewItems) btnViewItems.click(); // Switch back to items view
            populateDetailsTable(catItems);
        });
        detailsTableBody.appendChild(tr);
    });

    // Total footer
    const totalTr = document.createElement('tr');
    totalTr.className = 'total-row-summary';
    totalTr.innerHTML = `
        <td colspan="2" class="total-label">ยอดยกมาทั้งหมด (${sortedGroups.length} หมวดหมู่):</td>
        <td class="total-amount-val">${formatCurrency(grandTotal)}</td>
    `;
    detailsTableFooter.appendChild(totalTr);
}

// Populate details table (used by full-list and group-click) - Reusable
function populateDetailsTable(items = [], tbodyId = 'detailsTableBody', tfootId = 'detailsTableFooter') {
    const detailsTableBody = document.getElementById(tbodyId);
    const detailsTableFooter = document.getElementById(tfootId);
    const table = detailsTableBody ? detailsTableBody.closest('table') : null;
    const detailsTableHeader = table ? table.querySelector('thead') : null;

    if (table) {
        table.classList.remove('grouped-mode', 'grouped-mode-5col');
    }

    // Restore Original Table Header for Detailed View
    if (detailsTableHeader) {
        detailsTableHeader.innerHTML = `
            <tr>
                <th style="text-align: left;">เลขที่เอกสาร</th>
                <th style="text-align: left;">เจ้าหนี้</th>
                <th style="text-align: left;">รายละเอียด</th>
                <th style="text-align: left;">หมวดหมู่</th>
                <th style="text-align: left;">วันที่ครบกำหนดชำระ</th>
                <th style="text-align: center; white-space: nowrap;">ระยะเวลา</th>
                <th style="text-align: right;">จำนวนเงิน (฿)</th>
            </tr>
        `;
    }

    detailsTableBody.innerHTML = '';
    detailsTableFooter.innerHTML = '';

    // Sort items by date (ascending: earliest first), then by amount (descending)
    if (items && items.length > 0) {
        items.sort((a, b) => {
            // First sort by year (ascending)
            const yearA = parseInt(a.yearDue) || 9999;
            const yearB = parseInt(b.yearDue) || 9999;
            if (yearA !== yearB) return yearA - yearB;

            // Sort by month (ascending)
            const monthA = monthMap[a.monthDue] || 99;
            const monthB = monthMap[b.monthDue] || 99;
            if (monthA !== monthB) return monthA - monthB;

            // Sort by day (ascending)
            const dayA = parseInt(a.dayDue) || 99;
            const dayB = parseInt(b.dayDue) || 99;
            if (dayA !== dayB) return dayA - dayB;

            // If date is equal, sort by amount (descending)
            return (Number(b.amount) || 0) - (Number(a.amount) || 0);
        });
    }

    let totalSum = 0;
    if (!items || items.length === 0) {
        detailsTableBody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding: 32px; color: var(--text-muted);">ไม่มีข้อมูลสำหรับตัวกรองนี้</td></tr>`;
    } else {
        items.forEach(item => {
            const amount = Number(item.amount) || 0;
            totalSum += amount;

            const statusStr = (item.status || '').toString().trim();
            let statusColor = 'var(--text-muted)';
            if (statusStr.includes('เกินกำหนด')) statusColor = '#ef4444';
            else if (statusStr.includes('ตรงดิว')) statusColor = '#10b981';
            else if (statusStr.includes('ยังไม่ถึงกำหนด')) statusColor = '#3b82f6';
            else if (statusStr.includes('จ่ายก่อนกำหนด')) statusColor = '#a855f7';

            const tr = document.createElement('tr');
            const dueDateStr = [item.dayDue, item.monthDue, item.yearDue].filter(Boolean).join(' ') || '-';
            tr.innerHTML = `
                <td style="font-weight: 500; font-family: monospace; color: var(--accent-primary);">${item.creditor || '-'}</td> <!-- Doc No / ID -->
                <td style="font-weight: 700; color: #fff;">${item.docNo || '-'}</td> <!-- Creditor Name -->
                <td style="color: var(--text-muted); font-size: 13.5px;">${item.description || '-'}</td>
                <td>
                    <span class="cat-pill" style="background: rgba(255,255,255,0.06); padding: 4px 8px; border-radius: 4px; font-size: 12px; white-space: nowrap;">${item.category || '-'}</span>
                </td>
                <td class="due-date-cell">${dueDateStr}</td>
                <td style="text-align: center; font-variant-numeric: tabular-nums;">${item.overdueDays || '0'}</td>
                <td style="text-align: right; color: var(--color-total); font-weight: 600; white-space: nowrap;">${formatCurrency(amount)}</td>
            `;
            detailsTableBody.appendChild(tr);
        });
    }

    // Total footer
    const totalTr = document.createElement('tr');
    totalTr.className = 'total-row-summary';
    totalTr.innerHTML = `
        <td colspan="6" class="total-label">ยอดรวมทั้งหมด (Total):</td>
        <td class="total-amount-val">${formatCurrency(totalSum)}</td>
    `;
    detailsTableFooter.appendChild(totalTr);

    // shrink status text to fit after rendering
    window.requestAnimationFrame(() => shrinkStatusTextToFit(detailsTableBody));
}

function renderGroupSummary(items = [], barId = 'groupSummaryBar', targetTbodyId = 'detailsTableBody', targetTfootId = 'detailsTableFooter', btnViewItemsId = 'btnViewItems') {
    const bar = document.getElementById(barId);
    if (!bar) return;
    bar.innerHTML = '';
    if (!items || items.length === 0) { bar.hidden = true; return; }

    // aggregate by creditor
    const groups = {};
    items.forEach(it => {
        const key = it.docNo || 'ไม่ระบุชื่อ';
        const amt = Number(it.amount) || 0;
        if (!groups[key]) groups[key] = { total: 0, count: 0, items: [] };
        groups[key].total += amt;
        groups[key].count += 1;
        groups[key].items.push(it);
    });

    const arr = Object.entries(groups).map(([name, o]) => ({ name, total: o.total, count: o.count, items: o.items }));
    arr.sort((a, b) => (b.total - a.total) || (b.count - a.count));

    // Also render a print-friendly summary table (used for PDF / first page)
    renderPrintGroupTable(arr);

    // Add 'show all' card
    const allBtn = document.createElement('button');
    allBtn.type = 'button'; allBtn.className = 'group-card group-card-all';
    allBtn.innerHTML = `<div class="group-card-name">แสดงทั้งหมด</div><div class="group-card-meta"><span class="group-count">${items.length} รายการ</span><span class="group-total">${formatCurrency(items.reduce((s, i) => s + (Number(i.amount) || 0), 0))}</span></div>`;
    allBtn.addEventListener('click', () => {
        document.querySelectorAll(`#${barId} .group-card`).forEach(c => c.classList.remove('is-selected'));
        const itemsToRender = (barId === 'groupSummaryBarDate') ? lastDateDetailItems : lastDetailsItems;
        populateDetailsTable(itemsToRender, targetTbodyId, targetTfootId);
    });
    bar.appendChild(allBtn);

    arr.forEach(g => {
        const card = document.createElement('button');
        card.type = 'button';
        card.className = 'group-card';
        card.innerHTML = `<div class="group-card-name">${g.name}</div><div class="group-card-meta"><span class="group-count">${g.count} รายการ</span><span class="group-total">${formatCurrency(g.total)}</span></div>`;
        card.addEventListener('click', () => {
            document.querySelectorAll(`#${barId} .group-card`).forEach(c => c.classList.remove('is-selected'));
            card.classList.add('is-selected');
            populateDetailsTable(g.items, targetTbodyId, targetTfootId);
        });
        bar.appendChild(card);
    });

    bar.hidden = false;
}

// Render a compact, print-friendly summary table above the full table
function renderPrintGroupTable(groupsArr = []) {
    const container = document.getElementById('globalPrintGroupTable');
    const tbody = document.getElementById('globalPrintGroupTableBody');
    if (!container || !tbody) return;
    tbody.innerHTML = '';
    if (!groupsArr || groupsArr.length === 0) {
        tbody.innerHTML = '';
        return;
    }

    groupsArr.forEach(g => {
        const tr = document.createElement('tr');
        const nameTd = document.createElement('td');
        const countTd = document.createElement('td');
        const totalTd = document.createElement('td');

        nameTd.textContent = g.name || '-';
        countTd.textContent = `${g.count} รายการ`;
        countTd.style.textAlign = 'center';
        totalTd.textContent = formatCurrency(g.total || 0);
        totalTd.style.textAlign = 'right';

        tr.appendChild(nameTd);
        tr.appendChild(countTd);
        tr.appendChild(totalTd);
        tbody.appendChild(tr);
    });

    // Add Grand Total row to tfoot
    const tfoot = document.getElementById('globalPrintGroupTableFooter');
    if (tfoot) {
        tfoot.innerHTML = '';
        const totalAmount = groupsArr.reduce((sum, g) => sum + (g.total || 0), 0);
        const totalCount = groupsArr.reduce((sum, g) => sum + (g.count || 0), 0);

        const tr = document.createElement('tr');
        tr.className = 'total-row-summary';

        const labelTd = document.createElement('td');
        labelTd.textContent = 'ยอดรวมทั้งหมด';
        labelTd.style.fontWeight = 'bold';
        labelTd.style.textAlign = 'left';

        const countTd = document.createElement('td');
        countTd.textContent = `${totalCount} รายการ`;
        countTd.style.fontWeight = 'bold';
        countTd.style.textAlign = 'center';

        const amountTd = document.createElement('td');
        amountTd.textContent = formatCurrency(totalAmount);
        amountTd.style.fontWeight = 'bold';
        amountTd.style.textAlign = 'right';

        tr.appendChild(labelTd);
        tr.appendChild(countTd);
        tr.appendChild(amountTd);
        tfoot.appendChild(tr);
    }
}

// Setup Listeners
const setupEventListeners = () => {
    const creditorFilter = document.getElementById('creditorFilter');
    const searchClearBtn = document.querySelector('.search-clear');

    // Wrap calls so we control whether date-summary updates.
    if (paymentStatusFilter) paymentStatusFilter.addEventListener('change', () => updateDashboard());
    if (categoryFilter) categoryFilter.addEventListener('change', () => updateDashboard());
    if (dayFilter) dayFilter.addEventListener('change', () => updateDashboard());
    if (monthFilter) monthFilter.addEventListener('change', () => updateDashboard());
    if (yearFilter) yearFilter.addEventListener('change', () => updateDashboard());

    // For the creditor search input we intentionally skip updating the Date Summary
    // so that searching only affects the top area (cards + charts) as requested.
    if (creditorFilter) {
        // typing filters the dropdown and updates the top charts/cards only
        // do NOT open dropdown on empty input to avoid auto-popup
        creditorFilter.addEventListener('input', (e) => {
            const q = e.target.value || '';
            filterCreditorDropdown(q, false);
            searchClearBtn && (searchClearBtn.hidden = creditorFilter.value.trim() === '');
            updateDashboard({ skipDateSummary: true });
        });

        // don't open dropdown on mere click — show suggestions only when typing
        creditorFilter.addEventListener('click', (e) => {
            e.stopPropagation();
            filterCreditorDropdown(e.target.value || '', false);
        });

        // show / hide clear button and wire its behavior
        if (searchClearBtn) {
            searchClearBtn.hidden = creditorFilter.value.trim() === '';
            creditorFilter.addEventListener('input', () => {
                searchClearBtn.hidden = creditorFilter.value.trim() === '';
            });
            searchClearBtn.addEventListener('click', () => {
                creditorFilter.value = '';
                searchClearBtn.hidden = true;
                // update dropdown
                filterCreditorDropdown('');
                creditorFilter.focus();
                // top area updates
                updateDashboard({ skipDateSummary: true });
            });
            creditorFilter.addEventListener('keydown', (e) => {
                if (e.key === 'Escape') {
                    // close dropdown first
                    closeCreditorDropdown();
                    creditorFilter.value = '';
                    searchClearBtn.hidden = true;
                    filterCreditorDropdown('');
                    updateDashboard({ skipDateSummary: true });
                } else if (e.key === 'ArrowDown') {
                    // focus first checkbox if exists
                    const first = document.querySelector('#creditorListPanel .creditor-checkbox');
                    if (first) first.focus();
                } else if (e.key === 'Enter') {
                    // If user presses Enter in the input, pick the first filtered creditor (if any)
                    const firstChk = document.querySelector('#creditorListPanel .creditor-checkbox');
                    if (firstChk) {
                        // toggle and trigger change
                        firstChk.checked = !firstChk.checked;
                        firstChk.dispatchEvent(new Event('change', { bubbles: true }));
                        // keep focus in input
                        e.preventDefault();
                        creditorFilter.focus();
                    }
                }
            });
        }

        // Toolbar buttons inside dropdown: clear-selection
        const clearSelectionBtn = document.querySelector('.clear-selection-btn');
        if (clearSelectionBtn) {
            clearSelectionBtn.addEventListener('click', () => {
                selectedCreditors.clear();
                document.querySelectorAll('#creditorListPanel .creditor-checkbox').forEach(cb => cb.checked = false);
                updateSelectedChips();
                updateDashboard({ skipDateSummary: true });
            });
        }

        // Close dropdown when clicking outside
        document.addEventListener('click', (e) => {
            const headerSearch = document.querySelector('.header-search');
            if (headerSearch && !headerSearch.contains(e.target)) {
                closeCreditorDropdown();
            }
            const headerOverdue = document.querySelector('.header-overdue');
            if (headerOverdue && !headerOverdue.contains(e.target)) {
                closeOverdueDropdown();
            }
        });
    }

    // Overdue Range Filter Setup
    const overdueRangeFilter = document.getElementById('overdueRangeFilter');
    if (overdueRangeFilter) {
        overdueRangeFilter.addEventListener('click', (e) => {
            e.stopPropagation();
            const dd = document.getElementById('overdueDropdown');
            if (dd && dd.hidden) {
                closeAllDropdowns('overdueDropdown'); // Close others first
            }
            if (dd) dd.hidden = !dd.hidden;
            updateHeaderSpacing();
        });

        const clearOverdueBtn = document.querySelector('.clear-overdue-btn');
        if (clearOverdueBtn) {
            clearOverdueBtn.addEventListener('click', () => {
                selectedOverdueRanges.clear();
                document.querySelectorAll('#overdueListPanel .overdue-checkbox').forEach(cb => cb.checked = false);
                updateSelectedOverdueChips();
                updateDashboard({ skipDateSummary: true });
            });
        }

        const overdueClearBtn = document.querySelector('.overdue-clear');
        if (overdueClearBtn) {
            overdueClearBtn.addEventListener('click', (e) => {
                e.stopPropagation();
                selectedOverdueRanges.clear();
                document.querySelectorAll('#overdueListPanel .overdue-checkbox').forEach(cb => cb.checked = false);
                updateSelectedOverdueChips();
                updateDashboard({ skipDateSummary: true });
            });
        }
    }

    // PayDoc date section filters (independent)
    const payDocStatusFilter = document.getElementById('payDocStatusFilter');
    const payDocMonthFilter = document.getElementById('payDocMonthFilter');
    const payDocYearFilter = document.getElementById('payDocYearFilter');

    if (payDocStatusFilter) payDocStatusFilter.addEventListener('change', updateDateSummary);
    if (payDocMonthFilter) payDocMonthFilter.addEventListener('change', updateDateSummary);
    if (payDocYearFilter) payDocYearFilter.addEventListener('change', updateDateSummary);

    // New PayDoc multi-select search and urgency filters
    const payDocCreditorFilter = document.getElementById('payDocCreditorFilter');
    const payDocSearchClear = document.getElementById('payDocSearchClear');
    const payDocUrgencyFilter = document.getElementById('payDocUrgencyFilter');
    const payDocClearSelection = document.getElementById('payDocClearSelection');

        if (payDocCreditorFilter) {
        payDocCreditorFilter.addEventListener('input', (e) => {
            const q = e.target.value || '';
            filterPayDocCreditorDropdown(q, false);
            if (payDocSearchClear) payDocSearchClear.hidden = q.trim() === '';
            updateDateSummary();
        });

        payDocCreditorFilter.addEventListener('click', (e) => {
            e.stopPropagation();
            filterPayDocCreditorDropdown(e.target.value || '', false);
        });

        if (payDocSearchClear) {
            payDocSearchClear.addEventListener('click', (e) => {
                e.stopPropagation();
                payDocCreditorFilter.value = '';
                payDocSearchClear.hidden = true;
                filterPayDocCreditorDropdown('');
                payDocCreditorFilter.focus();
                updateDateSummary();
            });
        }

        if (payDocClearSelection) {
            payDocClearSelection.addEventListener('click', () => {
                selectedPayDocCreditors.clear();
                document.querySelectorAll('.paydoc-creditor-checkbox').forEach(cb => cb.checked = false);
                updatePayDocSelectedChips();
                updateDateSummary();
            });
        }

        // Ensure outside clicks reliably close all specific dropdowns
        document.addEventListener('click', (e) => {
            const creditorSearch = document.querySelector('.filter-group.header-search');
            const overdueSearch = document.querySelector('.filter-group.header-overdue');
            const dateSearchArea = document.querySelector('.date-search.multi-select-creditor');

            if (creditorSearch && !creditorSearch.contains(e.target)) {
                const dd = document.getElementById('creditorDropdown');
                if (dd) dd.hidden = true;
            }
            if (overdueSearch && !overdueSearch.contains(e.target)) {
                const dd = document.getElementById('overdueDropdown');
                if (dd) dd.hidden = true;
            }
            if (dateSearchArea && !dateSearchArea.contains(e.target)) {
                const dd = document.getElementById('payDocCreditorDropdown');
                if (dd) dd.hidden = true;
            }
            updateHeaderSpacing();
        });
    }

    if (payDocUrgencyFilter) {
        payDocUrgencyFilter.addEventListener('change', updateDateSummary);
    }
    const payDocCategoryFilter = document.getElementById('payDocCategoryFilter');
    if (payDocCategoryFilter) {
        payDocCategoryFilter.addEventListener('change', updateDateSummary);
    }

    // Advanced search segmented control and filter panel
    const segButtons = document.querySelectorAll('.seg-btn');
    segButtons.forEach(btn => {
        btn.addEventListener('click', (e) => {
            segButtons.forEach(s => s.classList.remove('is-active'));
            segButtons.forEach(s => s.setAttribute('aria-pressed', 'false'));
            e.currentTarget.classList.add('is-active');
            e.currentTarget.setAttribute('aria-pressed', 'true');
            // future: switch search mode
        });
    });

    // Details modal view toggle (items vs grouped) - buttons are present in DOM
    const btnViewItems = document.getElementById('btnViewItems');
    const btnViewGrouped = document.getElementById('btnViewGrouped');
    const groupSummaryBar = document.getElementById('groupSummaryBar');
    const tableResp = document.querySelector('.table-responsive');
    if (btnViewItems && btnViewGrouped && groupSummaryBar && tableResp) {
        btnViewItems.addEventListener('click', () => {
            btnViewItems.classList.add('is-active');
            btnViewGrouped.classList.remove('is-active');
            populateDetailsTable(lastDetailsItems);
            groupSummaryBar.hidden = false;
        });
        btnViewGrouped.addEventListener('click', () => {
            btnViewGrouped.classList.add('is-active');
            btnViewItems.classList.remove('is-active');
            populateGroupedDetailsTable(lastDetailsItems);
        });
    }

    // Date Detail Modal view toggle
    const btnDateViewItems = document.getElementById('btnDateViewItems');
    const btnDateViewGrouped = document.getElementById('btnDateViewGrouped');
    const groupSummaryBarDate = document.getElementById('groupSummaryBarDate');
    if (btnDateViewItems && btnDateViewGrouped && groupSummaryBarDate) {
        btnDateViewItems.addEventListener('click', () => {
            btnDateViewItems.classList.add('is-active');
            btnDateViewGrouped.classList.remove('is-active');
            populateDetailsTable(lastDateDetailItems, 'dateDetailTableBody', 'dateDetailTableFooter');
            groupSummaryBarDate.hidden = false;
        });
        btnDateViewGrouped.addEventListener('click', () => {
            btnDateViewGrouped.classList.add('is-active');
            btnDateViewItems.classList.remove('is-active');
            populateGroupedDetailsTable(lastDateDetailItems, 'dateDetailTableBody', 'dateDetailTableFooter', 'btnDateViewItems');
        });
    }

    const advFilterBtn = document.getElementById('advFilterBtn');
    const advFilterPanel = document.getElementById('advFilterPanel');
    const advFilterClose = advFilterPanel && advFilterPanel.querySelector('.btn-close');
    if (advFilterBtn && advFilterPanel) {
        advFilterBtn.addEventListener('click', (e) => {
            const expanded = advFilterBtn.getAttribute('aria-expanded') === 'true';
            advFilterBtn.setAttribute('aria-expanded', String(!expanded));
            advFilterPanel.hidden = expanded; // if expanded true -> hide
            updateHeaderSpacing();
        });
        if (advFilterClose) {
            advFilterClose.addEventListener('click', () => {
                advFilterPanel.hidden = true;
                advFilterBtn.setAttribute('aria-expanded', 'false');
                updateHeaderSpacing();
            });
        }
        // apply / reset buttons
        const applyBtn = advFilterPanel.querySelector('.btn-apply');
        const resetBtn = advFilterPanel.querySelector('.btn-reset');
        if (applyBtn) applyBtn.addEventListener('click', () => {
            // apply filters - for now, close panel and trigger dashboard update
            advFilterPanel.hidden = true;
            advFilterBtn.setAttribute('aria-expanded', 'false');
            updateHeaderSpacing();
            updateDashboard({ skipDateSummary: true });
        });
        if (resetBtn) resetBtn.addEventListener('click', () => {
            advFilterPanel.querySelectorAll('select, input').forEach(i => {
                if (i.type === 'checkbox' || i.type === 'radio') i.checked = false;
                else if (i.type === 'range') i.value = i.min || 0;
                else i.value = '';
            });
        });
    }

    refreshBtn.addEventListener('click', async () => {
        if (GOOGLE_APP_SCRIPT_URL === 'YOUR_WEB_APP_URL_HERE') {
            alert('กรุณาใส่ Web App URL ของคุณในไฟล์ script.js ก่อนครับ');
        } else {
            await fetchData();
        }
    });

    // Open Google Sheets button (direct to Spreadsheet UI if configured,
    // otherwise fallback to Apps Script web app endpoint or show setup modal)
    // (header openSheetBtn removed) — floating FAB below still handles opening the sheet
    // Floating button (bottom-right) opens the same target as header button
    const floatingOpenSheetBtn = document.getElementById('floatingOpenSheetBtn');
    if (floatingOpenSheetBtn) {
        floatingOpenSheetBtn.addEventListener('click', (e) => {
            if (GOOGLE_SHEET_URL && GOOGLE_SHEET_URL.trim() !== '') {
                window.open(GOOGLE_SHEET_URL, '_blank', 'noopener');
                return;
            }
            if (GOOGLE_APP_SCRIPT_URL && GOOGLE_APP_SCRIPT_URL !== 'YOUR_WEB_APP_URL_HERE') {
                window.open(GOOGLE_APP_SCRIPT_URL, '_blank', 'noopener');
                return;
            }
            alert('กรุณาตั้งค่า `GOOGLE_SHEET_URL` หรือ `GOOGLE_APP_SCRIPT_URL` ในไฟล์ script.js ก่อนครับ');
            const modal = document.getElementById('setupModal');
            if (modal) modal.classList.remove('hidden');
        });
    }

    // Modal Close
    const modal = document.getElementById('setupModal');
    const closeBtn = document.querySelector('.close-btn');
    const closeBtn2 = document.getElementById('closeModalBtn');

    closeBtn.addEventListener('click', () => modal.classList.add('hidden'));
    closeBtn2.addEventListener('click', () => modal.classList.add('hidden'));

    // Details Modal Setup
    const detailsModal = document.getElementById('detailsModal');
    const closeDetailsBtn = document.querySelector('.close-details-btn');
    const closeDetailsModalBtn = document.getElementById('closeDetailsModalBtn');

    closeDetailsBtn.addEventListener('click', () => detailsModal.classList.add('hidden'));
    closeDetailsModalBtn.addEventListener('click', () => detailsModal.classList.add('hidden'));

    // Date Detail Modal Setup
    const dateDetailModal = document.getElementById('dateDetailModal');
    const closeDateDetailBtn = document.querySelector('.close-date-detail-btn');
    const closeDateDetailModalBtn = document.getElementById('closeDateDetailModalBtn');
    const exportDatePdfBtn = document.getElementById('exportDatePdfBtn');

    if (closeDateDetailBtn) closeDateDetailBtn.addEventListener('click', () => dateDetailModal.classList.add('hidden'));
    if (closeDateDetailModalBtn) closeDateDetailModalBtn.addEventListener('click', () => dateDetailModal.classList.add('hidden'));
    if (exportDatePdfBtn) exportDatePdfBtn.addEventListener('click', exportDatePDF);

    // PDF Export
    if (exportPdfBtn) {
        exportPdfBtn.addEventListener('click', exportToPDF);
    }

    // Details Buttons
    document.querySelectorAll('.details-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            const statusType = e.currentTarget.getAttribute('data-status');
            openDetailsModal(statusType);
        });
    });

    window.addEventListener('click', (e) => {
        if (e.target === modal) modal.classList.add('hidden');
        if (e.target === detailsModal) detailsModal.classList.add('hidden');
        if (e.target === dateDetailModal) dateDetailModal.classList.add('hidden');
    });

};

// Open details Modal and populate table
const openDetailsModal = (type) => {
    const detailsModal = document.getElementById('detailsModal');
    const detailsModalTitle = document.getElementById('detailsModalTitle');
    const detailsTableBody = document.getElementById('detailsTableBody');

    // Define status mappings
    const statusMap = {
        'overdue': { text: 'จ่ายเกินกำหนด', key: 'เกินกำหนด' },
        'ontime': { text: 'จ่ายตรงดิว', key: 'ตรงดิว' },
        'notdue': { text: 'ยังไม่ถึงกำหนด', key: 'ยังไม่ถึงกำหนด' },
        'nodate': { text: 'ยังไม่กำหนดวันจ่าย', key: 'ยังไม่กำหนดวันจ่าย' },
        'early': { text: 'จ่ายก่อนกำหนด', key: 'จ่ายก่อนกำหนด' },
        'pending': { text: 'เกินกำหนด (รอพิจารณา)', key: 'เกินกำหนด (รอพิจารณา)' }
    };

    const config = statusMap[type];
    if (!config) return;

    // Filter and sort data based on current context
    const items = currentFilteredData.filter(item => {
        const s = (item.status || '').toString().trim();
        // ต้องแยก 'เกินกำหนด' ออกจาก 'เกินกำหนด (รอพิจารณา)' เพื่อให้ยอดตรงกับ Card หน้าหลัก
        if (type === 'overdue') {
            return s.includes('เกินกำหนด') && !s.includes('(รอพิจารณา)');
        }
        return s.includes(config.key);
    }).sort((a, b) => {
        // First sort by date (ascending)
        const yearA = parseInt(a.yearDue) || 9999;
        const yearB = parseInt(b.yearDue) || 9999;
        if (yearA !== yearB) return yearA - yearB;

        const monthA = monthMap[a.monthDue] || 99;
        const monthB = monthMap[b.monthDue] || 99;
        if (monthA !== monthB) return monthA - monthB;

        const dayA = parseInt(a.dayDue) || 99;
        const dayB = parseInt(b.dayDue) || 99;
        if (dayA !== dayB) return dayA - dayB;

        // If date is equal, sort by amount (descending)
        return (Number(b.amount) || 0) - (Number(a.amount) || 0);
    });

    // Save last items for re-render / group interactions
    lastDetailsItems = items.slice();

    // Reset view toggle to 'Items' by default
    const btnViewItems = document.getElementById('btnViewItems');
    const btnViewGrouped = document.getElementById('btnViewGrouped');
    if (btnViewItems && btnViewGrouped) {
        btnViewItems.classList.add('is-active');
        btnViewGrouped.classList.remove('is-active');
    }

    // Render grouped summary and full table
    renderGroupSummary(items);
    populateDetailsTable(items);

    // Update Title with Total Sum for immediate clarity (modal shows total,
    // but the print header should not include the total amount)
    const totalSum = items.reduce((s, it) => s + (Number(it.amount) || 0), 0);
    const finalTitle = `ประเภทรายงาน: ${config.text} (ยอดรวมทั้งหมด: ${formatCurrency(totalSum)})`;
    detailsModalTitle.innerText = finalTitle;

    // Append the Category to the Print Header Title if it is not "all"
    const catVal = document.getElementById('categoryFilter') ? document.getElementById('categoryFilter').value : 'all';
    const catText = catVal !== 'all' ? ` - ${catVal}` : '';
    const printHeader = document.getElementById('printReportHeaderGlobal');
    if (printHeader) printHeader.innerText = `ประเภทรายงาน: ${config.text}${catText}`;

    detailsModal.classList.remove('hidden');
    // After modal is visible, shrink any long status texts to fit their cells
    window.requestAnimationFrame(() => shrinkStatusTextToFit(detailsTableBody));
};

// Initializing empty charts
const initCharts = () => {
    // 1. Donut Chart (Status)
    const ctxStatus = document.getElementById('statusChart').getContext('2d');

    // Shared styling properties
    Chart.defaults.color = '#8e8e9e';
    Chart.defaults.font.family = "'Prompt', sans-serif";

    donutChart = new Chart(ctxStatus, {
        type: 'doughnut',
        data: {
            labels: ['จ่ายเกินกำหนด', 'จ่ายตรงดิว', 'ยังไม่ถึงกำหนด', 'ยังไม่กำหนดวันจ่าย', 'จ่ายก่อนกำหนด', 'เกินกำหนด (รอพิจารณา)'],
            datasets: [{
                data: [],
                backgroundColor: ['#ef4444', '#10b981', '#3b82f6', '#f59e0b', '#a855f7', '#f97316'],
                borderWidth: 0,
                hoverOffset: 12
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            cutout: '70%',
            plugins: {
                legend: {
                    position: 'bottom',
                    labels: {
                        color: 'rgba(255,255,255,0.7)',
                        usePointStyle: true,
                        padding: 20,
                        font: { family: 'Inter, Prompt, sans-serif', size: 12 }
                    }
                },
                tooltip: {
                    backgroundColor: 'rgba(15, 15, 25, 0.95)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    titleFont: { size: 14, weight: 'bold', family: 'Inter, Prompt, sans-serif' },
                    bodyFont: { size: 13, family: 'Inter, Prompt, sans-serif' },
                    padding: 12,
                    cornerRadius: 10,
                    borderColor: 'rgba(99, 102, 241, 0.3)',
                    borderWidth: 1,
                    displayColors: true,
                    boxPadding: 6,
                    callbacks: {
                        label: function (context) {
                            const total = context.dataset.data.reduce((a, b) => a + b, 0);
                            const val = context.raw;
                            const pct = total === 0 ? 0 : ((val / total) * 100).toFixed(2);
                            return [
                                ` ประเภท: ${context.label}`,
                                ` ยอดรวม: ${formatCurrency(val)}`,
                                ` สัดส่วน: ${pct}%`
                            ];
                        }
                    }
                }
            }
        }
    });

    // 2. Bar Chart (Top Expenses by Creditor/Category)
    const ctxCategory = document.getElementById('categoryChart').getContext('2d');

    // Custom inline plugin for bar value labels
    const barValueLabels = {
        id: 'barValueLabels',
        afterDatasetsDraw(chart) {
            const { ctx, data } = chart;
            data.datasets.forEach((dataset, datasetIndex) => {
                const meta = chart.getDatasetMeta(datasetIndex);
                meta.data.forEach((bar, index) => {
                    const value = dataset.data[index];
                    if (value === undefined || value === null || value === 0) return;

                    let label;
                    if (value >= 1000000) label = '\u0e3f' + (value / 1000000).toFixed(2) + 'M';
                    else if (value >= 1000) label = '\u0e3f' + (value / 1000).toFixed(1) + 'K';
                    else label = '\u0e3f' + value.toLocaleString();

                    const x = bar.x;
                    const y = bar.y - 12;

                    ctx.save();
                    ctx.font = 'bold 11px Inter, sans-serif';
                    const textWidth = ctx.measureText(label).width;
                    const padX = 8, padY = 4;
                    const pillW = textWidth + padX * 2;
                    const pillH = 22;
                    const pillX = x - pillW / 2;
                    const pillY = y - pillH;

                    ctx.beginPath();
                    ctx.roundRect(pillX, pillY, pillW, pillH, 6);
                    ctx.fillStyle = 'rgba(139, 92, 246, 0.9)';
                    ctx.shadowColor = 'rgba(139, 92, 246, 0.5)';
                    ctx.shadowBlur = 8;
                    ctx.fill();
                    ctx.shadowBlur = 0;

                    ctx.fillStyle = '#ffffff';
                    ctx.textAlign = 'center';
                    ctx.textBaseline = 'middle';
                    ctx.fillText(label, x, pillY + pillH / 2);
                    ctx.restore();
                });
            });
        }
    };

    barChart = new Chart(ctxCategory, {
        type: 'bar',
        data: {
            labels: [],
            datasets: [{
                label: '\u0e22\u0e2d\u0e14\u0e43\u0e0a\u0e49\u0e08\u0e48\u0e32\u0e22 (\u0e1a\u0e32\u0e17)',
                data: [],
                backgroundColor: function (context) {
                    const chart = context.chart;
                    const { ctx, chartArea } = chart;
                    if (!chartArea) return 'rgba(99, 102, 241, 0.8)';
                    const gradient = ctx.createLinearGradient(0, chartArea.top, 0, chartArea.bottom);
                    gradient.addColorStop(0, 'rgba(167, 139, 250, 0.95)');
                    gradient.addColorStop(1, 'rgba(99, 102, 241, 0.55)');
                    return gradient;
                },
                borderRadius: 8,
                borderSkipped: false,
                hoverBackgroundColor: 'rgba(192, 132, 252, 1)'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 44 } },
            scales: {
                y: {
                    beginAtZero: true,
                    grid: { color: 'rgba(255, 255, 255, 0.05)' },
                    ticks: {
                        color: 'rgba(255,255,255,0.45)',
                        callback: function (value) {
                            if (value >= 1000000) return '\u0e3f' + (value / 1000000).toFixed(1) + 'M';
                            if (value >= 1000) return '\u0e3f' + (value / 1000).toFixed(0) + 'K';
                            return '\u0e3f' + value;
                        }
                    }
                },
                x: {
                    grid: { display: false },
                    ticks: { color: 'rgba(255,255,255,0.6)', maxRotation: 30 }
                }
            },
            plugins: {
                legend: { display: false },
                tooltip: {
                    backgroundColor: 'rgba(15, 15, 25, 0.95)',
                    titleColor: '#fff',
                    bodyColor: '#fff',
                    titleFont: { size: 14, weight: 'bold', family: 'Inter, Prompt, sans-serif' },
                    bodyFont: { size: 13, family: 'Inter, Prompt, sans-serif' },
                    padding: 12,
                    cornerRadius: 10,
                    borderColor: 'rgba(99, 102, 241, 0.3)',
                    borderWidth: 1,
                    callbacks: {
                        title: function (context) {
                            return '👤 ชื่อเจ้าหนี้: ' + context[0].label;
                        },
                        label: function (context) {
                            const dataset = context.dataset;
                            const statusStr = dataset.statusData ? dataset.statusData[context.dataIndex] : 'ไม่ทราบกลุ่ม';
                            return [
                                ' 📦 ข้อมูลจากกลุ่ม: ' + statusStr,
                                ' 💰 ยอดเงิน: ' + formatCurrency(context.raw)
                            ];
                        }
                    }
                }
            }
        },
        plugins: [barValueLabels]
    });
};

// Fetch data from Google Sheets API
const fetchData = async () => {
    loading.classList.remove('hidden');
    try {
        const response = await fetch(GOOGLE_APP_SCRIPT_URL);
        const result = await response.json();


        if (result.status === 'success') {
            allData = result.data;

            // Populate Creditor Datalist + custom dropdown
            const creditors = [...new Set(allData.map(item => item.docNo))].filter(Boolean).sort();
            const datalist = document.getElementById('creditorList');
            if (datalist) {
                datalist.innerHTML = creditors.map(c => `<option value="${c}">`).join('');
            }
            creditorData = creditors;
            renderCreditorDropdown(creditors);

            // Populate bottom multi-select creditor dropdown and urgency filter
            renderPayDocCreditorDropdown(creditors);
            populateUrgencyDropdown(allData);

            updateDashboard();
            document.body.classList.add('dashboard-ready');
        } else {
            console.error('API Error:', result.message);
            alert('เกิดข้อผิดพลาดในการดึงข้อมูลจาก Google Sheets: ' + result.message);
        }
    } catch (error) {
        console.error('Fetch Error:', error);
        alert('เกิดข้อผิดพลาดในการเชื่อมต่อ กรุณาตรวจสอบ URL ของ Web App');
    } finally {
        loading.classList.add('hidden');
    }
};

// Populate the bottom 'Urgency/Priority' dropdown dynamically
function populateUrgencyDropdown(data) {
    const urgencyFilter = document.getElementById('payDocUrgencyFilter');
    if (!urgencyFilter) return;

    // Get unique values from column N (item.status)
    const uniqueValues = [...new Set(data.map(item => (item.status || '').toString().trim()))]
        .filter(Boolean)
        .sort();

    // Preserve 'all' option
    urgencyFilter.innerHTML = '<option value="all">ทั้งหมด</option>';
    uniqueValues.forEach(val => {
        const opt = document.createElement('option');
        opt.value = val;
        opt.textContent = val;
        urgencyFilter.appendChild(opt);
    });
}


// ---- Creditor dropdown helpers ----
function escapeHtml(str) {
    return String(str).replace(/[&<>"']/g, function (s) {
        return ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[s];
    });
}

function renderCreditorDropdown(list = [], query = '') {
    const panel = document.getElementById('creditorListPanel');
    if (!panel) return;
    panel.innerHTML = '';
    if (!list || list.length === 0) {
        panel.innerHTML = '<div class="creditor-empty"><i class="bx bx-search-alt" style="font-size:24px;display:block;margin-bottom:6px;opacity:0.4;"></i>ไม่พบชื่อเจ้าหนี้</div>';
        return;
    }
    const q = (query || '').toString().trim().toLowerCase();
    const frag = document.createDocumentFragment();
    list.forEach((name, idx) => {
        const label = document.createElement('label');
        label.className = 'creditor-item';
        if (selectedCreditors.has(name)) label.classList.add('is-checked');

        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.className = 'creditor-checkbox';
        chk.dataset.name = name;
        chk.id = `creditor_chk_${idx}`;
        if (selectedCreditors.has(name)) chk.checked = true;

        // Custom checkbox visual element
        const chkVisual = document.createElement('span');
        chkVisual.className = 'checkbox-visual';

        const span = document.createElement('span');
        span.className = 'creditor-name';

        if (q.length > 0) {
            const low = (name || '').toString().toLowerCase();
            const i = low.indexOf(q);
            if (i >= 0) {
                const before = escapeHtml(name.slice(0, i));
                const matchText = escapeHtml(name.slice(i, i + q.length));
                const after = escapeHtml(name.slice(i + q.length));
                span.innerHTML = `${before}<span class="match">${matchText}</span>${after}`;
            } else {
                span.textContent = name;
            }
        } else {
            span.textContent = name;
        }

        label.appendChild(chk);
        label.appendChild(chkVisual);
        label.appendChild(span);
        frag.appendChild(label);
    });
    panel.appendChild(frag);

    // wire checkbox change events
    panel.querySelectorAll('.creditor-checkbox').forEach(cb => {
        cb.addEventListener('change', (e) => {
            const nm = e.target.dataset.name;
            const parentLabel = e.target.closest('.creditor-item');
            if (e.target.checked) {
                selectedCreditors.add(nm);
                if (parentLabel) parentLabel.classList.add('is-checked');
            } else {
                selectedCreditors.delete(nm);
                if (parentLabel) parentLabel.classList.remove('is-checked');
            }
            updateSelectedChips();
            updateCreditorCountBadge('creditorDropdown');
            updateDashboard({ skipDateSummary: true });
        });
    });

    // Update count badge
    updateCreditorCountBadge('creditorDropdown');
}

// Update selected count badge in a given dropdown
function updateCreditorCountBadge(dropdownId) {
    const dropdown = document.getElementById(dropdownId);
    if (!dropdown) return;
    const header = dropdown.querySelector('.creditor-dropdown-header');
    if (!header) return;

    let badge = header.querySelector('.selected-count-badge');
    let count = 0;
    if (dropdownId === 'creditorDropdown') count = selectedCreditors.size;
    else if (dropdownId === 'payDocCreditorDropdown') count = selectedPayDocCreditors.size;

    if (count > 0) {
        if (!badge) {
            badge = document.createElement('span');
            badge.className = 'selected-count-badge';
            const strong = header.querySelector('strong');
            if (strong) strong.after(badge);
            else header.prepend(badge);
        }
        badge.textContent = `เลือก ${count} รายการ`;
    } else {
        if (badge) badge.remove();
    }
}

function filterCreditorDropdown(q = '', forceOpen = false) {
    const query = (q || '').toString().trim().toLowerCase();
    const filtered = creditorData.filter(n => n.toLowerCase().includes(query));
    renderCreditorDropdown(filtered, q);
    const dropdown = document.getElementById('creditorDropdown');
    if (!dropdown) return;
    // Only open the dropdown when forced (user clicked) or when user typed something (query length > 0)
    if (forceOpen || query.length > 0) {
        closeAllDropdowns('creditorDropdown'); // Close others
        dropdown.hidden = false;
    } else {
        dropdown.hidden = true;
    }
    // adjust header spacing to avoid overlapping content
    updateHeaderSpacing();
}

function updateSelectedChips() {
    const container = document.getElementById('selectedChips');
    if (!container) return;
    container.innerHTML = '';
    if (selectedCreditors.size === 0) {
        container.hidden = true;
        return;
    }
    container.hidden = false;
    selectedCreditors.forEach(name => {
        const chip = document.createElement('span');
        chip.className = 'chip';
        const text = document.createElement('span');
        text.className = 'chip-name';
        text.textContent = name;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'chip-remove';
        btn.setAttribute('aria-label', `ลบ ${name}`);
        btn.textContent = '✕';
        btn.addEventListener('click', () => {
            selectedCreditors.delete(name);
            const allChecks = Array.from(document.querySelectorAll('.creditor-checkbox'));
            const match = allChecks.find(c => c.dataset.name === name);
            if (match) match.checked = false;
            updateSelectedChips();
            updateDashboard({ skipDateSummary: true });
        });
        chip.appendChild(text);
        chip.appendChild(btn);
        container.appendChild(chip);
    });
}

function openCreditorDropdown() {
    const dd = document.getElementById('creditorDropdown');
    if (dd) dd.hidden = false;
    updateHeaderSpacing();
}

function closeCreditorDropdown() {
    const dd = document.getElementById('creditorDropdown');
    if (dd) dd.hidden = true;
    updateHeaderSpacing();
}

// Helper to close all dropdowns (mutual exclusivity)
function closeAllDropdowns(exceptId = '') {
    // Current dropdown and panel IDs that should be mutually exclusive
    const dropdownIds = [
        'creditorDropdown',
        'overdueDropdown',
        'payDocCreditorDropdown',
        'advFilterPanel' // Including advanced filter panel if it exists
    ];

    dropdownIds.forEach(id => {
        if (id !== exceptId) {
            const dd = document.getElementById(id);
            if (dd) {
                // Handle both .hidden and manual .style.display if necessary, 
                // but usually 'hidden' attribute is used here.
                dd.hidden = true;

                // If it's the advanced filter button, sync its aria-expanded state
                if (id === 'advFilterPanel') {
                    const btn = document.getElementById('advFilterBtn');
                    if (btn) btn.setAttribute('aria-expanded', 'false');
                }
            }
        }
    });
    updateHeaderSpacing();
}

// ---- Date Summary Creditor dropdown helpers ----
function renderPayDocCreditorDropdown(list = [], query = '') {
    const panel = document.getElementById('payDocCreditorListPanel');
    if (!panel) return;
    panel.innerHTML = '';
    if (!list || list.length === 0) {
        panel.innerHTML = '<div class="creditor-empty"><i class="bx bx-search-alt" style="font-size:24px;display:block;margin-bottom:6px;opacity:0.4;"></i>ไม่พบชื่อเจ้าหนี้</div>';
        return;
    }
    const q = (query || '').toString().trim().toLowerCase();
    const frag = document.createDocumentFragment();
    list.forEach((name, idx) => {
        const label = document.createElement('label');
        label.className = 'creditor-item';
        if (selectedPayDocCreditors.has(name)) label.classList.add('is-checked');

        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.className = 'paydoc-creditor-checkbox';
        chk.dataset.name = name;
        chk.id = `paydoc_creditor_chk_${idx}`;
        if (selectedPayDocCreditors.has(name)) chk.checked = true;

        // Custom checkbox visual element
        const chkVisual = document.createElement('span');
        chkVisual.className = 'checkbox-visual';

        const span = document.createElement('span');
        span.className = 'creditor-name';

        if (q.length > 0) {
            const low = (name || '').toString().toLowerCase();
            const i = low.indexOf(q);
            if (i >= 0) {
                const before = escapeHtml(name.slice(0, i));
                const matchText = escapeHtml(name.slice(i, i + q.length));
                const after = escapeHtml(name.slice(i + q.length));
                span.innerHTML = `${before}<span class="match">${matchText}</span>${after}`;
            } else {
                span.textContent = name;
            }
        } else {
            span.textContent = name;
        }

        label.appendChild(chk);
        label.appendChild(chkVisual);
        label.appendChild(span);
        frag.appendChild(label);
    });
    panel.appendChild(frag);

    // wire checkbox change events
    panel.querySelectorAll('.paydoc-creditor-checkbox').forEach(cb => {
        cb.addEventListener('change', (e) => {
            const nm = e.target.dataset.name;
            const parentLabel = e.target.closest('.creditor-item');
            if (e.target.checked) {
                selectedPayDocCreditors.add(nm);
                if (parentLabel) parentLabel.classList.add('is-checked');
            } else {
                selectedPayDocCreditors.delete(nm);
                if (parentLabel) parentLabel.classList.remove('is-checked');
            }
            updatePayDocSelectedChips();
            updateCreditorCountBadge('payDocCreditorDropdown');
            updateDateSummary();
        });
    });

    // Update count badge
    updateCreditorCountBadge('payDocCreditorDropdown');
}

function filterPayDocCreditorDropdown(q = '', forceOpen = false) {
    const query = (q || '').toString().trim().toLowerCase();
    const filtered = creditorData.filter(n => n.toLowerCase().includes(query));
    renderPayDocCreditorDropdown(filtered, q);
    const dropdown = document.getElementById('payDocCreditorDropdown');
    if (!dropdown) return;
    // Only open the dropdown when forced (user clicked) or when user typed something (query length > 0)
    if (forceOpen || query.length > 0) {
        closeAllDropdowns('payDocCreditorDropdown'); // Close others
        dropdown.hidden = false;
    } else {
        dropdown.hidden = true;
    }
}

function updatePayDocSelectedChips() {
    const container = document.getElementById('selectedPayDocChips');
    if (!container) return;
    container.innerHTML = '';
    if (selectedPayDocCreditors.size === 0) {
        container.hidden = true;
        return;
    }
    container.hidden = false;
    selectedPayDocCreditors.forEach(name => {
        const chip = document.createElement('span');
        chip.className = 'chip';
        const text = document.createElement('span');
        text.className = 'chip-name';
        text.textContent = name;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'chip-remove';
        btn.setAttribute('aria-label', `ลบ ${name}`);
        btn.textContent = '✕';
        btn.addEventListener('click', () => {
            selectedPayDocCreditors.delete(name);
            const allChecks = Array.from(document.querySelectorAll('.paydoc-creditor-checkbox'));
            const match = allChecks.find(c => c.dataset.name === name);
            if (match) match.checked = false;
            updatePayDocSelectedChips();
            updateDateSummary();
        });
        chip.appendChild(text);
        chip.appendChild(btn);
        container.appendChild(chip);
    });
}

// ---- Overdue Range dropdown helpers ----
function renderOverdueDropdown() {
    const panel = document.getElementById('overdueListPanel');
    if (!panel) return;
    panel.innerHTML = '';

    const frag = document.createDocumentFragment();
    overdueRanges.forEach((rng, idx) => {
        const label = document.createElement('label');
        label.className = 'creditor-item';

        const chk = document.createElement('input');
        chk.type = 'checkbox';
        chk.className = 'overdue-checkbox';
        chk.dataset.label = rng.label;
        chk.id = `overdue_chk_${idx}`;
        if (selectedOverdueRanges.has(rng.label)) chk.checked = true;

        const span = document.createElement('span');
        span.className = 'creditor-name';
        span.textContent = rng.label + ' วัน';

        label.appendChild(chk);
        label.appendChild(span);
        frag.appendChild(label);
    });
    panel.appendChild(frag);

    // wire checkbox change events
    panel.querySelectorAll('.overdue-checkbox').forEach(cb => {
        cb.addEventListener('change', (e) => {
            const lbl = e.target.dataset.label;
            if (e.target.checked) selectedOverdueRanges.add(lbl);
            else selectedOverdueRanges.delete(lbl);
            updateSelectedOverdueChips();
            updateDashboard({ skipDateSummary: true });
        });
    });
}

function updateSelectedOverdueChips() {
    const container = document.getElementById('selectedOverdueChips');
    const input = document.getElementById('overdueRangeFilter');
    const clearBtn = document.querySelector('.overdue-clear');

    if (!container || !input) return;
    container.innerHTML = '';

    if (selectedOverdueRanges.size === 0) {
        container.hidden = true;
        input.placeholder = "ทุกระยะเวลา...";
        if (clearBtn) clearBtn.hidden = true;
        return;
    }

    container.hidden = false;
    input.placeholder = ""; // hide placeholder when chips are present
    if (clearBtn) clearBtn.hidden = false;

    selectedOverdueRanges.forEach(label => {
        const chip = document.createElement('span');
        chip.className = 'chip';
        const text = document.createElement('span');
        text.className = 'chip-name';
        text.textContent = label;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'chip-remove';
        btn.textContent = '✕';
        btn.addEventListener('click', (e) => {
            e.stopPropagation();
            selectedOverdueRanges.delete(label);
            const allChecks = Array.from(document.querySelectorAll('.overdue-checkbox'));
            const match = allChecks.find(c => c.dataset.label === label);
            if (match) match.checked = false;
            updateSelectedOverdueChips();
            updateDashboard({ skipDateSummary: true });
        });
        chip.appendChild(text);
        chip.appendChild(btn);
        container.appendChild(chip);
    });
}

function closeOverdueDropdown() {
    const dd = document.getElementById('overdueDropdown');
    if (dd) dd.hidden = true;
    updateHeaderSpacing();
}

function updateHeaderSpacing() {
    const panels = [
        document.getElementById('creditorDropdown'),
        document.getElementById('advFilterPanel'),
        document.getElementById('overdueDropdown')
    ].filter(Boolean);

    const topNav = document.querySelector('.top-nav');
    if (!topNav) return;

    // Use a small delay to ensure DOM dimensions are updated if needed
    window.requestAnimationFrame(() => {
        let maxH = 0;
        panels.forEach(p => {
            if (p && !p.hidden) {
                // Get the height including margins/borders
                const rect = p.getBoundingClientRect();
                if (rect.height > maxH) maxH = rect.height;
            }
        });

        if (maxH > 0) {
            // Add a bit of extra breathing room (24px)
            topNav.style.paddingBottom = (maxH + 24) + 'px';
            topNav.style.transition = 'padding-bottom 0.3s cubic-bezier(0.4, 0, 0.2, 1)';
        } else {
            topNav.style.paddingBottom = '';
        }
    });
}

// Update Dashboard View based on selected filters
const updateDashboard = (opts = {}) => {
    const skipDateSummary = opts && opts.skipDateSummary === true;
    const selectedPaymentStatus = paymentStatusFilter.value;
    const selectedCategory = categoryFilter.value;
    const selectedDay = dayFilter.value;
    const selectedMonth = monthFilter.value;
    const selectedYear = yearFilter.value;
    const creditorVal = document.getElementById('creditorFilter')?.value.toLowerCase() || '';

    // Filter data
    currentFilteredData = allData.filter(item => {
        let matchPaymentStatus = selectedPaymentStatus === 'all' || (item.paymentStatus && item.paymentStatus.toString().includes(selectedPaymentStatus));
        let matchCategory = selectedCategory === 'all' || (item.category && item.category.toString().includes(selectedCategory));
        let matchDay = selectedDay === 'all' || (item.dayDue && parseInt(item.dayDue) === parseInt(selectedDay));
        let matchMonth = selectedMonth === 'all' || (item.monthDue && item.monthDue.toString() === selectedMonth);
        let matchYear = selectedYear === 'all' || (item.yearDue && parseInt(item.yearDue) === parseInt(selectedYear));
        let matchCreditor;
        if (selectedCreditors && selectedCreditors.size > 0) {
            matchCreditor = selectedCreditors.has(item.docNo);
        } else {
            matchCreditor = creditorVal === '' || (item.docNo && item.docNo.toLowerCase().includes(creditorVal));
        }

        let matchOverdue = true;
        if (selectedOverdueRanges.size > 0) {
            matchOverdue = false;
            const days = parseFloat(item.overdueDays) || 0;

            // ใช้ค่าสัมบูรณ์ของวัน (ติดลบ) มาเทียบกับช่วงที่กำหนด เช่น -172 จะอยู่ในกลุ่ม 151-180 วัน
            if (days < 0) {
                const absDays = Math.abs(days);
                for (let rangeLabel of selectedOverdueRanges) {
                    const rng = overdueRanges.find(r => r.label === rangeLabel);
                    if (rng && absDays >= rng.min && absDays <= rng.max) {
                        matchOverdue = true;
                        break;
                    }
                }
            }
        }

        return matchPaymentStatus && matchCategory && matchDay && matchMonth && matchYear && matchCreditor && matchOverdue;
    });

    // Calculate Summary numbers
    let total = 0, overdue = 0, ontime = 0, notdue = 0, nodate = 0, early = 0, pending = 0;

    // Temporary object to group data for bar chart (By Creditor - ชื่อเจ้าหนี้การค้า)
    const creditorSummary = {};

    currentFilteredData.forEach(item => {
        // Convert string to number just in case
        const amount = Number(item.amount) || 0;

        // Sum total directly to ensure 100% accuracy with Google Sheets
        total += amount;

        // Status calculation (N column = ความเร่งด่วน: เกินกำหนด/ตรงดิว/ยังไม่ถึงกำหนด)
        const statusStr = (item.status || '').toString().trim();

        // Count strictly correctly to avoid overlap bugs
        if (statusStr.includes('เกินกำหนด (รอพิจารณา)')) {
            pending += amount;
        } else if (statusStr.includes('เกินกำหนด')) {
            overdue += amount;
        } else if (statusStr.includes('ตรงดิว')) {
            ontime += amount;
        } else if (statusStr.includes('ยังไม่ถึงกำหนด')) {
            notdue += amount;
        } else if (statusStr.includes('ยังไม่กำหนดวันจ่าย')) {
            nodate += amount;
        } else if (statusStr.includes('จ่ายก่อนกำหนด')) {
            early += amount;
        }

        // Group by creditor Name (stored in item.docNo after column swap) and track their statuses for the bar chart
        const creditorName = item.docNo ? item.docNo : 'ไม่ระบุชื่อ';
        if (!creditorSummary[creditorName]) {
            creditorSummary[creditorName] = { amount: 0, statuses: new Set() };
        }
        creditorSummary[creditorName].amount += amount;

        // Map status strings to short readable box names
        let boxName = "อื่น ๆ";
        if (statusStr.includes('เกินกำหนด (รอพิจารณา)')) boxName = 'เกินกำหนด (รอพิจารณา)';
        else if (statusStr.includes('เกินกำหนด')) boxName = 'จ่ายเกินกำหนด';
        else if (statusStr.includes('ตรงดิว')) boxName = 'จ่ายตรงดิว';
        else if (statusStr.includes('ยังไม่ถึงกำหนด')) boxName = 'ยังไม่ถึงกำหนด';
        else if (statusStr.includes('ยังไม่กำหนดวันจ่าย')) boxName = 'ยังไม่กำหนดวันจ่าย';
        else if (statusStr.includes('จ่ายก่อนกำหนด')) boxName = 'จ่ายก่อนกำหนด';

        creditorSummary[creditorName].statuses.add(boxName);
    });

    // Total is already calculated in the loop above to include all items correctly
    // total = overdue + ontime + notdue + nodate + early + pending;

    // Update Text Elements with Counting Animation
    animateValue(totalAmountEl, 0, total, 1200, true);
    animateValue(overdueAmountEl, 0, overdue, 1200, true);
    animateValue(ontimeAmountEl, 0, ontime, 1200, true);
    animateValue(notdueAmountEl, 0, notdue, 1200, true);
    animateValue(nodateAmountEl, 0, nodate, 1200, true);
    animateValue(earlyAmountEl, 0, early, 1200, true);
    animateValue(pendingAmountEl, 0, pending, 1200, true);

    // Update Percentages
    totalPercentEl.innerText = `คิดเป็น 100.00%`;

    const overduePct = total === 0 ? 0 : (overdue / total) * 100;
    const ontimePct = total === 0 ? 0 : (ontime / total) * 100;
    const notduePct = total === 0 ? 0 : (notdue / total) * 100;
    const nodatePct = total === 0 ? 0 : (nodate / total) * 100;
    const earlyPct = total === 0 ? 0 : (early / total) * 100;
    const pendingPct = total === 0 ? 0 : (pending / total) * 100;

    animateValue(overduePercentEl, 0, overduePct, 1200, false);
    animateValue(ontimePercentEl, 0, ontimePct, 1200, false);
    animateValue(notduePercentEl, 0, notduePct, 1200, false);
    animateValue(nodatePercentEl, 0, nodatePct, 1200, false);
    animateValue(earlyPercentEl, 0, earlyPct, 1200, false);
    animateValue(pendingPercentEl, 0, pendingPct, 1200, false);

    // Update Donut Chart
    donutChart.data.datasets[0].data = [overdue, ontime, notdue, nodate, early, pending];
    donutChart.update();

    // Prepare Bar Chart Data (Sort by Highest Amount & take top 10)
    const sortedCreditors = Object.entries(creditorSummary)
        .sort((a, b) => b[1].amount - a[1].amount)
        .slice(0, 10);

    barChart.data.labels = sortedCreditors.map(item => item[0]);
    // Save metadata in the dataset for tooltip access
    barChart.data.datasets[0].data = sortedCreditors.map(item => item[1].amount);
    barChart.data.datasets[0].statusData = sortedCreditors.map(item => Array.from(item[1].statuses).join(', '));
    barChart.update();

    // Update Date Summary Section (skip when searching per user request)
    if (!skipDateSummary) updateDateSummary();
};

// ==========================================
// MOCK DATA: For demonstration during setup
// ==========================================
const loadMockData = () => {
    setTimeout(() => {
        allData = [
            { creditor: "สมปอง เซอร์วิส", amount: 15000, status: "ตรงดิว", paymentStatus: "จ่ายแล้ว", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "เจริญ ฮาร์ดแวร์", amount: 8500, status: "ยังไม่ถึงกำหนด", paymentStatus: "รอโอน", category: "รายสัปดาห์", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "การไฟฟ้า", amount: 2300, status: "เกินกำหนด", paymentStatus: "รอโอน", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "A Plus Company", amount: 12293699, status: "ตรงดิว", paymentStatus: "รอโอน", category: "ลิสซิ่ง", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "ค่าเช่าสำนักงาน", amount: 20000, status: "ยังไม่ถึงกำหนด", paymentStatus: "จ่ายแล้ว", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() },
            { creditor: "ผู้รับเหมา กริช", amount: 12000, status: "เกินกำหนด", paymentStatus: "ยกเลิก", category: "รายสัปดาห์", monthDue: "มิ.ย.", yearDue: new Date().getFullYear() },
            { creditor: "สมปอง เซอร์วิส", amount: 7000, status: "ตรงดิว", paymentStatus: "ตัดเช็คผ่าน", category: "เจ้าหนี้รายเดือน", monthDue: "พ.ค.", yearDue: new Date().getFullYear() }
        ];

        // Populate creditor list for mock mode as well
        const creditors = [...new Set(allData.map(item => item.docNo))].filter(Boolean).sort();
        const datalist = document.getElementById('creditorList');
        if (datalist) datalist.innerHTML = creditors.map(c => `<option value="${c}">`).join('');
        creditorData = creditors;
        renderCreditorDropdown(creditors);

        loading.classList.add('hidden');
        updateDashboard();
        document.body.classList.add('dashboard-ready');
    }, 1000);
};

// ==========================================
// DATE SUMMARY - รวมจำนวนเงินตามวันที่ทำเอกสารจ่าย (คอลัมน์ H)
// ==========================================
// Helper to get filtered data for the PayDoc/Date Summary section
const getFilteredForPayDoc = () => {
    const payDocStatusVal = document.getElementById('payDocStatusFilter')?.value || 'รอโอน';
    const payDocMonthVal = document.getElementById('payDocMonthFilter')?.value || 'all';
    const payDocYearVal = document.getElementById('payDocYearFilter')?.value || 'all';
    const payDocCreditorVal = document.getElementById('payDocCreditorFilter')?.value.toLowerCase() || '';
    const payDocUrgencyVal = document.getElementById('payDocUrgencyFilter')?.value || 'all';
    const payDocCategoryVal = document.getElementById('payDocCategoryFilter')?.value || 'all';

    return allData.filter(item => {
        const matchStatus = payDocStatusVal === 'all' || (item.paymentStatus && item.paymentStatus.toString().includes(payDocStatusVal));
        const matchMonth = payDocMonthVal === 'all' || (item.payDocMonth && item.payDocMonth === payDocMonthVal);
        const matchYear = payDocYearVal === 'all' || (item.payDocYear && parseInt(item.payDocYear) === parseInt(payDocYearVal));

        // Creditor match: support both typing AND multi-select chips
        let matchCreditor = true;
        if (selectedPayDocCreditors.size > 0) {
            matchCreditor = selectedPayDocCreditors.has(item.docNo);
        } else {
            matchCreditor = payDocCreditorVal === '' ||
                (item.docNo && item.docNo.toLowerCase().includes(payDocCreditorVal)) ||
                (item.creditor && item.creditor.toLowerCase().includes(payDocCreditorVal));
        }

        const matchUrgency = payDocUrgencyVal === 'all' || (item.status && item.status.toString().trim() === payDocUrgencyVal);
        const matchCategory = payDocCategoryVal === 'all' || (item.category && item.category === payDocCategoryVal);

        return matchStatus && matchMonth && matchYear && matchCreditor && matchUrgency && matchCategory;
    });
};

// ==========================================
// DATE SUMMARY - รวมจำนวนเงินตามวันที่ทำเอกสารจ่าย (คอลัมน์ H)
// ==========================================
const updateDateSummary = () => {
    const grid = document.getElementById('dateSummaryGrid');
    if (!grid) return;

    // Filter from ALL data (independent from top filters) by paymentStatus + payDoc month/year
    // Also include new creditor search and urgency filters.
    const filteredForPayDoc = getFilteredForPayDoc();

    // Group data by payDoc date (column H)
    const dateGroups = {};
    filteredForPayDoc.forEach(item => {
        const day = item.payDocDay || '';
        const month = item.payDocMonth || '';
        const year = item.payDocYear || '';
        const dateKey = [day, month, year].filter(Boolean).join(' ') || 'ไม่ระบุวันที่';

        if (!dateGroups[dateKey]) {
            dateGroups[dateKey] = {
                items: [],
                total: 0,
                day: parseInt(day) || 0,
                monthNum: monthMap[month] || 0,
                year: parseInt(year) || 0,
                statuses: new Set()
            };
        }
        const amount = Number(item.amount) || 0;
        dateGroups[dateKey].items.push(item);
        dateGroups[dateKey].total += amount;

        // Track statuses
        const statusStr = (item.status || '').toString().trim();
        if (statusStr.includes('เกินกำหนด (รอพิจารณา)')) dateGroups[dateKey].statuses.add('pending');
        else if (statusStr.includes('เกินกำหนด')) dateGroups[dateKey].statuses.add('overdue');
        else if (statusStr.includes('ตรงดิว')) dateGroups[dateKey].statuses.add('ontime');
        else if (statusStr.includes('ยังไม่ถึงกำหนด')) dateGroups[dateKey].statuses.add('notdue');
        else if (statusStr.includes('ยังไม่กำหนดวันจ่าย')) dateGroups[dateKey].statuses.add('nodate');
        else if (statusStr.includes('จ่ายก่อนกำหนด')) dateGroups[dateKey].statuses.add('early');
    });

    // Sort by date
    const sortedDates = Object.entries(dateGroups).sort((a, b) => {
        const da = a[1], db = b[1];
        if (da.year !== db.year) return da.year - db.year;
        if (da.monthNum !== db.monthNum) return da.monthNum - db.monthNum;
        return da.day - db.day;
    });

    // Render cards
    grid.innerHTML = '';

    if (sortedDates.length === 0) {
        grid.innerHTML = `
            <div class="date-summary-empty">
                <i class='bx bx-calendar-x'></i>
                <p>ไม่มีข้อมูลสำหรับตัวกรองที่เลือก</p>
            </div>`;
        return;
    }

    sortedDates.forEach(([dateKey, group], index) => {
        const statusBadgesHtml = Array.from(group.statuses).map(s => {
            const labels = {
                'overdue': 'เกินกำหนด',
                'ontime': 'ตรงดิว',
                'notdue': 'ยังไม่ถึง',
                'pending': 'รอพิจารณา',
                'nodate': 'ไม่กำหนด',
                'early': 'ก่อนกำหนด'
            };
            return `<span class="date-status-mini ${s}">${labels[s] || s}</span>`;
        }).join('');

        const card = document.createElement('div');
        card.className = 'date-card';
        card.style.animationDelay = `${index * 0.06}s`;
        card.innerHTML = `
            <div class="date-card-header">
                <div class="date-icon"><i class='bx bx-calendar'></i></div>
                <div class="date-header-text">
                    <span class="day-label">${dateKey}</span>
                    <span class="item-count">${group.items.length} รายการ</span>
                </div>
            </div>
            <div class="date-card-body">
                <div class="date-card-amount">${formatCurrency(group.total)}</div>
            </div>
            <div class="date-card-footer">
                <div class="date-status-badges">${statusBadgesHtml}</div>
                <button class="date-action-view" data-date-key="${dateKey}">
                    <i class='bx bx-show'></i> ดูข้อมูลเพิ่มเติม
                </button>
            </div>
        `;
        grid.appendChild(card);
    });

    // Attach event listeners to the new buttons
    grid.querySelectorAll('.date-action-view').forEach(btn => {
        btn.addEventListener('click', () => {
            const dateKey = btn.getAttribute('data-date-key');
            openDateDetailModal(dateKey);
        });
    });
};

// Open date detail modal and optionally trigger PDF export
const openDateDetailModal = (dateKey) => {
    const modal = document.getElementById('dateDetailModal');
    const title = document.getElementById('dateDetailModalTitle');
    const tbody = document.getElementById('dateDetailTableBody');
    const tfoot = document.getElementById('dateDetailTableFooter');

    // Find matching items by payDoc date (column H) + status filter
    // Now using the centralized filtering logic to ensure consistency
    const items = getFilteredForPayDoc().filter(item => {
        const day = item.payDocDay || '';
        const month = item.payDocMonth || '';
        const year = item.payDocYear || '';
        const itemDateKey = [day, month, year].filter(Boolean).join(' ') || 'ไม่ระบุวันที่';
        return itemDateKey === dateKey;
    }).sort((a, b) => {
        // Sort by year (ascending)
        const yearA = parseInt(a.yearDue) || 9999;
        const yearB = parseInt(b.yearDue) || 9999;
        if (yearA !== yearB) return yearA - yearB;

        // Sort by month (ascending)
        const monthA = monthMap[a.monthDue] || 99;
        const monthB = monthMap[b.monthDue] || 99;
        if (monthA !== monthB) return monthA - monthB;

        // Sort by day (ascending)
        const dayA = parseInt(a.dayDue) || 99;
        const dayB = parseInt(b.dayDue) || 99;
        if (dayA !== dayB) return dayA - dayB;

        // If date is equal, sort by amount (descending)
        return (Number(b.amount) || 0) - (Number(a.amount) || 0);
    });

    // Save last items for re-render / group interactions
    lastDateDetailItems = items.slice();

    // Reset view toggle to 'Items' by default
    const btnDateViewItems = document.getElementById('btnDateViewItems');
    const btnDateViewGrouped = document.getElementById('btnDateViewGrouped');
    if (btnDateViewItems && btnDateViewGrouped) {
        btnDateViewItems.classList.add('is-active');
        btnDateViewGrouped.classList.remove('is-active');
    }

    // Render grouped summary and full table
    renderGroupSummary(items, 'groupSummaryBarDate', 'dateDetailTableBody', 'dateDetailTableFooter', 'btnDateViewItems');
    populateDetailsTable(items, 'dateDetailTableBody', 'dateDetailTableFooter');

    const totalSum = items.reduce((s, it) => s + (Number(it.amount) || 0), 0);

    // Get currently selected status and urgency text for the title
    const statusFilter = document.getElementById('payDocStatusFilter');
    const statusText = statusFilter ? statusFilter.options[statusFilter.selectedIndex].text : '';

    const urgencyFilter = document.getElementById('payDocUrgencyFilter');
    const urgencyText = urgencyFilter ? urgencyFilter.options[urgencyFilter.selectedIndex].text : '';

    let suffixParts = [];
    if (urgencyText && urgencyText !== 'ทั้งหมด') suffixParts.push(`ความเร่งด่วน: ${urgencyText}`);
    if (statusText && statusText !== 'ทั้งหมด') suffixParts.push(`สถานะ: ${statusText}`);

    const statusSuffix = suffixParts.length > 0 ? ` (${suffixParts.join(', ')})` : '';

    const finalTitleText = `รายละเอียดวันที่: ${dateKey}${statusSuffix} (รายการทั้งหมด: ${items.length}, ยอดรวม: ${formatCurrency(totalSum)})`;
    const printTitleText = `สรุปรายการเบิกจ่ายประจำวันที่: ${dateKey}${statusSuffix}`;

    title.innerText = finalTitleText;

    // For printing: Use a more structured, professional layout
    const pTitle = document.getElementById('printDateReportTitle');
    const pSubtitle = document.getElementById('printDateReportSubtitle');
    if (pTitle) pTitle.innerText = "สรุปรายการเบิกจ่ายรายวัน";
    if (pSubtitle) {
        let sub = `ประจำวันที่: <span class="print-date-val">${dateKey}</span>`;
        if (statusSuffix) sub += ` <span class="print-filter-badge">${statusSuffix}</span>`;
        pSubtitle.innerHTML = sub;
    }

    // Save for any other PDF header logic
    currentModalDate = `${dateKey}${statusSuffix}`;

    modal.classList.remove('hidden');
    // After modal is visible, shrink any long status texts to fit their cells
    window.requestAnimationFrame(() => shrinkStatusTextToFit(tbody));
};

// Export date detail to PDF
const exportDatePDF = () => {
    const now = new Date();
    const docId = `PAY-${now.getTime().toString().slice(-6)}`;
    const dateStr = formatThaiDateTime(now);

    const printDocId = document.getElementById('printDocIdGlobal');
    const printIssueDate = document.getElementById('printIssueDateGlobal');
    const printDocIdDate = document.getElementById('printDocIdDate');
    const printIssueDateDate = document.getElementById('printIssueDateDate');

    if (printDocId) printDocId.innerText = docId;
    if (printIssueDate) printIssueDate.innerText = dateStr;
    if (printDocIdDate) printDocIdDate.innerText = docId;
    if (printIssueDateDate) printIssueDateDate.innerText = dateStr;

    // Using native print for the current modal state (Items vs Grouped view)
    window.print();
};

// Start application
document.addEventListener('DOMContentLoaded', init);

// Export to PDF function (Using browser's native print for perfect Thai font rendering)
const exportToPDF = () => {
    // บันทึกข้อมูลเลขที่เอกสารและวันที่
    const now = new Date();
    const docId = `RT-${now.getTime().toString().slice(-6)}`;
    const dateStr = formatThaiDateTime(now);

    // อัปเดตข้อมูลลงในธาตุ HTML สำหรับหน้าพิมพ์
    const printDocId = document.getElementById('printDocIdGlobal');
    const printIssueDate = document.getElementById('printIssueDateGlobal');
    const printDocIdDetails = document.getElementById('printDocIdDetails');
    const printIssueDateDetails = document.getElementById('printIssueDateDetails');
    const localHeader = document.getElementById('printReportHeaderLocal');

    if (printDocId) printDocId.innerText = docId;
    if (printIssueDate) printIssueDate.innerText = dateStr;
    if (printDocIdDetails) printDocIdDetails.innerText = docId;
    if (printIssueDateDetails) printIssueDateDetails.innerText = dateStr;

    // Sync report type to local header
    const globalHeader = document.getElementById('printReportHeaderGlobal');
    if (globalHeader && localHeader) {
        localHeader.innerText = globalHeader.innerText;
    }

    // Native print is generally more reliable for multi-page reports
    window.print();
};

/** 
 * High Quality PDF Export (Using html2pdf)
 * ถ้าต้องการใช้ตัวนี้ ให้เปลี่ยน window.print() ด้านบนเป็น runHtml2Pdf('.print-report-container')
 */
function runHtml2Pdf(selector) {
    const element = document.querySelector(selector);
    if (!element) return;

    const opt = {
        margin: [10, 10, 15, 10],
        filename: `Expense_Report_${new Date().getTime()}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };
    html2pdf().set(opt).from(element).save();
}

/* ========== Shrink status text to fit logic ========== */
function shrinkStatusTextToFit(root = document) {
    try {
        const container = root || document;
        const els = container.querySelectorAll('.status-text');
        els.forEach(el => {
            // reset to computed baseline
            el.style.fontSize = '';
            // get computed base font size
            const computed = window.getComputedStyle(el);
            let base = parseFloat(computed.fontSize) || 12;
            const minFont = 9; // minimum readable font size
            let size = base;

            // ensure single-line measurement
            el.style.whiteSpace = 'nowrap';
            el.style.overflow = 'hidden';

            // if the element is not in the layout (display:none), skip
            if (el.offsetParent === null) return;

            // shrink step until it fits or reaches minFont
            while (el.scrollWidth > el.clientWidth && size > minFont) {
                size = Math.max(minFont, size - 0.5);
                el.style.fontSize = size + 'px';
            }
        });
    } catch (err) {
        console.warn('shrinkStatusTextToFit error', err);
    }
}

// Debounced resize handler
let __shrinkResizeTimer = null;
window.addEventListener('resize', () => {
    clearTimeout(__shrinkResizeTimer);
    __shrinkResizeTimer = setTimeout(() => shrinkStatusTextToFit(), 160);
});

// Before printing, clear inline font sizes so print CSS can wrap naturally.
window.addEventListener('beforeprint', () => {
    document.querySelectorAll('.status-text').forEach(el => el.style.fontSize = '');

    // Adjust colspan for Detailed Table because we hide the Category column (4th) in CSS during print
    const detailsTotalLabel = document.querySelector('#detailsTableFooter .total-label');
    if (detailsTotalLabel && detailsTotalLabel.getAttribute('colspan') === '6') {
        detailsTotalLabel.setAttribute('colspan', '5');
        detailsTotalLabel.dataset.changedForPrint = 'true';
    }

    // Adjust colspan for Date Detail Table
    const dateTotalLabel = document.querySelector('#dateDetailTableFooter .total-label');
    if (dateTotalLabel && dateTotalLabel.getAttribute('colspan') === '6') {
        dateTotalLabel.setAttribute('colspan', '5');
        dateTotalLabel.dataset.changedForPrint = 'true';
    }
});
window.addEventListener('afterprint', () => {
    // restore shrink after printing
    setTimeout(() => shrinkStatusTextToFit(), 80);

    // Restore colspan for details table
    const detailsTotalLabel = document.querySelector('#detailsTableFooter .total-label');
    if (detailsTotalLabel && detailsTotalLabel.dataset.changedForPrint === 'true') {
        detailsTotalLabel.setAttribute('colspan', '6');
        delete detailsTotalLabel.dataset.changedForPrint;
    }

    // Restore colspan for date detail table
    const dateTotalLabel = document.querySelector('#dateDetailTableFooter .total-label');
    if (dateTotalLabel && dateTotalLabel.dataset.changedForPrint === 'true') {
        dateTotalLabel.setAttribute('colspan', '6');
        delete dateTotalLabel.dataset.changedForPrint;
    }
});

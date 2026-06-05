/* ========================================
   飞利浦监护仪报警管理器 - 应用逻辑
   ======================================== */

// ==================== 常量配置 ====================

const CATEGORY_LABELS = {
    arrhythmia: '心律失常',
    pressure: '压力相关',
    respiratory: '呼吸',
    spo2: '血氧(SpO₂)',
    heartrate: '心率(HR)',
    st: 'ST段',
    temperature: '体温',
    technical: '技术报警',
    operation: '操作记录',
    settings: '设置变更',
    other: '其他'
};

const CATEGORY_COLORS = {
    arrhythmia: '#ef4444',
    pressure: '#f97316',
    respiratory: '#06b6d4',
    spo2: '#3d8bfd',
    heartrate: '#f59e0b',
    st: '#8b5cf6',
    temperature: '#ec4899',
    technical: '#6b7280',
    operation: '#10b981',
    settings: '#a3a3a3',
    other: '#525252'
};

const SEVERITY_LABELS = { 3: '危急 (***)', 2: '严重 (**)', 1: '警告 (*)', 0: '提示' };
const SEVERITY_COLORS = { 3: '#ef4444', 2: '#f97316', 1: '#f59e0b', 0: '#6b7280' };

const DAY_NAMES = ['周日', '周一', '周二', '周三', '周四', '周五', '周六'];

const CHART_TEXT_STYLE = { color: '#7e93aa', fontFamily: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif' };
const TOOLTIP_STYLE = { backgroundColor: '#1a3050', borderColor: '#1e3a5f', textStyle: { color: '#e2e8f0', fontSize: 13 } };
const AXIS_STYLE = { axisLine: { lineStyle: { color: '#1e3a5f' } }, splitLine: { lineStyle: { color: '#1e3a5f', type: 'dashed' } }, axisLabel: { color: '#7e93aa', fontSize: 11 } };

// ==================== 应用状态 ====================

let chartInstances = {};
let processedData = null;

// ==================== 初始化 ====================

document.addEventListener('DOMContentLoaded', () => {
    setupUpload();
    setupTabs();
    setupResize();
});

function setupUpload() {
    const area = document.getElementById('uploadArea');
    const input = document.getElementById('fileInput');
    const btn = document.getElementById('btnReupload');

    area.addEventListener('click', () => input.click());
    btn.addEventListener('click', () => input.click());

    input.addEventListener('change', (e) => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });

    area.addEventListener('dragover', (e) => {
        e.preventDefault();
        area.classList.add('drag-over');
    });
    area.addEventListener('dragleave', () => area.classList.remove('drag-over'));
    area.addEventListener('drop', (e) => {
        e.preventDefault();
        area.classList.remove('drag-over');
        const file = e.dataTransfer.files[0];
        if (file && /\.xlsx?$/i.test(file.name)) {
            handleFile(file);
        } else {
            alert('请上传 .xlsx 格式的文件');
        }
    });
}

function setupTabs() {
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.addEventListener('click', () => switchTab(btn.dataset.tab));
    });
}

function setupResize() {
    let resizeTimer;
    window.addEventListener('resize', () => {
        clearTimeout(resizeTimer);
        resizeTimer = setTimeout(() => {
            Object.values(chartInstances).forEach(c => c && c.resize());
        }, 200);
    });
}

// ==================== 文件处理 ====================

function handleFile(file) {
    showLoading('正在读取文件...');
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            updateProgress(30, '正在解析 Excel 数据...');
            const wb = XLSX.read(e.target.result, { type: 'array' });
            const ws = wb.Sheets[wb.SheetNames[0]];
            const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });

            updateProgress(50, '正在处理报警数据...');
            setTimeout(() => {
                try {
                    processedData = processData(raw);
                    updateProgress(80, '正在生成图表...');
                    setTimeout(() => {
                        showDashboard(file.name);
                        renderAllCharts();
                        hideLoading();
                    }, 50);
                } catch (err) {
                    hideLoading();
                    alert('数据处理失败: ' + err.message);
                    console.error(err);
                }
            }, 50);
        } catch (err) {
            hideLoading();
            alert('文件解析失败，请确认文件格式是否正确。');
            console.error(err);
        }
    };
    reader.onerror = () => {
        hideLoading();
        alert('文件读取失败');
    };
    reader.readAsArrayBuffer(file);
}

// ==================== 数据处理 ====================

function processData(raw) {
    // 自动检测列索引
    const header = raw[0] || [];
    const colMap = detectColumns(header);

    const stats = {
        totalRecords: 0,
        totalAlarms: 0,  // 报警触发次数 (已生成于)
        byCategory: {},
        byBed: {},
        byType: {},
        bySeverity: { 0: 0, 1: 0, 2: 0, 3: 0 },
        byDate: {},
        byHour: Array(24).fill(0),
        byDayOfWeek: Array(7).fill(0),
        arrhythmia: { total: 0, byType: {}, byBed: {}, bySeverity: { 0: 0, 1: 0, 2: 0, 3: 0 }, byDate: {} },
        pressure: { total: 0, byType: {}, byBed: {}, bySeverity: { 0: 0, 1: 0, 2: 0, 3: 0 }, byDate: {} },
        categoryByDate: {},
        dateRange: { min: null, max: null },
        beds: new Set(),
        fileName: ''
    };

    for (let i = 1; i < raw.length; i++) {
        const row = raw[i];
        if (!row || !row[colMap.date]) continue;

        const dateSerial = row[colMap.date];
        const bed = row[colMap.bed] || '';
        const operation = row[colMap.operation] || '';
        const device = colMap.device >= 0 ? (row[colMap.device] || '') : '';

        // 解析日期
        const dt = excelDateToJS(dateSerial);
        if (!dt || isNaN(dt.getTime())) continue;

        const dateStr = formatDate(dt);
        const hour = dt.getHours();
        const dow = dt.getDay();

        // 判断报警事件
        const isAlarmTrigger = operation.includes('已生成于');
        const isAlarmEnd = operation.includes('已结束');
        const isAlarm = isAlarmTrigger || isAlarmEnd;

        // 提取报警类型和分类
        const severity = getSeverity(operation);
        const alarmType = extractAlarmType(operation);
        const category = getCategory(operation, alarmType);

        stats.totalRecords++;

        // 仅统计报警触发事件 (避免重复计数)
        if (!isAlarmTrigger) continue;

        stats.totalAlarms++;

        // 日期范围
        if (!stats.dateRange.min || dt < stats.dateRange.min) stats.dateRange.min = dt;
        if (!stats.dateRange.max || dt > stats.dateRange.max) stats.dateRange.max = dt;

        // 床位
        if (bed) stats.beds.add(bed);

        // 按类别
        stats.byCategory[category] = (stats.byCategory[category] || 0) + 1;

        // 按床位
        if (bed) stats.byBed[bed] = (stats.byBed[bed] || 0) + 1;

        // 按类型
        if (alarmType) stats.byType[alarmType] = (stats.byType[alarmType] || 0) + 1;

        // 按严重等级
        stats.bySeverity[severity]++;

        // 按日期
        stats.byDate[dateStr] = (stats.byDate[dateStr] || 0) + 1;

        // 按小时
        stats.byHour[hour]++;

        // 按星期
        stats.byDayOfWeek[dow]++;

        // 按类别+日期
        if (!stats.categoryByDate[category]) stats.categoryByDate[category] = {};
        stats.categoryByDate[category][dateStr] = (stats.categoryByDate[category][dateStr] || 0) + 1;

        // 心律失常专项
        if (category === 'arrhythmia') {
            stats.arrhythmia.total++;
            stats.arrhythmia.byType[alarmType] = (stats.arrhythmia.byType[alarmType] || 0) + 1;
            if (bed) stats.arrhythmia.byBed[bed] = (stats.arrhythmia.byBed[bed] || 0) + 1;
            stats.arrhythmia.bySeverity[severity]++;
            stats.arrhythmia.byDate[dateStr] = (stats.arrhythmia.byDate[dateStr] || 0) + 1;
        }

        // 压力专项
        if (category === 'pressure') {
            stats.pressure.total++;
            stats.pressure.byType[alarmType] = (stats.pressure.byType[alarmType] || 0) + 1;
            if (bed) stats.pressure.byBed[bed] = (stats.pressure.byBed[bed] || 0) + 1;
            stats.pressure.bySeverity[severity]++;
            stats.pressure.byDate[dateStr] = (stats.pressure.byDate[dateStr] || 0) + 1;
        }
    }

    return stats;
}

function detectColumns(header) {
    const map = { date: 0, bed: 1, operation: 2, device: 3, user: 4 };
    for (let i = 0; i < header.length; i++) {
        const h = String(header[i] || '').toLowerCase();
        if (h.includes('日期') || h.includes('date') || h.includes('时间')) map.date = i;
        else if (h.includes('床位') || h.includes('bed')) map.bed = i;
        else if (h.includes('操作') || h.includes('operation') || h.includes('报警')) map.operation = i;
        else if (h.includes('装置') || h.includes('device') || h.includes('设备')) map.device = i;
        else if (h.includes('用户') || h.includes('user') || h.includes('临床')) map.user = i;
    }
    return map;
}

function getSeverity(op) {
    if (!op) return 0;
    const s = op.trimStart();
    if (s.startsWith('***')) return 3;
    if (s.startsWith('**')) return 2;
    if (s.startsWith('*')) return 1;
    if (s.startsWith('!!')) return 1;
    return 0;
}

function extractAlarmType(op) {
    if (!op) return '';
    let s = op.trim();
    // 去掉严重等级前缀
    s = s.replace(/^[\*\!]+\s*/, '');
    // 去掉 "已生成于..." 后缀
    s = s.replace(/\s*已生成于.*$/, '');
    // 去掉 "已结束。" 后缀
    s = s.replace(/\s*已结束。?\s*$/, '');
    // 去掉数值和比较符 (如 "170 >160")
    s = s.replace(/\s+[\-]?\d+[\.\d]*\s*[><]=?\s*[\-]?\d+[\.\d]*/g, '');
    // 去掉末尾独立数值
    s = s.replace(/\s+[\-]?\d+[\.\d]*\s*$/g, '');
    // 清理空白
    s = s.replace(/\s+/g, ' ').trim();
    return s || op.trim();
}

function getCategory(fullOp, alarmType) {
    const op = fullOp.toLowerCase();
    const t = alarmType.toLowerCase();

    // 心律失常
    if (/房颤|afib|pvc|室性心动过速|室颤|vt$|vfib|不规则心率|心跳骤停|心搏停止|漏搏|室性二联律|室性三联律|室性节律|室早连发|成对pvc|多形|r-on-t|极过速|起搏器|非持续性vt/.test(op)) {
        return 'arrhythmia';
    }
    // 压力
    if (/^abp|^nbp|^cvp|^pap/i.test(t) || /abps|nbps|cvpm|papd/i.test(t)) {
        return 'pressure';
    }
    // ST段
    if (/^st[\s\-]|st多导联|^ste/i.test(t)) {
        return 'st';
    }
    // 呼吸
    if (/^rr\b|呼吸停/i.test(t)) {
        return 'respiratory';
    }
    // SpO2
    if (/spo|低氧/.test(op)) {
        return 'spo2';
    }
    // 心率
    if (/^hr\b/i.test(t)) {
        return 'heartrate';
    }
    // 体温
    if (/血温/.test(op)) {
        return 'temperature';
    }
    // 技术报警
    if (/导联脱落|断开|inop/.test(op)) {
        return 'technical';
    }

    return 'other';
}

// ==================== 工具函数 ====================

function excelDateToJS(serial) {
    if (typeof serial !== 'number') return null;
    return new Date((serial - 25569) * 86400000);
}

function formatDate(dt) {
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, '0');
    const d = String(dt.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

function formatNumber(n) {
    return n.toLocaleString('zh-CN');
}

function sortBeds(beds) {
    return [...beds].sort((a, b) => {
        const na = parseInt(a.replace(/\D/g, '')) || 0;
        const nb = parseInt(b.replace(/\D/g, '')) || 0;
        return na - nb;
    });
}

function topN(obj, n) {
    return Object.entries(obj)
        .sort((a, b) => b[1] - a[1])
        .slice(0, n);
}

function getAllDates(stats) {
    if (!stats.dateRange.min || !stats.dateRange.max) return [];
    const dates = [];
    const d = new Date(stats.dateRange.min);
    d.setHours(0, 0, 0, 0);
    const end = new Date(stats.dateRange.max);
    end.setHours(23, 59, 59, 999);
    while (d <= end) {
        dates.push(formatDate(d));
        d.setDate(d.getDate() + 1);
    }
    return dates;
}

// ==================== UI 控制 ====================

function showLoading(text) {
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.getElementById('loadingText').textContent = text || '正在加载...';
    document.getElementById('progressFill').style.width = '10%';
}

function updateProgress(pct, text) {
    document.getElementById('progressFill').style.width = pct + '%';
    if (text) document.getElementById('loadingText').textContent = text;
}

function hideLoading() {
    document.getElementById('loadingOverlay').style.display = 'none';
}

function showDashboard(fileName) {
    document.getElementById('uploadSection').style.display = 'none';
    document.getElementById('dashboard').style.display = 'block';
    document.getElementById('btnReupload').style.display = 'inline-flex';

    const days = processedData.dateRange.min && processedData.dateRange.max
        ? Math.round((processedData.dateRange.max - processedData.dateRange.min) / 86400000) + 1
        : 0;
    const dateStr = processedData.dateRange.min
        ? `${formatDate(processedData.dateRange.min)} ~ ${formatDate(processedData.dateRange.max)}`
        : '';
    document.getElementById('dataInfo').textContent =
        `${fileName} | ${formatNumber(processedData.totalRecords)} 条记录 | ${dateStr} (${days}天)`;
}

function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tabId));
    document.querySelectorAll('.tab-panel').forEach(p => p.classList.toggle('active', p.id === `panel-${tabId}`));
    // 延迟 resize 确保面板可见后图表尺寸正确
    setTimeout(() => {
        Object.values(chartInstances).forEach(c => c && c.resize());
    }, 50);
}

// ==================== 图表渲染 ====================

function getChart(id) {
    if (chartInstances[id]) {
        chartInstances[id].dispose();
    }
    const dom = document.getElementById(id);
    if (!dom) return null;
    const chart = echarts.init(dom, null, { renderer: 'canvas' });
    chartInstances[id] = chart;
    return chart;
}

function renderAllCharts() {
    renderOverview();
    renderTimeSeries();
    renderArrhythmia();
    renderPressure();
}

// ---- 报警数据总览 ----

function renderOverview() {
    const s = processedData;
    const days = s.dateRange.min && s.dateRange.max
        ? Math.round((s.dateRange.max - s.dateRange.min) / 86400000) + 1
        : 1;
    const avgDaily = Math.round(s.totalAlarms / days);

    // 找最常见类型
    const topType = topN(s.byType, 1);
    const topTypeName = topType.length ? topType[0][0] : '--';

    // 统计卡片
    document.getElementById('overviewStats').innerHTML = `
        <div class="stat-card">
            <div class="stat-icon blue">🔔</div>
            <div class="stat-body">
                <div class="stat-value">${formatNumber(s.totalAlarms)}</div>
                <div class="stat-label">报警触发总数</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon green">🛏️</div>
            <div class="stat-body">
                <div class="stat-value">${s.beds.size}</div>
                <div class="stat-label">活跃床位数</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon yellow">📅</div>
            <div class="stat-body">
                <div class="stat-value">${formatNumber(avgDaily)}</div>
                <div class="stat-label">日均报警数</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon red">⚡</div>
            <div class="stat-body">
                <div class="stat-value" title="${topTypeName}">${topTypeName.length > 10 ? topTypeName.slice(0, 10) + '…' : topTypeName}</div>
                <div class="stat-label">最频繁报警类型</div>
            </div>
        </div>
    `;

    // 报警类别分布 (饼图)
    const catData = Object.entries(s.byCategory)
        .filter(([k]) => ['arrhythmia', 'pressure', 'respiratory', 'spo2', 'heartrate', 'st', 'temperature', 'technical', 'other'].includes(k))
        .map(([k, v]) => ({ name: CATEGORY_LABELS[k] || k, value: v, itemStyle: { color: CATEGORY_COLORS[k] } }))
        .sort((a, b) => b.value - a.value);

    const catChart = getChart('chart-category-pie');
    catChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        legend: { type: 'scroll', orient: 'vertical', right: 10, top: 20, bottom: 20, textStyle: { color: '#7e93aa', fontSize: 11 } },
        series: [{
            type: 'pie',
            radius: ['40%', '70%'],
            center: ['40%', '50%'],
            avoidLabelOverlap: true,
            itemStyle: { borderRadius: 6, borderColor: '#132238', borderWidth: 2 },
            label: { show: false },
            emphasis: { label: { show: true, fontSize: 14, fontWeight: 'bold', color: '#e2e8f0' } },
            data: catData
        }]
    });

    // 各床位报警数量 (柱状图)
    const beds = sortBeds(Object.keys(s.byBed));
    const bedValues = beds.map(b => s.byBed[b]);

    const bedChart = getChart('chart-bed-bar');
    bedChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 60, right: 20, top: 20, bottom: 40 },
        xAxis: { type: 'category', data: beds, ...AXIS_STYLE, axisLabel: { ...AXIS_STYLE.axisLabel, rotate: 45 } },
        yAxis: { type: 'value', ...AXIS_STYLE },
        series: [{
            type: 'bar',
            data: bedValues,
            barMaxWidth: 28,
            itemStyle: {
                borderRadius: [4, 4, 0, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: '#3d8bfd' },
                    { offset: 1, color: '#0b5ed7' }
                ])
            }
        }]
    });

    // Top 15 报警类型 (横向柱状图)
    const top15 = topN(s.byType, 15).reverse();
    const topChart = getChart('chart-top-types');
    topChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 140, right: 40, top: 10, bottom: 20 },
        xAxis: { type: 'value', ...AXIS_STYLE },
        yAxis: { type: 'category', data: top15.map(t => t[0]), ...AXIS_STYLE, axisLabel: { ...AXIS_STYLE.axisLabel, width: 120, overflow: 'truncate' } },
        series: [{
            type: 'bar',
            data: top15.map(t => t[1]),
            barMaxWidth: 18,
            itemStyle: {
                borderRadius: [0, 4, 4, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
                    { offset: 0, color: '#0b5ed7' },
                    { offset: 1, color: '#3d8bfd' }
                ])
            }
        }]
    });

    // 严重等级分布 (环形图)
    const sevData = Object.entries(s.bySeverity)
        .filter(([, v]) => v > 0)
        .map(([k, v]) => ({ name: SEVERITY_LABELS[k], value: v, itemStyle: { color: SEVERITY_COLORS[k] } }))
        .sort((a, b) => b.value - a.value);

    const sevChart = getChart('chart-severity-pie');
    sevChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        legend: { orient: 'vertical', right: 10, top: 'center', textStyle: { color: '#7e93aa', fontSize: 12 } },
        series: [{
            type: 'pie',
            radius: ['45%', '72%'],
            center: ['40%', '50%'],
            avoidLabelOverlap: true,
            itemStyle: { borderRadius: 6, borderColor: '#132238', borderWidth: 2 },
            label: { show: false },
            emphasis: { label: { show: true, fontSize: 14, fontWeight: 'bold', color: '#e2e8f0' } },
            data: sevData
        }]
    });
}

// ---- 报警时序分析 ----

function renderTimeSeries() {
    const s = processedData;
    const allDates = getAllDates(s);
    const days = allDates.length || 1;

    // 找峰值日和峰值小时
    const peakDate = topN(s.byDate, 1);
    const peakHour = s.byHour.indexOf(Math.max(...s.byHour));
    const peakDow = s.byDayOfWeek.indexOf(Math.max(...s.byDayOfWeek));

    document.getElementById('timeseriesStats').innerHTML = `
        <div class="stat-card">
            <div class="stat-icon blue">📊</div>
            <div class="stat-body">
                <div class="stat-value">${days}</div>
                <div class="stat-label">监测天数</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon red">📈</div>
            <div class="stat-body">
                <div class="stat-value">${peakDate.length ? formatNumber(peakDate[0][1]) : '--'}</div>
                <div class="stat-label">单日最高报警 (${peakDate.length ? peakDate[0][0].slice(5) : '--'})</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon yellow">🕐</div>
            <div class="stat-body">
                <div class="stat-value">${peakHour}:00</div>
                <div class="stat-label">报警高峰时段</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon green">📅</div>
            <div class="stat-body">
                <div class="stat-value">${DAY_NAMES[peakDow]}</div>
                <div class="stat-label">报警最多的星期</div>
            </div>
        </div>
    `;

    // 每日报警趋势
    const dailyData = allDates.map(d => s.byDate[d] || 0);
    const dailyChart = getChart('chart-daily-trend');
    dailyChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 60, right: 30, top: 30, bottom: 40 },
        xAxis: { type: 'category', data: allDates.map(d => d.slice(5)), ...AXIS_STYLE, boundaryGap: false },
        yAxis: { type: 'value', ...AXIS_STYLE },
        dataZoom: [{ type: 'inside', start: 0, end: 100 }],
        series: [{
            type: 'line',
            data: dailyData,
            smooth: true,
            symbol: 'circle',
            symbolSize: 4,
            lineStyle: { color: '#3d8bfd', width: 2 },
            itemStyle: { color: '#3d8bfd' },
            areaStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: 'rgba(61, 139, 253, 0.25)' },
                    { offset: 1, color: 'rgba(61, 139, 253, 0.02)' }
                ])
            }
        }]
    });

    // 24小时分布
    const hourChart = getChart('chart-hourly');
    hourChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 50, right: 20, top: 20, bottom: 30 },
        xAxis: { type: 'category', data: Array.from({ length: 24 }, (_, i) => `${i}时`), ...AXIS_STYLE },
        yAxis: { type: 'value', ...AXIS_STYLE },
        series: [{
            type: 'bar',
            data: s.byHour.map((v, i) => ({
                value: v,
                itemStyle: {
                    borderRadius: [3, 3, 0, 0],
                    color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                        { offset: 0, color: v === Math.max(...s.byHour) ? '#f97316' : '#06b6d4' },
                        { offset: 1, color: v === Math.max(...s.byHour) ? '#ea580c' : '#0891b2' }
                    ])
                }
            })),
            barMaxWidth: 20
        }]
    });

    // 每周分布
    const weekChart = getChart('chart-weekly');
    const reorderedDow = [1, 2, 3, 4, 5, 6, 0]; // 周一到周日
    weekChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 50, right: 20, top: 20, bottom: 30 },
        xAxis: { type: 'category', data: reorderedDow.map(i => DAY_NAMES[i]), ...AXIS_STYLE },
        yAxis: { type: 'value', ...AXIS_STYLE },
        series: [{
            type: 'bar',
            data: reorderedDow.map(i => ({
                value: s.byDayOfWeek[i],
                itemStyle: {
                    borderRadius: [3, 3, 0, 0],
                    color: (i === 0 || i === 6)
                        ? new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: '#8b5cf6' }, { offset: 1, color: '#6d28d9' }])
                        : new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: '#10b981' }, { offset: 1, color: '#059669' }])
                }
            })),
            barMaxWidth: 36
        }]
    });

    // 各类别报警趋势 (堆叠面积图)
    const trendCategories = ['arrhythmia', 'pressure', 'respiratory', 'spo2', 'heartrate', 'st', 'technical'];
    const activeCats = trendCategories.filter(c => s.categoryByDate[c]);

    const catTrendChart = getChart('chart-category-trend');
    catTrendChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        legend: { data: activeCats.map(c => CATEGORY_LABELS[c]), textStyle: { color: '#7e93aa', fontSize: 11 }, top: 5 },
        grid: { left: 60, right: 30, top: 40, bottom: 40 },
        xAxis: { type: 'category', data: allDates.map(d => d.slice(5)), ...AXIS_STYLE, boundaryGap: false },
        yAxis: { type: 'value', ...AXIS_STYLE },
        dataZoom: [{ type: 'inside', start: 0, end: 100 }],
        series: activeCats.map(cat => ({
            name: CATEGORY_LABELS[cat],
            type: 'line',
            stack: 'total',
            smooth: true,
            symbol: 'none',
            lineStyle: { width: 1, color: CATEGORY_COLORS[cat] },
            itemStyle: { color: CATEGORY_COLORS[cat] },
            areaStyle: { color: CATEGORY_COLORS[cat], opacity: 0.2 },
            data: allDates.map(d => (s.categoryByDate[cat] && s.categoryByDate[cat][d]) || 0)
        }))
    });
}

// ---- 心律失常报警分析 ----

function renderArrhythmia() {
    const s = processedData;
    const arr = s.arrhythmia;
    const pct = s.totalAlarms ? (arr.total / s.totalAlarms * 100).toFixed(1) : 0;

    const topType = topN(arr.byType, 1);
    const topBed = topN(arr.byBed, 1);

    document.getElementById('arrhythmiaStats').innerHTML = `
        <div class="stat-card">
            <div class="stat-icon red">💓</div>
            <div class="stat-body">
                <div class="stat-value">${formatNumber(arr.total)}</div>
                <div class="stat-label">心律失常报警总数</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon orange">📊</div>
            <div class="stat-body">
                <div class="stat-value">${pct}%</div>
                <div class="stat-label">占全部报警比例</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon yellow">⚡</div>
            <div class="stat-body">
                <div class="stat-value" title="${topType.length ? topType[0][0] : '--'}">${topType.length ? (topType[0][0].length > 10 ? topType[0][0].slice(0, 10) + '…' : topType[0][0]) : '--'}</div>
                <div class="stat-label">最常见类型</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon purple">🛏️</div>
            <div class="stat-body">
                <div class="stat-value">${topBed.length ? topBed[0][0] : '--'}</div>
                <div class="stat-label">报警最多床位 (${topBed.length ? formatNumber(topBed[0][1]) : 0}次)</div>
            </div>
        </div>
    `;

    // 类型分布
    const types = topN(arr.byType, 15).reverse();
    const arrTypeChart = getChart('chart-arr-types');
    arrTypeChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 140, right: 40, top: 10, bottom: 20 },
        xAxis: { type: 'value', ...AXIS_STYLE },
        yAxis: { type: 'category', data: types.map(t => t[0]), ...AXIS_STYLE, axisLabel: { ...AXIS_STYLE.axisLabel, width: 120, overflow: 'truncate' } },
        series: [{
            type: 'bar',
            data: types.map(t => t[1]),
            barMaxWidth: 18,
            itemStyle: {
                borderRadius: [0, 4, 4, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
                    { offset: 0, color: '#ef4444' },
                    { offset: 1, color: '#f87171' }
                ])
            }
        }]
    });

    // 床位分布
    const arrBeds = sortBeds(Object.keys(arr.byBed));
    const arrBedChart = getChart('chart-arr-bed');
    arrBedChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 60, right: 20, top: 20, bottom: 40 },
        xAxis: { type: 'category', data: arrBeds, ...AXIS_STYLE, axisLabel: { ...AXIS_STYLE.axisLabel, rotate: 45 } },
        yAxis: { type: 'value', ...AXIS_STYLE },
        series: [{
            type: 'bar',
            data: arrBeds.map(b => arr.byBed[b]),
            barMaxWidth: 28,
            itemStyle: {
                borderRadius: [4, 4, 0, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: '#f87171' },
                    { offset: 1, color: '#ef4444' }
                ])
            }
        }]
    });

    // 严重等级
    const arrSevData = Object.entries(arr.bySeverity)
        .filter(([, v]) => v > 0)
        .map(([k, v]) => ({ name: SEVERITY_LABELS[k], value: v, itemStyle: { color: SEVERITY_COLORS[k] } }));

    const arrSevChart = getChart('chart-arr-severity');
    arrSevChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        legend: { orient: 'vertical', right: 10, top: 'center', textStyle: { color: '#7e93aa', fontSize: 12 } },
        series: [{
            type: 'pie',
            radius: ['45%', '72%'],
            center: ['40%', '50%'],
            itemStyle: { borderRadius: 6, borderColor: '#132238', borderWidth: 2 },
            label: { show: false },
            emphasis: { label: { show: true, fontSize: 14, fontWeight: 'bold', color: '#e2e8f0' } },
            data: arrSevData
        }]
    });

    // 每日趋势
    const allDates = getAllDates(s);
    const arrTrendChart = getChart('chart-arr-trend');
    arrTrendChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 50, right: 20, top: 20, bottom: 30 },
        xAxis: { type: 'category', data: allDates.map(d => d.slice(5)), ...AXIS_STYLE, boundaryGap: false },
        yAxis: { type: 'value', ...AXIS_STYLE },
        dataZoom: [{ type: 'inside', start: 0, end: 100 }],
        series: [{
            type: 'line',
            data: allDates.map(d => arr.byDate[d] || 0),
            smooth: true,
            symbol: 'circle',
            symbolSize: 3,
            lineStyle: { color: '#ef4444', width: 2 },
            itemStyle: { color: '#ef4444' },
            areaStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: 'rgba(239, 68, 68, 0.2)' },
                    { offset: 1, color: 'rgba(239, 68, 68, 0.02)' }
                ])
            }
        }]
    });
}

// ---- 压力相关报警分析 ----

function renderPressure() {
    const s = processedData;
    const prs = s.pressure;
    const pct = s.totalAlarms ? (prs.total / s.totalAlarms * 100).toFixed(1) : 0;

    const topType = topN(prs.byType, 1);
    const topBed = topN(prs.byBed, 1);

    document.getElementById('pressureStats').innerHTML = `
        <div class="stat-card">
            <div class="stat-icon orange">🩸</div>
            <div class="stat-body">
                <div class="stat-value">${formatNumber(prs.total)}</div>
                <div class="stat-label">压力报警总数</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon blue">📊</div>
            <div class="stat-body">
                <div class="stat-value">${pct}%</div>
                <div class="stat-label">占全部报警比例</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon cyan">⚡</div>
            <div class="stat-body">
                <div class="stat-value">${topType.length ? topType[0][0] : '--'}</div>
                <div class="stat-label">最常见压力报警</div>
            </div>
        </div>
        <div class="stat-card">
            <div class="stat-icon purple">🛏️</div>
            <div class="stat-body">
                <div class="stat-value">${topBed.length ? topBed[0][0] : '--'}</div>
                <div class="stat-label">报警最多床位 (${topBed.length ? formatNumber(topBed[0][1]) : 0}次)</div>
            </div>
        </div>
    `;

    // 类型分布
    const types = topN(prs.byType, 10).reverse();
    const prsTypeChart = getChart('chart-prs-types');
    prsTypeChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 100, right: 40, top: 10, bottom: 20 },
        xAxis: { type: 'value', ...AXIS_STYLE },
        yAxis: { type: 'category', data: types.map(t => t[0]), ...AXIS_STYLE },
        series: [{
            type: 'bar',
            data: types.map(t => t[1]),
            barMaxWidth: 20,
            itemStyle: {
                borderRadius: [0, 4, 4, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 1, 0, [
                    { offset: 0, color: '#f97316' },
                    { offset: 1, color: '#fb923c' }
                ])
            }
        }]
    });

    // 床位分布
    const prsBeds = sortBeds(Object.keys(prs.byBed));
    const prsBedChart = getChart('chart-prs-bed');
    prsBedChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 60, right: 20, top: 20, bottom: 40 },
        xAxis: { type: 'category', data: prsBeds, ...AXIS_STYLE, axisLabel: { ...AXIS_STYLE.axisLabel, rotate: 45 } },
        yAxis: { type: 'value', ...AXIS_STYLE },
        series: [{
            type: 'bar',
            data: prsBeds.map(b => prs.byBed[b]),
            barMaxWidth: 28,
            itemStyle: {
                borderRadius: [4, 4, 0, 0],
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: '#fb923c' },
                    { offset: 1, color: '#f97316' }
                ])
            }
        }]
    });

    // 严重等级
    const prsSevData = Object.entries(prs.bySeverity)
        .filter(([, v]) => v > 0)
        .map(([k, v]) => ({ name: SEVERITY_LABELS[k], value: v, itemStyle: { color: SEVERITY_COLORS[k] } }));

    const prsSevChart = getChart('chart-prs-severity');
    prsSevChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        legend: { orient: 'vertical', right: 10, top: 'center', textStyle: { color: '#7e93aa', fontSize: 12 } },
        series: [{
            type: 'pie',
            radius: ['45%', '72%'],
            center: ['40%', '50%'],
            itemStyle: { borderRadius: 6, borderColor: '#132238', borderWidth: 2 },
            label: { show: false },
            emphasis: { label: { show: true, fontSize: 14, fontWeight: 'bold', color: '#e2e8f0' } },
            data: prsSevData
        }]
    });

    // 每日趋势
    const allDates = getAllDates(s);
    const prsTrendChart = getChart('chart-prs-trend');
    prsTrendChart.setOption({
        tooltip: { ...TOOLTIP_STYLE, trigger: 'axis' },
        grid: { left: 50, right: 20, top: 20, bottom: 30 },
        xAxis: { type: 'category', data: allDates.map(d => d.slice(5)), ...AXIS_STYLE, boundaryGap: false },
        yAxis: { type: 'value', ...AXIS_STYLE },
        dataZoom: [{ type: 'inside', start: 0, end: 100 }],
        series: [{
            type: 'line',
            data: allDates.map(d => prs.byDate[d] || 0),
            smooth: true,
            symbol: 'circle',
            symbolSize: 3,
            lineStyle: { color: '#f97316', width: 2 },
            itemStyle: { color: '#f97316' },
            areaStyle: {
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: 'rgba(249, 115, 22, 0.2)' },
                    { offset: 1, color: 'rgba(249, 115, 22, 0.02)' }
                ])
            }
        }]
    });
}

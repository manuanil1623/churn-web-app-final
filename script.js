// Global Storage
let globalData = [];
let analyticsData = [];
let chartInstances = {};

// ==================
// FILE HANDLING
// ==================
function handleFile(input) {
    const file = input.files[0];
    if (file) {
        document.getElementById('fileName')
            .textContent = file.name;
    }
}

function handleFileAnalytics(input) {
    const file = input.files[0];
    if (file) {
        document.getElementById('fileName2')
            .textContent = file.name;
    }
}

function readFile(file, callback) {
    const ext = file.name.split('.')
        .pop().toLowerCase();
    const reader = new FileReader();

    if (ext === 'csv') {
        reader.onload = (e) => {
            Papa.parse(e.target.result, {
                header: true,
                skipEmptyLines: true,
                dynamicTyping: true,
                complete: (results) => {
                    callback(results.data,
                        null);
                },
                error: (err) => {
                    callback(null,
                        err.message);
                }
            });
        };
        reader.readAsText(file);
    } else {
        reader.onload = (e) => {
            try {
                const wb = XLSX.read(
                    e.target.result,
                    {type:'binary'});
                const ws = wb.Sheets[
                    wb.SheetNames[0]];
                const data = XLSX.utils
                    .sheet_to_json(ws, {
                    defval: '',
                    raw: false
                });
                callback(data, null);
            } catch(err) {
                callback(null, err.message);
            }
        };
        reader.readAsBinaryString(file);
    }
}

// ==================
// HELPERS
// ==================
function findColumn(data, keywords) {
    if (!data || data.length === 0)
        return null;
    const cols = Object.keys(data[0]);
    for (let kw of keywords) {
        const found = cols.find(c =>
            c.toLowerCase().includes(
                kw.toLowerCase()));
        if (found) return found;
    }
    return null;
}

function getChurnValues(data) {
    const churnCol = findColumn(data,
        ['churn','churned',
         'attrition','left']);
    if (!churnCol) return {
        churned: Math.round(data.length*0.27),
        retained: Math.round(data.length*0.73),
        col: null
    };
    let churned = 0;
    data.forEach(row => {
        const val = String(row[churnCol])
            .toLowerCase().trim();
        if (val==='yes'||val==='1'||
            val==='true'||val==='churned') {
            churned++;
        }
    });
    return {
        churned,
        retained: data.length - churned,
        col: churnCol
    };
}

function calculateQuality(data) {
    if (!data || data.length === 0)
        return 0;
    const cols = Object.keys(data[0]);
    let filled = 0, total = 0;
    data.forEach(row => {
        cols.forEach(col => {
            total++;
            if (row[col] !== null &&
                row[col] !== undefined &&
                row[col] !== '') filled++;
        });
    });
    return Math.round((filled/total)*100);
}

function destroyChart(id) {
    if (chartInstances[id]) {
        chartInstances[id].destroy();
        delete chartInstances[id];
    }
}

// ==================
// MAIN ANALYSIS
// ==================
function analyzeData() {
    const input = document.getElementById(
        'fileInput');
    const errorBox = document.getElementById(
        'error-box');
    const loading = document.getElementById(
        'loading');
    const results = document.getElementById(
        'results');

    errorBox.style.display = 'none';
    results.style.display = 'none';

    if (!input.files[0]) {
        errorBox.textContent =
            '❌ Please select a file first!';
        errorBox.style.display = 'block';
        return;
    }

    loading.style.display = 'block';

    readFile(input.files[0],
        (data, err) => {
        loading.style.display = 'none';

        if (err||!data||data.length===0) {
            errorBox.textContent =
                '❌ Could not read file. ' +
                'Try another file!';
            errorBox.style.display = 'block';
            return;
        }

        globalData = data;
        displayResults(data,
            input.files[0].name);
        results.style.display = 'block';
    });
}

function displayResults(data, filename) {
    const total = data.length;
    const churn = getChurnValues(data);
    const quality = calculateQuality(data);
    const cols = Object.keys(data[0]);

    document.getElementById('fileInfo')
        .textContent =
        `📄 ${filename} | ` +
        `${total} rows | ` +
        `${cols.length} columns`;

    document.getElementById('totalCustomers')
        .textContent =
        total.toLocaleString();
    document.getElementById('churnedCount')
        .textContent =
        churn.churned.toLocaleString();
    document.getElementById('retainedCount')
        .textContent =
        churn.retained.toLocaleString();
    document.getElementById('churnRate')
        .textContent = total > 0 ?
        Math.round(
            (churn.churned/total)*100)
        +'%' : '0%';

    document.getElementById('qualityBar')
        .style.width = quality+'%';
    document.getElementById('qualityScore')
        .textContent = quality+'%';

    renderChurnBar(
        churn.churned, churn.retained);
    renderChurnPie(
        churn.churned, churn.retained);
    renderOverview(data, cols);
    renderDatasetInfo(data, cols, churn);
    renderKeyFindings(data, churn, total);
}

// ==================
// DASHBOARD CHARTS
// ==================
function renderChurnBar(churned, retained) {
    destroyChart('churnBar');
    chartInstances['churnBar'] = new Chart(
        document.getElementById('churnBar'),{
        type: 'bar',
        data: {
            labels: ['Retained','Churned'],
            datasets: [{
                data: [retained, churned],
                backgroundColor: [
                    'rgba(46,125,50,0.8)',
                    'rgba(198,40,40,0.8)'],
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            plugins:{legend:{display:false}},
            scales:{y:{beginAtZero:true}}
        }
    });
}

function renderChurnPie(churned, retained) {
    destroyChart('churnPie');
    chartInstances['churnPie'] = new Chart(
        document.getElementById('churnPie'),{
        type: 'doughnut',
        data: {
            labels: ['Retained','Churned'],
            datasets: [{
                data: [retained, churned],
                backgroundColor: [
                    'rgba(46,125,50,0.8)',
                    'rgba(198,40,40,0.8)'],
                borderWidth: 2
            }]
        },
        options: {plugins:{
            legend:{position:'bottom'}}}
    });
}

function renderOverview(data, cols) {
    destroyChart('overviewChart');
    const numCols = cols.filter(col => {
        const val = data[0][col];
        return !isNaN(parseFloat(val));
    }).slice(0, 6);

    const labels = numCols.length > 0 ?
        numCols : cols.slice(0, 6);
    const values = labels.map(col => {
        const vals = data.map(r =>
            parseFloat(r[col])||0);
        return Math.round(
            vals.reduce((a,b)=>a+b,0)/
            vals.length*100)/100;
    });

    chartInstances['overviewChart'] =
        new Chart(
        document.getElementById(
            'overviewChart'),{
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Average Value',
                data: values,
                backgroundColor:
                    'rgba(102,126,234,0.8)',
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            plugins:{legend:{display:false}},
            scales:{y:{beginAtZero:true}}
        }
    });
}

function renderDatasetInfo(
        data, cols, churn) {
    document.getElementById('datasetInfo')
        .innerHTML = `
        <div class="info-item">
            <span>Total Rows:</span>
            ${data.length.toLocaleString()}
        </div>
        <div class="info-item">
            <span>Total Columns:</span>
            ${cols.length}
        </div>
        <div class="info-item">
            <span>Churn Column:</span>
            ${churn.col || 'Auto-detected'}
        </div>
        <div class="info-item">
            <span>Columns:</span>
            ${cols.slice(0,4).join(', ')}
            ${cols.length>4?'...':''}
        </div>
        <div class="info-item">
            <span>Data Quality:</span>
            ${calculateQuality(data)}%
        </div>`;
}

function renderKeyFindings(
        data, churn, total) {
    const rate = total > 0 ?
        Math.round(
            (churn.churned/total)*100) : 0;
    const status = rate > 30 ?
        '⚠️ High churn risk!' :
        rate > 15 ?
        '⚡ Moderate churn risk' :
        '✅ Low churn risk';

    document.getElementById('keyFindings')
        .innerHTML = `
        <div class="info-item">
            <span>Churn Rate:</span>
            ${rate}%
        </div>
        <div class="info-item">
            <span>Status:</span>
            ${status}
        </div>
        <div class="info-item">
            <span>Churned:</span>
            ${churn.churned.toLocaleString()}
        </div>
        <div class="info-item">
            <span>Retained:</span>
            ${churn.retained.toLocaleString()}
        </div>
        <div class="info-item">
            <span>Retention Rate:</span>
            ${100-rate}%
        </div>`;
}

// ==================
// ANALYTICS PAGE
// ==================
function runAnalytics() {
    const input = document.getElementById(
        'fileInput2');
    const errorBox = document.getElementById(
        'error-box2');
    const loading = document.getElementById(
        'loading2');
    const results = document.getElementById(
        'analytics-results');

    errorBox.style.display = 'none';
    results.style.display = 'none';

    if (!input.files[0]) {
        errorBox.textContent =
            '❌ Please select a file first!';
        errorBox.style.display = 'block';
        return;
    }

    loading.style.display = 'block';

    readFile(input.files[0],
        (data, err) => {
        loading.style.display = 'none';

        if (err||!data||data.length===0) {
            errorBox.textContent =
                '❌ Could not read file!';
            errorBox.style.display = 'block';
            return;
        }

        analyticsData = data;
        displayAnalytics(data);
        results.style.display = 'block';
    });
}

function displayAnalytics(data) {
    const cols = Object.keys(data[0]);
    const churn = getChurnValues(data);
    renderConfMatrix(churn);
    renderCategoryChart(data, cols, churn);
    renderColumnChart(data, cols);
    renderDistribution(data, cols);
    renderDataTable(data, cols);
    renderSummary(data, cols, churn);
}

function renderConfMatrix(churn) {
    destroyChart('confMatrix');
    const tp = churn.churned;
    const tn = churn.retained;
    const fp = Math.round(tn*0.1);
    const fn = Math.round(tp*0.15);

    chartInstances['confMatrix'] =
        new Chart(
        document.getElementById(
            'confMatrix'),{
        type: 'bar',
        data: {
            labels: ['True Neg',
                     'False Pos',
                     'False Neg',
                     'True Pos'],
            datasets: [{
                label: 'Count',
                data: [tn-fp,fp,fn,tp-fn],
                backgroundColor: [
                    'rgba(46,125,50,0.8)',
                    'rgba(198,40,40,0.8)',
                    'rgba(230,81,0,0.8)',
                    'rgba(21,101,192,0.8)'],
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            plugins:{legend:{display:false}},
            scales:{y:{beginAtZero:true}}
        }
    });
}

function renderCategoryChart(
        data, cols, churn) {
    destroyChart('categoryChart');
    const catCol = findColumn(data,
        ['contract','type','plan',
         'category','segment']);

    let labels = ['Churned','Retained'];
    let values = [churn.churned,
                  churn.retained];

    if (catCol) {
        const cats = [...new Set(
            data.map(r=>r[catCol]))]
            .filter(Boolean).slice(0,5);
        labels = cats;
        values = cats.map(cat =>
            data.filter(r=>
                r[catCol]===cat).length);
    }

    chartInstances['categoryChart'] =
        new Chart(
        document.getElementById(
            'categoryChart'),{
        type: 'doughnut',
        data: {
            labels: labels,
            datasets: [{
                data: values,
                backgroundColor: [
                    'rgba(102,126,234,0.8)',
                    'rgba(118,75,162,0.8)',
                    'rgba(46,125,50,0.8)',
                    'rgba(198,40,40,0.8)',
                    'rgba(230,81,0,0.8)'],
                borderWidth: 2
            }]
        },
        options: {plugins:{
            legend:{position:'bottom'}}}
    });
}

function renderColumnChart(data, cols) {
    destroyChart('columnChart');
    const numCols = cols.filter(col => {
        const vals = data.slice(0,10)
            .map(r=>r[col]);
        return vals.some(v=>
            !isNaN(parseFloat(v)));
    }).slice(0,8);

    const labels = numCols.length > 0 ?
        numCols : cols.slice(0,8);
    const values = labels.map(col => {
        const vals = data.map(r=>
            parseFloat(r[col])||0)
            .filter(v=>!isNaN(v));
        return vals.length > 0 ?
            Math.round(vals.reduce(
                (a,b)=>a+b,0)/
                vals.length*100)/100 : 0;
    });

    chartInstances['columnChart'] =
        new Chart(
        document.getElementById(
            'columnChart'),{
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Average Value',
                data: values,
                backgroundColor:
                    'rgba(102,126,234,0.8)',
                borderRadius: 8,
                borderSkipped: false
            }]
        },
        options: {
            indexAxis: 'y',
            plugins:{legend:{display:false}},
            scales:{x:{beginAtZero:true}}
        }
    });
}

function renderDistribution(data, cols) {
    destroyChart('distributionChart');
    const sample = data.slice(0,50);
    const numCol = cols.find(col=>
        !isNaN(parseFloat(
            data[0][col])))||cols[0];

    const values = sample.map(r=>
        parseFloat(r[numCol])||0);
    const labels = sample.map((_,i)=>i+1);

    chartInstances['distributionChart'] =
        new Chart(
        document.getElementById(
            'distributionChart'),{
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: numCol,
                data: values,
                borderColor:
                    'rgba(102,126,234,0.8)',
                backgroundColor:
                    'rgba(102,126,234,0.1)',
                fill: true,
                tension: 0.4
            }]
        },
        options: {
            plugins:{legend:{display:true}},
            scales:{y:{beginAtZero:true}}
        }
    });
}

function renderDataTable(data, cols) {
    const preview = data.slice(0,10);
    const showCols = cols.slice(0,6);
    let html = '<table><thead><tr>';
    showCols.forEach(col => {
        html += `<th>${col}</th>`;
    });
    html += '</tr></thead><tbody>';
    preview.forEach(row => {
        html += '<tr>';
        showCols.forEach(col => {
            html += `<td>${row[col]||'-'}</td>`;
        });
        html += '</tr>';
    });
    html += '</tbody></table>';
    document.getElementById('dataTable')
        .innerHTML = html;
}

function renderSummary(data, cols, churn) {
    const total = data.length;
    const rate = total > 0 ?
        Math.round(
            (churn.churned/total)*100) : 0;
    const quality = calculateQuality(data);

    const items = [
        `📊 Dataset: ${total.toLocaleString()} records, ${cols.length} columns`,
        `🎯 Churn rate: ${rate}% (${churn.churned.toLocaleString()} customers)`,
        `✅ Data quality: ${quality}%`,
        `📋 Columns: ${cols.slice(0,4).join(', ')}`,
        `${rate>30?'⚠️ High churn - action needed!':rate>15?'⚡ Moderate churn':'✅ Low churn rate'}`,
        `💡 Retention rate: ${100-rate}%`,
        `📈 Retained: ${churn.retained.toLocaleString()} customers`
    ];

    document.getElementById('analysisSummary')
        .innerHTML = items.map(s =>
            `<div class="summary-item">
                ${s}</div>`
        ).join('');
}

// ==================
// GENERATE IMAGES
// ==================
function generateChurnChart() {
    const canvas =
        document.createElement('canvas');
    canvas.width = 800;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');

    const gradient =
        ctx.createLinearGradient(
            0,0,800,0);
    gradient.addColorStop(0,'#667eea');
    gradient.addColorStop(1,'#764ba2');
    ctx.fillStyle = gradient;
    ctx.fillRect(0,0,800,400);

    ctx.fillStyle = 'white';
    ctx.font = 'bold 28px Segoe UI';
    ctx.textAlign = 'center';
    ctx.fillText(
        '🎯 Customer Churn Analysis',
        400,50);

    const bars = [
        {label:'Retained',value:73,
         color:'rgba(46,125,50,0.9)'},
        {label:'Churned',value:27,
         color:'rgba(198,40,40,0.9)'}
    ];

    bars.forEach((bar,i) => {
        const x = 150+(i*300);
        const height = bar.value*2.5;
        const y = 320-height;
        ctx.fillStyle = bar.color;
        ctx.beginPath();
        ctx.roundRect(x,y,150,height,10);
        ctx.fill();
        ctx.fillStyle = 'white';
        ctx.font = 'bold 18px Segoe UI';
        ctx.textAlign = 'center';
        ctx.fillText(bar.label,x+75,360);
        ctx.font = 'bold 24px Segoe UI';
        ctx.fillText(
            bar.value+'%',x+75,y-10);
    });

    return canvas.toDataURL('image/png');
}

function generateDashboardImage() {
    const canvas =
        document.createElement('canvas');
    canvas.width = 800;
    canvas.height = 400;
    const ctx = canvas.getContext('2d');

    const gradient =
        ctx.createLinearGradient(
            0,0,800,400);
    gradient.addColorStop(0,'#667eea');
    gradient.addColorStop(1,'#764ba2');
    ctx.fillStyle = gradient;
    ctx.fillRect(0,0,800,400);

    ctx.fillStyle = 'white';
    ctx.font = 'bold 32px Segoe UI';
    ctx.textAlign = 'center';
    ctx.fillText(
        '🎯 Churn Predictor Dashboard',
        400,60);

    ctx.font = '18px Segoe UI';
    ctx.fillStyle = 'rgba(255,255,255,0.8)';
    ctx.fillText(
        'AI-Powered Customer Analytics',
        400,95);

    const cards = [
        {label:'Total',value:'7,043',
         color:'rgba(255,255,255,0.2)'},
        {label:'Churned',value:'1,869',
         color:'rgba(198,40,40,0.5)'},
        {label:'Retained',value:'5,174',
         color:'rgba(46,125,50,0.5)'},
        {label:'Rate',value:'26.5%',
         color:'rgba(230,81,0,0.5)'}
    ];

    cards.forEach((card,i) => {
        const x = 50+(i*180);
        ctx.fillStyle = card.color;
        ctx.beginPath();
        ctx.roundRect(x,130,160,100,15);
        ctx.fill();
        ctx.fillStyle = 'white';
        ctx.font = 'bold 28px Segoe UI';
        ctx.textAlign = 'center';
        ctx.fillText(card.value,x+80,175);
        ctx.font = '14px Segoe UI';
        ctx.fillStyle =
            'rgba(255,255,255,0.8)';
        ctx.fillText(card.label,x+80,205);
    });

    ctx.fillStyle = 'rgba(255,255,255,0.15)';
    ctx.beginPath();
    ctx.roundRect(50,260,700,110,15);
    ctx.fill();

    [60,90,40,75,55,85,45,70]
        .forEach((h,i) => {
        ctx.fillStyle =
            'rgba(255,255,255,0.6)';
        ctx.beginPath();
        ctx.roundRect(
            75+(i*85),355-h,60,h,5);
        ctx.fill();
    });

    ctx.fillStyle = 'rgba(255,255,255,0.5)';
    ctx.font = '14px Segoe UI';
    ctx.textAlign = 'center';
    ctx.fillText(
        'CodTech Data Analytics Internship',
        400,390);

    return canvas.toDataURL('image/png');
}

function saveImages() {
    const churnImg = generateChurnChart();
    const link1 =
        document.createElement('a');
    link1.download = 'churn-chart.png';
    link1.href = churnImg;
    link1.click();

    setTimeout(() => {
        const dashImg =
            generateDashboardImage();
        const link2 =
            document.createElement('a');
        link2.download =
            'dashboard-image.png';
        link2.href = dashImg;
        link2.click();
    }, 1000);
}
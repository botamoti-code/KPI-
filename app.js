document.addEventListener('DOMContentLoaded', () => {
    const tableBody = document.getElementById('table-body');
    const saveBtn = document.getElementById('save-btn');
    const exportBtn = document.getElementById('export-btn');
    const importBtn = document.getElementById('import-btn');
    const importFile = document.getElementById('import-file');
    const clearBtn = document.getElementById('clear-btn');
    const toast = document.getElementById('toast');

    // データ構造の初期化
    let kpiData = [];
    const STORAGE_KEY = 'wafu_kpi_tracker_v1';

    // データの読み込み
    function loadData() {
        const stored = localStorage.getItem(STORAGE_KEY);
        if (stored) {
            try {
                kpiData = JSON.parse(stored);
                // 足りない月があれば補完
                if (kpiData.length < 12) {
                    initEmptyData();
                }
            } catch (e) {
                console.error("データの読み込みに失敗しました", e);
                initEmptyData();
            }
        } else {
            initEmptyData();
        }
    }

    // 空のデータ12ヶ月分を作成
    function initEmptyData() {
        kpiData = [];
        for (let i = 1; i <= 12; i++) {
            kpiData.push({
                month: i,
                instaTotal: '',
                instaInc: '',
                lineTotal: '',
                lineInc: '',
                videoViews: '',
                consults: '',
                contracts: ''
            });
        }
    }

    // 計算ロジック
    function calculateRates(row) {
        const instaInc = parseFloat(row.instaInc) || 0;
        const lineInc = parseFloat(row.lineInc) || 0;
        const videoViews = parseFloat(row.videoViews) || 0;
        const consults = parseFloat(row.consults) || 0;
        const contracts = parseFloat(row.contracts) || 0;

        let lineRate = 0;
        let videoRate = 0;
        let consultRate = 0;
        let contractRate = 0;

        if (instaInc > 0) lineRate = (lineInc / instaInc) * 100;
        if (lineInc > 0) videoRate = (videoViews / lineInc) * 100;
        if (videoViews > 0) consultRate = (consults / videoViews) * 100;
        if (consults > 0) contractRate = (contracts / consults) * 100;

        return {
            lineRate: Math.round(lineRate),
            videoRate: Math.round(videoRate),
            consultRate: Math.round(consultRate),
            contractRate: Math.round(contractRate)
        };
    }

    // テーブルの描画
    function renderTable() {
        tableBody.innerHTML = '';
        kpiData.forEach((row, index) => {
            const rates = calculateRates(row);
            
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <th class="sticky-col">${row.month}月</th>
                <td><input type="number" data-index="${index}" data-field="instaTotal" value="${row.instaTotal}"></td>
                <td><input type="number" data-index="${index}" data-field="instaInc" value="${row.instaInc}"></td>
                <td><input type="number" data-index="${index}" data-field="lineTotal" value="${row.lineTotal}"></td>
                <td><input type="number" data-index="${index}" data-field="lineInc" value="${row.lineInc}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.lineRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="videoViews" value="${row.videoViews}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.videoRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="consults" value="${row.consults}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.consultRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="contracts" value="${row.contracts}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.contractRate}%"></td>
            `;
            tableBody.appendChild(tr);
        });

        // 入力イベントの付与
        const inputs = tableBody.querySelectorAll('input:not([readonly])');
        inputs.forEach(input => {
            input.addEventListener('input', (e) => {
                const index = e.target.getAttribute('data-index');
                const field = e.target.getAttribute('data-field');
                kpiData[index][field] = e.target.value;
                
                // 同じ行の計算セルのみ更新する（フォーカスを外さないため）
                updateRowDisplay(index, tr => {
                    const rowInputs = e.target.closest('tr').querySelectorAll('td.col-percent input');
                    const newRates = calculateRates(kpiData[index]);
                    if(rowInputs.length === 4) {
                        rowInputs[0].value = newRates.lineRate + '%';
                        rowInputs[1].value = newRates.videoRate + '%';
                        rowInputs[2].value = newRates.consultRate + '%';
                        rowInputs[3].value = newRates.contractRate + '%';
                    }
                });
                
                // 自動保存
                saveData(false);
            });
        });
    }

    function updateRowDisplay(index, updateFn) {
        updateFn();
    }

    // データの保存
    function saveData(showToast = true) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(kpiData));
        if (showToast) displayToast('データを保存しました');
    }

    // トースト表示
    function displayToast(msg) {
        toast.textContent = msg;
        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    // 初期処理
    loadData();
    renderTable();

    // イベントリスナー
    saveBtn.addEventListener('click', () => saveData(true));

    clearBtn.addEventListener('click', () => {
        if (confirm('すべてのデータをリセットしますか？この操作は元に戻せません。')) {
            initEmptyData();
            saveData(true);
            renderTable();
            displayToast('データをリセットしました');
        }
    });

    // CSV出力
    exportBtn.addEventListener('click', () => {
        let csvContent = "data:text/csv;charset=utf-8,\uFEFF"; // BOM for Excel
        csvContent += "月,Instagram累計,Instagram増加数,LINE登録累計,LINE登録増加数,動画視聴数,個別相談数,成約数\r\n";
        
        kpiData.forEach(row => {
            const r = [
                row.month,
                row.instaTotal,
                row.instaInc,
                row.lineTotal,
                row.lineInc,
                row.videoViews,
                row.consults,
                row.contracts
            ];
            csvContent += r.join(",") + "\r\n";
        });

        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", "kpi_data.csv");
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    });

    // CSV読込
    importBtn.addEventListener('click', () => {
        importFile.click();
    });

    importFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = function(e) {
            const content = e.target.result;
            const lines = content.split('\n');
            if (lines.length > 1) {
                try {
                    // ヘッダー行を飛ばす
                    for (let i = 1; i <= 12; i++) {
                        if (i < lines.length) {
                            const cols = lines[i].split(',');
                            if (cols.length >= 8 && kpiData[i-1]) {
                                kpiData[i-1].instaTotal = cols[1].trim();
                                kpiData[i-1].instaInc = cols[2].trim();
                                kpiData[i-1].lineTotal = cols[3].trim();
                                kpiData[i-1].lineInc = cols[4].trim();
                                kpiData[i-1].videoViews = cols[5].trim();
                                kpiData[i-1].consults = cols[6].trim();
                                kpiData[i-1].contracts = cols[7].trim();
                            }
                        }
                    }
                    saveData(false);
                    renderTable();
                    displayToast('CSVを読み込みました');
                } catch (err) {
                    alert('CSVフォーマットが不正です。');
                }
            }
        };
        reader.readAsText(file);
        importFile.value = ''; // reset
    });
});

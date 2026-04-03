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
    const STORAGE_KEY = 'wafu_kpi_tracker_v3_april'; // データをリセットして4月始まりを強制

    let mediaNames = {
        media1: 'インスタグラム',
        media2: 'スレッズ',
        media3: 'ユーチューブ',
        media4: 'その他'
    };
    const MEDIA_STORAGE_KEY = 'wafu_kpi_media_names_v2';

    // グルコンメモ用
    let memoData = [];
    const MEMO_STORAGE_KEY = 'wafu_kpi_memos_v2';
    const memoContainer = document.getElementById('memo-container');
    const addMemoBtn = document.getElementById('add-memo-btn');

    // 設定管理・現在の目標用
    const SETTINGS_STORAGE_KEY = 'wafu_kpi_settings_v1';
    let appSettings = {
        clientName: '',
        webhookUrl: '',
        currentGoal: ''
    };
    const clientNameInput = document.getElementById('client-name');
    const webhookUrlInput = document.getElementById('webhook-url');
    const submitAdminBtn = document.getElementById('submit-admin-btn');
    const currentGoalInput = document.getElementById('current-goal');

    // タブ切り替え制御
    const tabBtns = document.querySelectorAll('.tab-btn');
    const tabPanels = document.querySelectorAll('.tab-panel');

    tabBtns.forEach(btn => {
        btn.addEventListener('click', () => {
            tabBtns.forEach(b => b.classList.remove('active'));
            tabPanels.forEach(p => p.classList.remove('active'));
            btn.classList.add('active');
            document.getElementById(btn.getAttribute('data-target')).classList.add('active');
        });
    });

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

        const storedMedia = localStorage.getItem(MEDIA_STORAGE_KEY);
        if (storedMedia) {
            try {
                mediaNames = { ...mediaNames, ...JSON.parse(storedMedia) };
            } catch (e) {}
        }
        document.getElementById('media1-name').value = mediaNames.media1;
        document.getElementById('media2-name').value = mediaNames.media2;
        document.getElementById('media3-name').value = mediaNames.media3;
        document.getElementById('media4-name').value = mediaNames.media4;

        // メモデータの読み込み
        const storedMemos = localStorage.getItem(MEMO_STORAGE_KEY);
        if (storedMemos) {
            try {
                memoData = JSON.parse(storedMemos);
            } catch (e) {
                console.error("メモデータの読み込みに失敗しました", e);
                memoData = [];
            }
        }

        // 設定データの読み込み
        const storedSettings = localStorage.getItem(SETTINGS_STORAGE_KEY);
        if (storedSettings) {
            try {
                appSettings = { ...appSettings, ...JSON.parse(storedSettings) };
            } catch (e) {}
        }

        // URLパラメータからの招待リンク自動設定
        const urlParams = new URLSearchParams(window.location.search);
        const hookUrl = urlParams.get('hook');
        if (hookUrl) {
            appSettings.webhookUrl = hookUrl;
            localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(appSettings));
            // 画面をスッキリさせるため、URLからクエリパラメータをこっそり消去する
            const cleanUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
            window.history.replaceState({path: cleanUrl}, '', cleanUrl);
        }

        clientNameInput.value = appSettings.clientName || '';
        if (webhookUrlInput) {
            webhookUrlInput.value = appSettings.webhookUrl || '';
        }
        currentGoalInput.value = appSettings.currentGoal || '';
    }

    // 空のデータ12ヶ月分を作成（4月始まり）
    function initEmptyData() {
        kpiData = [];
        const months = [4, 5, 6, 7, 8, 9, 10, 11, 12, 1, 2, 3];
        months.forEach(m => {
            kpiData.push({
                month: m,
                instaTotal: '',
                instaInc: '',
                threadsTotal: '',
                threadsInc: '',
                youtubeTotal: '',
                youtubeInc: '',
                otherTotal: '',
                otherInc: '',
                lineTotal: '',
                lineInc: '',
                videoViews: '',
                consults: '',
                contracts: '',
                sales: ''
            });
        });
    }

    // 計算ロジック
    function calculateRates(row) {
        const instaInc = parseFloat(row.instaInc) || 0;
        const threadsInc = parseFloat(row.threadsInc) || 0;
        const youtubeInc = parseFloat(row.youtubeInc) || 0;
        const otherInc = parseFloat(row.otherInc) || 0;
        
        const lineInc = parseFloat(row.lineInc) || 0;
        const videoViews = parseFloat(row.videoViews) || 0;
        const consults = parseFloat(row.consults) || 0;
        const contracts = parseFloat(row.contracts) || 0;

        const totalMediaInc = instaInc + threadsInc + youtubeInc + otherInc;

        let lineRate = 0;
        let videoRate = 0;
        let consultRate = 0;
        let contractRate = 0;

        if (totalMediaInc > 0) lineRate = (lineInc / totalMediaInc) * 100;
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
                <td><input type="number" data-index="${index}" data-field="instaTotal" value="${row.instaTotal || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="instaInc" value="${row.instaInc || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="threadsTotal" value="${row.threadsTotal || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="threadsInc" value="${row.threadsInc || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="youtubeTotal" value="${row.youtubeTotal || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="youtubeInc" value="${row.youtubeInc || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="otherTotal" value="${row.otherTotal || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="otherInc" value="${row.otherInc || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="lineTotal" value="${row.lineTotal || ''}"></td>
                <td><input type="number" data-index="${index}" data-field="lineInc" value="${row.lineInc || ''}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.lineRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="videoViews" value="${row.videoViews || ''}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.videoRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="consults" value="${row.consults || ''}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.consultRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="contracts" value="${row.contracts || ''}"></td>
                <td class="col-percent"><input type="text" readonly value="${rates.contractRate}%"></td>
                <td><input type="number" data-index="${index}" data-field="sales" value="${row.sales || ''}"></td>
            `;
            tableBody.appendChild(tr);
        });

        // 入力イベントの付与
        const inputs = tableBody.querySelectorAll('input:not([readonly])');
        inputs.forEach(input => {
            input.addEventListener('input', (e) => {
                const index = parseInt(e.target.getAttribute('data-index'));
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

            // 褒め言葉・励ましの判定（フォーカスが外れた時＝値が確定した時）
            input.addEventListener('change', (e) => {
                const index = parseInt(e.target.getAttribute('data-index'));
                const field = e.target.getAttribute('data-field');
                const currentValue = parseFloat(e.target.value);
                
                // 1行目（4月）は比較先の「先月」がないのでスキップ
                if (index > 0) {
                    const prevValue = parseFloat(kpiData[index - 1][field]);
                    if (!isNaN(currentValue) && !isNaN(prevValue)) {
                        if (currentValue > prevValue) {
                            // 先月より増えていたら褒める
                            showFeedbackTooltip(e.target, 'praise');
                        } else if (currentValue < prevValue) {
                            // 先月より下がっていたら励ます
                            showFeedbackTooltip(e.target, 'encourage');
                        }
                    }
                }
            });
        });
    }

    // 吹き出しを表示する関数 (種類: 'praise' または 'encourage')
    function showFeedbackTooltip(targetElement, type) {
        let messages = [];
        let bgColor = '#ffb3c1'; // デフォルトはピンク(増えた時用)

        if (type === 'praise') {
            messages = [
                "すごい！🌸", "頑張ったね！✨", "素晴らしい！🎉", "さすがです！💕", "いい調子！😊",
                "天才かも！💖", "その調子！📈", "成長してるね！🌱", "パーフェクト！🌟", "才能の塊！🤩",
                "限界突破！🔥", "数字が伸びてる！🙌", "このまま行こう！🚀", "最高です！👑", "エクセレント！💎",
                "努力の賜物！✨", "圧倒的成長！💪", "絶好調だね！😆", "神がかってます！👼", "止まらないね！💨"
            ];
        } else if (type === 'encourage') {
            // 下がった時の励ましバリエーション
            messages = [
                "また頑張ろうね！🍀", "焦らずいこう！🐌", "こんな月もあるよ☕️", "次に期待！🌟", "一休みして次へ！🍵",
                "ここから挽回！✊", "継続が大事！🌷", "マイペースでOK！🦥", "次はきっと伸びる！🎈"
            ];
            bgColor = '#90cdf4'; // 下がった時は少し落ち着いた水色などに変更
        }

        const randomMsg = messages[Math.floor(Math.random() * messages.length)];

        const tooltip = document.createElement('div');
        tooltip.className = 'praise-tooltip';
        tooltip.textContent = randomMsg;
        tooltip.style.backgroundColor = bgColor;
        
        // 吹き出しの尻尾の色も合わせるためのインライン追加スタイル要素
        const tailStyle = document.createElement('style');
        const uniqueId = 'tooltip-' + Math.random().toString(36).substr(2, 9);
        tooltip.id = uniqueId;
        tailStyle.innerHTML = `#${uniqueId}::after { border-color: ${bgColor} transparent transparent transparent; }`;
        tooltip.appendChild(tailStyle);

        // セル（td）の相対位置に配置するために、親tdにpositionを設定
        const parentTd = targetElement.closest('td');
        if (parentTd) {
            parentTd.style.position = 'relative';
            parentTd.appendChild(tooltip);

            // アニメーションのため少し遅れて表示
            requestAnimationFrame(() => {
                tooltip.classList.add('show');
            });

            // 約1.5秒後に消す
            setTimeout(() => {
                tooltip.classList.remove('show');
                setTimeout(() => {
                    tooltip.remove();
                }, 300); // fadeOutアニメーションの時間待つ
            }, 1500);
        }
    }

    function updateRowDisplay(index, updateFn) {
        updateFn();
    }

    // データの保存（KPI）
    function saveData(showToast = true) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(kpiData));
        if (showToast) displayToast('データを保存しました');
    }

    // ===== グルコンメモの処理 =====

    function saveMemos(showToast = false) {
        localStorage.setItem(MEMO_STORAGE_KEY, JSON.stringify(memoData));
        if (showToast) displayToast('メモを保存しました');
    }

    function renderMemos() {
        memoContainer.innerHTML = '';
        if (memoData.length === 0) {
            memoContainer.innerHTML = '<p style="text-align: center; color: var(--secondary-color); margin-top: 20px;">まだメモがありません。「新しいメモを追加」から記録を始めましょう！</p>';
            return;
        }

        memoData.forEach((memo, index) => {
            const card = document.createElement('div');
            card.className = 'memo-card';
            
            card.innerHTML = `
                <div class="memo-header">
                    <input type="date" class="memo-date" data-index="${index}" value="${memo.date || ''}">
                    <button class="memo-delete-btn" data-index="${index}" title="このメモを削除">✖</button>
                </div>
                <div class="memo-body">
                    <div class="memo-field">
                        <label class="memo-label">質問したこと</label>
                        <textarea class="memo-textarea" data-index="${index}" data-field="question" placeholder="グルコンで質問したこと、相談したことを記入...">${memo.question || ''}</textarea>
                    </div>
                    <div class="memo-field">
                        <label class="memo-label">もらったアドバイス</label>
                        <textarea class="memo-textarea" data-index="${index}" data-field="advice" placeholder="もらった回答や、気づき、次にやることを記入...">${memo.advice || ''}</textarea>
                    </div>
                </div>
            `;
            memoContainer.appendChild(card);
        });

        // 削除イベントの付与
        document.querySelectorAll('.memo-delete-btn').forEach(btn => {
            btn.addEventListener('click', (e) => {
                const idx = e.target.getAttribute('data-index');
                if (confirm('このメモを削除してもよろしいですか？')) {
                    memoData.splice(idx, 1);
                    saveMemos(true);
                    renderMemos();
                }
            });
        });

        // 入力イベントの付与（自動保存）
        document.querySelectorAll('.memo-date, .memo-textarea').forEach(input => {
            input.addEventListener('input', (e) => {
                const idx = e.target.getAttribute('data-index');
                if (e.target.classList.contains('memo-date')) {
                    memoData[idx].date = e.target.value;
                } else {
                    const field = e.target.getAttribute('data-field');
                    memoData[idx][field] = e.target.value;
                }
                saveMemos(false);
            });
        });
    }

    addMemoBtn.addEventListener('click', () => {
        const today = new Date().toISOString().split('T')[0];
        memoData.unshift({
            date: today,
            question: '',
            advice: ''
        });
        saveMemos(false);
        renderMemos();
    });

    function displayToast(msg) {
        toast.textContent = msg;
        toast.classList.add('show');
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    // 設定とデータ提出ロジック
    function saveSettings() {
        if (clientNameInput) appSettings.clientName = clientNameInput.value;
        if (webhookUrlInput) appSettings.webhookUrl = webhookUrlInput.value;
        if (currentGoalInput) appSettings.currentGoal = currentGoalInput.value;
        localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(appSettings));
    }

    if (clientNameInput) clientNameInput.addEventListener('input', saveSettings);
    if (webhookUrlInput) webhookUrlInput.addEventListener('input', saveSettings);
    if (currentGoalInput) currentGoalInput.addEventListener('input', saveSettings);

    submitAdminBtn.addEventListener('click', async () => {
        saveData(false);
        saveMemos(false);
        saveSettings();

        if (!appSettings.clientName) {
            alert('ご自身の名前（クライアント名）を入力してください');
            return;
        }

        if (!appSettings.webhookUrl) {
            alert('管理者用のWebhook URLが設定されていません');
            return;
        }

        submitAdminBtn.textContent = '⏱️ 送信中...';
        submitAdminBtn.disabled = true;

        const payload = {
            clientName: appSettings.clientName,
            timestamp: new Date().toISOString(),
            currentGoal: appSettings.currentGoal,
            kpiData: kpiData,
            memoData: memoData,
            mediaNames: mediaNames
        };

        try {
            // Google Apps ScriptへPOST。CORS制限回避のためmode: 'no-cors'を使用
            await fetch(appSettings.webhookUrl, {
                method: 'POST',
                mode: 'no-cors',
                headers: {
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify(payload)
            });
            displayToast('データを管理者に送信しました！');
        } catch (error) {
            console.error('送信エラー', error);
            alert('送信に失敗しました。URLが正しいか、ネットワーク接続を確認してください。');
        } finally {
            submitAdminBtn.innerHTML = '🚀 管理者へデータを送信する';
            submitAdminBtn.disabled = false;
        }
    });

    // 初期処理
    loadData();
    renderTable();
    renderMemos();

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
        csvContent += `月,${mediaNames.media1}累計,${mediaNames.media1}増加数,${mediaNames.media2}累計,${mediaNames.media2}増加数,${mediaNames.media3}累計,${mediaNames.media3}増加数,${mediaNames.media4}累計,${mediaNames.media4}増加数,LINE登録累計,LINE登録増加数,動画視聴数,個別相談数,成約数\r\n`;
        
        kpiData.forEach(row => {
            const r = [
                row.month,
                row.instaTotal || '',
                row.instaInc || '',
                row.threadsTotal || '',
                row.threadsInc || '',
                row.youtubeTotal || '',
                row.youtubeInc || '',
                row.otherTotal || '',
                row.otherInc || '',
                row.lineTotal || '',
                row.lineInc || '',
                row.videoViews || '',
                row.consults || '',
                row.contracts || ''
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
                            if (cols.length >= 14 && kpiData[i-1]) {
                                kpiData[i-1].instaTotal = cols[1].trim();
                                kpiData[i-1].instaInc = cols[2].trim();
                                kpiData[i-1].threadsTotal = cols[3].trim();
                                kpiData[i-1].threadsInc = cols[4].trim();
                                kpiData[i-1].youtubeTotal = cols[5].trim();
                                kpiData[i-1].youtubeInc = cols[6].trim();
                                kpiData[i-1].otherTotal = cols[7].trim();
                                kpiData[i-1].otherInc = cols[8].trim();
                                kpiData[i-1].lineTotal = cols[9].trim();
                                kpiData[i-1].lineInc = cols[10].trim();
                                kpiData[i-1].videoViews = cols[11].trim();
                                kpiData[i-1].consults = cols[12].trim();
                                kpiData[i-1].contracts = cols[13].trim();
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

    // ヘッダー入力の保存
    const headerInputs = document.querySelectorAll('.header-input');
    headerInputs.forEach(input => {
        input.addEventListener('input', (e) => {
            const id = e.target.id;
            const key = id.split('-')[0]; // media1
            mediaNames[key] = e.target.value;
            localStorage.setItem(MEDIA_STORAGE_KEY, JSON.stringify(mediaNames));
        });
    });

    // -----------------------------------------
    // 運営者用（管理者）メニューの制御
    // -----------------------------------------
    const showAdminBtn = document.getElementById('show-admin-login-btn');
    const adminPanel = document.getElementById('admin-secret-panel');
    const adminWebhookUrlInput = document.getElementById('admin-webhook-url');
    const adminGenerateBtn = document.getElementById('generate-btn');
    const shareArea = document.getElementById('share-area');
    const shareUrlText = document.getElementById('share-url-text');

    if (showAdminBtn) {
        showAdminBtn.addEventListener('click', () => {
            // すでに開いている場合は閉じる
            if (adminPanel.style.display === 'block') {
                adminPanel.style.display = 'none';
                showAdminBtn.textContent = '運営者用メニューを開く';
                return;
            }

            // パスワードの入力要求
            const password = prompt('運営者用パスワードを入力してください:');
            if (password === 'admin') {
                // パスワード正解
                adminPanel.style.display = 'block';
                showAdminBtn.textContent = '運営者用メニューを閉じる';
                // 既存の保存済み設定があれば表示
                const storedSettings = localStorage.getItem(SETTINGS_STORAGE_KEY);
                if (storedSettings) {
                    try {
                        const settings = JSON.parse(storedSettings);
                        if (settings.webhookUrl) {
                            adminWebhookUrlInput.value = settings.webhookUrl;
                        }
                    } catch (e) {}
                }
            } else if (password !== null) {
                // キャンセル以外で不正解の場合
                alert('パスワードが違います。');
            }
        });
    }

    if (adminGenerateBtn) {
        adminGenerateBtn.addEventListener('click', () => {
            const hookUrl = adminWebhookUrlInput.value.trim();
            if (!hookUrl) {
                alert('Webhook URLを入力してください。');
                return;
            }

            // 現在のURL（index.html）をベースに招待リンクを生成
            let baseUrl = window.location.protocol + "//" + window.location.host + window.location.pathname;
            const inviteUrl = `${baseUrl}?hook=${encodeURIComponent(hookUrl)}`;

            // UIに表示
            shareUrlText.textContent = inviteUrl;
            shareArea.style.display = 'block';

            // クリップボードにコピー
            navigator.clipboard.writeText(inviteUrl).then(() => {
                displayToast('招待用URLをコピーしました！');
            }).catch(err => {
                console.error('コピー失敗', err);
                displayToast('コピーに失敗しました。手動でコピーしてください。');
            });
            
            // ついでに管理者のブラウザにも設定を保存しておく
            let settings = { clientName: '', webhookUrl: hookUrl, currentGoal: '' };
            const stored = localStorage.getItem(SETTINGS_STORAGE_KEY);
            if (stored) {
                try {
                    settings = { ...JSON.parse(stored), webhookUrl: hookUrl };
                } catch(e) {}
            }
            localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(settings));
            // 変数も更新
            appSettings.webhookUrl = hookUrl;
        });
    }
});

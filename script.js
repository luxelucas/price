document.addEventListener('DOMContentLoaded', () => {
    // --- State ---
    let sheetData = []; // Array of { thickness, type, stock, price }
    let currentResults = []; // Currently displayed items 
    let selections = {}; // itemID -> quantity selected

    // --- DOM Elements ---
    const connectionStatus = document.getElementById('connection-status');
    const searchBtn = document.getElementById('search-btn');
    const thicknessInputs = document.querySelectorAll('.thick-input');
    const resultsPanel = document.getElementById('results-panel');
    const resultsBody = document.getElementById('results-body');
    const summaryQuoteBtn = document.getElementById('summary-quote-btn');
    const resultStats = document.getElementById('result-stats');
    
    // Modal
    const modal = document.getElementById('quote-modal');
    const closeModalBtn = document.getElementById('close-modal-btn');
    const quoteSummaryBody = document.getElementById('quote-summary-body');

    const toast = document.getElementById('toast');

    // --- Currency Formatter ---
    const formatCurrency = new Intl.NumberFormat('zh-TW', {
        style: 'currency', currency: 'TWD', minimumFractionDigits: 0
    }).format;

    // --- Hardcoded Google Sheet ID ---
    const sheetId = '1942RnYmUIFoWnVgLMQr36sh3dlQguTBKQ_FmpMqGmlo';

    // --- Fetch Google Sheet Data ---
    const fetchSheet = () => {
        connectionStatus.textContent = '連線中...';
        connectionStatus.className = 'status-msg';

        // Setup JSONP callback
        window.google = window.google || {};
        window.google.visualization = window.google.visualization || {};
        window.google.visualization.Query = window.google.visualization.Query || {};
        
        window.google.visualization.Query.setResponse = function(data) {
            try {
                if (data.status === 'error') {
                    throw new Error(data.errors[0].message);
                }
                
                const cols = data.table.cols;
                let colMap = { thickness: -1, stock: -1, type: -1, price: -1 };
                
                cols.forEach((col, idx) => {
                    if(!col.label) return;
                    const label = col.label.toLowerCase();
                    if (label.includes('厚')) colMap.thickness = idx;
                    if (label.includes('片') || label.includes('量')) colMap.stock = idx;
                    if (label.includes('種') || label.includes('類')) colMap.type = idx;
                    if (label.includes('價')) colMap.price = idx;
                });

                // If header matching fails, assume standard order: Thickness(0), Stock(1), Type(2), Price(3)
                if (colMap.thickness === -1) colMap = { thickness: 0, stock: 1, type: 2, price: 3 };

                sheetData = [];
                data.table.rows.forEach((row, idx) => {
                    // skip empty rows
                    if (!row.c || !row.c[colMap.thickness] || row.c[colMap.thickness].v === null) return;

                    sheetData.push({
                        id: `item_${idx}`,
                        thickness: parseFloat(row.c[colMap.thickness].v) || 0,
                        stock: parseInt(row.c[colMap.stock]?.v) || 0,
                        type: String(row.c[colMap.type]?.v || ''),
                        price: parseFloat(row.c[colMap.price]?.v) || 0,
                    });
                });

                connectionStatus.textContent = `連線成功！共載入 ${sheetData.length} 筆資料。`;
                connectionStatus.className = 'status-msg success';
                showToast('資料載入成功');

            } catch (error) {
                console.error(error);
                connectionStatus.textContent = error.message || '連線失敗，請確認該表單是否已設定為公開。';
                connectionStatus.className = 'status-msg text-danger';
            } finally {
                // Cleanup script tag
                const oldScript = document.getElementById('gviz-script');
                if (oldScript) oldScript.remove();
            }
        };

        // Inject script tag
        const oldScript = document.getElementById('gviz-script');
        if (oldScript) oldScript.remove();
        
        const script = document.createElement('script');
        script.id = 'gviz-script';
        script.src = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:json`;
        script.onerror = () => {
            connectionStatus.textContent = '網路異常或表單未正確公開，無法建立連線。';
            connectionStatus.className = 'status-msg text-danger';
        };
        document.body.appendChild(script);
    };

    // --- Auto Fetch on Load ---
    fetchSheet();

    // --- Search Functionality ---
    searchBtn.addEventListener('click', () => {
        if (sheetData.length === 0) {
            showToast('請先連線並載入 Google Sheet 資料！');
            return;
        }

        const targets = [];
        thicknessInputs.forEach(input => {
            const val = parseFloat(input.value);
            if (!isNaN(val)) targets.push(val);
        });

        if (targets.length === 0) {
            showToast('請至少輸入一個厚度進行搜尋。');
            return;
        }

        // Filter data match
        currentResults = sheetData.filter(item => targets.includes(item.thickness));
        
        // Sort by thickness then type
        currentResults.sort((a,b) => a.thickness - b.thickness || a.type.localeCompare(b.type));

        renderResults(currentResults, targets);
    });

    const renderResults = (results, targetArr) => {
        resultsPanel.style.display = 'block';
        resultsBody.innerHTML = '';
        
        resultStats.textContent = `針對厚度 (${targetArr.join(', ')}) 找到 ${results.length} 筆符合項`;

        if (results.length === 0) {
            resultsBody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding: 2rem; color: #94a3b8;">找不到符合厚度的鋼板項目</td></tr>`;
            return;
        }

        results.forEach(item => {
            const tr = document.createElement('tr');
            
            // Previous Selection recovery (if mapped by ID)
            const prevQty = parseInt(selections[item.id]) || 0;
            const maxOptions = Math.min(10, item.stock);
            
            let optionsHTML = `<option value="0">0</option>`;
            for(let i = 1; i <= maxOptions; i++) {
                optionsHTML += `<option value="${i}" ${prevQty === i ? 'selected' : ''}>${i}</option>`;
            }

            tr.innerHTML = `
                <td style="font-weight:bold; color: #3b82f6;">${item.thickness} mm</td>
                <td>${item.type}</td>
                <td>${item.stock} <span style="font-size:0.8rem; color:#94a3b8">片</span></td>
                <td style="color:#10b981;">${formatCurrency(item.price)}</td>
                <td>
                    <select class="order-input" data-id="${item.id}">
                        ${optionsHTML}
                    </select>
                </td>
            `;
            resultsBody.appendChild(tr);
        });

        // Add Listeners to inputs
        document.querySelectorAll('.order-input').forEach(inp => {
            inp.addEventListener('change', (e) => {
                const id = e.target.dataset.id;
                let val = parseInt(e.target.value);
                
                if (isNaN(val) || val === 0) {
                    delete selections[id];
                } else {
                    selections[id] = val;
                }
            });
        });

        // Scroll to results
        resultsPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
    };

    // --- Modal / Quote Summary ---
    summaryQuoteBtn.addEventListener('click', () => {
        const selectedItems = Object.keys(selections).map(id => {
            const item = sheetData.find(i => i.id === id);
            return {
                ...item,
                orderQty: selections[id]
            };
        });

        if (selectedItems.length === 0) {
            showToast('請先搜尋並輸入查閱數量！');
            return;
        }

        quoteSummaryBody.innerHTML = '';

        selectedItems.forEach(item => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${item.thickness}</td>
                <td>${item.type}</td>
                <td>${item.orderQty}</td>
                <td>${formatCurrency(item.price)}</td>
            `;
            quoteSummaryBody.appendChild(tr);
        });

        modal.classList.add('active');
    });

    closeModalBtn.addEventListener('click', () => {
        modal.classList.remove('active');
    });

    const emailBtn = document.getElementById('email-btn');
    emailBtn.addEventListener('click', () => {
        const selectedItems = Object.keys(selections).map(id => {
            const item = sheetData.find(i => i.id === id);
            return {
                ...item,
                orderQty: selections[id]
            };
        });

        if(selectedItems.length === 0) return;

        let bodyText = "您好，這是我預訂的鋼板清單：\n\n";
        selectedItems.forEach((item, index) => {
            bodyText += `${index + 1}. 厚度 ${item.thickness}mm | 種類 ${item.type} | 數量 ${item.orderQty} 片\n`;
        });
        bodyText += "\n再請協助確認，謝謝！";

        const subject = "鋼板預訂需求";
        window.location.href = `mailto:?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(bodyText)}`;
    });

    // --- Toast ---
    let toastTimeout;
    function showToast(msg) {
        toast.textContent = msg;
        toast.classList.add('show');
        clearTimeout(toastTimeout);
        toastTimeout = setTimeout(() => toast.classList.remove('show'), 3000);
    }
});

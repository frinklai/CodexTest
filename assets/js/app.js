        const fileInput = document.getElementById('excel-file-input');
        const fileNameDisplay = document.getElementById('file-name-display');
        const excelContent = document.getElementById('excel-content');
        const ganttControls = document.getElementById('gantt-controls');
        const generateGanttBtn = document.getElementById('generate-gantt-btn');
        const exportExcelBtn = document.getElementById('export-excel-btn');
        const startDateInput = document.getElementById('start-date');
        const endDateInput = document.getElementById('end-date');
        const errorMessage = document.getElementById('error-message');
        const dataSummary = document.getElementById('data-summary');
        const errorModal = document.getElementById('error-modal');
        const modalConfirmBtn = document.getElementById('modal-confirm-btn');
        const highlightToggle = document.getElementById('highlight-toggle');
        const freezeColsInput = document.getElementById('freeze-cols');
        const collapseAllBtn = document.getElementById('collapse-all-btn');
        const expandAllBtn = document.getElementById('expand-all-btn');
        const contextMenu = document.getElementById('context-menu');
        const viewChangesBtn = document.getElementById('view-changes-btn');
        const changesModal = document.getElementById('changes-modal');
        const changesList = document.getElementById('changes-list');
        const changesModalCloseBtn = document.getElementById('changes-modal-close-btn');
        const downloadChangesBtn = document.getElementById('download-changes-btn');
        const collapsedState = new Map();

        // 用於儲存全域資料
        let currentWorkbook = null;
        let currentDataArray = [];     // 包含標頭的原始資料
        let generatedDateColumns = []; // 動態日期欄
        let dateChanges = {};          // 日期變更紀錄
        let devStartDateColIndex = -1;
        let dueDateColIndex = -1;
        let activeFilters = {};
        let isSelecting = false;
        let selectionStartCell = null;
        let contextMenuTargetCol = null;
        let manualDateCellTexts = {};  // key: `${rowIndex}-${colIndex}` -> 使用者輸入的字串
        let dateCellTexts = {};
        let currentRowMeta = [];  // 對應每一列的 outline level（Excel 群組層級）
        let keepNonmatchingParents = true; // 勾：保留未命中父列；不勾：不保留
        let selectionAnchorCell = null; // Shift 連選錨點

        // ===== 配色主題 =====
        const THEME = {
            row: {
                topLevelHex: '#bbf7d0',   // 深綠（Tailwind green-200）
                childHex:    '#dcfce7',   // 淺綠（Tailwind green-100）
            },
            gantt: {
                rootHex:  '#84C1FF',  // 深藍
                childHex: '#C4E1FF',  // 淺藍
            }
        };

        // 小工具：HEX -> ARGB（ExcelJS用），預設不透明 FF
        function hexToARGB(hex, alpha='FF'){
            const h = hex.replace('#','').trim();
            const r = h.length===3 ? h[0]+h[0] : h.slice(0,2);
            const g = h.length===3 ? h[1]+h[1] : h.slice(2,4);
            const b = h.length===3 ? h[2]+h[2] : h.slice(4,6);
            return (alpha + r + g + b).toUpperCase();
        }

        // 只有「白底且非甘特藍」的格子才允許吃黃色高亮
        function canHighlight(cell) {
            if (!cell) return false;
            const isWhiteBase = cell.classList.contains('bg-white');
            const isGanttColored = cell.dataset && (cell.dataset.gantt === 'root' || cell.dataset.gantt === 'child');
            return isWhiteBase && !isGanttColored;
        }


        // 保存欄位顯示/寬度狀態：data=原始資料欄，date=動態日期欄（用表頭文字做 key）
        const colState = {
            data: { hidden: new Set(), widths: {} },  // key: 資料欄 index
            date: { hidden: new Set(), widths: {} },  // key: 日期表頭文字 e.g. "2025/01/05"
        };
        function isDataCol(index) {
            const totalOriginal = currentDataArray?.[0]?.length || 0;
            return index < totalOriginal;
        }
        function getHeaderTextByIndex(table, colIndex) {
            const th = table.querySelector('thead tr')?.children[colIndex];
            return (th ? th.textContent.trim() : '') || '';
        }

        // 渲染後把隱藏/寬度套回去（一定要在 append table 之後、sticky 前）
        function applySavedColumnStates() {
            const table = document.getElementById('data-table');
            if (!table) return;
            const headerRow = table.querySelector('thead tr');
            if (!headerRow) return;

            const totalCols = headerRow.children.length;
            for (let i = 0; i < totalCols; i++) {
                const isData = isDataCol(i);
                const key = isData ? i : getHeaderTextByIndex(table, i);

                // 隱藏
                const hiddenSet = isData ? colState.data.hidden : colState.date.hidden;
                if (hiddenSet.has(key)) {
                    table.querySelectorAll('tr').forEach(row => {
                        if (row.children[i]) row.children[i].style.display = 'none';
                    });
                }

                // 寬度
                const widthMap = isData ? colState.data.widths : colState.date.widths;
                const w = widthMap[key];
                if (w) {
                    table.querySelectorAll('tr').forEach(row => {
                        const cell = row.children[i];
                        if (!cell) return;
                        cell.style.width = `${w}px`;
                        cell.style.minWidth = `${w}px`;
                        cell.style.maxWidth = `${w}px`;
                    });
                }
            }
        }

        // =================================================================================
        // ============================ 核心功能：讀取與渲染 ============================
        // =================================================================================
        function renderTable() {
            try {
                if (currentDataArray.length === 0) {
                excelContent.innerHTML = '';
                return;
                }

                const table = document.createElement('table');
                table.id = 'data-table';
                const thead = document.createElement('thead');
                const tbody = document.createElement('tbody');

                // 表頭：原始 + 週區間（使用 HTML 換行顯示）
                const headerTr = document.createElement('tr');

                // 原始欄
                (currentDataArray[0] || []).forEach(cellData => {
                const th = document.createElement('th');
                th.innerHTML = getCellText(cellData);
                headerTr.appendChild(th);
                });

                // 週區間欄（顯示用 label，資料對應用 weekKey）
                generatedDateColumns.forEach(week => {
                const th = document.createElement('th');
                th.innerHTML = weekLabelHTML(week); // 顯示兩行
                headerTr.appendChild(th);
                });
                thead.appendChild(headerTr);

                // 內容列
                const originalColumnCount = currentDataArray[0].length;
                currentDataArray.slice(1).forEach((rowData, rowIndex) => {
                const tr = document.createElement('tr');
                tr.dataset.originalIndex = rowIndex;

                // 原始資料儲存格
                rowData.forEach(cellData => {
                    const td = document.createElement('td');
                    const cellText = getCellText(cellData);
                    if (cellData && cellData.hyperlink) {
                    const link = document.createElement('a');
                    link.href = cellData.hyperlink;
                    link.target = '_blank';
                    link.rel = 'noopener noreferrer';
                    link.textContent = cellText;
                    link.classList.add('text-blue-600','underline');
                    td.appendChild(link);
                    td.dataset.hyperlink = cellData.hyperlink;
                    } else {
                    td.innerHTML = cellText || '&nbsp;';
                    }
                    tr.appendChild(td);
                });

                // 動態「週」儲存格：用 weekKey 取值（不受表頭文字影響）
                for (let i = 0; i < generatedDateColumns.length; i++) {
                    const td = document.createElement('td');
                    const w = generatedDateColumns[i];
                    const key = weekKey(w);
                    const rowMap = dateCellTexts[rowIndex] || {};
                    const saved = rowMap[key];
                    td.textContent = (saved && saved.trim() !== '') ? saved : '\u00A0';
                    tr.appendChild(td);
                }

                tbody.appendChild(tr);
                });

                table.appendChild(thead);
                table.appendChild(tbody);
                excelContent.innerHTML = '';
                excelContent.appendChild(table);

                // 套回狀態 + 互動
                applySavedColumnStates();
                styleGeneratedTable();
                applyStickyLayout();
                addFilterIconsToHeader();
                addColumnResizeHandles();
                addDateEditing();
                addGeneratedDateTextEditing(); // 讓週欄位可輸入並以 weekKey 保存
            } catch (error) {
                console.error("渲染表格時發生錯誤:", error);
                excelContent.innerHTML = `<p class="text-red-500 p-4">渲染表格時發生錯誤。</p>`;
            }
        }



        function addGeneratedDateTextEditing() {
            const table = document.getElementById('data-table');
            if (!table || currentDataArray.length === 0) return;

            const headerRow = table.querySelector('thead tr');
            const originalColumnCount = currentDataArray[0].length;
            const bodyRows = table.querySelectorAll('tbody tr');

            bodyRows.forEach((row) => {
                const originalRowIndex = parseInt(row.dataset.originalIndex, 10);
                if (isNaN(originalRowIndex)) return;

                for (let col = originalColumnCount; col < headerRow.children.length; col++) {
                const cell = row.children[col];
                if (!cell) continue;

                cell.addEventListener('dblclick', () => {
                    // 用欄位對應的「週物件」產生穩定 key
                    const weekIdx = col - originalColumnCount;
                    const w = generatedDateColumns[weekIdx];
                    if (!w) return;
                    const key = weekKey(w); // ← 穩定 key（YYYY-MM-DD~YYYY-MM-DD）

                    const oldText = (dateCellTexts[originalRowIndex]?.[key] ?? cell.textContent).trim();
                    const input = document.createElement('input');
                    input.type = 'text';
                    input.className = 'w-full h-full border border-gray-300 rounded px-1 outline-none';
                    input.value = oldText;

                    cell.innerHTML = '';
                    cell.appendChild(input);
                    input.focus();
                    input.select();

                    const commit = () => {
                    const newVal = input.value.trim();
                    if (!dateCellTexts[originalRowIndex]) dateCellTexts[originalRowIndex] = {};
                    if (newVal === '') delete dateCellTexts[originalRowIndex][key];
                    else dateCellTexts[originalRowIndex][key] = newVal;
                    cell.textContent = newVal || '\u00A0';
                    };
                    const cancel = () => { cell.textContent = oldText || '\u00A0'; };

                    input.addEventListener('keydown', (e) => {
                    if (e.key === 'Enter') commit();
                    else if (e.key === 'Escape') cancel();
                    });
                    input.addEventListener('blur', commit);
                });
                }
            });
        }




        // 讓 tbody 的每一列有 hover 效果（不蓋掉已上色格，因為 CSS 只針對 .bg-white）
        excelContent.addEventListener('mouseover', (e) => {
        const tr = e.target.closest('tbody tr');
        if (!tr) return;
        tr.classList.add('row-hovering');
        });
        excelContent.addEventListener('mouseout', (e) => {
        const tr = e.target.closest('tbody tr');
        if (!tr) return;
        tr.classList.remove('row-hovering');
        });





        fileInput.addEventListener('change', (event) => {
            const file = event.target.files && event.target.files[0];

            // ▶ 沒選檔或按了「取消」：什麼都不做，維持現況
            if (!file) {
                // 可選：重置 input 值，避免下次選同一個檔案時不觸發 change
                // event.target.value = '';
                return;
            }

            // ▶ 有選檔，這裡才做重置/清空
            errorMessage.textContent = '';
            dataSummary.textContent = '';
            highlightToggle.checked = false;
            activeFilters = {};
            generatedDateColumns = [];
            dateChanges = {};
            selectionStartCell = null;
            freezeColsInput.value = 2;

            // 換檔案時清空欄狀態
            colState.data.hidden.clear(); colState.data.widths = {};
            colState.date.hidden.clear(); colState.date.widths = {};

            fileNameDisplay.textContent = `已選擇檔案：${file.name}`;
            fileNameDisplay.classList.remove('text-gray-500');
            fileNameDisplay.classList.add('text-gray-700', 'font-medium');

            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                const data = new Uint8Array(e.target.result);
                currentWorkbook = XLSX.read(data, { type: 'array', cellDates: true, cellNF: true });
                const firstSheetName = currentWorkbook.SheetNames[0];
                if (!firstSheetName) {
                    excelContent.innerHTML = `<p class="text-red-500 p-4">錯誤：找不到任何工作表。</p>`;
                    return;
                }
                const worksheet = currentWorkbook.Sheets[firstSheetName];

                const wsRows = worksheet['!rows'] || [];
                currentRowMeta = wsRows.map(r => (r && (r.level ?? r.outlineLevel)) ?? 0);

                const dataArray = [];
                const range = XLSX.utils.decode_range(worksheet['!ref']);
                for (let R = range.s.r; R <= range.e.r; ++R) {
                    const row = [];
                    for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cell_ref = XLSX.utils.encode_cell({ c: C, r: R });
                    const cellObj = worksheet[cell_ref];
                    if (!cellObj) { row.push(null); continue; }

                    if (cellObj.l && cellObj.l.Target) {
                        row.push({ text: cellObj.w || cellObj.v, hyperlink: cellObj.l.Target });
                    } else if (cellObj.t === 'n' && cellObj.w && cellObj.w.includes('%')) {
                        row.push(cellObj.w);
                    } else if (cellObj.t === 'd') {
                        row.push(cellObj.v);
                    } else {
                        row.push(cellObj.v);
                    }
                    }
                    dataArray.push(row);
                }

                processAndCleanData(dataArray);

                if (currentDataArray.length <= 1) {
                    dataSummary.textContent = '檔案中沒有資料或只有標題列。';
                    renderTable();
                    ganttControls.classList.add('hidden');
                    return;
                }

                const itemCount = currentDataArray.length - 1;
                dataSummary.textContent = `總共 ${itemCount} 項需求`;

                renderTable();
                ganttControls.classList.remove('hidden');

                } catch (error) {
                console.error("解析 Excel 檔案時發生錯誤:", error);
                excelContent.innerHTML = `<p class="text-red-500 p-4">無法解析此檔案。請確認檔案格式是否正確。</p>`;
                currentWorkbook = null;
                ganttControls.classList.add('hidden');
                } finally {
                // 可選：成功/失敗後都重置 input 值，讓下次選同一檔也能觸發 change
                // event.target.value = '';
                }
            };
            reader.readAsArrayBuffer(file);
        });


        function processAndCleanData(data){
            const headerRow = data[0] || [];
            // 記住欄位位置，但不做任何自動填值
            devStartDateColIndex = headerRow.indexOf('Dev Start Date');
            dueDateColIndex = headerRow.indexOf('Due Date');

            // 直接保存原始資料：不改動日期欄
            currentDataArray = data;
            }


        // =================================================================================
        // ============================ UI 控制與事件處理 ============================
        // =================================================================================
        generateGanttBtn.addEventListener('click', () => {
            if (currentDataArray.length === 0) return;

            const startDate = startDateInput.value;
            const endDate = endDateInput.value;

            if (!startDate || !endDate) { errorMessage.textContent = '請同時選擇起始與結束日期。'; return; }
            const start = new Date(startDate);
            const end = new Date(endDate);
            if (start > end) { errorMessage.textContent = '起始日期不可晚於結束日期。'; return; }
            errorMessage.textContent = '';

            // 取得「包含起始日」那週的週一
            const startMonday = new Date(start);
            const dayS = startMonday.getDay();              // 0=Sun..6=Sat
            const diffS = (dayS + 6) % 7;                   // 往回移到週一
            startMonday.setDate(startMonday.getDate() - diffS);

            // 取得「包含結束日」那週的週五
            const endFriday = new Date(end);
            const dayE = endFriday.getDay();
            const mondayToFriday = 4;                        // 週一+4天=週五
            const diffE = (dayE + 6) % 7;                   // 距離週一的天數
            endFriday.setDate(endFriday.getDate() - diffE + mondayToFriday);

            // 重新產生「週」欄位：每一格是 { start: 週一, end: 週五 }
            generatedDateColumns = [];
            let cur = new Date(startMonday);
            while (cur <= endFriday) {
                const weekStart = new Date(cur);
                const weekEnd = new Date(cur); weekEnd.setDate(weekEnd.getDate() + 4); // 週五
                generatedDateColumns.push({ start: weekStart, end: weekEnd });
                cur.setDate(cur.getDate() + 7); // 下一週
            }

            renderTable();
            applyFiltersToTable();
            colorGanttCells();
        });


        exportExcelBtn.addEventListener('click', async () => {
            const table = document.getElementById('data-table');
            if (!table) {
                document.getElementById('modal-error-text').textContent = '表格不存在';
                errorModal.classList.remove('hidden');
                return;
            }

            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Gantt Chart');
            const headerRowEl = table.querySelector('thead tr');
            if (!headerRowEl) return;

            // 只輸出可見欄
            const visibleColIndices = Array.from(headerRowEl.children)
                .map((th, index) => (th.style.display !== 'none' ? index : -1))
                .filter(index => index !== -1);

            if (visibleColIndices.length === 0) {
                document.getElementById('modal-error-text').textContent = '沒有可見的資料以供匯出。';
                errorModal.classList.remove('hidden');
                return;
            }

            // 欄寬
            const colWidths = visibleColIndices.map(index => headerRowEl.children[index].getBoundingClientRect().width);
            worksheet.columns = colWidths.map(width => ({ width: Math.max(width / 7.5, 10) }));

            // 方便判斷哪些是「產生的日期欄」
            const originalColumnCount = (window.currentDataArray?.[0]?.length) || 0;

            // 寫入表格
            const rowsToExport = Array.from(table.querySelectorAll('tr')); // 列：全部輸出

            rowsToExport.forEach((htmlRow, rowIndex) => {
                const row = worksheet.getRow(rowIndex + 1);
                let excelCellIndex = 1;

                const isHeader = htmlRow.parentElement.tagName === 'THEAD';

                visibleColIndices.forEach(colIndex => {
                const htmlCell = htmlRow.children[colIndex];
                if (!htmlCell) return;

                const cell = row.getCell(excelCellIndex++);
                const text = (htmlCell.textContent || '').trim();
                const innerHTML = htmlCell.innerHTML || '';
                const hyperlink = htmlCell.dataset.hyperlink;

                if (isHeader) {
                    // === 表頭：把 <br> 轉為 \n，並啟用換行 ===
                    // 1) 從 innerHTML 取用，處理 <br> 與 &nbsp;，再刪 HTML tag
                    let headerText = innerHTML
                    .replace(/<br\s*\/?>/gi, '\n')
                    .replace(/&nbsp;/gi, ' ')
                    .replace(/<\/?[^>]+>/g, '')
                    .trim();

                    // 若 innerHTML 沒有 <br>，headerText 會等於純文字；保險起見，也允許 text 當備援
                    if (!headerText) headerText = text;

                    cell.value = headerText || null;
                    cell.font = { bold: true };
                    cell.alignment = { wrapText: true, horizontal: 'center', vertical: 'middle' };
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEFEFEF' } };

                } else {
                    // === 內容列 ===
                    if (htmlCell.innerHTML.includes('&nbsp;') || text === "") {
                    cell.value = null;
                    } else if (hyperlink) {
                    cell.value = { text: text, hyperlink: hyperlink };
                    } else if (text.endsWith('%')) {
                    cell.value = parseFloat(text.replace('%','')) / 100;
                    cell.numFmt = '0.00%';
                    } else if (/^\d{4}\/\d{2}\/\d{2}$/.test(text)) {
                    const d = new Date(text);
                    cell.value = !isNaN(d.getTime()) ? d : text;
                    } else if (!isNaN(Number(text)) && text !== '') {
                    cell.value = Number(text);
                    } else {
                    cell.value = text;
                    }
                }

                // 邊框
                cell.border = {
                    top:   { style: 'thin' },
                    left:  { style: 'thin' },
                    bottom:{ style: 'thin' },
                    right: { style: 'thin' }
                };

                if (!isHeader) {
                    const role = htmlCell.dataset.gantt; // 'root' | 'child' | undefined
                    if (role === 'root') {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToARGB(THEME.gantt.rootHex) } };
                    } else if (role === 'child') {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToARGB(THEME.gantt.childHex) } };
                    } else if (htmlCell.classList.contains('bg-green-200')) {
                        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBBF7D0' } };
                    }
                }




                });

                // 把表頭列高度加高一點，讓換行後不擠在一起
                if (isHeader) {
                row.height = 32; // 可依需要調整 28~40
                }
            });

            // === 讓 Excel 顯示群組大綱（summary 在上方）===
            worksheet.properties.outlineProperties = {
            summaryBelow: false, // 父摘要列在上方；若你希望在下方可改 true
            summaryRight: true
            };

            // 小工具：檢查 DOM 列是否是某祖先的子孫
            function isDescendantOfDom(rowEl, ancestorId) {
            let cur = rowEl;
            while (cur && cur.dataset && cur.dataset.parentId) {
                if (cur.dataset.parentId === ancestorId) return true;
                cur = document.getElementById(cur.dataset.parentId);
            }
            return false;
            }

            // 小工具：這一列是否有任一祖先在 UI 當下是 collapsed
            function hasAnyCollapsedAncestor(rowEl) {
            let cur = rowEl;
            while (cur && cur.dataset && cur.dataset.parentId) {
                const p = document.getElementById(cur.dataset.parentId);
                if (p && p.dataset && p.dataset.collapsed === 'true') return true;
                cur = p;
            }
            return false;
            }

            // === 對每一個「可見且輸出」的 HTML 列，設定 Excel 的 outline level ===
            // 注意：rowsToExport 包含 thead，所以資料列要 +1；你的寫入邏輯是從第一列開始寫在 Excel 的 rowIndex 1
            const dataTable = document.getElementById('data-table');
            const htmlBodyRows = Array.from(dataTable.querySelectorAll('tbody tr')); // 列：全部參與大綱/隱藏判斷

            let excelRowCursor = 2; // Excel 第1列是表頭，資料從第2列開始
            htmlBodyRows.forEach((htmlRow) => {
            // 取這一列的原始 rowIndex 與階層等級
            const originalRowIndex = parseInt(htmlRow.dataset.originalIndex, 10);
            const level0 = (currentRowMeta[originalRowIndex + 1] || 0); // 你的 meta：0 表最上層
            const outlineLevel = Math.min(Math.max(level0, 0), 7); // Excel 允許 0~7

            const xRow = worksheet.getRow(excelRowCursor);
            if (outlineLevel > 0) {
                // Excel 的大綱層級是從 1 起算較直觀，0 代表無群組
                xRow.outlineLevel = outlineLevel; 
            }

            // 如果 UI 當下此列有任何祖先被收折，先把它設為 hidden，Excel 打開時就像是收起狀態
            if (hasAnyCollapsedAncestor(htmlRow)) {
                xRow.hidden = true;
            }

            excelRowCursor++;
            });


            const buffer = await workbook.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });

            if (window.showSaveFilePicker) {
                try {
                const handle = await window.showSaveFilePicker({
                    suggestedName: 'gantt_export.xlsx',
                    types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
                });
                const writable = await handle.createWritable();
                await writable.write(blob);
                await writable.close();
                } catch (err) {
                console.log('使用者取消存檔或發生錯誤。', err);
                }
            } else {
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url; a.download = 'gantt_export.xlsx';
                document.body.appendChild(a); a.click();
                window.URL.revokeObjectURL(url); document.body.removeChild(a);
            }
        });


        modalConfirmBtn.addEventListener('click', () => { errorModal.classList.add('hidden'); });

        highlightToggle.addEventListener('change', clearAllHighlights);
        freezeColsInput.addEventListener('input', applyStickyLayout);

        // =================================================================================
        // ============================ 樣式與排版輔助函式 ============================
        // =================================================================================
        function styleCell(cell, isHeader) {
            const baseClasses = ['px-4', 'py-3', 'whitespace-nowrap', 'border', 'border-black'];
            const headerClasses = ['font-semibold', 'text-gray-700', 'bg-green-100', 'relative', 'cursor-pointer'];
            const bodyClasses = ['text-gray-800', 'bg-white'];
            cell.classList.add(...baseClasses);
            if (isHeader) cell.classList.add(...headerClasses);
            else cell.classList.add(...bodyClasses);
        }

        function styleGeneratedTable() {
            const table = excelContent.querySelector('#data-table');
            if (!table) return;

            // 表格基礎樣式
            table.classList.add('min-w-full', 'text-base', 'border-collapse');

            // 先清 class，再套表頭/內文樣式
            table.querySelectorAll('th, td').forEach(cell => cell.className = '');
            table.querySelectorAll('thead th').forEach(th => styleCell(th, true));
            table.querySelectorAll('tbody td').forEach(td => styleCell(td, false));

            // 以前的「第一欄空白就上綠」規則請刪除，不再需要

            // 建立階層（插入 +/-、設定 parentId/level 等）
            buildRowHierarchy();

            // 只將根節點（level=0 且不是任何人的子節點）上色
            markRootRowColors();
        }


        function markRootRowColors() {
            const table = document.getElementById('data-table');
            if (!table) return;

            const headerRow = table.querySelector('thead tr');
            const bodyRows  = table.querySelectorAll('tbody tr');

            bodyRows.forEach(row => {
                // 先清標記
                row.dataset.tier = ''; // 'top' | 'child' | ''
                // 清除舊綠色，保留甘特與 sticky 在其他流程處理
                Array.from(row.children).forEach((cell, colIdx) => {
                cell.style.removeProperty('background-color');
                cell.classList.remove('bg-green-200', 'bg-green-100'); // 舊 class 也清掉
                });

                // 判斷階層
                const isParent = (row.dataset.isParent === 'true');
                const hasParent = !!row.dataset.parentId;
                const isTopLevel = isParent && !hasParent;   // 有 child 且沒有 parent
                const isChild    = hasParent;                // 任何有 parent 的列

                if (isTopLevel) {
                row.dataset.tier = 'top';
                } else if (isChild) {
                row.dataset.tier = 'child';
                } else {
                row.dataset.tier = ''; // 單純一般列（無 parent & 無 child）
                }

                // 逐格上綠（僅「一般欄位」，遇到「日期直行」或「空白表頭」就跳過）
                Array.from(row.children).forEach((cell, colIdx) => {
                    // 只對「原始欄」上綠；日期欄（動態產生的）直接跳過
                    if (isGeneratedDateCol(colIdx)) return;

                    // 若你仍想保護空白表頭，也可保留這兩行：
                    const th = headerRow?.children[colIdx];
                    const headerText = th ? (th.textContent || '').trim() : '';
                    if (headerText === '') return;

                    if (row.dataset.tier === 'top') {
                        cell.style.backgroundColor = THEME.row.topLevelHex; // 深綠
                    } else if (row.dataset.tier === 'child') {
                        cell.style.backgroundColor = THEME.row.childHex;    // 淺綠
                    }
                });

            });
        }






        function applyStickyLayout() {
            const table = document.getElementById('data-table');
            if (!table) return;

            // 先把所有 sticky 與接縫標記清掉
            table.querySelectorAll('th, td').forEach(cell => {
                cell.classList.remove('sticky', 'z-10', 'z-20', 'z-30', 'top-0', 'freeze-edge');
                cell.style.left = '';
            });

            const freezeCount = parseInt(freezeColsInput.value, 10);
            if (isNaN(freezeCount) || freezeCount < 0) return;

            const rows = Array.from(table.querySelectorAll('tr'));
            if (rows.length === 0 || !rows[0].children[0]) return;

            const originalColumnCount = (currentDataArray?.[0]?.length) || 0;
            const headerRow = table.querySelector('thead tr');

            requestAnimationFrame(() => {
                // 計算每個可見凍結欄的 left 偏移
                const headerVisibleCells = Array.from(rows[0].children).filter(c => c.style.display !== 'none');
                const offsets = [];
                let currentOffset = 0;

                headerVisibleCells.forEach((cell, i) => {
                offsets[cell.cellIndex] = currentOffset;
                if (i < freezeCount) currentOffset += cell.getBoundingClientRect().width;
                });

                // 找到「最後一個凍結且可見」欄位的 cellIndex
                let lastFrozenCellIndex = -1;
                if (freezeCount > 0 && headerVisibleCells[freezeCount - 1]) {
                lastFrozenCellIndex = headerVisibleCells[freezeCount - 1].cellIndex;
                }

                rows.forEach(row => {
                const isHeader = row.parentElement.tagName === 'THEAD';

                Array.from(row.children).forEach((cell) => {
                    if (cell.style.display === 'none') return;

                    // 只算可見欄位的索引來判斷是否在凍結範圍
                    const visibleIndex = Array
                    .from(cell.parentElement.children)
                    .filter(c => c.style.display !== 'none')
                    .indexOf(cell);

                    const isFrozenCol = visibleIndex > -1 && visibleIndex < freezeCount;

                    // 套 sticky 與定位
                    if (isHeader) cell.classList.add('sticky', 'top-0');
                    if (isFrozenCol) {
                    cell.classList.add('sticky');
                    cell.style.left = `${offsets[cell.cellIndex] || 0}px`;
                    }

                    // z-index 疊層
                    if (isHeader && isFrozenCol)      cell.classList.add('z-30');
                    else if (isFrozenCol)             cell.classList.add('z-20');
                    else if (isHeader)                cell.classList.add('z-10');

                    // sticky 背景（只處理 sticky 的格子，以免覆蓋甘特藍）
                    if (cell.classList.contains('sticky')) {
                    // 先移除我們會動到的底色類
                    cell.classList.remove('bg-green-200', 'bg-green-100', 'bg-green-50', 'bg-white');

                    if (isHeader) {
                        // 表頭底色
                        cell.classList.add('bg-green-100');
                    } else {
                        // 資料列 sticky 底色：日期區不吃綠；資料區按 isTop/isChild 顯示
                        const isDateCol = cell.cellIndex >= originalColumnCount;

                        if (isDateCol) {
                        cell.classList.add('bg-white'); // 日期區一律白
                        } else {
                        if (row.dataset.isTop === 'true') {
                            // 最上層 parent（且非別人的 child）
                            cell.classList.add('bg-green-200');
                        } else if (row.dataset.isChild === 'true') {
                            // 子節點
                            cell.classList.add('bg-green-50');
                        } else {
                            // 一般列
                            cell.classList.add('bg-white');
                        }
                        }
                    }
                    }
                });

                // 在每一列的「最後一個凍結欄」掛上接縫線 class
                if (lastFrozenCellIndex >= 0 && row.children[lastFrozenCellIndex]) {
                    row.children[lastFrozenCellIndex].classList.add('freeze-edge');
                }
                });
            });
        }




        // =================================================================================
        // ============================ 甘特圖與篩選相關功能 ============================
        // =================================================================================
        function getCellText(cellData) {
            if (cellData === null || cellData === undefined) return '';
            if (cellData.hyperlink) return String(cellData.text || '').trim();
            if (cellData instanceof Date) {
                const year = cellData.getFullYear();
                const month = String(cellData.getMonth() + 1).padStart(2, '0');
                const day = String(cellData.getDate()).padStart(2, '0');
                return `${year}/${month}/${day}`;
            }
            return String(cellData).trim();
        }

        function addFilterIconsToHeader() {
            const table = document.getElementById('data-table');
            if (!table || currentDataArray.length === 0) return;
            const headerCells = table.querySelectorAll('thead th');
            const dataColumnCount = currentDataArray[0].length;

            headerCells.forEach((th, colIndex) => {
                if (colIndex >= dataColumnCount) return; // Skip generated date columns
                if(th.querySelector('.filter-icon')) return;
                const filterIcon = document.createElement('span');
                filterIcon.textContent = '▼';
                filterIcon.classList.add('filter-icon', 'cursor-pointer', 'text-xs', 'ml-2', 'text-gray-400', 'hover:text-blue-600');
                filterIcon.dataset.colIndex = colIndex;
                th.appendChild(filterIcon);
            });
            table.querySelector('thead').addEventListener('click', handleFilterIconClick);
            updateFilterIcons();
        }

        function handleFilterIconClick(event) {
            const target = event.target;
            if (!target.classList.contains('filter-icon')) return;

            event.stopPropagation();

            const colIndex = parseInt(target.dataset.colIndex, 10);
            const filterPopup = document.getElementById('filter-popup');
            const filterOptions = document.getElementById('filter-options');
            const searchInput = document.getElementById('filter-search-input');

            // 取得本欄所有唯一值
            const uniqueValues = new Set();
            currentDataArray.slice(1).forEach(row => {
                const cellValue = row[colIndex];
                uniqueValues.add(getCellText(cellValue));
            });

            // 讀取當前過濾狀態
            const AF = activeFilters[colIndex];

            // === 建立清單（清單模式時反映勾選；文字模式時全勾） ===
            const currentValues = (AF && AF.mode === 'values') ? AF.values : new Set(uniqueValues);

            filterOptions.innerHTML =
                `<div class="flex items-center border-b pb-1 mb-1">
                    <input type="checkbox" id="select-all-filter" class="h-4 w-4" ${currentValues.size === uniqueValues.size ? 'checked' : ''}>
                    <label for="select-all-filter" class="ml-2 font-bold">(全選)</label>
                </div>`;
            Array.from(uniqueValues).sort().forEach(value => {
                const isChecked = currentValues.has(value);
                const displayValue = value || '(空白)';
                filterOptions.innerHTML +=
                `<div class="flex items-center filter-option-item">
                    <input type="checkbox" value="${escapeHTML(value)}" class="filter-option h-4 w-4" ${isChecked ? 'checked' : ''}>
                    <label class="ml-2">${escapeHTML(displayValue)}</label>
                </div>`;
            });

            // === 帶回「文字模式」的字串到輸入框 ===
            searchInput.value =
                (AF && AF.mode === 'text_any')
                    ? (AF.terms || []).join(', ')
                    : (AF && AF.mode === 'text')   // 舊資料相容
                    ? (AF.text || '')
                    : '';
                searchInput.placeholder = '輸入關鍵字，可用逗號分隔；支援 "片語"';


            // === 是否保留未命中父列的開關 ===
            let keepParentWrap = document.getElementById('keep-parents-wrap');
            if (!keepParentWrap) {
            keepParentWrap = document.createElement('div');
            keepParentWrap.id = 'keep-parents-wrap';
            keepParentWrap.className = 'flex items-center gap-2 mb-2';
            // 插到選項清單區塊上方（你也可以換位置）
            filterPopup.insertBefore(keepParentWrap, filterOptions);
            // <div class="text-xs text-gray-500 mt-1">可輸入多個關鍵字，逗號/空白/| 分隔；支援 "片語"</div>
            }
            keepParentWrap.innerHTML = `
            <input type="checkbox" id="keep-parents-checkbox" class="h-4 w-4" ${keepNonmatchingParents ? 'checked' : ''}>
            <label for="keep-parents-checkbox" class="text-sm select-none">保留未命中父列</label>
            `;
            const keepParentsCheckbox = document.getElementById('keep-parents-checkbox');
            keepParentsCheckbox.onchange = () => {
            keepNonmatchingParents = keepParentsCheckbox.checked;
            };


            // 位置與顯示
            const rect = target.getBoundingClientRect();
            filterPopup.style.left = `${rect.left}px`;
            filterPopup.style.top = `${rect.bottom + window.scrollY}px`;
            filterPopup.classList.remove('hidden');
            searchInput.focus();

            // 文字有輸入 → 停用清單勾選（但可視覺搜尋清單）
            const setOptionsDisabled = (disabled) => {
                filterOptions.style.opacity = disabled ? '0.5' : '';
                const allInputs = filterOptions.querySelectorAll('input[type="checkbox"]');
                allInputs.forEach(cb => cb.disabled = disabled);
            };
            setOptionsDisabled(searchInput.value.trim().length > 0);

            searchInput.oninput = () => {
                const hasText = searchInput.value.trim().length > 0;
                setOptionsDisabled(hasText);
                const term = searchInput.value.toLowerCase();
                const optionItems = filterOptions.querySelectorAll('.filter-option-item');
                optionItems.forEach(item => {
                    const label = item.querySelector('label').textContent.toLowerCase();
                    item.style.display = label.includes(term) ? 'flex' : 'none';
                });
            };

            document.getElementById('select-all-filter').onchange = (e) => {
                filterOptions.querySelectorAll('.filter-option').forEach(checkbox => checkbox.checked = e.target.checked);
            };

            // 套用：有字串 → 文字模式；沒字串 → 清單模式
            document.getElementById('filter-apply-btn').onclick = () => {
                const searchTerm = (searchInput.value || '').trim();
                const raw = (searchInput.value || '').trim();
                const terms = parseSearchTerms(raw);
                if (terms.length > 0) {
                    // 新的多關鍵字（OR）模式
                    activeFilters[colIndex] = { mode: 'text_any', terms };
                } 
                else {
                    // 沒輸入 → 回到清單模式
                    const selectedValues = new Set();
                    filterOptions.querySelectorAll('.filter-option:checked')
                        .forEach(input => selectedValues.add(input.value));
                    if (selectedValues.size === uniqueValues.size) {
                        delete activeFilters[colIndex];
                    } 
                    else {
                        activeFilters[colIndex] = { mode: 'values', values: selectedValues };
                    }
                }

                applyFiltersToTable();
                filterPopup.classList.add('hidden');
                updateFilterIcons();
            };

            document.getElementById('filter-cancel-btn').onclick = () => filterPopup.classList.add('hidden');

            // 清除：如果沒有就動態建立
            let clearBtn = document.getElementById('filter-clear-btn');
            if (!clearBtn) {
                const footer = filterPopup.querySelector('div.flex.justify-between');
                clearBtn = document.createElement('button');
                clearBtn.id = 'filter-clear-btn';
                clearBtn.className = 'text-sm bg-gray-100 px-3 py-1 rounded-md hover:bg-gray-200';
                clearBtn.textContent = '清除';
                footer.insertBefore(clearBtn, document.getElementById('filter-cancel-btn'));
            }
            clearBtn.onclick = () => {
                delete activeFilters[colIndex];
                searchInput.value = '';
                applyFiltersToTable();
                filterPopup.classList.add('hidden');
                updateFilterIcons();
            };
        }

        function parseSearchTerms(input) {
            // 支援引號包住的片語，和多種分隔符：逗號, 分號; 頓號、 直線| 空白/換行
            // 例:  Car, Dog | "New York" ；會得到 ["car","dog","new york"]
            if (!input) return [];
            const s = String(input).trim();
            if (!s) return [];
            const terms = [];
            const re = /"([^"]+)"|([^,;、|\s]+)/g; // 先抓 "片語"，再抓分隔的 token
            let m;
            while ((m = re.exec(s)) !== null) {
                const token = (m[1] ?? m[2] ?? '').trim().toLowerCase();
                if (token) terms.push(token);
            }
            return Array.from(new Set(terms)); // 去重
        }


        function applyFiltersToTable() {
            const table = document.getElementById('data-table');
            if (!table) return;

            const bodyRows = Array.from(table.querySelectorAll('tbody tr'));
            const rowVisibility = new Map();

            // 依每列套用所有欄位的過濾條件
            bodyRows.forEach(row => {
                let matchesFilter = true;
                const originalRowIndex = parseInt(row.dataset.originalIndex, 10);
                const rowData = currentDataArray[originalRowIndex + 1];

                if (rowData) {
                    for (const colIndex in activeFilters) {
                        const filter = activeFilters[colIndex];
                        const rawValue = rowData[colIndex];
                        const cellText = (getCellText(rawValue) || '').toLowerCase();

                        if (!filter) continue;
                        if (filter.mode === 'text_any') {
                            // 多關鍵字 OR：任一關鍵字命中即可
                            const terms = Array.isArray(filter.terms) ? filter.terms : [];
                            const hit = terms.some(t => t && cellText.includes(t));
                            if (!hit) { matchesFilter = false; break; }
                        } 
                        else if (filter.mode === 'text') {
                            // 相容舊的單一字串
                            const term = (filter.text || '').toLowerCase();
                            if (term && !cellText.includes(term)) { matchesFilter = false; break; }
                        } 
                        else if (filter.mode === 'values') {
                            const v = getCellText(rawValue);
                            if (!filter.values.has(v)) { matchesFilter = false; break; }
                        }

                    }
                }
                rowVisibility.set(row, matchesFilter);
            });

            // 子列可見 → 父列也可見（可選）
            if (keepNonmatchingParents) {
                for (let i = bodyRows.length - 1; i >= 0; i--) {
                    const row = bodyRows[i];
                    if (row.dataset.isParent === 'true') {
                        const parentId = row.id;
                        const children = bodyRows.filter(child => child.dataset.parentId === parentId);
                        const anyChildVisible = children.some(child => rowVisibility.get(child) === true);
                    if (anyChildVisible) 
                        rowVisibility.set(row, true);
                    }
                }
            }


            // 套用可見性（含收合狀態）
            bodyRows.forEach(row => {
                let isVisible = rowVisibility.get(row);
                const parentId = row.dataset.parentId;
                if (parentId) {
                    const parentRow = document.getElementById(parentId);
                    const parentCollapsed = parentRow && parentRow.dataset.collapsed === 'true';

                    // 1) 永遠尊重收合狀態：若任何可見祖先收合就隱藏
                    if (parentCollapsed) {
                        isVisible = false;
                    } 
                    else if (keepNonmatchingParents) {
                        // 2) 只有在「保留未命中父列」為 true 時，才讓「父列不顯示」影響到子列
                        if (parentRow && rowVisibility.get(parentRow) === false) {
                            isVisible = false;
                        }
                    }
            }
            row.style.display = isVisible ? '' : 'none';
            });


            updateVisibleRowCount();
        }

        function updateVisibleRowCount() {
            const table = document.getElementById('data-table');
            if (!table) return;
            const visibleRows = Array.from(table.querySelectorAll('tbody tr')).filter(row => row.style.display !== 'none');
            dataSummary.textContent = `總共 ${visibleRows.length} 項需求`;
        }

        function fmtISO(d){ // YYYY-MM-DD
            const y = d.getFullYear();
            const m = String(d.getMonth()+1).padStart(2,'0');
            const da = String(d.getDate()).padStart(2,'0');
            return `${y}-${m}-${da}`;
        }
        function fmtSlash(d){ // YYYY/MM/DD（用來顯示）
            const y = d.getFullYear();
            const m = String(d.getMonth()+1).padStart(2,'0');
            const da = String(d.getDate()).padStart(2,'0');
            return `${y}/${m}/${da}`;
        }
            // 週的穩定 key（用來保存/讀取資料，不受表頭格式影響）
        function weekKey(week){ return `${fmtISO(week.start)}~${fmtISO(week.end)}`; }
            // 表頭顯示用（含換行）
        function weekLabelHTML(week){ return `${fmtSlash(week.start)}<br>${fmtSlash(week.end)}`; }
        function isDateHeader(text) {
            const t = (text || '').trim();
            if (!t) return false;
            // 你表頭用 2025/01/06 這類：YYYY/MM/DD
            return /^\d{4}\/\d{2}\/\d{2}$/.test(t);
        }
        function isGeneratedDateCol(colIdx){
            const originalColumnCount = (currentDataArray?.[0]?.length) || 0;
            return colIdx >= originalColumnCount;
        }




        function updateFilterIcons() {
            const table = document.getElementById('data-table');
            if (!table) return;
            table.querySelectorAll('thead th').forEach((th) => {
                const icon = th.querySelector('.filter-icon');
                if (icon) {
                    const dataColIndex = parseInt(icon.dataset.colIndex, 10);
                    if (activeFilters[dataColIndex]) {
                        icon.style.color = 'rgb(37 99 235)'; // 藍色
                        th.classList.add('filter-active');
                    } else {
                        icon.style.color = '';
                        th.classList.remove('filter-active');
                    }
                }
            });
        }

        function colorGanttCells() {
            const table = document.getElementById('data-table');
            if (!table || devStartDateColIndex === -1 || dueDateColIndex === -1 || currentDataArray.length <= 1) return;
            if (!generatedDateColumns || generatedDateColumns.length === 0) return;

            const headerRow = table.querySelector('thead tr');
            const bodyRows  = Array.from(table.querySelectorAll('tbody tr'));
            const originalColumnCount = currentDataArray[0].length;

            bodyRows.forEach((row) => {
                const originalRowIndex = parseInt(row.dataset.originalIndex, 10);
                if (isNaN(originalRowIndex)) return;

                const rowData = currentDataArray[originalRowIndex + 1];
                if (!rowData) return;

                // 讀取區間
                const rawDevStartDate = rowData[devStartDateColIndex];
                const rawDueDate      = rowData[dueDateColIndex];

                const devStartDate = rawDevStartDate instanceof Date ? rawDevStartDate : new Date(getCellText(rawDevStartDate));
                const dueDate      = rawDueDate      instanceof Date ? rawDueDate      : new Date(getCellText(rawDueDate));

                if (isNaN(devStartDate.getTime()) || isNaN(dueDate.getTime())) return;

                // 是否為「根節點」：有 child（isParent=true）且沒有 parentId
                const isTopLevel = (!row.dataset.parentId || row.dataset.parentId === '');

                // 逐一處理每一個甘特週欄位
                for (let j = originalColumnCount; j < headerRow.children.length; j++) {
                const weekIdx = j - originalColumnCount;
                const week    = generatedDateColumns[weekIdx];
                if (!week) continue;

                const cell = row.children[j];
                if (!cell) continue;

                // 先清掉舊的樣式/標記（避免殘留）
                cell.classList.remove('bg-blue-300', 'bg-blue-600'); // 舊 Tailwind 類別保險移除
                cell.style.backgroundColor = '';
                cell.dataset.gantt = '';

                // 與任一週有重疊就上色
                const overlap = (week.start <= dueDate) && (week.end >= devStartDate);
                if (overlap) {
                    cell.style.backgroundColor = isTopLevel ? THEME.gantt.rootHex : THEME.gantt.childHex;
                    cell.dataset.gantt = isTopLevel ? 'root' : 'child';

                }
                }
            });
        }





        // =================================================================================
        // ============================ 階層、高亮、互動功能 ============================
        // =================================================================================
        function buildRowHierarchy() {
            const table = document.getElementById('data-table');
            if (!table) return;

            const bodyRows = table.querySelectorAll('tbody tr');
            if (!bodyRows.length) return;

            const hasAnyLevel = currentRowMeta.some(l => l && l > 0);
            if (!hasAnyLevel) return buildRowHierarchy_heuristic();

            let stack = []; // stack[level] = latest parent row id on this level

            bodyRows.forEach((row) => {
                const idx = getOriginalRowIndexSafe(row);
                if (idx < 0) return;

                const level = (currentRowMeta[idx + 1] || 0);

                // keep stack to current level
                stack = stack.slice(0, level);

                // link parent
                if (level > 0 && stack[level - 1]) {
                row.dataset.parentId = stack[level - 1];
                }

                const nextLevel = (currentRowMeta[idx + 2] || 0);
                const isParent = nextLevel > level;

                const rowId = `row-${idx}`;
                row.id = rowId;
                row.dataset.isParent = String(isParent);

                const savedCollapsed = collapsedState.get(getStableRowKeyByIndex(idx));
                row.dataset.collapsed = typeof savedCollapsed === 'boolean'
                ? String(savedCollapsed)
                : (row.dataset.collapsed || 'false');

                // === 原本：先清舊元素 ===
                const firstCell = row.children[0];
                if (firstCell) {
                    const oldIndent = firstCell.querySelector('.tree-indent');
                    if (oldIndent) oldIndent.remove();
                    const oldPlaceholder = firstCell.querySelector('.toggle-placeholder');
                    if (oldPlaceholder) oldPlaceholder.remove();
                    const oldToggle = firstCell.querySelector('.toggle-collapse');
                    if (oldToggle) oldToggle.remove();

                    // ★★★ 這裡改成用 padding-left 做整體縮排（含 +/- 與文字） ★★★
                    // 每層 16px，可自行調整（例如 20px）
                    const BASE_PAD = 4;              // 若想多一點基準內縮就調大，例如 4 或 8
                    const INDENT_PER_LEVEL = 16;     // 每層的縮排距離
                    firstCell.style.paddingLeft = `${BASE_PAD + level * INDENT_PER_LEVEL}px`;

                    // 再放 +/- 或等寬占位（確保文字起點一致）
                    if (isParent) {
                        const toggle = document.createElement('span');
                        toggle.textContent = row.dataset.collapsed === 'true' ? '+' : '−';
                        toggle.classList.add('toggle-collapse', 'cursor-pointer');
                        firstCell.prepend(toggle);
                    } 
                    else {
                        const placeholder = document.createElement('span');
                        placeholder.className = 'toggle-placeholder';
                        firstCell.prepend(placeholder);
                    }
                }


                // stack 紀錄最新父節點
                if (isParent) stack[level] = rowId;

                // 標記層級（若之後想用 CSS :has 或 data 屬性）
                row.dataset.level = String(level);
            });
        }


        function isDescendantOf(row, ancestorId) {
            // 走訪祖先鏈，若遇到 ancestorId 就回 true
            let cur = row;
            while (cur && cur.dataset && cur.dataset.parentId) {
                if (cur.dataset.parentId === ancestorId) return true;
                cur = document.getElementById(cur.dataset.parentId);
            }
            return false;
        }

        function cascadeCollapseDescendants(rootRow) {
            const table = document.getElementById('data-table');
            if (!table || !rootRow || !rootRow.id) return;

            const allRows = Array.from(table.querySelectorAll('tbody tr'));
            allRows.forEach(r => {
                if (isDescendantOf(r, rootRow.id)) {
                // 對所有子孫標記為 collapsed=true
                r.dataset.collapsed = 'true';
                const t = r.querySelector('.toggle-collapse');
                if (t) t.textContent = '+';

                // 把狀態寫回你的穩定 key（避免重建時遺失）
                const rIdx = getOriginalRowIndexSafe(r);
                if (rIdx >= 0) {
                    collapsedState.set(getStableRowKeyByIndex(rIdx), true);
                }
                }
            });
        }


        function getOriginalRowIndexSafe(row) {
            if (!row) return -1;
            // 先從 dataset 讀
            const v = row.dataset ? row.dataset.originalIndex : undefined;
            if (v !== undefined && v !== '') {
                const n = Number(v);
                if (!Number.isNaN(n)) return n;
            }
            // 退路：用 DOM 的 rowIndex（第一列是 header，所以要 -1）
            const fallback = Math.max(0, (row.rowIndex || 1) - 1);
            // 補回 dataset，後續就穩了
            if (row.dataset) row.dataset.originalIndex = String(fallback);
            return fallback;
        }



        function buildRowHierarchy_heuristic() {
            const table = document.getElementById('data-table');
            if (!table) return;
            let currentParentId = null;
            let parentCounter = 0;
            const bodyRows = table.querySelectorAll('tbody tr');

            bodyRows.forEach((row) => {
                const firstCell = row.children[0];
                if (!firstCell) return;
                const originalIndex = parseInt(row.dataset.originalIndex, 10);
                const originalRowData = currentDataArray[originalIndex + 1] || [];
                const isFirstCellBlank = isBlankText(firstCell);
                const isDataRow = originalRowData.some(c => c !== null && c !== undefined && c !== '');

                if (isFirstCellBlank && isDataRow) {
                parentCounter++;
                currentParentId = `parent-row-${parentCounter}`;
                row.id = currentParentId;
                row.dataset.isParent = 'true';
                row.dataset.collapsed = row.dataset.collapsed || 'false';
                const toggle = document.createElement('span');
                toggle.textContent = row.dataset.collapsed === 'true' ? '+' : '−';
                toggle.classList.add('toggle-collapse','cursor-pointer');
                firstCell.prepend(toggle);
                } else if (currentParentId && !isFirstCellBlank) {
                row.dataset.parentId = currentParentId;
                } else {
                currentParentId = null;
                }
            });
            }

            // 更穩定的空白判斷（NBSP、零寬空白都剔除）
            function isBlankText(node){
            const t = (node.textContent || '')
                .replace(/\u00A0/g, ' ')
                .replace(/[\u200B-\u200D\uFEFF]/g,'')
                .trim();
            return t === '';
            }


        function clearAllHighlights() {
            document.querySelectorAll('.highlight').forEach(cell => cell.classList.remove('highlight'));
        }
        function highlightHeaders(minCol, maxCol) {
            const table = document.getElementById('data-table');
            if (!table) return;
            const headerRow = table.querySelector('thead tr');
            if (headerRow) {
                for (let i = minCol; i <= maxCol; i++) {
                const th = headerRow.children[i];
                // 表頭通常是灰底；若你要完全不動表頭，可直接 continue
                if (th && th.classList.contains('bg-white')) th.classList.add('highlight');
                }
            }
            }

        function highlightRowAndColumn(cell) {
            const table = cell.closest('table');
            const colIndex = cell.cellIndex;
            const row = cell.parentElement;

            Array.from(row.children).forEach(c => {
                if (canHighlight(c)) c.classList.add('highlight');
            });

            Array.from(table.querySelectorAll('tr')).forEach(r => {
                if (r.children[colIndex] && canHighlight(r.children[colIndex])) {
                r.children[colIndex].classList.add('highlight');
                }
            });
        }


        function highlightSelectionRectangle(start, end) {
            clearAllHighlights();
            const table = start.closest('table');
            if (!table) return;
            const startCoords = { row: start.parentElement.rowIndex, col: start.cellIndex };
            const endCoords   = { row: end.parentElement.rowIndex,   col: end.cellIndex };
            const minRow = Math.min(startCoords.row, endCoords.row);
            const maxRow = Math.max(startCoords.row, endCoords.row);
            const minCol = Math.min(startCoords.col, endCoords.col);
            const maxCol = Math.max(startCoords.col, endCoords.col);

            // 表頭仍可高亮顯示（保留原行為）
            highlightHeaders(minCol, maxCol);

            const rows = table.querySelectorAll('tr');
            const freezeCount = parseInt(freezeColsInput.value, 10);

            for (let i = minRow; i <= maxRow; i++) {
                if (!rows[i]) continue;
                const cells = rows[i].children;

                // 凍結欄（只白底且非甘特）
                if (i > 0 && !isNaN(freezeCount) && freezeCount > 0) {
                for (let j = 0; j < freezeCount; j++) {
                    if (cells[j] && canHighlight(cells[j])) {
                    cells[j].classList.add('highlight');
                    }
                }
                }
                // 實際選取區（只白底且非甘特）
                for (let j = minCol; j <= maxCol; j++) {
                if (cells[j] && canHighlight(cells[j])) {
                    cells[j].classList.add('highlight');
                }
                }
            }
        }


        /**
         * 一次展開或收折所有父列
         * @param {boolean} collapse - true=全部收折（顯示 +），false=全部展開（顯示 −）
         */
        function toggleAllRows(collapse) {
            const table = document.getElementById('data-table');
            if (!table) return;

            const bodyParents = table.querySelectorAll('tbody tr[data-is-parent="true"]');

            bodyParents.forEach(row => {
                const idx = getOriginalRowIndexSafe(row);
                if (idx >= 0) collapsedState.set(getStableRowKeyByIndex(idx), collapse);

                row.dataset.collapsed = String(collapse);
                const t = row.querySelector('.toggle-collapse');
                if (t) t.textContent = collapse ? '+' : '−';

                // ★ 若是要全部收折，也把各父節點的子孫一併級聯收折
                if (collapse) cascadeCollapseDescendants(row);
            });

            applyFiltersToTable();
        }



        // =================================================================================
        // ============================ v1.5.1 新增/修改功能 =======================
        // =================================================================================

        // 右鍵選單
        excelContent.addEventListener('contextmenu', e => {
            const targetTh = e.target.closest('thead th');
            if (targetTh) {
                e.preventDefault();
                contextMenuTargetCol = targetTh;

                const colIndex = contextMenuTargetCol.cellIndex;
                const totalCols = contextMenuTargetCol.parentElement.children.length;

                const moveLeftItem = contextMenu.querySelector('#context-move-left');
                const moveRightItem = contextMenu.querySelector('#context-move-right');

                moveLeftItem.classList.toggle('disabled', colIndex === 0);
                moveRightItem.classList.toggle('disabled', colIndex === totalCols - 1);

                contextMenu.style.top = `${e.pageY}px`;
                contextMenu.style.left = `${e.pageX}px`;
                contextMenu.style.display = 'block';
            } else {
                contextMenu.style.display = 'none';
            }
        });

        document.addEventListener('click', (event) => {
            if (!event.target.closest('#excel-content-container') && event.target.id !== 'highlight-toggle' && !event.target.closest('#filter-popup') && !event.target.closest('#context-menu')) {
                clearAllHighlights();
            }
            if (contextMenu.style.display === 'block' && !event.target.closest('#context-menu')) {
                contextMenu.style.display = 'none';
            }
        });

        // 隱藏直行（寫入 colState）
        document.getElementById('context-hide-col').addEventListener('click', () => {
            contextMenu.style.display = 'none';
            if (contextMenuTargetCol) {
                const table = document.getElementById('data-table');
                const colIndex = contextMenuTargetCol.cellIndex;
                if (colIndex > -1) {
                    table.querySelectorAll('tr').forEach(row => {
                        if (row.children[colIndex]) row.children[colIndex].style.display = 'none';
                    });
                    // 寫入狀態
                    if (isDataCol(colIndex)) colState.data.hidden.add(colIndex);
                    else {
                        const key = getHeaderTextByIndex(table, colIndex);
                        if (key) colState.date.hidden.add(key);
                    }
                    applyStickyLayout();
                }
            }
        });

        // 取消隱藏全部（清空 colState.hidden）
        document.getElementById('context-unhide-all-cols').addEventListener('click', () => {
            contextMenu.style.display = 'none';
            const table = document.getElementById('data-table');
            if (table) {
                table.querySelectorAll('th, td').forEach(cell => { cell.style.display = ''; });
                colState.data.hidden.clear(); colState.date.hidden.clear();
                applyStickyLayout();
            }
        });

        function moveColumn(fromIndex, toIndex) {
            const totalOriginalCols = currentDataArray.length > 0 ? currentDataArray[0].length : 0;
            const totalGeneratedCols = generatedDateColumns.length;
            const totalCols = totalOriginalCols + totalGeneratedCols;

            if(fromIndex < 0 || toIndex < 0 || fromIndex >= totalCols || toIndex >= totalCols) return;

            const fromIsData = fromIndex < totalOriginalCols;
            const toIsData = toIndex < totalOriginalCols;

            if (fromIsData !== toIsData) {
                console.warn("不支援在資料欄與日期欄之間移動。");
                return;
            }

            if (fromIsData) { // 資料欄互換：資料、filter、欄寬/隱藏都要重映射
                currentDataArray.forEach(row => {
                    const temp = row[fromIndex];
                    row[fromIndex] = row[toIndex];
                    row[toIndex] = temp;
                });

                // filters 對調
                const tempFilter = activeFilters[fromIndex];
                activeFilters[fromIndex] = activeFilters[toIndex];
                activeFilters[toIndex] = tempFilter;
                if (activeFilters[fromIndex] === undefined) delete activeFilters[fromIndex];
                if (activeFilters[toIndex] === undefined) delete activeFilters[toIndex];

                // 欄寬狀態對調
                const w1 = colState.data.widths[fromIndex];
                const w2 = colState.data.widths[toIndex];
                if (w1 === undefined) delete colState.data.widths[toIndex]; else colState.data.widths[toIndex] = w1;
                if (w2 === undefined) delete colState.data.widths[fromIndex]; else colState.data.widths[fromIndex] = w2;

                // 隱藏索引重映射
                const newHidden = new Set();
                for (const idx of colState.data.hidden) {
                    if (idx === fromIndex) newHidden.add(toIndex);
                    else if (idx === toIndex) newHidden.add(fromIndex);
                    else newHidden.add(idx);
                }
                colState.data.hidden = newHidden;

                // 更新 dev/due index
                const header = currentDataArray[0];
                devStartDateColIndex = header.indexOf('Dev Start Date');
                dueDateColIndex = header.indexOf('Due Date');

            } else { // 日期欄互換：純換順序（key 用日期字串，不用重映射）
                const fromDateIndex = fromIndex - totalOriginalCols;
                const toDateIndex = toIndex - totalOriginalCols;
                const tempDate = generatedDateColumns[fromDateIndex];
                generatedDateColumns[fromDateIndex] = generatedDateColumns[toDateIndex];
                generatedDateColumns[toDateIndex] = tempDate;
            }

            renderTable();
            applyFiltersToTable();
            colorGanttCells();
        }

        document.getElementById('context-move-left').addEventListener('click', (e) => {
            contextMenu.style.display = 'none';
            if (e.target.classList.contains('disabled') || !contextMenuTargetCol) return;
            const colIndex = contextMenuTargetCol.cellIndex;
            moveColumn(colIndex, colIndex - 1);
        });

        document.getElementById('context-move-right').addEventListener('click', (e) => {
            contextMenu.style.display = 'none';
            if (e.target.classList.contains('disabled') || !contextMenuTargetCol) return;
            const colIndex = contextMenuTargetCol.cellIndex;
            moveColumn(colIndex, colIndex + 1);
        });

        // ===== 欄寬量測工具 & 計算預設寬度 =====
        function getTextMeasurer() {
            let m = document.getElementById('__col_width_measurer__');
            if (m) return m;
            m = document.createElement('div');
            m.id = '__col_width_measurer__';
            m.style.position = 'absolute';
            m.style.left = '-99999px';
            m.style.top = '0';
            m.style.visibility = 'hidden';
            // 對齊 styleCell：text-sm + px-4（左右 1rem）
            m.style.fontFamily = 'Inter, sans-serif';
            m.style.fontSize = '0.875rem';
            m.style.whiteSpace = 'nowrap';
            m.style.boxSizing = 'border-box';
            m.style.paddingLeft = '1rem';
            m.style.paddingRight = '1rem';
            document.body.appendChild(m);
            return m;
        }
        function measureTextWidthPx(text) {
            const m = getTextMeasurer();
            m.textContent = text || '';
            const w = m.getBoundingClientRect().width;
            m.textContent = '';
            return Math.ceil(w);
        }

        /**
         * 依原始列索引回傳一個「穩定的 row key」：
         * 盡量用第一欄（通常是 Issue Key / Summary）作為 key；
         * 取不到時退回 "#<index>"，避免排序/插入列讓狀態漂移。
         */
        function getStableRowKeyByIndex(originalRowIndex) {
            const row = currentDataArray[originalRowIndex + 1] || [];
            const raw = row[0]; // 取第一欄
            let text = '';

            if (raw && typeof raw === 'object' && 'text' in raw) {
                // 例如含超連結時你的資料結構是 { text, hyperlink }
                text = String(raw.text || '');
            } else {
                text = String(raw ?? '');
            }

            text = text.trim();
            return text ? text : `#${originalRowIndex + 1}`;
        }


        function computeDefaultColWidthPx(table, colIndex) {
            const rows = Array.from(table.querySelectorAll('tr'))
                .filter(r => r.style.display !== 'none');

            let maxW = 0;
            rows.forEach(row => {
                const cell = row.children[colIndex];
                if (!cell || cell.style.display === 'none') return;

                let text = '';
                const a = cell.querySelector('a');
                if (a) text = a.textContent.trim();
                else text = (cell.textContent || '').replace(/\u00a0/g, ' ').trim();

                if (text === '') {
                    const headerCell = table.querySelector('thead tr')?.children[colIndex];
                    if (headerCell) text = (headerCell.textContent || '').trim();
                }
                const w = measureTextWidthPx(text);
                if (w > maxW) maxW = w;
            });

            const PADDING = 8; // ← 修正掉原本的多餘 'a'
            const MIN_W = 24;
            const MAX_W = 1200;
            return Math.min(Math.max(maxW + PADDING, MIN_W), MAX_W);
        }
        function applyColumnWidthPx(table, colIndex, widthPx) {
            const rows = table.querySelectorAll('tr');
            rows.forEach(row => {
                const cell = row.children[colIndex];
                if (!cell) return;
                cell.style.width = `${widthPx}px`;
                cell.style.minWidth = `${widthPx}px`;
                cell.style.maxWidth = `${widthPx}px`;
            });
        }

        // 加入調整欄寬 + 雙擊自動寬（並寫入 colState）
        function addColumnResizeHandles() {
            const table = document.getElementById('data-table');
            if (!table) return;

            table.querySelectorAll('thead th').forEach((th, thIndex) => {
                // 刷新把手
                const old = th.querySelector('.resize-handle');
                if (old) old.remove();

                // 雙擊欄頭：恢復預設寬度並寫入狀態
                th.addEventListener('dblclick', (e) => {
                    if (e.target.classList.contains('filter-icon') || e.target.classList.contains('resize-handle')) return;
                    const width = computeDefaultColWidthPx(table, thIndex);
                    applyColumnWidthPx(table, thIndex, width);

                    if (isDataCol(thIndex)) {
                        colState.data.widths[thIndex] = width;
                    } else {
                        const key = getHeaderTextByIndex(table, thIndex);
                        if (key) colState.date.widths[key] = width;
                    }
                    applyStickyLayout(); // 重新計算凍結偏移
                });

                // 拖曳把手
                const handle = document.createElement('div');
                handle.classList.add('resize-handle');
                th.appendChild(handle);

                handle.addEventListener('mousedown', (e) => {
                    e.preventDefault();
                    e.stopPropagation();

                    const startX = e.pageX;
                    const startWidth = th.getBoundingClientRect().width;

                    const rows = table.querySelectorAll('tr');
                    const colCells = Array.from(rows).map(row => row.children[thIndex]).filter(Boolean);

                    const onMove = (moveEvent) => {
                        const delta = moveEvent.pageX - startX;
                        let newWidth = startWidth + delta;
                        const MIN_W = 24;
                        if (newWidth < MIN_W) newWidth = MIN_W;

                        colCells.forEach(cell => {
                            cell.style.width = `${newWidth}px`;
                            cell.style.minWidth = `${newWidth}px`;
                            cell.style.maxWidth = `${newWidth}px`;
                        });
                    };

                    const onUp = () => {
                        document.removeEventListener('mousemove', onMove);
                        document.removeEventListener('mouseup', onUp);

                        // 最終寬度存起來
                        const finalWidth = Math.round(th.getBoundingClientRect().width);
                        if (isDataCol(thIndex)) {
                            colState.data.widths[thIndex] = finalWidth;
                        } else {
                            const key = getHeaderTextByIndex(table, thIndex);
                            if (key) colState.date.widths[key] = finalWidth;
                        }
                        applyStickyLayout();
                    };

                    document.addEventListener('mousemove', onMove);
                    document.addEventListener('mouseup', onUp);
                });
            });
        }

        // 主要互動事件 (儲存格高亮)
        excelContent.addEventListener('mousedown', (event) => {
            // 正在編輯（點在 input/textarea）時，不啟動任何選取/高亮流程
            if (event.target.matches('input, textarea') || event.target.closest('input, textarea')) return;
            if (event.button === 2) return; // 右鍵略過

            const targetCell = event.target.closest('td, th');
            if (
                event.target.closest('.toggle-collapse') ||
                event.target.classList.contains('filter-icon') ||
                event.target.tagName === 'A' ||
                event.target.closest('#filter-popup') ||
                event.target.closest('.resize-handle')
            ) {
                return;
            }
            if (!targetCell) return;

            // ========== A) Shift 連選（重點新增） ==========
            if (event.shiftKey) {
                // 沒有錨點 → 設定錨點並高亮當前 cell
                if (!selectionAnchorCell || selectionAnchorCell.closest('table') !== targetCell.closest('table')) {
                selectionAnchorCell = targetCell;
                clearAllHighlights();
                // 只高亮自己
                highlightSelectionRectangle(targetCell, targetCell);
                return;
                }

                // 有錨點 → 用錨點到目前 cell 的矩形/直線範圍高亮
                clearAllHighlights();
                highlightSelectionRectangle(selectionAnchorCell, targetCell);
                // 不改變錨點，方便你再 Shift 點第三個 cell 重新劃範圍
                return;
            }

            // ========== B) 非 Shift：更新錨點，照你原本邏輯 ==========
            // 每次正常點擊都把「錨點」移到當前 cell，方便下一次 Shift 連選
            selectionAnchorCell = targetCell;

            const isCtrlPressed = event.ctrlKey || event.metaKey;

            if (highlightToggle.checked) {
                clearAllHighlights();
                highlightRowAndColumn(targetCell);
            } else {
                if (!isCtrlPressed) clearAllHighlights();

                if (isCtrlPressed) {
                    if (canHighlight(targetCell)) {
                        targetCell.classList.toggle('highlight');
                    }
                } else {

                isSelecting = true;
                selectionStartCell = targetCell;
                highlightSelectionRectangle(selectionStartCell, targetCell);
                }

                const handleMouseMove = (e) => {
                if (!isSelecting) return;
                const currentCell = e.target.closest('td, th');
                if (!currentCell || !selectionStartCell) return;
                highlightSelectionRectangle(selectionStartCell, currentCell);
                };

                const handleMouseUp = () => {
                isSelecting = false;
                document.removeEventListener('mousemove', handleMouseMove);
                document.removeEventListener('mouseup', handleMouseUp);
                };

                if (isSelecting) {
                document.addEventListener('mousemove', handleMouseMove);
                document.addEventListener('mouseup', handleMouseUp);
                }
            }
        });




        // =================================================================================
        // ============================ 日期欄雙擊編輯（保留） ============================
        // =================================================================================
        function addDateEditing() {
            const table = document.getElementById('data-table');
            if (!table || currentDataArray.length === 0) return;

            table.querySelectorAll('tbody tr').forEach(row => {
                const originalRowIndex = parseInt(row.dataset.originalIndex, 10);
                if (isNaN(originalRowIndex)) return;

                // 供變更紀錄用的 key（第一欄文字，或預設 Row N）
                const keyCellText = getCellText(
                    currentDataArray[originalRowIndex + 1][0] || `Row ${originalRowIndex + 1}`
                );

                // ===== 僅綁定：Dev Start Date =====
                if (devStartDateColIndex > -1 && row.children[devStartDateColIndex]) {
                    const cell = row.children[devStartDateColIndex];
                    cell.title = '雙擊可編輯 Dev Start Date';
                    cell.addEventListener(
                        'dblclick',
                        (e) => {
                            e.stopPropagation(); // 避免外層選取機制干擾
                            handleDateEdit(
                                cell,
                                originalRowIndex,
                                devStartDateColIndex,
                                'Dev Start Date',
                                keyCellText
                            );
                        },
                        { once: false }
                    );
                }

                // ===== 僅綁定：Due Date =====
                if (dueDateColIndex > -1 && row.children[dueDateColIndex]) {
                    const cell = row.children[dueDateColIndex];
                    cell.title = '雙擊可編輯 Due Date';
                    cell.addEventListener(
                        'dblclick',
                        (e) => {
                            e.stopPropagation(); // 避免外層選取機制干擾
                            handleDateEdit(
                                cell,
                                originalRowIndex,
                                dueDateColIndex,
                                'Due Date',
                                keyCellText
                            );
                        },
                        { once: false }
                    );
                }

                // ⚠️ 不在這裡處理「產生的日期欄」文字編輯！
                // 產生的日期欄文字編輯請交給 addGeneratedDateTextEditing()。
            });
        }



        function handleDateEdit(cell, rowIndex, colIndex, colName, key) {
            const originalValue = currentDataArray[rowIndex + 1][colIndex];
            const originalDateText = getCellText(originalValue);

            const tempInput = document.createElement('input');
            tempInput.type = 'date';
            tempInput.className = 'w-full h-full border-none outline-none';

            if (originalValue instanceof Date) {
                tempInput.value = originalValue.toISOString().split('T')[0];
            } else if (String(originalDateText).match(/^\d{4}\/\d{2}\/\d{2}$/)) {
                tempInput.value = originalDateText.replace(/\//g, '-');
            }

            cell.innerHTML = '';
            cell.appendChild(tempInput);
            tempInput.focus();

            const cleanup = () => {
                tempInput.removeEventListener('blur', onBlur);
                tempInput.removeEventListener('change', onChange);
            };

            const onBlur = () => {
                cell.innerHTML = originalDateText || '&nbsp;';
                cleanup();
            };

            const onChange = () => {
                const newDateValue = tempInput.value;
                if (newDateValue) {
                    const parts = newDateValue.split('-');
                    const newDate = new Date(parts[0], parts[1] - 1, parts[2]);
                    const newDateText = getCellText(newDate);

                    const changeKey = `${rowIndex}-${colIndex}`;

                    if (!dateChanges[changeKey]) {
                        dateChanges[changeKey] = {
                            key: key,
                            rowIndex: rowIndex,
                            column: colName,
                            from: originalDateText,
                            to: newDateText,
                        };
                    } else {
                        dateChanges[changeKey].to = newDateText;
                    }

                    currentDataArray[rowIndex + 1][colIndex] = newDate;

                    // 重渲染：會保持隱藏/寬度（因為 applySavedColumnStates）
                    renderTable();
                    applyFiltersToTable();
                    colorGanttCells();
                } else {
                    cell.innerHTML = originalDateText || '&nbsp;';
                }
                cleanup();
            };

            tempInput.addEventListener('blur', onBlur);
            tempInput.addEventListener('change', onChange);
        }

        viewChangesBtn.addEventListener('click', () => {
            const changes = Object.values(dateChanges);
            if (changes.length === 0) {
                changesList.innerHTML = '<p class="text-gray-500 text-center">目前沒有任何變更。</p>';
            } else {
                changes.sort((a, b) => {
                    if (a.rowIndex !== b.rowIndex) return a.rowIndex - b.rowIndex;
                    return a.column.localeCompare(b.column);
                });
                let html = '<ul class="divide-y divide-gray-200">';
                changes.forEach(change => {
                    html += `<li class="py-2"><span class="font-bold">${change.key}</span> 的 <span class="font-semibold text-blue-600">${change.column}</span> 從 <span class="text-red-500">${change.from}</span> 改為 <span class="text-green-600">${change.to}</span></li>`;
                });
                html += '</ul>';
                changesList.innerHTML = html;
            }
            changesModal.classList.remove('hidden');
        });

        changesModalCloseBtn.addEventListener('click', () => {
            changesModal.classList.add('hidden');
        });

        downloadChangesBtn.addEventListener('click', () => {
            const changes = Object.values(dateChanges);
            if (changes.length === 0) { alert('沒有變更紀錄可供下載。'); return; }

            changes.sort((a, b) => {
                if (a.rowIndex !== b.rowIndex) return a.rowIndex - b.rowIndex;
                return a.column.localeCompare(b.column);
            });

            let textContent = '變更紀錄\n==========\n\n';
            changes.forEach(change => {
                textContent += `議題 "${change.key}" 的欄位 "${change.column}" 從 "${change.from}" 改為 "${change.to}"\n`;
            });

            const blob = new Blob([textContent], { type: 'text/plain;charset=utf-8' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = 'change_log.txt';
            document.body.appendChild(a); a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        });

        // =================================================================================
        // ============================ 頁面初始化與輔助 ============================
        // =================================================================================
        excelContent.addEventListener('click', (event) => {
            const toggle = event.target.closest('.toggle-collapse');
            if (!toggle) return;

            const parentRow = toggle.closest('tr');
            if (!parentRow || parentRow.dataset.isParent !== 'true') return;

            const idx = getOriginalRowIndexSafe(parentRow);
            if (idx < 0) return;

            const key = getStableRowKeyByIndex(idx);
            const wasCollapsed = collapsedState.has(key)
                ? collapsedState.get(key)
                : (parentRow.dataset.collapsed === 'true');

            const nowCollapsed = !wasCollapsed;
            collapsedState.set(key, nowCollapsed);
            parentRow.dataset.collapsed = String(nowCollapsed);
            toggle.textContent = nowCollapsed ? '+' : '−';

            // ★ 只在「收折」時做級聯，把所有子孫也標記為收折
            if (nowCollapsed === true) {
                cascadeCollapseDescendants(parentRow);
            }

            applyFiltersToTable();
        });




        highlightToggle.addEventListener('change', clearAllHighlights);
        collapseAllBtn.addEventListener('click', () => toggleAllRows(true));
        expandAllBtn.addEventListener('click', () => toggleAllRows(false));

        function setDefaultDates() {
            const today = new Date();
            const year = today.getFullYear();
            const month = String(today.getMonth() + 1).padStart(2, '0');
            const day = String(today.getDate()).padStart(2, '0');
            startDateInput.value = `${year}-${month}-${day}`;
            endDateInput.value = `${year}-12-31`;
        }

        function escapeHTML(str) {
            if (typeof str !== 'string') str = String(str);
            return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#039;');
        }


        // ======================= 統計分析（表1 / 表2） =======================

        // 取欄位索引（大小寫一致即可）
        function idxOf(headerName){
            const h = (currentDataArray[0] || []).map(x => (x == null ? '' : String(x)));
            return h.findIndex(v => String(v).trim().toLowerCase() === headerName.trim().toLowerCase());
        }

        // 解析日期：支援 Date 或 YYYY/MM/DD 字串
        function parseCellDate(raw){
            if (!raw) return null;
            if (raw instanceof Date && !isNaN(raw)) return raw;
            const s = getCellText(raw);
            if (!s) return null;
            const m = s.match(/^(\d{4})\/(\d{2})\/(\d{2})$/);
            if (!m) return null;
            const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
            return isNaN(d) ? null : d;
        }

        // 日期區間是否相交（含邊界）
        function overlapInclusive(aStart, aEnd, bStart, bEnd){
        return aStart <= bEnd && aEnd >= bStart;
        }

        // 拆 Value Category（逗號/頓號/中文逗號）並正規化大小寫；空白視為 "Issue Bug"
        function splitCategories(raw){
            const txt = getCellText(raw);
            if (!txt) return ["Issue Bug"];
            const parts = txt.split(/[,，、;]/).map(s => s.trim()).filter(Boolean);
            if (parts.length === 0) return ["Issue Bug"];
            // 規一化：首字母大寫，其餘小寫（Class1 / Class2...）
            return parts.map(s => s[0] ? (s[0].toUpperCase() + s.slice(1)) : s).map(x => x.trim());
        }

        // 正規化狀態顯示（大小寫不敏感）
        function normStatus(raw){
            const s = (getCellText(raw) || '').trim();
            const t = s.toLowerCase();
            const map = {
                "to do":"To Do",
                "todevelop":"To Develop","to develop":"To Develop",
                "in progress":"In Progress","inprogress":"In Progress",
                "done":"Done",
                "v&v":"V&V","v & v":"V&V"
            };
            return map[t] || s || "(未填)";
        }

        // 拆分 Assignees 並去重（逗號/頓號/中文逗號）
        function splitAssignees(raw){
            const s = getCellText(raw);
            if (!s) return [];
            return Array.from(new Set(
                s.split(/[,，、;]/).map(x => x.trim()).filter(Boolean)
            ));
        }

        // 建構表格資料：
        // isTopMode = true  -> 表格1（Top level）
        // isTopMode = false -> 表格2（非 Top level）
        // 依 isTopMode 產生統計（true=表格1 Top level；false=表格2 Child rows）
        // 依 isTopMode 產生統計（true=表格1 Top level；false=表格2 Child rows）
        function buildStats(isTopMode){
            if (!currentDataArray || currentDataArray.length <= 1) {
                return { headers: [], rows: [], totalRow: [] };
            }

            // 主要欄位索引
            const COL = {
                status:    idxOf('Status'),
                valueCat:  idxOf('Value Category'),
                assignees: idxOf('Backend Assignees'),
                devStart:  idxOf('Dev Start Date'),
                due:       idxOf('Due Date'),
            };
            if (COL.status < 0 || COL.valueCat < 0 || COL.assignees < 0 || COL.devStart < 0 || COL.due < 0) {
                return { headers: [], rows: [], totalRow: [] };
            }

            // 固定狀態順序（兩表不同）
            const fixedOrder = isTopMode
                ? ['To Do', 'To Develop', 'In Progress', 'Done']  // 表格1
                : ['To Do', 'In Progress', 'V&V', 'Done'];        // 表格2

            // 狀態正規化
            const norm = (s) => {
                const t = String(s || '').trim();
                const u = t.toUpperCase();
                if (u === 'TO DO') return 'To Do';
                if (u === 'TO DEVELOP' || u === 'TODEVELOP') return 'To Develop';
                if (u === 'IN PROGRESS' || u === 'WIP') return 'In Progress';
                if (u === 'DONE') return 'Done';
                if (u === 'V&V' || u === 'V & V' || u === 'VNV' || u === 'V V') return 'V&V';
                return t; // 未預期狀態保留原樣（動態欄）
            };

            // 分割工具
            const splitCats = (v) => {
                const raw = getCellText(v);
                const list = raw ? raw.split(/[,，、;]/).map(s => s.trim()).filter(Boolean) : [];
                return list.length ? list : ['Issue Bug'];
            };
            const splitAssignees = (v) => {
                const raw = getCellText(v);
                return raw ? raw.split(/[,，、;]/).map(s => s.trim()).filter(Boolean) : [];
            };

            // UI 日期
            const sd = startDateInput?.value ? new Date(startDateInput.value) : null;
            const ed = endDateInput?.value ? new Date(endDateInput.value) : null;

            // 聚合器：Map<cat, { jiraCount:number, people:Set<string>, status:Map<status, count> }>
            const agg = new Map();
            // 動態狀態（未在 fixedOrder 的），按首次出現先後
            const seenDynamic = [];

            // 逐列掃描
            for (let r = 1; r < currentDataArray.length; r++) {
                const level = currentRowMeta?.[r] ?? 0;
                if (isTopMode && level > 0) continue;     // 表1：僅 top level
                if (!isTopMode && level === 0) continue;  // 表2：僅 child

                const row = currentDataArray[r];
                const d1  = parseCellDate(row[COL.devStart]);
                const d2  = parseCellDate(row[COL.due]);
                if (!d1 || !d2) continue;     // 限制2：不得空白
                if (+d1 > +d2) continue;      // 防呆

                // 限制3：兩個日期區間有交集才納入
                let pass = false;
                if (sd && ed) {
                // 若 devStart <= 結束 && due >= 起始 → 表示有交集
                if (d1 <= ed && d2 >= sd) pass = true;
                } else if (sd && !ed) {
                // 只有起始日期：若 due >= 起始日期，表示仍有交集
                if (d2 >= sd) pass = true;
                } else if (!sd && ed) {
                // 只有結束日期：若 devStart <= 結束日期，表示仍有交集
                if (d1 <= ed) pass = true;
                } else {
                // 起始與結束都沒設定 → 全部算入
                pass = true;
                }
                if (!pass) continue;

                // 類別 + 狀態 + 人力
                const cats = splitCats(row[COL.valueCat]);
                const st   = norm(row[COL.status]);
                const ppl  = splitAssignees(row[COL.assignees]);

                // 動態狀態登錄
                if (!fixedOrder.includes(st) && !seenDynamic.includes(st)) {
                seenDynamic.push(st);
                }

                // 聚合到每個類別
                cats.forEach(cat => {
                if (!agg.has(cat)) {
                    agg.set(cat, { jiraCount: 0, people: new Set(), status: new Map() });
                }
                const a = agg.get(cat);
                a.jiraCount += 1;                                // JIRA 總數（多分類各+1）
                ppl.forEach(p => a.people.add(p));               // 人力去重
                a.status.set(st, (a.status.get(st) || 0) + 1);   // 狀態次數
                });
            }

            // 表頭：固定狀態在前、動態狀態在後，最後是「人力」
            const headers = ['Value Category', 'JIRA總數', ...fixedOrder, ...seenDynamic, '人力'];

            // 排序規則：一般類別在前 → 含「專案」或「技術研究」的類別 → 最後才是 Issue Bug
            const catPriority = (name) => {
                const s = String(name || '').trim();
                if (s === 'Issue Bug') return 1;                     // 倒數第二群
                if (/專案|技術研究/i.test(s)) return 2;               // 最底
                return 0;                                              // 一般在最前面
            };

            const catsSorted = Array.from(agg.keys()).sort((a, b) => {
                const pa = catPriority(a), pb = catPriority(b);
                if (pa !== pb) return pa - pb;                         // 先比群組
                return a.localeCompare(b, 'zh-Hant');                  // 同群依名稱排序
            });


            // 輸出 rows 與總計
            const rows = [];
            const totals = new Array(headers.length).fill(0);
            totals[0] = 'Total';

            catsSorted.forEach(cat => {
                const a = agg.get(cat);
                const line = new Array(headers.length).fill(0);
                line[0] = cat;

                // JIRA 總數
                line[1] = a.jiraCount;
                totals[1] += a.jiraCount;

                // 固定狀態
                fixedOrder.forEach((s, i) => {
                const cnt = a.status.get(s) || 0;
                line[2 + i] = cnt;
                totals[2 + i] += cnt;
                });

                // 動態狀態
                seenDynamic.forEach((s, k) => {
                const idx = 2 + fixedOrder.length + k;
                const cnt = a.status.get(s) || 0;
                line[idx] = cnt;
                totals[idx] += cnt;
                });

                // 人力（唯一人數）
                const pplIdx = headers.length - 1;
                const pplCount = a.people.size;
                line[pplIdx] = pplCount;
                totals[pplIdx] += pplCount; // 你的樣例是做直行加總，我沿用
                rows.push(line);
            });

            return { headers, rows, totalRow: totals };
            }



        // 將資料渲染成表格（可點數字 + 明細觸發）
        function renderStatsTable(tblEl, data){
            if (!tblEl) return;
            tblEl.innerHTML = "";

            const { headers, rows, totalRow } = data || { headers:[], rows:[], totalRow:[] };
            if (!headers.length){
                tblEl.innerHTML =
                '<thead><tr><th>無法產生</th></tr></thead><tbody><tr><td>缺少必要欄位或日期區間未設定。</td></tr></tbody>';
                return;
            }

            // colgroup：第一欄較寬，其餘欄齊右
            const colgroup = document.createElement('colgroup');
            headers.forEach((_, i) => {
                const col = document.createElement('col');
                if (i > 0) col.className = 'num';
                colgroup.appendChild(col);
            });
            tblEl.appendChild(colgroup);

            // thead
            const thead = document.createElement('thead');
            const trh = document.createElement('tr');
            headers.forEach((h, i) => {
                const th = document.createElement('th');
                th.textContent = h;
                if (i > 0) th.classList.add('num');
                trh.appendChild(th);
            });
            thead.appendChild(trh);

            // tbody（資料列）
            const tbody = document.createElement('tbody');
            rows.forEach(r => {
                const tr = document.createElement('tr');
                r.forEach((v, i) => {
                const td = document.createElement('td');
                td.textContent = (i === 0) ? String(v) : String(v ?? 0);
                if (i > 0) {
                    // ★ 讓數字可點，並記下欄名與類別
                    td.classList.add('num', 'stats-clickable');
                    td.dataset.col = headers[i];
                    td.dataset.cat = r[0];
                }
                tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });

            // 總計列（不加可點，避免 cat=Total 無對應明細）
            if (totalRow && totalRow.length){
                const trt = document.createElement('tr');
                trt.classList.add('row-total');
                totalRow.forEach((v, i) => {
                const td = document.createElement('td');
                td.textContent = (i === 0) ? String(v) : String(v ?? 0);
                if (i > 0) td.classList.add('num');
                trt.appendChild(td);
                });
                tbody.appendChild(trt);
            }

            tblEl.appendChild(thead);
            tblEl.appendChild(tbody);

            // ★ 事件代理：點任一數字即顯示明細
            // 用 onassign 取代 addEventListener，避免重複綁定造成多次觸發
            tblEl.onclick = (e) => {
                const td = e.target.closest('td.stats-clickable');
                if (!td) return;
                const colName = td.dataset.col;
                const catName = td.dataset.cat;
                const isTopMode = (tblEl.id === 'stats-table-1'); // 表1=Top level；表2=child
                showDetailForCell({ catName, colName, isTopMode });
            };
        }



        // 需要用到既有的工具函式：idxOf / getCellText / parseCellDate
        // 也會用到：currentDataArray, currentRowMeta, startDateInput, endDateInput

        const statsDetailWrap  = document.getElementById('stats-detail');
        const statsDetailTitle = document.getElementById('stats-detail-title');
        const statsDetailTable = document.getElementById('stats-detail-table');
        const statsDetailClose = document.getElementById('stats-detail-close');

        // 小工具：時間是否落在區間內（含邊界）
        
    // ---- Date helpers (compare by YYYY-MM-DD; avoids timezone issues) ----
    function __dateToYMD(d){
    if (!d || isNaN(+d)) return null;
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const dd = String(d.getDate()).padStart(2,'0');
    return `${y}-${m}-${dd}`;
    }
    function __inRangeYMD(d, l, r){
    if (!d || !l || !r) return false;
    const ds = __dateToYMD(d), ls = __dateToYMD(l), rs = __dateToYMD(r);
    if (!ds || !ls || !rs) return false;
    return (ds >= ls) && (ds <= rs); // inclusive
    }
    function __eqYMD(a, b){
    const as = __dateToYMD(a), bs = __dateToYMD(b);
    return !!(as && bs && as === bs);
    }
    function inRange(day, a, b){ 
        if (!day || !a || !b) return false; 
        const t = +day, ta = +a, tb = +b; 
        return t >= ta && t <= tb; 
        }

        // 顯示明細主函式
        // 顯示明細主函式（點統計數字後彈出；含「包含式 OR」日期邏輯）
        function showDetailForCell({ catName, colName, isTopMode }) {
            // 主要欄位索引
            const COL = {
                status:    idxOf('Status'),
                valueCat:  idxOf('Value Category'),
                assignees: idxOf('Backend Assignees'),
                devStart:  idxOf('Dev Start Date'),
                due:       idxOf('Due Date')
            };
            if (COL.status < 0 || COL.valueCat < 0 || COL.assignees < 0 || COL.devStart < 0 || COL.due < 0) {
                return;
            }

            // 狀態正規化（需與 buildStats 一致）
            const norm = (s) => {
                const t = String(s || '').trim();
                const u = t.toUpperCase();
                if (u === 'TO DO') return 'To Do';
                if (u === 'TO DEVELOP' || u === 'TODEVELOP') return 'To Develop';
                if (u === 'IN PROGRESS' || u === 'WIP') return 'In Progress';
                if (u === 'DONE') return 'Done';
                if (u === 'V&V' || u === 'V & V' || u === 'VNV' || u === 'V V') return 'V&V';
                return t; // 未預期狀態維持原樣
            };

            // UI 篩選區間
            const sd = startDateInput.value ? new Date(startDateInput.value) : null;
            const ed = endDateInput.value ? new Date(endDateInput.value) : null;

            // 明細欄位 = 主畫面「凍結欄數」前 N 欄
            const freezeN = Math.max(1, parseInt(document.getElementById('freeze-cols').value || '1', 10));
            const header  = (currentDataArray[0] || []).slice(0, freezeN).map(getCellText);

            // 掃描資料行
            const rows = [];
            for (let r = 1; r < currentDataArray.length; r++) {
                const level = currentRowMeta[r] || 0;
                if (isTopMode && level > 0) continue;    // 表1：只收 top
                if (!isTopMode && level === 0) continue; // 表2：只收 child

                const row  = currentDataArray[r];
                const d1   = parseCellDate(row[COL.devStart]);
                const d2   = parseCellDate(row[COL.due]);
                if (!d1 || !d2 || (+d1 > +d2)) continue; // 限制2 + 防呆

                // 限制3：兩個日期區間有交集才納入
                let pass = false;
                if (sd && ed) {
                // 若 devStart <= 結束 && due >= 起始 → 表示有交集
                if (d1 <= ed && d2 >= sd) pass = true;
                } else if (sd && !ed) {
                // 只有起始日期：若 due >= 起始日期，表示仍有交集
                if (d2 >= sd) pass = true;
                } else if (!sd && ed) {
                // 只有結束日期：若 devStart <= 結束日期，表示仍有交集
                if (d1 <= ed) pass = true;
                } else {
                // 起始與結束都沒設定 → 全部算入
                pass = true;
                }
                if (!pass) continue;


                // Value Category 比對（支援多值；空值視為 "Issue Bug"）
                const cats = (() => {
                const raw = getCellText(row[COL.valueCat]);
                const list = raw ? raw.split(/[,，、;]/).map(s => s.trim()).filter(Boolean) : [];
                return list.length ? list : ['Issue Bug'];
                })();

                // 欄位型別：JIRA總數 / 人力 / 其餘視為狀態欄
                const isJiraTotalCol = (colName === 'JIRA總數');
                const isPeopleCol    = (colName === '人力');
                const rowStatus      = norm(row[COL.status]);

                // 要符合被點的 Value Category
                if (!isJiraTotalCol && !isPeopleCol) {
                // 狀態欄：狀態要命中 + 類別要含 catName
                if (rowStatus !== colName) continue;
                }
                if (!cats.includes(catName) && catName !== 'Total') continue;

                // 收集「凍結欄」資料作為明細
                rows.push((row || []).slice(0, freezeN));
            }

            // 欄寬自動（以各欄最大字元數推 min-width:ch）
            const widths = new Array(header.length).fill(0);
            const textRows = rows.map(r => r.map(getCellText));
            textRows.forEach(r => r.forEach((cell, i) => widths[i] = Math.max(widths[i], String(cell ?? '').length)));
            header.forEach((h, i) => widths[i] = Math.max(widths[i], String(h ?? '').length));

            // 繪製明細表
            statsDetailTitle.textContent = `明細：${catName} / ${colName}（共 ${rows.length} 筆）`;
            const thead = '<thead><tr>' + header.map((h,i)=>`<th style="min-width:${Math.min(60, widths[i])}ch;">${h}</th>`).join('') + '</tr></thead>';
            const tbody = '<tbody>' + textRows.map(r => '<tr>' + r.map((c,i)=>`<td style="min-width:${Math.min(60, widths[i])}ch;">${getCellText(c)}</td>`).join('') + '</tr>').join('') + '</tbody>';
            statsDetailTable.innerHTML = thead + tbody;

            // 顯示
            statsDetailWrap.classList.remove('hidden');
            statsDetailWrap.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
            }


        statsDetailClose?.addEventListener('click', () => {
        statsDetailWrap.classList.add('hidden');
        });



        // 將兩張表做 CSV 下載（合併到一個檔案，兩表之間空一行）
        function downloadStatsCSV(){
            const t1 = buildStats(true);
            const t2 = buildStats(false);

            function toCSV({headers, rows, totalRow}){
                if (!headers?.length) return "無法產生";
                const encode = val => `"${String(val ?? '').replace(/"/g,'""')}"`;
                const lines = [];
                lines.push(headers.map(encode).join(','));
                rows.forEach(r => lines.push(r.map(encode).join(',')));
                if (totalRow && totalRow.length) lines.push(totalRow.map(encode).join(','));
                return lines.join('\n');
            }

            const csvBody = [
                '表格1（Top level）',
                toCSV(t1),
                '',
                '表格2（非 Top level）',
                toCSV(t2)
            ].join('\n');

            // ★ 重要：加入 UTF-8 BOM，避免 Excel 開啟後中文亂碼
            const bom = "\uFEFF";
            const blob = new Blob([bom + csvBody], { type: 'text/csv;charset=utf-8' });

            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; 
            a.download = 'stats.csv';
            document.body.appendChild(a); 
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }


        // 綁定按鈕事件
        const statsBtn = document.getElementById('stats-btn');
        const statsModal = document.getElementById('stats-modal');
        const statsCloseBtn = document.getElementById('stats-close-btn');
        const statsDownloadBtn = document.getElementById('stats-download-btn');

        function openStats(){
            // 依目前 UI 區間重算兩表
            const data1 = buildStats(true);
            const data2 = buildStats(false);
            renderStatsTable(document.getElementById('stats-table-1'), data1);
            renderStatsTable(document.getElementById('stats-table-2'), data2);

            // 顯示範圍提示
            const sd = startDateInput.value || '';
            const ed = endDateInput.value || '';
            const hint1 = document.getElementById('stats1-range-hint');
            const hint2 = document.getElementById('stats2-range-hint');
            const txt = (sd && ed)
                ? `統計區間：${sd} ~ ${ed}`
                : '請先設定起訖日期';
            hint1.textContent = txt; hint2.textContent = txt;

            statsModal.classList.remove('hidden');
        }
        function closeStats(){ statsModal.classList.add('hidden'); }

        if (statsBtn) statsBtn.addEventListener('click', openStats);
        if (statsCloseBtn) statsCloseBtn.addEventListener('click', closeStats);
        if (statsDownloadBtn) statsDownloadBtn.addEventListener('click', downloadStatsCSV);

        // 若使用者改了日期或重產日期欄，再次打開也會重算；（不用即時監聽）





        setDefaultDates();
        
        // ---- 高亮目前選取的儲存格（表格1/2 共用） ----
        let __activeStatsCell = null;
        function __setActiveCell(td){
        if (__activeStatsCell && __activeStatsCell !== td) {
            __activeStatsCell.classList.remove('cell-active');
        }
        __activeStatsCell = td || null;
        if (td) td.classList.add('cell-active');
        }
        function __clearActiveCell(){
        if (__activeStatsCell) {
            __activeStatsCell.classList.remove('cell-active');
            __activeStatsCell = null;
        }
        }
        // 事件代理：點數字 → 設定高亮（同時會觸發明細）
        (function bindHighlightDelegation(){
            const statsModalEl = document.getElementById('stats-modal');
            statsModalEl?.addEventListener('click', (e) => {
                const td = e.target.closest('td.stats-clickable');
                if (td) {
                __setActiveCell(td);
                return;
                }
                // 點到空白處（非可點儲存格）：清除高亮
                const withinTables = e.target.closest('#stats-table-1, #stats-table-2');
                if (!withinTables) __clearActiveCell();
            });
            
            // 切換分頁按鈕時也清除
            document.getElementById('tab-1')?.addEventListener('click', __clearActiveCell);
            document.getElementById('tab-2')?.addEventListener('click', __clearActiveCell);
            // 關閉統計視窗時清除
            document.getElementById('stats-close-btn')?.addEventListener('click', __clearActiveCell);
        })();
        

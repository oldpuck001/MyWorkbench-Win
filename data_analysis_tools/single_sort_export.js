// single_sort_export.js

const { ipcRenderer } = require('electron');

let filePath = '';

export async function single_sort_exportFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.style.fontFamily = 'Arial, sans-serif';
    contentDiv.style.margin = '20px';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <style>
            .row {
                display: flex;
                margin-bottom: 10px;
            }
            .row input, .row select {
                margin-right: 10px;
            }
            .row button {
                margin-left: 10px;
            }
            .readonly {
                cursor: not-allowed;
            }
            .export {
                margin-top: 20px;
            }
        </style>

        <h1 style="text-align: center; width: 100%;">分类导出表格（单列金额）</h1>
        
        <div id="mainLayout" style="display: flex; flex-direction: column; align-items: center;">
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 60px; text-align: left;">源表格</label>
                <input id="source_path" type="text" style="width: 708px;">
            </div>
            <br>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label for="sheetDropdown" style="margin-right: 5px;">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown" style="width: 70px; margin-right: 20px;">
                </select>

                <label for="exportDropdown" style="margin-right: 5px;">选择导出分类列</label>
                <select id="exportDropdown" name="exportDropdown" style="width: 70px; margin-right: 20px;">
                </select>

                <label for="statisticsDropdown" style="margin-right: 5px;">选择统计分类列</label>
                <select id="statisticsDropdown" name="statisticsDropdown" style="width: 70px; margin-right: 20px;">
                </select>

                <label for="valueDropdown" style="margin-right: 5px;">选择统计数值列</label>
                <select id="valueDropdown" name="valueDropdown" style="width: 70px;">
                </select>
            </div>
        </div>

        <div class="export" style="text-align: center;">
            <button id="importButton" style="width: 180px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 15px;">导入</button>
            <button id="exportButton" style="width: 180px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 15px;">导出</button>
        </div>
    `;

    var input = document.getElementById('source_path');
    input.classList.add('readonly');
    input.readOnly = true;

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        filePath = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!filePath) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'single_sort_export_import', data: {'filePath': filePath }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'single_sort_export_index', data: {'filePath': filePath, 'sheetName': selectedSheet }});
    });

    // 自动触发工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('exportButton').addEventListener('click', async () => {
        if (!filePath) {
            alert('请先导入文件！');
            return;
        }

        const data = {
            sheet_name: document.getElementById('sheetDropdown').value,
            sort_column: document.getElementById('exportDropdown').value,
            secondary_column: document.getElementById('statisticsDropdown').value,
            value_column: document.getElementById('valueDropdown').value,
            filePath: filePath,
        };
        ipcRenderer.send('asynchronous-message', { command: 'single_sort_export_export', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'single_sort_export_import') {
            // 将文件路径显示在输入框中
            document.getElementById(`source_path`).value = filePath;

            // 获取 select 元素
            const sheetDropdown = document.getElementById('sheetDropdown');
            // 清空旧的选项
            sheetDropdown.innerHTML = '';

            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheetDropdown.appendChild(option);
            });

            triggerChange(sheetDropdown);       // 触发 change 事件，自动加载列数据

            alert('导入成功！');
        }

        if (result[0] === 'single_sort_export_index') {
            const exportDropdown = document.getElementById('exportDropdown');
            const statisticsDropdown = document.getElementById('statisticsDropdown');
            const valueDropdown = document.getElementById('valueDropdown');
    
            // 清空旧的选项
            exportDropdown.innerHTML = '';
            statisticsDropdown.innerHTML = '';
            valueDropdown.innerHTML = '';
    
            result[1].forEach(column => {
                const optionRow = document.createElement('option');
                optionRow.value = column;
                optionRow.text = column;
                exportDropdown.appendChild(optionRow);
    
                const optionColumn = document.createElement('option');
                optionColumn.value = column;
                optionColumn.text = column;
                statisticsDropdown.appendChild(optionColumn);
    
                const optionValue = document.createElement('option');
                optionValue.value = column;
                optionValue.text = column;
                valueDropdown.appendChild(optionValue);
            });
        }

        if (result[0] === 'single_sort_export_export') {
            alert('导出成功！');
        }

        if (result[0] === 'single_sort_export_no') {
            alert('导出失败！');
        }
    });
}
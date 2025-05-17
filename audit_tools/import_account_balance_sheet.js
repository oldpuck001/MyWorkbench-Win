// import_account_balance_sheet.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function import_account_balance_sheet() {

    window.project_folder = await ipcRenderer.invoke('get-project-folder');

    if (window.project_folder) {
        ipcRenderer.send('asynchronous-message', { command: 'import_config', data: { project_folder: window.project_folder } });
    }

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>导入本期科目余额表</h1>
        
        <div class="import">

            <div>
                <label>源数据文件</label>
                <input id="file_path" type="text">
            </div>

            <div>
                <label for="sheetDropdown">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown"></select>
            </div>

            <div>
                <label for="account_idDropdown">选择科目代码列</label>
                <select id="account_idDropdown" name="account_idDropdown"></select>
            </div>

            <div>
                <label for="account_nameDropdown">选择科目名称列</label>
                <select id="account_nameDropdown" name="account_nameDropdown"></select>
            </div>

            <div>
                <label for="begin_debitDropdown">选择期初借方列</label>
                <select id="begin_debitDropdown" name="begin_debitDropdown"></select>
            </div>

            <div>
                <label for="begin_creditDropdown">选择期初贷方列</label>
                <select id="begin_creditDropdown" name="begin_creditDropdown"></select>
            </div>

            <div>
                <label for="this_debitDropdown">选择本期借方列</label>
                <select id="this_debitDropdown" name="this_debitDropdown"></select>
            </div>

            <div>
                <label for="this_creditDropdown">选择本期贷方列</label>
                <select id="this_creditDropdown" name="this_creditDropdown"></select>
            </div>

            <div>
                <label for="end_debitDropdown">选择期末借方列</label>
                <select id="end_debitDropdown" name="end_debitDropdown"></select>
            </div>

            <div>
                <label for="end_creditDropdown">选择期末贷方列</label>
                <select id="end_creditDropdown" name="end_creditDropdown"></select>
            </div>

            <div>
                <button id="importButton">导入源文件</button>
                <button id="exportButton">导出余额表</button>
            </div>
        </div>
    `;

    var file_path_input = document.getElementById('file_path');
    file_path_input.classList.add('readonly');
    file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const account_idDropdown = document.getElementById('account_idDropdown');
    const account_nameDropdown = document.getElementById('account_nameDropdown');
    const begin_debitDropdown = document.getElementById('begin_debitDropdown');
    const begin_creditDropdown = document.getElementById('begin_creditDropdown');
    const this_debitDropdown = document.getElementById('this_debitDropdown');
    const this_creditDropdown = document.getElementById('this_creditDropdown');
    const end_debitDropdown = document.getElementById('end_debitDropdown');
    const end_creditDropdown = document.getElementById('end_creditDropdown');

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'import_account_balance_sheet', data: {'file_path': file_path }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'index_account_balance_sheet', data: {'file_path': file_path, 'sheetName': selectedSheet }});
    });

    // 自动触发工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('exportButton').addEventListener('click', async () => {
        if (!file_path) {
            alert('请先导入文件！');
            return;
        }

        const data = {
            project_folder: window.project_folder,
            file_path: file_path,
            sheet_name: sheetDropdown.value,
            account_id: account_idDropdown.value,
            account_name: account_nameDropdown.value,
            begin_debit: begin_debitDropdown.value,
            begin_credit: begin_creditDropdown.value,
            this_debit: this_debitDropdown.value,
            this_credit: this_creditDropdown.value,
            end_debit: end_debitDropdown.value,
            end_credit: end_creditDropdown.value,
        };

        ipcRenderer.send('asynchronous-message', { command: 'export_account_balance_sheet', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'import_account_balance_sheet') {
            // 将文件路径显示在输入框中
            file_path_input.value = file_path;

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

        if (result[0] === 'index_account_balance_sheet') {
    
            // 清空旧的选项
            account_idDropdown.innerHTML = '';
            account_nameDropdown.innerHTML = '';
            begin_debitDropdown.innerHTML = '';
            begin_creditDropdown.innerHTML = '';
            this_debitDropdown.innerHTML = '';
            this_creditDropdown.innerHTML = '';
            end_debitDropdown.innerHTML = '';
            end_creditDropdown.innerHTML = '';

            result[1].forEach(column => {
                const account_idOption = document.createElement('option');
                account_idOption.value = column;
                account_idOption.text = column;
                account_idDropdown.appendChild(account_idOption);
    
                const account_nameOption = document.createElement('option');
                account_nameOption.value = column;
                account_nameOption.text = column;
                account_nameDropdown.appendChild(account_nameOption);
    
                const begin_debitOption = document.createElement('option');
                begin_debitOption.value = column;
                begin_debitOption.text = column;
                begin_debitDropdown.appendChild(begin_debitOption);
    
                const begin_creditOption = document.createElement('option');
                begin_creditOption.value = column;
                begin_creditOption.text = column;
                begin_creditDropdown.appendChild(begin_creditOption);

                const this_debitOption = document.createElement('option');
                this_debitOption.value = column;
                this_debitOption.text = column;
                this_debitDropdown.appendChild(this_debitOption);

                const this_creditOption = document.createElement('option');
                this_creditOption.value = column;
                this_creditOption.text = column;
                this_creditDropdown.appendChild(this_creditOption);
    
                const end_debitOption = document.createElement('option');
                end_debitOption.value = column;
                end_debitOption.text = column;
                end_debitDropdown.appendChild(end_debitOption);

                const end_creditOption = document.createElement('option');
                end_creditOption.value = column;
                end_creditOption.text = column;
                end_creditDropdown.appendChild(end_creditOption);
            });
        }

        if (result[0] === 'export_account_balance_sheet') {
            alert('导出成功！');
        }
    });
}
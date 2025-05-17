// import_chronological_account.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function import_chronological_account() {

    window.project_folder = await ipcRenderer.invoke('get-project-folder');

    if (window.project_folder) {
        ipcRenderer.send('asynchronous-message', { command: 'import_config', data: { project_folder: window.project_folder } });
    }

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>导入本期序时账</h1>
        
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
                <label for="account_dateDropdown">选择凭证日期列</label>
                <select id="account_dateDropdown" name="account_dateDropdown"></select>
            </div>

            <div>
                <label for="account_numberDropdown">选择凭证字号列</label>
                <select id="account_numberDropdown" name="account_numberDropdown"></select>
            </div>

            <div>
                <label for="account_nameDropdown">选择科目名称列</label>
                <select id="account_nameDropdown" name="account_nameDropdown"></select>
            </div>

            <div>
                <label for="account_summaryDropdown">选择摘要文本列</label>
                <select id="account_summaryDropdown" name="account_summaryDropdown"></select>
            </div>

            <div>
                <label for="account_debitDropdown">选择借方金额列</label>
                <select id="account_debitDropdown" name="account_debitDropdown"></select>
            </div>

            <div>
                <label for="account_creditDropdown">选择贷方金额列</label>
                <select id="account_creditDropdown" name="account_creditDropdown"></select>
            </div>

            <div>
                <button id="importButton">导入源文件</button>
                <button id="exportButton">导出序时账</button>
            </div>
        </div>
    `;

    var file_path_input = document.getElementById('file_path');
    file_path_input.classList.add('readonly');
    file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const account_dateDropdown = document.getElementById('account_dateDropdown');
    const account_numberDropdown = document.getElementById('account_numberDropdown');
    const account_nameDropdown = document.getElementById('account_nameDropdown');
    const account_summaryDropdown = document.getElementById('account_summaryDropdown');
    const account_debitDropdown = document.getElementById('account_debitDropdown');
    const account_creditDropdown = document.getElementById('account_creditDropdown');

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'import_chronological_account', data: {'file_path': file_path }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'index_chronological_account', data: {'file_path': file_path, 'sheetName': selectedSheet }});
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
            account_date: account_dateDropdown.value,
            account_number: account_numberDropdown.value,
            account_name: account_nameDropdown.value,
            account_summary: account_summaryDropdown.value,
            account_debit: account_debitDropdown.value,
            account_credit: account_creditDropdown.value,
        };

        ipcRenderer.send('asynchronous-message', { command: 'export_chronological_account', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'import_chronological_account') {
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

        if (result[0] === 'index_chronological_account') {
    
            // 清空旧的选项
            account_dateDropdown.innerHTML = '';
            account_numberDropdown.innerHTML = '';
            account_nameDropdown.innerHTML = '';
            account_summaryDropdown.innerHTML = '';
            account_debitDropdown.innerHTML = '';
            account_creditDropdown.innerHTML = '';

            result[1].forEach(column => {
                const account_dateOption = document.createElement('option');
                account_dateOption.value = column;
                account_dateOption.text = column;
                account_dateDropdown.appendChild(account_dateOption);
    
                const account_numberOption = document.createElement('option');
                account_numberOption.value = column;
                account_numberOption.text = column;
                account_numberDropdown.appendChild(account_numberOption);

                const account_nameOption = document.createElement('option');
                account_nameOption.value = column;
                account_nameOption.text = column;
                account_nameDropdown.appendChild(account_nameOption);

                const account_summaryOption = document.createElement('option');
                account_summaryOption.value = column;
                account_summaryOption.text = column;
                account_summaryDropdown.appendChild(account_summaryOption);

                const account_debitOption = document.createElement('option');
                account_debitOption.value = column;
                account_debitOption.text = column;
                account_debitDropdown.appendChild(account_debitOption);
    
                const account_creditOption = document.createElement('option');
                account_creditOption.value = column;
                account_creditOption.text = column;
                account_creditDropdown.appendChild(account_creditOption);
            });
        }

        if (result[0] === 'export_chronological_account') {
            alert('导出成功！');
        }
    });
}
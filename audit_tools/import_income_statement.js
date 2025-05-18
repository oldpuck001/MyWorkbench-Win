// import_income_statement.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function import_income_statement() {

    window.project_folder = await ipcRenderer.invoke('get-project-folder');

    if (window.project_folder) {
        ipcRenderer.send('asynchronous-message', { command: 'import_config', data: { project_folder: window.project_folder } });
    }

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>导入未审利润表</h1>
        
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
                <label for="items_nameDropdown">报表项目名称列</label>
                <select id="items_nameDropdown" name="items_nameDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                </select>
            </div>

            <div>
                <label for="items_thisDropdown">本期金额列</label>
                <select id="items_thisDropdown" name="items_thisDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                </select>
            </div>

            <div>
                <label for="items_previousDropdown">上期金额列</label>
                <select id="items_previousDropdown" name="items_previousDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                </select>
            </div>

            <div>
                <button id="selectButton">选择文件</button>
                <button id="importButton">导入数据</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>导入结果：</label>
            </div>
            <div>
                <textarea id="result_output" rows="18"></textarea>
            </div>
        </div>
    `;

    var file_path_input = document.getElementById('file_path');
    file_path_input.classList.add('readonly');
    file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const items_nameDropdown = document.getElementById('items_nameDropdown');
    const items_thisDropdown = document.getElementById('items_thisDropdown');
    const items_previousDropdown = document.getElementById('items_previousDropdown');
    const result_outputTextarea = document.getElementById('result_output');

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'select_income_statement', data: {'file_path': file_path }});
    });

    document.getElementById('importButton').addEventListener('click', async () => {
        if (!file_path) {
            alert('请先导入文件！');
            return;
        }

        const data = {
            project_folder: window.project_folder,
            file_path: file_path,
            sheet_name: sheetDropdown.value,
            items_name: items_nameDropdown.value,
            items_this: items_thisDropdown.value,
            items_previous: items_previousDropdown.value,
        };

        ipcRenderer.send('asynchronous-message', { command: 'import_income_statement', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'select_income_statement') {
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

            alert('选择文件成功！');
        }

        if (result[0] === 'import_income_statement') {
            
            alert('导入数据成功！');
            result_outputTextarea.value = JSON.stringify(result[1], null, 2);       // 2 是缩进空格数
        }
    });
}
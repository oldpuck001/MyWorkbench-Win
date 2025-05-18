// import_balance_sheet.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function import_balance_sheet() {

    window.project_folder = await ipcRenderer.invoke('get-project-folder');

    if (window.project_folder) {
        ipcRenderer.send('asynchronous-message', { command: 'import_config', data: { project_folder: window.project_folder } });
    }

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>导入未审资产负债表</h1>
        
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
                <label for="assets_nameDropdown">资产项目名称列</label>
                <select id="assets_nameDropdown" name="assets_nameDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                    <option value=5>F列</option>
                    <option value=6>G列</option>
                    <option value=7>H列</option>
                    <option value=8>I列</option>
                    <option value=9>J列</option>
                </select>
            </div>

            <div>
                <label for="assets_thisDropdown">资产本期期末列</label>
                <select id="assets_thisDropdown" name="assets_thisDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                    <option value=5>F列</option>
                    <option value=6>G列</option>
                    <option value=7>H列</option>
                    <option value=8>I列</option>
                    <option value=9>J列</option>
                </select>
            </div>

            <div>
                <label for="assets_previousDropdown">资产本期期初列</label>
                <select id="assets_previousDropdown" name="assets_previousDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                    <option value=5>F列</option>
                    <option value=6>G列</option>
                    <option value=7>H列</option>
                    <option value=8>I列</option>
                    <option value=9>J列</option>
                </select>
            </div>

            <div>
                <label for="liabilities_nameDropdown">负债权益名称列</label>
                <select id="liabilities_nameDropdown" name="liabilities_nameDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                    <option value=5>F列</option>
                    <option value=6>G列</option>
                    <option value=7>H列</option>
                    <option value=8>I列</option>
                    <option value=9>J列</option>
                </select>
            </div>

            <div>
                <label for="liabilities_thisDropdown">负债权益本期期末列</label>
                <select id="liabilities_thisDropdown" name="liabilities_thisDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                    <option value=5>F列</option>
                    <option value=6>G列</option>
                    <option value=7>H列</option>
                    <option value=8>I列</option>
                    <option value=9>J列</option>
                </select>
            </div>

            <div>
                <label for="liabilities_previousDropdown">负债权益本期期初列</label>
                <select id="liabilities_previousDropdown" name="liabilities_previousDropdown">
                    <option value=0>A列</option>
                    <option value=1>B列</option>
                    <option value=2>C列</option>
                    <option value=3>D列</option>
                    <option value=4>E列</option>
                    <option value=5>F列</option>
                    <option value=6>G列</option>
                    <option value=7>H列</option>
                    <option value=8>I列</option>
                    <option value=9>J列</option>
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
    const assets_nameDropdown = document.getElementById('assets_nameDropdown');
    const assets_thisDropdown = document.getElementById('assets_thisDropdown');
    const assets_previousDropdown = document.getElementById('assets_previousDropdown');
    const liabilities_nameDropdown = document.getElementById('liabilities_nameDropdown');
    const liabilities_thisDropdown = document.getElementById('liabilities_thisDropdown');
    const liabilities_previousDropdown = document.getElementById('liabilities_previousDropdown');
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
        ipcRenderer.send('asynchronous-message', { command: 'select_balance_sheet', data: {'file_path': file_path }});
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
            assets_name: assets_nameDropdown.value,
            assets_this: assets_thisDropdown.value,
            assets_previous: assets_previousDropdown.value,
            liabilities_name: liabilities_nameDropdown.value,
            liabilities_this: liabilities_thisDropdown.value,
            liabilities_previous: liabilities_previousDropdown.value,
        };

        ipcRenderer.send('asynchronous-message', { command: 'import_balance_sheet', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'select_balance_sheet') {
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

        if (result[0] === 'import_balance_sheet') {
            
            alert('导入数据成功！');
            result_outputTextarea.value = JSON.stringify(result[1], null, 2);       // 2 是缩进空格数
        }
    });
}
// regex.js

const { ipcRenderer } = require('electron');

let filePath = '';              // 用來存儲文件路徑
let savePath = '';              // 用來存儲文件路徑

export async function regexFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1>使用正则表达式筛选</h1>
        
        <div class="import">
            <div>
                <label style="width: 25%;">源表格</label>
                <input id="source_path" type="text">
            </div>

            <div>
                <label for="sheetDropdown" style="width: 25%;">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown"></select>
            </div>

            <div>
                <label for="columnDropdown" style="width: 25%;">选择要筛选的列</label>
                <select id="columnDropdown" name="columnDropdown"></select>
            </div>

            <div>
                <label style="width: 25%;">正则表达式（Python环境）</label>
                <input id="regex_text" type="text">
            </div>

            <div>
                <button id="importButton">导入</button>
                <button id="generateButton">生成</button>
            </div>
        </div>
    `;

    var path_input = document.getElementById('source_path');
    path_input.classList.add('readonly');
    path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const columnDropdown = document.getElementById('columnDropdown');
    const regex_input = document.getElementById('regex_text');

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        filePath = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!filePath) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'regex_import', data: {'filePath': filePath }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'regex_index', data: {'filePath': filePath, 'sheetName': selectedSheet }});
    });

    // 自动触发工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('generateButton').addEventListener('click', async () => {
        if (!filePath) {
            alert('请先导入文件！');
            return;
        }
        // 動態設置過濾器和默認文件名
        const saveFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        const defaultFileName = 'xlsx_output.xlsx';
        savePath = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);

        const data = {
            sheet_value: sheetDropdown.value,
            column_value: columnDropdown.value,
            regex_pattern: regex_input.value,
            filePath: filePath,
            savePath: savePath
        };
        ipcRenderer.send('asynchronous-message', { command: 'regex_generate', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'regex_import') {
            // 将文件路径显示在输入框中
            path_input.value = filePath;

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

        if (result[0] === 'regex_index') {
    
            // 清空旧的选项
            columnDropdown.innerHTML = '';
    
            result[1].forEach(column => {    
                const optionColumn = document.createElement('option');
                optionColumn.value = column;
                optionColumn.text = column;
                columnDropdown.appendChild(optionColumn);
            });
        }

        if (result[0] === 'regex_generate') {
            alert('生成成功！');
        }
    });
}
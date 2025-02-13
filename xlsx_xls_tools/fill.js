// fill.js

const { ipcRenderer } = require('electron');

let filePath = '';              // 用來存儲文件路徑
let savePath = '';              // 用來存儲文件路徑

export async function fillFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1>自动填充空行</h1>
        
        <div class="import">
            <div>
                <label>源表格</label>
                <input id="source_path" type="text">
            </div>

            <div>
                <label for="sheetDropdown">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="columnDropdown">选择要填充的列</label>
                <select id="columnDropdown" name="columnDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label>选择填充方式</label>
                <input type="radio" id="repeatOption" name="action" value="repeat" checked>
                <label for="repeatOption">重复上一行</label>

                <input type="radio" id="zeroOption" name="action" value="zero">
                <label for="zeroOption">填充：0</label>
            </div>

            <div>
                <button id="importButton">导入</button>
                <button id="generateButton">生成</button>
            </div>
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
        ipcRenderer.send('asynchronous-message', { command: 'fill_import', data: {'filePath': filePath }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'fill_index', data: {'filePath': filePath, 'sheetName': selectedSheet }});
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

        // 获取所有 name 为 "action" 的单选按钮
        const radios = document.getElementsByName('action');
        let selectedValue = '';
        
        // 遍历所有单选按钮，找到被选中的那个
        for (const radio of radios) {
            if (radio.checked) {
                selectedValue = radio.value;
                break;
            }
        }

        const data = {
            sheet_value: document.getElementById('sheetDropdown').value,
            column_value: document.getElementById('columnDropdown').value,
            select_value: selectedValue,
            filePath: filePath,
            savePath: savePath
        };
        ipcRenderer.send('asynchronous-message', { command: 'fill_generate', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'fill_import') {
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

        if (result[0] === 'fill_index') {
            const columnDropdown = document.getElementById('columnDropdown');
    
            // 清空旧的选项
            columnDropdown.innerHTML = '';
    
            result[1].forEach(column => {    
                const optionColumn = document.createElement('option');
                optionColumn.value = column;
                optionColumn.text = column;
                columnDropdown.appendChild(optionColumn);
            });
        }

        if (result[0] === 'fill_generate') {
            alert('生成成功！');
        }
    });
}
// xlsx_comparison.js

const { ipcRenderer } = require('electron');

let xlsx_path_1 = '';
let xlsx_path_2 = '';

export async function xlsx_comparison() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>Excel表格对比</h1>
        
        <div class="import">

            <div>
                <label>表格文件1</label>
                <input id="xlsx_path_1" type="text">
            </div>

            <div>
                <label for="sheet1Dropdown">选择工作表</label>
                <select id="sheet1Dropdown" name="sheet1Dropdown"></select>
            </div>

            <div>
                <label>表格文件2</label>
                <input id="xlsx_path_2" type="text">
            </div>

            <div>
                <label for="sheet2Dropdown">选择工作表</label>
                <select id="sheet2Dropdown" name="sheet2Dropdown"></select>
            </div>

            <div>
                <button id="select1Button">选择文件1</button>
                <button id="select2Button">选择文件2</button>
                <button id="comparisonButton">开始对比</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>对比结果：</label>
            </div>
            <div>
                <textarea id="result_output" rows="23"></textarea>
            </div>
        </div>
    `;

    var xlsx_path_1Input = document.getElementById('xlsx_path_1');
    xlsx_path_1Input.classList.add('readonly');
    xlsx_path_1Input.readOnly = true;

    var xlsx_path_2Input = document.getElementById('xlsx_path_2');
    xlsx_path_2Input.classList.add('readonly');
    xlsx_path_2Input.readOnly = true;

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    const sheet1Dropdown = document.getElementById('sheet1Dropdown');
    const sheet2Dropdown = document.getElementById('sheet2Dropdown');

    document.getElementById('select1Button').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        xlsx_path_1 = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!xlsx_path_1) {
            console.log('File selection was canceled.');
            return;
        }

        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'xlsx_comparision_sheetnames', data: {'file': 1, 'file_path': xlsx_path_1 }});
    });

    document.getElementById('select2Button').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        xlsx_path_2 = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!xlsx_path_2) {
            console.log('File selection was canceled.');
            return;
        }

        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'xlsx_comparision_sheetnames', data: {'file': 2, 'file_path': xlsx_path_2 }});
    });

    // 选择文件夹按钮js代码
    document.getElementById('comparisonButton').addEventListener('click', async () => {

        if (!xlsx_path_1 || !xlsx_path_2) {
            alert('请先选择文件！');
            return;
        }

        const data = {
            'xlsx_path_1': xlsx_path_1Input.value,
            'sheet_name_1': sheet1Dropdown.value,
            'xlsx_path_2': xlsx_path_2Input.value,
            'sheet_name_2': sheet2Dropdown.value,
        }

        ipcRenderer.send('asynchronous-message', { command: 'xlsx_comparison', data: data});
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'xlsx_comparision_sheetnames_1') {
            
            // 将文件路径显示在输入框中
            xlsx_path_1Input.value = xlsx_path_1;

            // 清空旧的选项
            sheet1Dropdown.innerHTML = '';

            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheet1Dropdown.appendChild(option);
            });
        }

        if (result[0] === 'xlsx_comparision_sheetnames_2') {
            
            // 将文件路径显示在输入框中
            xlsx_path_2Input.value = xlsx_path_2;

            // 清空旧的选项
            sheet2Dropdown.innerHTML = '';

            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheet2Dropdown.appendChild(option);
            });
        }

        if (result[0] === 'xlsx_comparison') {

            document.getElementById(`result_output`).value = result[1]['result_message'];

        };
    });
}
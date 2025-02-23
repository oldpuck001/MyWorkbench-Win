// export.js

const { ipcRenderer } = require('electron');

let folderPath = '';                // 用來存儲文件路徑
let savePath = '';                  // 用來存儲文件路徑

export async function exportFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>输出文件名汇总</h1>

        <div class="import">
            <div>
                <label style="width: 25%;">文件夹路径</label>
                <input id="folder_path" type="text">
            </div>

            <div>
                <label style="width: 25%;">请选择是否检索子文件夹</label>
                <input type="radio" id="yes" name="action" value="yes" checked>
                <label for="yes">是</label>
                <input type="radio" id="no" name="action" value="copy">
                <label for="no">否</label>
            </div>

            <div>
                <button id="selectButton">选择文件夹</button>
                <button id="outputButton">汇总与输出</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>查找/修改结果：</label>
            </div>
            <div>
                <textarea id="result_output" rows="20"></textarea>
            </div>
        </div>
    `;

    var input = document.getElementById('folder_path');
    input.classList.add('readonly');
    input.readOnly = true;

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        folderPath = await ipcRenderer.invoke('dialog:openDirectory');

        if (!folderPath) {
            console.log('Folder selection was canceled.');
            return;
        }

        document.getElementById(`folder_path`).value = folderPath;
    });

    document.getElementById('outputButton').addEventListener('click', async () => {
        if (!folderPath) {
            alert('请先选择文件夹！');
            return;
        }
        // 動態設置過濾器和默認文件名，保存 xlsx 文件
        const saveFilters = [{ name: 'Excel Files', extensions: ['xlsx'] }];
        const defaultFileName = 'xlsx_output.xlsx';
        savePath = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);

        if (!savePath) {
            return;
        }

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
            folderPath: folderPath,
            savePath: savePath,
            yes_or_no: selectedValue
        };
        ipcRenderer.send('asynchronous-message', { command: 'filename_export', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_export') {
            document.getElementById(`result_output`).value += '导出成功！\n';
        }
    });
}
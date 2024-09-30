// export.js

const { ipcRenderer } = require('electron');

let folderPath = '';  // 用來存儲文件路徑
let savePath = '';  // 用來存儲文件路徑

export async function exportFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.style.fontFamily = 'Arial, sans-serif';
    contentDiv.style.margin = '20px';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1 style="text-align: center; width: 100%;">输出文件名汇总</h1>

        <div id="mainLayout" style="display: flex; flex-direction: column; align-items: center;">
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 90px; text-align: left;">文件夹路径</label>
                <input id="folder_path" type="text" style="width: 500px;">
            </div>
        </div>
        <br>
        <div class="export" style="text-align: center;">
            <button id="selectButton" style="width: 180px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 10px;">选择文件夹</button>
            <button id="outputButton" style="width: 180px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 10px;">汇总与输出</button>
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

        const data = {
            folderPath: folderPath,
            savePath: savePath
        };
        ipcRenderer.send('asynchronous-message', { command: 'filename_export', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_export') {
            alert('导出成功！');
        }
    });
}
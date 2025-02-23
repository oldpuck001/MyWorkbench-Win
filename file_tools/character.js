// character.js

const { ipcRenderer } = require('electron');

let folderPath = '';                // 用來存儲文件路徑

export async function characterFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>按前n个字符分类修改文件名</h1>

        <div class="import">
            <div>
                <label>按前n个字符分类修改文件名，请输入n的值</label>
                <input id="location_character" type="text">
            </div>

            <div>
                <label>文件夹路径</label>
                <input id="folder_path" type="text">
            </div>

            <div>
                <button id="selectButton">选择文件夹</button>
                <button id="modifyButton">修改文件名</button>
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

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        folderPath = await ipcRenderer.invoke('dialog:openDirectory');

        if (!folderPath) {
            console.log('Folder selection was canceled.');
            return;
        }

        document.getElementById(`folder_path`).value = folderPath;
    });

    document.getElementById('modifyButton').addEventListener('click', async () => {
        if (!folderPath) {
            alert('请先选择文件夹！');
            return;
        }

        const data = {
            location_character: document.getElementById('location_character').value,
            folderPath: folderPath
        };

        ipcRenderer.send('asynchronous-message', { command: 'filename_character', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_character') {
            document.getElementById(`result_output`).value = '';
            document.getElementById(`result_output`).value = result[1]
        }
    });
}
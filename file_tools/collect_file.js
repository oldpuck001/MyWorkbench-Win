// collect_file.js

const { ipcRenderer } = require('electron');

let sourcePath = '';
let targetPath = '';

export async function collect_fileFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>将文件汇集至单一文件夹内</h1>

        <div class="import">
            <div>
                <label style="width: 30%;">选择源文件所在的文件夹路径</label>
                <input id="source_path" type="text">
            </div>

            <div>
                <label style="width: 30%;">选择目标文件所在的文件夹路径</label>
                <input id="target_path" type="text">
            </div>

            <div>
                <button id="sourceButton">选择源文件夹</button>
                <button id="targetButton">选择目标文件夹</button>
                <button id="outputButton">汇集文件</button>
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

    var input = document.getElementById('source_path');
    input.classList.add('readonly');
    input.readOnly = true;

    var input = document.getElementById('target_path');
    input.classList.add('readonly');
    input.readOnly = true;

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    document.getElementById('sourceButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        sourcePath = await ipcRenderer.invoke('dialog:openDirectory');

        if (!sourcePath) {
            console.log('Folder selection was canceled.');
            return;
        }

        document.getElementById(`source_path`).value = sourcePath;
    });

    document.getElementById('targetButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        targetPath = await ipcRenderer.invoke('dialog:openDirectory');

        if (!targetPath) {
            console.log('Folder selection was canceled.');
            return;
        }

        document.getElementById(`target_path`).value = targetPath;
    });

    document.getElementById('outputButton').addEventListener('click', async () => {
        if (!sourcePath) {
            alert('请先选择源文件夹！');
            return;
        }

        if (!targetPath) {
            alert('请先选择目标文件夹！');
            return;
        }

        const data = {
            sourcePath: sourcePath,
            targetPath: targetPath,
        };
        ipcRenderer.send('asynchronous-message', { command: 'collect_file', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'collect_file') {
            document.getElementById(`result_output`).value = '';
            document.getElementById(`result_output`).value = result[1]
        }
    });
}
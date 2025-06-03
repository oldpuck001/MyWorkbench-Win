// docx_comparison.js

const { ipcRenderer } = require('electron');

let docx_path_1 = '';
let docx_path_2 = '';

export async function docx_comparison() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>Word文档对比</h1>
        
        <div class="import">

            <div>
                <label>文档文件1</label>
                <input id="docx_path_1" type="text">
            </div>

            <div>
                <label>文档文件2</label>
                <input id="docx_path_2" type="text">
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
                <textarea id="result_output" rows="27"></textarea>
            </div>
        </div>
    `;

    var docx_path_1Input = document.getElementById('docx_path_1');
    docx_path_1Input.classList.add('readonly');
    docx_path_1Input.readOnly = true;

    var docx_path_2Input = document.getElementById('docx_path_2');
    docx_path_2Input.classList.add('readonly');
    docx_path_2Input.readOnly = true;

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    document.getElementById('select1Button').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Word Files', extensions: ['docx'] }];
        docx_path_1 = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!docx_path_1) {
            console.log('File selection was canceled.');
            return;
        }

        docx_path_1Input.value = docx_path_1;
    });

    document.getElementById('select2Button').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Word Files', extensions: ['docx'] }];
        docx_path_2 = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!docx_path_2) {
            console.log('File selection was canceled.');
            return;
        }

        docx_path_2Input.value = docx_path_2;
    });

    // 选择文件夹按钮js代码
    document.getElementById('comparisonButton').addEventListener('click', async () => {

        if (!docx_path_1 || !docx_path_2) {
            alert('请先选择文件！');
            return;
        }

        const data = {
            'docx_path_1': docx_path_1Input.value,
            'docx_path_2': docx_path_2Input.value
        }

        ipcRenderer.send('asynchronous-message', { command: 'docx_comparison', data: data});
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'docx_comparison') {

            document.getElementById(`result_output`).value = result[1]['result_message'];

        };
    });
}
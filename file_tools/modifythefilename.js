// modifythefilename.js

const { ipcRenderer } = require('electron');

let folderPath = '';                // 用來存儲文件路徑

export async function modifythefilenameFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>自动修改文件名</h1>

        <div class="import">
            <div>
                <label style="width: 30%;">指定字符（不检查扩展名）</label>
                <input id="source_character" type="text">
            </div>

            <div>
                <label style="width: 30%;">指定字符的第一个字符的位置（从左至右，第几个字符）</label>
                <input id="location_character" type="text">
            </div>

            <div>
                <label style="width: 30%;">替换成（不替换扩展名）</label>
                <input id="target_character" type="text">
            </div>

            <div>
                <label style="width: 30%;">文件夹路径</label>
                <input id="folder_path" type="text">
            </div>

            <div style="display: flex; justify-content: center;">
                <button id="selectButton">选择文件夹</button>
                <button id="findButton">查找指定字符</button>
                <button id="addButton">添加指定字符</button>
                <button id="delButton">删除指定字符</button>
                <button id="replaceButton">替换指定字符</button>
                <button id="regexButton">使用正则表达式</button>
            </div>

            <div style="display: flex; flex-direction: column; align-items: center;">
                <label style="width: 100%;">&nbsp;&nbsp;说明：</label>
                <label style="width: 100%;">&nbsp;&nbsp;1.使用查找功能时，若指定字符的位置不填写，则检查文件名所有位置（不含扩展名）；</label>
                <label style="width: 100%;">&nbsp;&nbsp;2.使用添加功能时，若需要将指定字符插入原有文件名之前，则输入“0”；</label>
                <label style="width: 100%;">&nbsp;&nbsp;3.使用删除功能时，若指定字符的位置不填写，则对文件名中的所有位置（不含扩展名）进行删除；</label>
                <label style="width: 100%;">&nbsp;&nbsp;4.使用删除功能时，若填写指定字符的位置，则指定字符处需填写要删除的字符数；</label>
                <label style="width: 100%;">&nbsp;&nbsp;5.使用替换功能时，若位置不填写，则对文件名中的所有位置（不含扩展名）进行替换；</label>
                <label style="width: 100%;">&nbsp;&nbsp;6.使用正则表达式功能时，请使用Python语言的正则表达式规范。</label>
            </div>

        </div>

        <div class="export">
            <div>
                <label>查找/修改结果：</label>
            </div>
            <div>
                <textarea id="result_output" rows="10"></textarea>
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

    const handleButtonClick = async (select_function) => {
        if (!folderPath) {
            alert('请先选择文件夹！');
            return;
        }

        const data = {
            select_function: select_function,
            source_character: document.getElementById('source_character').value,
            location_character: document.getElementById('location_character').value,
            target_character: document.getElementById('target_character').value,
            folderPath: folderPath
        };

        ipcRenderer.send('asynchronous-message', { command: 'filename_modify', data: data });
    };

    document.getElementById('findButton').addEventListener('click', () => handleButtonClick('find'));
    document.getElementById('addButton').addEventListener('click', () => handleButtonClick('add'));
    document.getElementById('delButton').addEventListener('click', () => handleButtonClick('del'));
    document.getElementById('replaceButton').addEventListener('click', () => handleButtonClick('replace'));
    document.getElementById('regexButton').addEventListener('click', () => handleButtonClick('regex'));

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_modify') {
            document.getElementById(`result_output`).value = '';
            document.getElementById(`result_output`).value = result[1]
        }
    });
}
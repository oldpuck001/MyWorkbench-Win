// modifythefilename.js

const { ipcRenderer } = require('electron');

let folderPath = '';  // 用來存儲文件路徑

export async function modifythefilenameFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.style.fontFamily = 'Arial, sans-serif';
    contentDiv.style.margin = '20px';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1 style="text-align: center; width: 100%;">自动修改文件名</h1>

        <div id="mainLayout" style="display: flex; flex-direction: column; align-items: center;">
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 198px; text-align: left;">指定字符（不检查扩展名）</label>
                <input id="source_character" type="text" style="width: 642px;">
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 414px; text-align: left;">指定字符的第一个字符的位置（从左至右，第几个字符）</label>
                <input id="location_character" type="text" style="width: 426px;">
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 180px; text-align: left;">替换成（不替换扩展名）</label>
                <input id="target_character" type="text" style="width: 660px;">
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 90px; text-align: left;">文件夹路径</label>
                <input id="folder_path" type="text" style="width: 750px;">
            </div>
        </div>
        <br>
        <div class="export" style="text-align: center;">
            <button id="selectButton" style="width: 130px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 10px;">选择文件夹</button>
            <button id="findButton" style="width: 130px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 10px;">查找指定字符</button>
            <button id="addButton" style="width: 130px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 5px;">添加指定字符</button>
            <button id="delButton" style="width: 130px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 5px;">删除指定字符</button>
            <button id="replaceButton" style="width: 130px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 10px;">替换指定字符</button>
            <button id="regexButton" style="width: 130px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 10px;">使用正则表达式</button>
        </div>
        <br>
        <div style="display: flex; flex-direction: column; align-items: center;">
            <label style="width: 850px; text-align: left;">说明：</label>
            <label style="width: 850px; text-align: left;">1.使用查找功能时，若指定字符的位置不填写，则检查文件名所有位置（不含扩展名）；</label>
            <label style="width: 850px; text-align: left;">2.使用添加功能时，若需要将指定字符插入原有文件名之前，则输入“0”；</label>
            <label style="width: 850px; text-align: left;">3.使用删除功能时，若指定字符的位置不填写，则对文件名中的所有位置（不含扩展名）进行删除；</label>
            <label style="width: 850px; text-align: left;">4.使用删除功能时，若填写指定字符的位置，则指定字符处需填写要删除的字符数；</label>
            <label style="width: 850px; text-align: left;">5.使用替换功能时，若位置不填写，则对文件名中的所有位置（不含扩展名）进行替换；</label>
            <label style="width: 850px; text-align: left;">6.使用正则表达式功能时，请使用Python语言的正则表达式规范。</label>
        </div>
        <br>
        <div style="display: flex; flex-direction: column; align-items: center;">
            <label style="width: 850px; text-align: left;">查找/修改结果：</label>
            <textarea id="result_output" rows="12" style="width: 850px;"></textarea>
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

    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_modify') {
            document.getElementById(`result_output`).value = '';
            document.getElementById(`result_output`).value = result[1]
        }
    });
}
// character.js

const { ipcRenderer } = require('electron');

let folderPath = '';  // 用來存儲文件路徑

export async function characterFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.style.fontFamily = 'Arial, sans-serif';
    contentDiv.style.margin = '20px';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1 style="text-align: center; width: 100%;">按前n个字符分类修改文件名</h1>

        <div id="mainLayout" style="display: flex; flex-direction: column; align-items: center;">
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 315px; text-align: left;">按前n个字符分类修改文件名，请输入n的值</label>
                <input id="location_character" type="text" style="width: 275px;">
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 90px; text-align: left;">文件夹路径</label>
                <input id="folder_path" type="text" style="width: 500px;">
            </div>
        </div>
        <br>
        <div class="export" style="text-align: center;">
            <button id="selectButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 10px;">选择文件夹</button>
            <button id="modifyButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 10px;">修改文件名</button>
        </div>
        <br>
        <div style="display: flex; flex-direction: column; align-items: center;">
            <label style="width: 590px; text-align: left;">查找/修改结果：</label>
            <textarea id="result_output" rows="12" style="width: 590px;"></textarea>
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
    
        ipcRenderer.on('asynchronous-reply', (event, result) => {
            if (result[0] === 'filename_character') {
                document.getElementById(`result_output`).value = '';
                document.getElementById(`result_output`).value = result[1]
            }
        });
}
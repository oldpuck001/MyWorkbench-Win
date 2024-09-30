// sort.js

const { ipcRenderer } = require('electron');

let sourcePath = '';
let targetPath = '';

export async function sortFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.style.fontFamily = 'Arial, sans-serif';
    contentDiv.style.margin = '20px';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1 style="text-align: center; width: 100%;">按文件名分类</h1>

        <div id="mainLayout" style="display: flex; flex-direction: column; align-items: center;">
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 410px; text-align: left;">按前n个字符将文件剪切或复制至新文件夹，请输入n的值</label>
                <input id="location_character" type="text" style="width: 180px;">
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label for="cutOption" style="width: 160px; text-align: left; margin-right: 10px;">请选择剪切或者复制</label>
                <input type="radio" id="cutOption" name="action" value="cut" style="margin-right: 5px;" checked>
                <label for="cutOption" style="margin-right: 20px;">剪切</label>
                <input type="radio" id="copyOption" name="action" value="copy" style="margin-right: 5px;">
                <label for="copyOption" style="margin-right: 300px;">複製</label>
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 218px; text-align: left;">选择源文件所在的文件夹路径</label>
                <input id="source_path" type="text" style="width: 372px;">
            </div>
            <div style="display: flex; align-items: center; margin-bottom: 10px;">
                <label style="width: 233px; text-align: left;">选择目标文件所在的文件夹路径</label>
                <input id="target_path" type="text" style="width: 357px;">
            </div>
        </div>
        <br>
        <div class="export" style="text-align: center;">
            <button id="sourceButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 15px;">选择源文件夹</button>
            <button id="targetButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer;">选择目标文件夹</button>
            <button id="sortButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 15px;">分类文件</button>
        </div>
        <br>
        <div style="display: flex; flex-direction: column; align-items: center;">
            <label style="width: 590px; text-align: left;">查找/修改结果：</label>
            <textarea id="result_output" rows="12" style="width: 590px;"></textarea>
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

    document.getElementById('sortButton').addEventListener('click', async () => {
        if (!sourcePath) {
            alert('请先选择源文件夹！');
            return;
        }

        if (!targetPath) {
            alert('请先选择目标文件夹！');
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
            location_character: document.getElementById('location_character').value,
            cut_or_copy: selectedValue,
            sourcePath: sourcePath,
            targetPath: targetPath
        };

        ipcRenderer.send('asynchronous-message', { command: 'filename_sort', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_sort') {
            document.getElementById(`result_output`).value = '';
            document.getElementById(`result_output`).value = result[1]
        }
    });
}
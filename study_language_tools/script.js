// script.js

export async function vocabulary_to_audio() {

    let filePath = '';  // 用來存儲文件路徑
    let savePath = '';  // 用來存儲文件路徑

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1 style="text-align: center; width: 100%;">单词表转音频</h1>
        
        <div style="display: flex; flex-direction: column; align-items: center;">

            <div class="form-group">
                <label for="speechKey" style="display: inline-block; width: 140px;">SPEECH KEY</label>
                <input type="text" id="speechKey" name="speechKey" style="width: 500px; height: 30px; border: none; background-color: #F0F4F8; outline: none;">
            </div>
            <br>
            <div class="form-group">
                <label for="speechRegion" style="display: inline-block; width: 140px;">SPEECH REGION</label>
                <input type="text" id="speechRegion" name="speechRegion" style="width: 500px; height: 30px; border: none; background-color: #F0F4F8; outline: none;">
            </div>
            <br>
            <div class="form-group">
                <input type="checkbox" id="maleVoice" name="maleVoice" checked>
                <label for="maleVoice" style="margin-right: 5px;">男生朗读</label>
                <select id="maleDropdown" name="maleDropdown" style="margin-right: 20px;">
                    <option value="1">1遍</option>
                    <option value="2" selected>2遍</option>
                    <option value="3">3遍</option>
                </select>
                <input type="checkbox" id="maleTranslate" name="maleTranslate">
                <label for="maleTranslate" style="margin-right: 10px;">男生朗读译文</label>
                <input type="checkbox" id="maleExample" name="maleExample">
                <label for="maleExample">男生朗读例句</label>
            </div>
            <br>
            <div class="form-group">
                <input type="checkbox" id="femaleVoice" name="femaleVoice">
                <label for="femaleVoice" style="margin-right: 5px;">女生朗读</label>
                <select id="femaleDropdown" name="femaleDropdown" style="margin-right: 20px;" disabled>
                    <option value="1">1遍</option>
                    <option value="2" selected>2遍</option>
                    <option value="3">3遍</option>
                </select>
                <input type="checkbox" id="femaleTranslate" name="femaleTranslate">
                <label for="femaleTranslate" style="margin-right: 10px;">女生朗读译文</label>
                <input type="checkbox" id="femaleExample" name="femaleExample">
                <label for="femaleExample">女生朗读例句</label>
            </div>
            <br>
            <div class="form-group">
                <label for="languageDropdown" style="margin-right: 5px;">语种</label>
                <select id="languageDropdown" name="languageDropdown" style="width: 50px; margin-right: 20px;">
                    <option value="1" selected>英语</option>
                    <option value="2">日语</option>
                    <option value="3">泰语</option>
                </select>
                <label for="sheetDropdown" style="margin-right: 5px;">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown" style="width: 70px; margin-right: 20px;">
                </select>
                <input type="checkbox" id="singleAudio" name="singleAudio">
                <label for="singleAudio">生成单字音频</label>
            </div>
            <br>
            <div class="form-buttons">
                <button id="importButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-right: 10px;">导入</button>
                <button id="generateButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer; margin-left: 10px;">生成</button>
            </div>
        </div>
    `;

    // 初始化复选框状态
    toggleDropdown('maleDropdown', 'maleVoice');
    toggleDropdown('femaleDropdown', 'femaleVoice');

    // 绑定事件监听器到男生和女生朗读的复选框
    document.getElementById('maleVoice').addEventListener('change', function() {
        toggleDropdown('maleDropdown', 'maleVoice');
    });

    document.getElementById('femaleVoice').addEventListener('change', function() {
        toggleDropdown('femaleDropdown', 'femaleVoice');
    });
    
    const { ipcRenderer } = require('electron');

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        filePath = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!filePath) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'importButton', data: {'filePath': filePath }});
    });

    document.getElementById('generateButton').addEventListener('click', async () => {
        if (!filePath) {
            alert('請先導入文件');
            return;
        }
        // 動態設置過濾器和默認文件名，保存 MP3 文件
        const saveFilters = [{ name: 'MP3 Files', extensions: ['mp3'] }];
        const defaultFileName = 'audio_output.mp3';
        savePath = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);

        const data = {
            speechKey: document.getElementById('speechKey').value,
            speechRegion: document.getElementById('speechRegion').value,
            maleVoice: document.getElementById('maleVoice').checked,
            maleRepeats: document.getElementById('maleDropdown').value,
            maleTranslate: document.getElementById('maleTranslate').checked,
            maleExample: document.getElementById('maleExample').checked,
            femaleVoice: document.getElementById('femaleVoice').checked,
            femaleRepeats: document.getElementById('femaleDropdown').value,
            femaleTranslate: document.getElementById('femaleTranslate').checked,
            femaleExample: document.getElementById('femaleExample').checked,
            language: document.getElementById('languageDropdown').value,
            sheet: document.getElementById('sheetDropdown').value,
            singleAudio: document.getElementById('singleAudio').checked,
            filePath: filePath,
            savePath: savePath
        };
        ipcRenderer.send('asynchronous-message', { command: 'generateButton', data: data });
    });
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'import') {
            // 获取 select 元素
            const sheetDropdown = document.getElementById('sheetDropdown');
            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheetDropdown.appendChild(option);
            });
            alert('导入成功！');
        }

        if (result[0] === 'generate') {
            alert('生成成功！');
        }
    });

    // 切换下拉菜单的启用/禁用状态的内部函数
    function toggleDropdown(dropdownId, checkboxId) {
        const dropdown = document.getElementById(dropdownId);
        const checkbox = document.getElementById(checkboxId);
        
        if (checkbox.checked) {
            dropdown.disabled = false; // 启用下拉框
        } else {
            dropdown.disabled = true; // 禁用下拉框
        }
    }
}
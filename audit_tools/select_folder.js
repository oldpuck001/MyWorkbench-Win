// select_folder.js

const { ipcRenderer } = require('electron');

export async function select_folder() {
    
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>选择项目文件夹</h1>

        <div class="import">
            <div>
                <label>文件夹路径</label>
                <input id="folder_path" type="text">
            </div>

            <div>
                <button id="selectButton">选择文件夹</button>
            </div>
        </div>
    `;

    const folder_input = document.getElementById('folder_path');
    folder_input.classList.add('readonly');
    folder_input.readOnly = true;

    window.project_folder = await ipcRenderer.invoke('get-project-folder');
    folder_input.value = window.project_folder || '';

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        window.project_folder = await ipcRenderer.invoke('dialog:openDirectory');

        if (!window.project_folder) {
            console.log('Folder selection was canceled.');
            return;
        }

        folder_input.value = window.project_folder;

        const data = {
            project_folder: window.project_folder
        };

        ipcRenderer.send('asynchronous-message', { command: 'select_folder_path', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'select_folder_path') {
            alert(result[1]['result_message'])
        }
    });
}
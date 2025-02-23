// sql_sqlite.js

const { ipcRenderer } = require('electron');

let folder_path = '';
let save_path = '';

export async function sql_sqlite_Function() {

    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>SQLite数据库</h1>

        <div class="import">

            <div>
                <label style="width: 10%;">文件夹路径</label>
                <input id="folder_path" type="text">
            </div>

            <div>
                <label style="width: 10%;">数据库指令</label>
                <textarea id="sql_command" rows="10"></textarea>
            </div>

            <div>
                <button id="folderButton">选择文件夹</button>
                <button id="sqlButton">操作数据库</button>
                <button id="backupButton">建立还原点</button>
                <button id="selectButton">查询并导出</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>操作反馈：</label>
            </div>
            <div>
                <textarea id="result_output" rows="20"></textarea>
            </div>
        </div>
        `;

    var folder_path_input = document.getElementById('folder_path');
    folder_path_input.classList.add('readonly');
    folder_path_input.readOnly = true;

    const sql_command = document.getElementById('sql_command');

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    // 选择文件夹按钮js代码
    document.getElementById('folderButton').addEventListener('click', async () => {

        folder_path = await ipcRenderer.invoke('dialog:openDirectory');
        if (!folder_path) { return };

        ipcRenderer.send('asynchronous-message', { command: 'sql_sqlite_folder', data: {'folder_path': folder_path}});
    });

    // 操作数据库按钮js代码
    document.getElementById('sqlButton').addEventListener('click', async () => {

        folder_path = folder_path_input.value;
        if (!folder_path) { return };

        const data = {
            folder_path: folder_path,
            sql_command: sql_command.value
        };

        ipcRenderer.send('asynchronous-message', { command: 'sql_sqlite_sql', data: data});
    });

    // 建立还原点按钮js代码
    document.getElementById('backupButton').addEventListener('click', async () => {

        folder_path = folder_path_input.value;
        if (!folder_path) { return };

        ipcRenderer.send('asynchronous-message', { command: 'sql_sqlite_backup', data: {'folder_path': folder_path}});
    });

    // 查询并导出按钮js代码
    document.getElementById('selectButton').addEventListener('click', async () => {

        folder_path = folder_path_input.value;
        if (!folder_path) { return };

        const saveFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        const defaultFileName = 'xlsx_output.xlsx';
        save_path = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);
        if (!save_path) { return };

        const data = {
            folder_path: folder_path,
            save_path: save_path,
            sql_command: sql_command.value
        };

        ipcRenderer.send('asynchronous-message', { command: 'sql_sqlite_select', data: data});
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'sql_sqlite_folder') {

            folder_path_input.value = folder_path;
            document.getElementById(`result_output`).value += result[1];

        } else if (result[0] === 'sql_sqlite_sql') {

            document.getElementById(`result_output`).value += result[1];

        } else if (result[0] === 'sql_sqlite_backup') {

            document.getElementById(`result_output`).value += result[1];

        } else if (result[0] === 'sql_sqlite_select') {

            document.getElementById(`result_output`).value += result[1];
        
        };
    });
}
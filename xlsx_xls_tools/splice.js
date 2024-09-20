// script.js

const { ipcRenderer } = require('electron');

export async function spliceFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');
    contentDiv.style.border = 'none';
    contentDiv.style.fontFamily = 'Arial, sans-serif';
    contentDiv.style.margin = '20px';

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <style>
            .row {
                display: flex;
                margin-bottom: 10px;
            }
            .row input, .row select {
                margin-right: 10px;
            }
            .row button {
                margin-left: 10px;
            }
            .readonly {
                cursor: not-allowed;
            }
            .export {
                margin-top: 20px;
            }
        </style>

        <h1 style="text-align: center; width: 100%;">电子表格拼接</h1>
        
        <div id="mainLayout" style="display: flex; flex-direction: column; align-items: center;">
            <!-- Rows will be added dynamically via JavaScript -->
        </div>
        
        <div class="export" style="text-align: center;">
            <button id="exportButton" style="width: 150px; background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer;">导出</button>
        </div>
    `;

    // 创建12行
    const mainLayout = document.getElementById('mainLayout');
    const numRows = 12;

    for (let i = 0; i < numRows; i++) {
        createRow(i);
    }

    function createRow(index) {
        const rowDiv = document.createElement('div');
        rowDiv.classList.add('row');

        const label = document.createElement('label');
        label.textContent = `表格${index + 1}`;
        label.style.width = '60px';  // 设置宽度以保持布局一致
        label.style.textAlign = 'left';  // 右对齐文本

        const input = document.createElement('input');
        input.type = 'text';
        input.style.width = '350px';
        input.classList.add('readonly');
        input.readOnly = true;
        input.id = `sheetInput-${index}`;

        const select = document.createElement('select');
        select.style.width = '120px';
        select.id = `sheetDropdown-${index}`;  // 为每个 select 分配唯一的 ID

        const addButton = document.createElement('button');
        addButton.textContent = '添加';
        addButton.addEventListener('click', () => addSheet(index));

        const delButton = document.createElement('button');
        delButton.textContent = '删除';
        delButton.addEventListener('click', () => delSheet(index));

        rowDiv.appendChild(label);
        rowDiv.appendChild(input);
        rowDiv.appendChild(select);
        rowDiv.appendChild(addButton);
        rowDiv.appendChild(delButton);
        mainLayout.appendChild(rowDiv);
    }

    async function addSheet(index) {
        // 定义文件筛选器
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
    
        // 使用 ipcRenderer.invoke 与主进程通信，打开文件对话框
        const filePath = await ipcRenderer.invoke('dialog:openFile', { filters: fileFilters });

        if (!filePath) {
            console.log('File selection was canceled.');
            return;
        }

        // 将文件路径显示在输入框中
        document.getElementById(`sheetInput-${index}`).value = filePath;

        // 将文件路径发送到主进程
        ipcRenderer.send('asynchronous-message', { command: 'splice_sheet_input', data: { filePath: filePath }, num: index });
    }

    function delSheet(index) {
        document.getElementById(`sheetInput-${index}`).value = '';
        document.getElementById(`sheetDropdown-${index}`).innerHTML = '';
    }

    document.getElementById('exportButton').addEventListener('click', async () => {
        
        const data = {};
        const numRows = 12;  // 假设总共有12行
        let allEmpty = true; // 用于检查是否所有输入都是空
        
        for (let i = 0; i <= numRows; i++) {
            const fileInput = document.getElementById(`sheetInput-${i}`);
            const sheetDropdown = document.getElementById(`sheetDropdown-${i}`);
        
            if (fileInput && sheetDropdown) {
                const fileValue = fileInput.value.trim();
                const sheetValue = sheetDropdown.value.trim();

                data[`file${i}`] = fileValue;
                data[`sheet${i}`] = sheetValue;
        
                // 检查是否有任意一个 file 或 sheet 不为空
                if (fileValue !== '' || sheetValue !== '') {
                    allEmpty = false;
                }
            } else {
                console.error(`Element sheetInput-${i} or sheetDropdown-${i} not found.`);
            }
        }
        
        // 如果所有输入都是空的，返回或者执行某个逻辑
        if (allEmpty) {
            console.log('所有输入均为空');
            return;
        }
        
        // 動態設置過濾器和默認文件名，保存 xlsx 文件
        const saveFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        const defaultFileName = 'xlsx_output.xlsx';
        const savePath = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);

        if (!savePath) {
            console.log('保存操作取消');
            return;
        }

        // 如果你还有 savePath 变量
        data['savePath'] = savePath;
        
        ipcRenderer.send('asynchronous-message', { command: 'splice_sheet_output', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'input') {
            
            const sheetIndex = result[2];

            // 确保获取到正确的 sheetDropdown 元素
            const sheetDropdown = document.getElementById(`sheetDropdown-${sheetIndex}`);

            // 先清空已有的选项
            sheetDropdown.innerHTML = ''; 

            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheetDropdown.appendChild(option);
            });
        }
    
        if (result[0] === 'output') {
            alert('拼接成功！');
        }
    });
}
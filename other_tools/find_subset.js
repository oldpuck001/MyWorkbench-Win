// find_subset.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function find_subset() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1>凑数工具</h1>
        
        <div class="import">

            <div>
                <label>源数据文件</label>
                <input id="file_path" type="text">
            </div>

            <div style="width: 100%;">
                <div style="width: 50%;">
                    <label for="sheetDropdown" style="width: 40%;">选择工作表</label>
                    <select id="sheetDropdown" name="sheetDropdown" style="width: 45%;"></select>
                </div>
                <div style="width: 40%;">
                    <label for="target_valueDropdown" style="width: 40%;">选择目标值列</label>
                    <select id="target_valueDropdown" name="target_valueDropdown" style="width: 50%;"></select>
                </div>
            </div>

            <div style="width: 100%;">
                <div style="width: 50%;">
                    <label for="value_nameDropdown" style="width: 40%;">选择数据名称列</label>
                    <select id="value_nameDropdown" name="value_nameDropdown" style="width: 45%;"></select>
                </div>
                <div style="width: 40%;">
                    <label for="value_numDropdown" style="width: 40%;">选择数据数值列</label>
                    <select id="value_numDropdown" name="value_numDropdown" style="width: 50%;"></select>
                </div>
            </div>

            <div></div>

        </div>

        <div id="rowsContainer" class="rows">
            <div>
                <label style="width: 8%; text-align: left;">目标值</label>
                <input id="target_nameInput" type="text" style="width: 18%; margin-right: 3%">
                <input id="target_valueInput" type="text" style="width: 18%; text-align: right; margin-right: 6%">

                <label style="width: 8%; text-align: left;">数值1</label>
                <input id="name_numInput-1" type="text" style="width: 18%; margin-right: 3%">
                <input id="value_numInput-1" type="text" style="width: 18%; text-align: right;">
            </div>

            <!-- Rows will be added dynamically via JavaScript -->

        </div>

        <div class="import">
            <div>
                <button id="selectButton">选择文件</button>
                <button id="importButton">导入数据</button>
                <button id="exportButton">导出结果</button>
            </div>
        </div>
    `;

    // 创建12行
    const rowsContainer = document.getElementById('rowsContainer');
    const numRows = 13;

    for (let i = 1; i < numRows; i++) {
        createRow(i);
    }

    function createRow(index) {
        const rowDiv = document.createElement('div');
        rowDiv.classList.add('row');

        const label_left = document.createElement('label');
        label_left.textContent = `数值${index * 2}`;
        label_left.style.width = '8%';
        label_left.style.textAlign = 'left';

        const input_left_name = document.createElement('input');
        input_left_name.type = 'text';
        input_left_name.style.width = '18%';
        input_left_name.style.marginRight = '3%';
        input_left_name.id = `name_numInput-${index * 2}`;

        const input_left_value = document.createElement('input');
        input_left_value.type = 'text';
        input_left_value.style.width = '18%';
        input_left_value.style.marginRight = '6%';
        input_left_value.style.textAlign = 'right';
        input_left_value.id = `value_numInput-${index * 2}`;

        const label_right = document.createElement('label');
        label_right.textContent = `数值${index * 2 + 1}`;
        label_right.style.width = '8%';
        label_right.style.textAlign = 'left';

        const input_right_name = document.createElement('input');
        input_right_name.type = 'text';
        input_right_name.style.width = '18%';
        input_right_name.style.marginRight = '3%';
        input_right_name.id = `name_numInput-${index * 2 + 1}`;

        const input_right_value = document.createElement('input');
        input_right_value.type = 'text';
        input_right_value.style.width = '18%';
        input_right_value.style.textAlign = 'right';
        input_right_value.id = `value_numInput-${index * 2 + 1}`;

        rowDiv.appendChild(label_left);
        rowDiv.appendChild(input_left_name);
        rowDiv.appendChild(input_left_value);
        rowDiv.appendChild(label_right);
        rowDiv.appendChild(input_right_name);
        rowDiv.appendChild(input_right_value);
        rowsContainer.appendChild(rowDiv);
    }

    var file_path_input = document.getElementById('file_path');
    file_path_input.classList.add('readonly');
    file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const target_valueDropdown = document.getElementById('target_valueDropdown');
    const value_nameDropdown = document.getElementById('value_nameDropdown');
    const value_numDropdown = document.getElementById('value_numDropdown');

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'find_subset_sheetnames_import', data: {'file_path': file_path }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'find_subset_columns_index', data: {'file_path': file_path, 'sheet_name': selectedSheet }});
    });

    // 自动触发工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('importButton').addEventListener('click', async () => {
        if (!file_path) {
            alert('请先导入文件！');
            return;
        }

        const data = {
            file_path: file_path,
            sheet_name: sheetDropdown.value,
            target_value: target_valueDropdown.value,
            value_name: value_nameDropdown.value,
            value_num: value_numDropdown.value,            
        };

        ipcRenderer.send('asynchronous-message', { command: 'find_subset_import', data: data });
    });

    document.getElementById('exportButton').addEventListener('click', async () => {
        
        const data = {
            target_name: document.getElementById('target_nameInput').value,
            target_value: document.getElementById('target_valueInput').value,

            value_name: [document.getElementById('name_numInput-1').value,
                         document.getElementById('name_numInput-2').value,
                         document.getElementById('name_numInput-3').value,
                         document.getElementById('name_numInput-4').value,
                         document.getElementById('name_numInput-5').value,
                         document.getElementById('name_numInput-6').value,
                         document.getElementById('name_numInput-7').value,
                         document.getElementById('name_numInput-8').value,
                         document.getElementById('name_numInput-9').value,
                         document.getElementById('name_numInput-10').value,
                         document.getElementById('name_numInput-11').value,
                         document.getElementById('name_numInput-12').value,
                         document.getElementById('name_numInput-13').value,
                         document.getElementById('name_numInput-14').value,
                         document.getElementById('name_numInput-15').value,
                         document.getElementById('name_numInput-16').value,
                         document.getElementById('name_numInput-17').value,
                         document.getElementById('name_numInput-18').value,
                         document.getElementById('name_numInput-19').value,
                         document.getElementById('name_numInput-20').value,
                         document.getElementById('name_numInput-21').value,
                         document.getElementById('name_numInput-22').value,
                         document.getElementById('name_numInput-23').value,
                         document.getElementById('name_numInput-24').value,
                         document.getElementById('name_numInput-25').value
            ],

            value_num: [document.getElementById('value_numInput-1').value,
                        document.getElementById('value_numInput-2').value,
                        document.getElementById('value_numInput-3').value,
                        document.getElementById('value_numInput-4').value,
                        document.getElementById('value_numInput-5').value,
                        document.getElementById('value_numInput-6').value,
                        document.getElementById('value_numInput-7').value,
                        document.getElementById('value_numInput-8').value,
                        document.getElementById('value_numInput-9').value,
                        document.getElementById('value_numInput-10').value,
                        document.getElementById('value_numInput-11').value,
                        document.getElementById('value_numInput-12').value,
                        document.getElementById('value_numInput-13').value,
                        document.getElementById('value_numInput-14').value,
                        document.getElementById('value_numInput-15').value,
                        document.getElementById('value_numInput-16').value,
                        document.getElementById('value_numInput-17').value,
                        document.getElementById('value_numInput-18').value,
                        document.getElementById('value_numInput-19').value,
                        document.getElementById('value_numInput-20').value,
                        document.getElementById('value_numInput-21').value,
                        document.getElementById('value_numInput-22').value,
                        document.getElementById('value_numInput-23').value,
                        document.getElementById('value_numInput-24').value,
                        document.getElementById('value_numInput-25').value
            ]
        };
        
        // 動態設置過濾器和默認文件名，保存 xlsx 文件
        const saveFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        const defaultFileName = 'xlsx_output.xlsx';
        const savePath = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);

        if (!savePath) {
            console.log('保存操作取消！');
            return;
        }

        // 如果你还有 savePath 变量
        data['savePath'] = savePath;
        
        ipcRenderer.send('asynchronous-message', { command: 'find_subset_export', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'find_subset_sheetnames_import') {
            // 将文件路径显示在输入框中
            file_path_input.value = file_path;

            // 清空旧的选项
            sheetDropdown.innerHTML = '';

            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheetDropdown.appendChild(option);
            });

            triggerChange(sheetDropdown);       // 触发 change 事件，自动加载列数据

            alert('导入成功！');
        }

        if (result[0] === 'find_subset_columns_index') {
    
            // 清空旧的选项
            target_valueDropdown.innerHTML = '';
    
            result[1].forEach(column => {    
                const optionColumn = document.createElement('option');
                optionColumn.value = column;
                optionColumn.text = column;
                target_valueDropdown.appendChild(optionColumn);
            });

            // 清空旧的选项
            value_nameDropdown.innerHTML = '';
    
            result[1].forEach(column => {    
                const optionColumn = document.createElement('option');
                optionColumn.value = column;
                optionColumn.text = column;
                value_nameDropdown.appendChild(optionColumn);
            });

            // 清空旧的选项
            value_numDropdown.innerHTML = '';
    
            result[1].forEach(column => {    
                const optionColumn = document.createElement('option');
                optionColumn.value = column;
                optionColumn.text = column;
                value_numDropdown.appendChild(optionColumn);
            });
        }

        if (result[0] === 'find_subset_import') {

            document.getElementById('target_nameInput').value = target_valueDropdown.value;
            document.getElementById('target_valueInput').value = Number(result[1]['target_value'][0]).toLocaleString('en-US', {minimumFractionDigits: 2});

            for (let i = 0; i < 25; i++) {
                document.getElementById(`name_numInput-${i + 1}`).value = result[1]['name_value'][i];
                document.getElementById(`value_numInput-${i + 1}`).value = Number(result[1]['num_value'][i]).toLocaleString('en-US', {minimumFractionDigits: 2});
            }

        }
    
        if (result[0] === 'find_subset_export') {
            alert(result[1]['result_message']);
        }
    });
}
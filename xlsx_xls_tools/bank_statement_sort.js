// bank_statement_sort.js

const { ipcRenderer } = require('electron');

let file_path = '';

export async function bank_statement_sortFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>银行对账单分类</h1>
        
        <div class="import">
            <div>
                <label>源表格</label>
                <input id="source_path" type="text">
            </div>

            <div>
                <label for="sheetDropdown">选择工作表</label>
                <select id="sheetDropdown" name="sheetDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="debit_or_creditDropdown">选择借贷标识列</label>
                <select id="debit_or_creditDropdown" name="debit_or_creditDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="creditDropdown">选择贷方/转入标识</label>
                <select id="creditDropdown" name="creditDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="debitDropdown">选择借方/转出标识</label>
                <select id="debitDropdown" name="debitDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="nameDropdown">选择对方户名列</label>
                <select id="nameDropdown" name="nameDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="bankDropdown">选择对方银行列</label>
                <select id="bankDropdown" name="bankDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="numberDropdown">选择对方账号列</label>
                <select id="numberDropdown" name="numberDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="valueDropdown">选择统计数值列</label>
                <select id="valueDropdown" name="valueDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label>转入优先显示项目：</label>
                <textarea id="credit_priority" rows="9"></textarea>
            </div>

            <div>
                <label>转出优先显示项目：</label>
                <textarea id="debit_priority" rows="9"></textarea>
            </div>

            <div>
                <button id="importButton">导入</button>
                <button id="exportButton">导出</button>
            </div>
        </div>
    `;

    var input = document.getElementById('source_path');
    input.classList.add('readonly');
    input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const debit_or_creditDropdown = document.getElementById('debit_or_creditDropdown')
    const creditDropdown = document.getElementById('creditDropdown');
    const debitDropdown = document.getElementById('debitDropdown');
    const nameDropdown = document.getElementById('nameDropdown');
    const bankDropdown = document.getElementById('bankDropdown');
    const numberDropdown = document.getElementById('numberDropdown');
    const valueDropdown = document.getElementById('valueDropdown');
    const credit_priority = document.getElementById('credit_priority');
    const debit_priority = document.getElementById('debit_priority');

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'bank_statement_sort_import', data: {'file_path': file_path }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'bank_statement_sort_index', data: {'file_path': file_path, 'sheetName': selectedSheet }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('debit_or_creditDropdown').addEventListener('change', async (event) => {
        const selectedSheet = sheetDropdown.value;
        const columnName = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'bank_statement_sort_debit_or_credit', data: {'file_path': file_path, 'sheetName': selectedSheet, 'columnName': columnName}});
    });

    // 自动触发工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('exportButton').addEventListener('click', async () => {
        if (!file_path) {
            alert('请先导入文件！');
            return;
        }

        const data = {
            file_path: file_path,
            sheet_name: sheetDropdown.value,
            debit_or_credit_column: debit_or_creditDropdown.value,
            credit_column: creditDropdown.value,
            debit_column: debitDropdown.value,
            name_column: nameDropdown.value,
            bank_column: bankDropdown.value,
            number_column: numberDropdown.value,
            value_column: valueDropdown.value,
            credit_priority: credit_priority.value,
            debit_priority: debit_priority.value
        };

        ipcRenderer.send('asynchronous-message', { command: 'bank_statement_sort_export', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'bank_statement_sort_import') {
            // 将文件路径显示在输入框中
            document.getElementById(`source_path`).value = file_path;

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

        if (result[0] === 'bank_statement_sort_index') {
    
            // 清空旧的选项
            debit_or_creditDropdown.innerHTML = '';
            nameDropdown.innerHTML = '';
            bankDropdown.innerHTML = '';
            numberDropdown.innerHTML = '';
            valueDropdown.innerHTML = '';
    
            result[1].forEach(column => {
                const debit_or_creditOption = document.createElement('option');
                debit_or_creditOption.value = column;
                debit_or_creditOption.text = column;
                debit_or_creditDropdown.appendChild(debit_or_creditOption);
    
                const nameOption = document.createElement('option');
                nameOption.value = column;
                nameOption.text = column;
                nameDropdown.appendChild(nameOption);
    
                const bankOption = document.createElement('option');
                bankOption.value = column;
                bankOption.text = column;
                bankDropdown.appendChild(bankOption);
    
                const numberOption = document.createElement('option');
                numberOption.value = column;
                numberOption.text = column;
                numberDropdown.appendChild(numberOption);

                const valueOption = document.createElement('option');
                valueOption.value = column;
                valueOption.text = column;
                valueDropdown.appendChild(valueOption);
            });
        }

        if (result[0] === 'bank_statement_sort_debit_or_credit') {

            // 清空旧的选项
            creditDropdown.innerHTML = '';
            debitDropdown.innerHTML = '';
    
            result[1].forEach(column => {
                const creditOption = document.createElement('option');
                creditOption.value = column;
                creditOption.text = column;
                creditDropdown.appendChild(creditOption);
    
                const debitOption = document.createElement('option');
                debitOption.value = column;
                debitOption.text = column;
                debitDropdown.appendChild(debitOption);
            });
        }

        if (result[0] === 'bank_statement_sort_export') {
            alert('导出成功！');
        }

        if (result[0] === 'bank_statement_sort_no') {
            alert('导出失败！');
        }
    });
}
// generate_chronogical_account.js

const { ipcRenderer } = require('electron');

let file_path = '';
let save_path = '';

export async function generate_chronological_accountFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>银行交易记录生成序时账</h1>
        
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
                <label for="numberDropdown">选择银行账号列</label>
                <select id="numberDropdown" name="numberDropdown" style="width: 30%;"></select>
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
                <label for="dateDropdown">选择日期列</label>
                <select id="dateDropdown" name="dateDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="nameDropdown">选择对方户名列</label>
                <select id="nameDropdown" name="nameDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="summaryDropdown">选择摘要列</label>
                <select id="summaryDropdown" name="summaryDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="currencyDropdown">选择币种列</label>
                <select id="currencyDropdown" name="currencyDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label for="valueDropdown">选择交易金额列</label>
                <select id="valueDropdown" name="valueDropdown" style="width: 30%;"></select>
            </div>

            <div>
                <label>纳入应收账款的户名</label>
                <textarea id="yszk_name" rows="7"></textarea>
            </div>

            <div>
                <label>纳入应付账款的户名</label>
                <textarea id="yfzk_name" rows="7"></textarea>
            </div>

            <div>
                <button id="importButton">导入</button>
                <button id="exportButton">导出</button>
            </div>
        </div>
    `;

    var file_path_input = document.getElementById('source_path');
    file_path_input.classList.add('readonly');
    file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('sheetDropdown');
    const numberDropdown = document.getElementById('numberDropdown');
    const debit_or_creditDropdown = document.getElementById('debit_or_creditDropdown')
    const creditDropdown = document.getElementById('creditDropdown');
    const debitDropdown = document.getElementById('debitDropdown');
    const dateDropdown = document.getElementById('dateDropdown');
    const nameDropdown = document.getElementById('nameDropdown');
    const summaryDropdown = document.getElementById('summaryDropdown')
    const currencyDropdown = document.getElementById('currencyDropdown')
    const valueDropdown = document.getElementById('valueDropdown');
    const yszk_name = document.getElementById('yszk_name');
    const yfzk_name = document.getElementById('yfzk_name');

    document.getElementById('importButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }
        // 將文件路徑發送到主進程
        ipcRenderer.send('asynchronous-message', { command: 'generate_chronological_account_import', data: {'file_path': file_path }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('sheetDropdown').addEventListener('change', async (event) => {
        const selectedSheet = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'generate_chronological_account_index', data: {'file_path': file_path, 'sheet_name': selectedSheet }});
    });

    // 当用户选择工作表时，获取该工作表的列索引
    document.getElementById('debit_or_creditDropdown').addEventListener('change', async (event) => {
        const selectedSheet = sheetDropdown.value;
        const columnName = event.target.value;

        // 将选择的工作表名发送到主进程，并请求该工作表的列
        ipcRenderer.send('asynchronous-message', { command: 'generate_chronological_account_debit_or_credit', data: {'file_path': file_path, 'sheet_name': selectedSheet, 'column_name': columnName}});
    });

    // 自动触发工作表选择的更改事件
    const triggerChange = (element) => {
        var event = new Event('change');
        element.dispatchEvent(event);
    };

    document.getElementById('exportButton').addEventListener('click', async () => {
        // 動態設置過濾器和默認文件名
        const saveFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];
        const defaultFileName = 'xlsx_output.xlsx';
        save_path = await ipcRenderer.invoke('dialog:saveFile', saveFilters, defaultFileName);

        if (!file_path) {
            alert('请先导入文件！');
            return;
        }

        if (!save_path) {
            alert('保存操作取消！');
            return;
        }

        const data = {
            file_path: file_path,
            save_path: save_path,
            sheet_name: sheetDropdown.value,
            number_column: numberDropdown.value,
            debit_or_credit_column: debit_or_creditDropdown.value,
            credit_column: creditDropdown.value,
            debit_column: debitDropdown.value,
            date_column: dateDropdown.value,
            name_column: nameDropdown.value,
            summary_column: summaryDropdown.value,
            currency_column: currencyDropdown.value,
            value_column: valueDropdown.value,
            yszk_name: yszk_name.value,
            yfzk_name: yfzk_name.value
        };

        ipcRenderer.send('asynchronous-message', { command: 'generate_chronological_account_export', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'generate_chronological_account_import') {
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

        if (result[0] === 'generate_chronological_account_index') {
    
            // 清空旧的选项
            numberDropdown.innerHTML = '';
            debit_or_creditDropdown.innerHTML = '';
            dateDropdown.innerHTML = '';
            nameDropdown.innerHTML = '';
            summaryDropdown.innerHTML = '';
            currencyDropdown.innerHTML = '';
            valueDropdown.innerHTML = '';
    
            result[1].forEach(column => {
                const numberOption = document.createElement('option');
                numberOption.value = column;
                numberOption.text = column;
                numberDropdown.appendChild(numberOption);

                const debit_or_creditOption = document.createElement('option');
                debit_or_creditOption.value = column;
                debit_or_creditOption.text = column;
                debit_or_creditDropdown.appendChild(debit_or_creditOption);
    
                const dateOption = document.createElement('option');
                dateOption.value = column;
                dateOption.text = column;
                dateDropdown.appendChild(dateOption);

                const nameOption = document.createElement('option');
                nameOption.value = column;
                nameOption.text = column;
                nameDropdown.appendChild(nameOption);

                const summaryOption = document.createElement('option');
                summaryOption.value = column;
                summaryOption.text = column;
                summaryDropdown.appendChild(summaryOption);

                const currencyOption = document.createElement('option');
                currencyOption.value = column;
                currencyOption.text = column;
                currencyDropdown.appendChild(currencyOption);

                const valueOption = document.createElement('option');
                valueOption.value = column;
                valueOption.text = column;
                valueDropdown.appendChild(valueOption);
            });
        }

        if (result[0] === 'generate_chronological_account_debit_or_credit') {

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

        if (result[0] === 'generate_chronological_account_export') {
            alert('导出成功！');
        }
    });
}
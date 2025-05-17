// set_up.js

const { ipcRenderer } = require('electron');

export async function set_up() {

    window.project_folder = await ipcRenderer.invoke('get-project-folder');

    if (window.project_folder) {
        ipcRenderer.send('asynchronous-message', { command: 'import_config', data: { project_folder: window.project_folder } });
    }

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    // 使用 configData 中的數據設置內容
    contentDiv.innerHTML = `

        <h1>项目信息设置</h1>

        <div class="import">

            <div>
                <label>被审计会计期间</label>
                <input id="period" type="text">
            </div>
            
            <div>
                <label>被审计会计报表截止日</label>
                <input id="deadline" type="text">
            </div>
            
            <div>
                <label>会计师事务所名称</label>
                <input id="accounting_firm" type="text">
            </div>

            <div>
                <label>保护单元格工作表密码</label>
                <input id="password" type="text">
            </div>

            <div>
                <label>企业名称</label>
                <input id="enterprise_name" type="text">
            </div>

            <div>
                <label>成立日期</label>
                <input id="date_of_establishment" type="text">
            </div>

            <div>
                <label>核准日期</label>
                <input id="approval_date" type="text">
            </div>

            <div>
                <label>统一社会信用代码</label>
                <input id="unified_social_credit_code" type="text">
            </div>

            <div>
                <label>注册资本</label>
                <input id="registered_capital" type="text">
            </div>

            <div>
                <label>法定代表人</label>
                <input id="legal_representative" type="text">
            </div>

            <div>
                <label>注册地址</label>
                <input id="registered_address" type="text">
            </div>

            <div>
                <label>经营范围</label>
                <textarea id="business_scope" rows="8"></textarea>
            </div>

            <div>
                <label>基本信息文件路径</label>
                <input id="basic_file_path" type="text">
            </div>

            <div>
                <label for="basic_sheetDropdown">选择基本信息工作表</label>
                <select id="basic_sheetDropdown" name="basic_sheetDropdown"></select>
            </div>

            <div>
                <button id="selectButton">选择基本信息文件</button>
                <button id="openButton">读取基本信息文件</button>
                <button id="saveButton">保存项目信息</button>
            </div>
        </div>
    `;

    const period_input = document.getElementById('period');
    const deadline_input = document.getElementById('deadline');
    const accounting_firm_input = document.getElementById('accounting_firm');
    const password_input = document.getElementById('password');
    const enterprise_name_input = document.getElementById('enterprise_name');
    const date_of_establishment_input = document.getElementById('date_of_establishment');
    const approval_date_input = document.getElementById('approval_date');
    const unified_social_credit_code_input = document.getElementById('unified_social_credit_code');
    const registered_capital_input = document.getElementById('registered_capital');
    const legal_representative_input = document.getElementById('legal_representative');
    const registered_address_input = document.getElementById('registered_address');
    const business_scope_input = document.getElementById('business_scope');

    const basic_file_path_input = document.getElementById('basic_file_path');
    basic_file_path_input.classList.add('readonly');
    basic_file_path_input.readOnly = true;

    const sheetDropdown = document.getElementById('basic_sheetDropdown');

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 定义文件筛选器
        const fileFilters = [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }];

        // 使用 ipcRenderer.invoke 与主进程通信，打开文件对话框
        const basic_file_path = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!basic_file_path) {
            console.log('File selection was canceled.');
            return;
        }

        // 将文件路径显示在输入框中
        basic_file_path_input.value = basic_file_path;

        const data = {
            basic_file_path: basic_file_path
        }
    
        ipcRenderer.send('asynchronous-message', { command: 'select_basic_file', data: data });
    });

    document.getElementById('openButton').addEventListener('click', async () => {

        const file_path = basic_file_path_input.value;

        if (!file_path) {
            console.log('File selection was canceled.');
            return;
        }

        const data = {
            basic_file_path: file_path,
            sheet_name: sheetDropdown.value
        }
    
        ipcRenderer.send('asynchronous-message', { command: 'import_basic', data: data });
    });

    document.getElementById('saveButton').addEventListener('click', async () => {
        if (!window.project_folder) {
            alert('请先选择项目文件夹！');
            return;
        }

        const data = {
            project_folder: window.project_folder,
            period: period_input.value,
            deadline: deadline_input.value,
            accounting_firm: accounting_firm_input.value,
            password: password_input.value,
            enterprise_name: enterprise_name_input.value,
            date_of_establishment: date_of_establishment_input.value,
            approval_date: approval_date_input.value,
            unified_social_credit_code: unified_social_credit_code_input.value,
            registered_capital: registered_capital_input.value,
            legal_representative: legal_representative_input.value,
            registered_address: registered_address_input.value,
            business_scope: business_scope_input.value
        };

        ipcRenderer.send('asynchronous-message', { command: 'save_settings', data: data });
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');
    
    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'import_config') {

            // 将配置文件中的数据回填到表单
            period_input.value = result[1]['被审计会计期间'];
            deadline_input.value = result[1]['被审计会计报表截止日'];
            accounting_firm_input.value = result[1]['会计师事务所名称'];
            password_input.value = result[1]['保护单元格工作表密码'];
            enterprise_name_input.value = result[1]['企业名称'];
            date_of_establishment_input.value = result[1]['成立日期'];
            approval_date_input.value = result[1]['核准日期'];
            unified_social_credit_code_input.value = result[1]['统一社会信用代码'];
            registered_capital_input.value = result[1]['注册资本'];
            legal_representative_input.value = result[1]['法定代表人'];
            registered_address_input.value = result[1]['注册地址'];
            business_scope_input.value = result[1]['经营范围'];

        } else if (result[0] === 'select_basic_file') {

            // 清空旧的选项
            sheetDropdown.innerHTML = '';

            // 遍历数据并创建 option 元素
            result[1].forEach(item => {
                const option = document.createElement('option');
                option.value = item;
                option.text = item;
                sheetDropdown.appendChild(option);
            });

            alert('选择文件成功！');

        } else if (result[0] === 'import_basic') {

            enterprise_name_input.value = result[1]['企业名称'];
            date_of_establishment_input.value = result[1]['成立日期'];
            approval_date_input.value = result[1]['核准日期'];
            unified_social_credit_code_input.value = result[1]['统一社会信用代码'];
            registered_capital_input.value = result[1]['注册资本'];
            legal_representative_input.value = result[1]['法定代表人'];
            registered_address_input.value = result[1]['注册地址'];
            business_scope_input.value = result[1]['经营范围'];

            alert('读取企查查信用报告（Excel操作版）成功！');

        } else if (result[0] === 'save_settings') {

            alert('项目信息设置页面保存项目信息成功！');

        }
    });
}
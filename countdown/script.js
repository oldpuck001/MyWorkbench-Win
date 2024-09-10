// script.js

const fs = require('fs');
const path = require('path');

export async function dashboardFunction() {

    const configFilePath = path.join(__dirname, 'countdown', 'countdown_config.json');

    // 讀取 countdown_config.json 文件
    try {
        const response = await fetch('./countdown/countdown_config.json');
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        let configData = await response.json();  // 解析 JSON 文件內容

        // 檢查讀取的數據結構是否正確
        if (!Array.isArray(configData.task_list)) {
            configData.task_list = [];
        }
        if (!Array.isArray(configData.days_off)) {
            configData.days_off = [];
        }
        if (!Array.isArray(configData.work_weekends)) {
            configData.work_weekends = [];
        }

        // 獲取 DOM 元素
        const contentDiv = document.getElementById('content');
        contentDiv.style.border = 'none';

        // 使用 configData 中的數據設置內容
        contentDiv.innerHTML = `

            <h1 style="text-align: center; width: 100%;">倒计时看板</h1>
            
            <div>
                <div class="import" style="display: flex; justify-content: center;">
                    <input class="input_date" type="date" style="width: 150px; border: none; background-color: #F0F4F8; padding: 10px 15px; outline: none; margin-right: 10px; text-align: center;" placeholder="目标日期">
                    <input class="input_task" style="width: 445px; border: none; background-color: #F0F4F8; padding: 10px 15px; outline: none; margin-right: 10px;" placeholder="新增项目">
                    <button id="saveBtn" class="button_task" style="background-color: #00c787; border: none; color: white; padding: 10px 15px; cursor: pointer;">+</button>
                </div>
                <div style="display: flex; justify-content: center;">
                    <ul class="list_task" style="list-style-type: none; padding: 0;"></ul>
                </div>
            </div>
        `;

        // 獲取 .list_task 元素
        const taskListUl = document.querySelector('.list_task');
        
        // 渲染初始任務列表
        renderTaskList(configData.task_list, configData.days_off, configData.work_weekends, taskListUl);

        // 計算剩餘自然日和工作日的函數
        function calculateDaysLeft(targetDate, daysOff, workWeekends) {
            const now = new Date();
            const endDate = new Date(targetDate);
            const diffTime = endDate - now;
            const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24)); // 剩餘自然日

            // 剩餘工作日的計算
            let workDaysLeft = 0;
            let currentDate = new Date(now);

            // 轉換 daysOff 和 workWeekends 列表為 Set（提升查詢效率）
            const daysOffSet = new Set(daysOff);
            const workWeekendsSet = new Set(workWeekends);

            while (currentDate <= endDate) {
                const currentDateStr = currentDate.toISOString().split('T')[0]; // 獲取日期的 YYYY-MM-DD 格式
                const dayOfWeek = currentDate.getDay();

                if (workWeekendsSet.has(currentDateStr)) {
                    // 如果是休息日，但這天在工作週末列表中，計入工作日
                    workDaysLeft++;
                } else if (dayOfWeek !== 0 && dayOfWeek !== 6 && !daysOffSet.has(currentDateStr)) {
                    // 如果是工作日，且不在 daysOff 列表中，計入工作日
                    workDaysLeft++;
                }

                currentDate.setDate(currentDate.getDate() + 1); // 增加一天
            }
            return { naturalDaysLeft: diffDays, workDaysLeft: workDaysLeft };
        }

        // 渲染任務列表的函數
        function renderTaskList(taskList, daysOff, workWeekends, taskListUl) {
            taskListUl.innerHTML = '';  // 清空列表，避免重複渲染
            taskList.forEach(task => {
                const daysLeft = calculateDaysLeft(task.targetDate, daysOff, workWeekends);
                insertTaskToList(task, daysLeft, taskListUl);
            });
        }

        // 創建並插入任務的函數，將任務信息插入到列表中
        function insertTaskToList(task, daysLeft, taskListUl) {

            // 創建一個 <li> 元素
            const li = document.createElement('li');
            li.style.display = 'flex';
            li.style.alignItems = 'center';
            li.style.padding = '10px 0';
            li.style.borderBottom = '1px solid #eee';
            
            // 將 targetDate 和 targetTask 加入到 <li> 中
            li.innerHTML = `
                <span class="countdown" style="width: 170px;">目标日期：${task.targetDate}</span>
                <span class="countdown" style="width: 250px;">目标任务：${task.targetTask}</span>
                <span class="countdown" style="width: 130px;">剩余自然日: ${daysLeft.naturalDaysLeft}</span>
                <span class="countdown" style="width: 130px;">剩余工作日: ${daysLeft.workDaysLeft}</span>
                <button class="del_task" aria-label="del_task" style="margin-left: auto;">🗑️</button>
            `;

            // 將 <li> 添加到 <ul> 中
            taskListUl.appendChild(li);

            // 綁定刪除按鈕的事件
            bindDeleteTaskEvent(li, task);
        };

        // 綁定刪除按鈕的事件
        function bindDeleteTaskEvent(taskItem, task) {
            const deleteButton = taskItem.querySelector('.del_task');
            deleteButton.addEventListener('click', function() {
                // 從 configData 中刪除任務
                configData.task_list = configData.task_list.filter(t => t.targetTask !== task.targetTask);

                // 更新 JSON 文件
                fs.writeFile(configFilePath, JSON.stringify(configData, null, 2), (err) => {
                    if (err) {
                        console.error('Error writing to countdown_config.json:', err);
                        return;
                    }

                    // 重新渲染任務列表
                    renderTaskList(configData.task_list, configData.days_off, configData.work_weekends, taskListUl);
                });
            });
        }

        // 保存按鈕點擊事件
        document.getElementById('saveBtn').addEventListener('click', saveTask);

        document.querySelector('.input_task').addEventListener('keydown', function(event) {
            if (event.key === 'Enter') {
                saveTask(); // 如果按下回車鍵，則執行保存
            }
        });

        // 將保存邏輯抽取到 saveTask 函數中
        function saveTask() {

            // 獲取目標日期的值與新增项目的值
            const inputDate = document.querySelector('.input_date').value;
            const inputTask = document.querySelector('.input_task').value;

            // 確認輸入的目標日期與新增项目不為空
            if (inputDate && inputTask) {

                // 創建新的任務對象
                const newTask = {
                    targetDate: inputDate,
                    targetTask: inputTask
                };

                // 將新項目添加到 configData 中
                configData.task_list.push(newTask);

                // 按照日期排序 task_list，確保是升序排列
                configData.task_list.sort((a, b) => new Date(a.targetDate) - new Date(b.targetDate));

                // 更新 JSON 文件
                fs.writeFile(configFilePath, JSON.stringify(configData, null, 2), (err) => {
                    if (err) {
                        console.error('Error writing to countdown_config.json:', err);
                        return;
                    }

                    // 成功保存後，重新渲染任務列表
                    renderTaskList(configData.task_list, configData.days_off, configData.work_weekends, taskListUl);

                    // 清空輸入框
                    document.querySelector('.input_date').value = '';
                    document.querySelector('.input_task').value = '';
                });
            } else {
                // 如果輸入值為空，則提示用戶
                alert('請填寫目標日期和新增項目');
            }
        }

    } catch (error) {
        console.error('Error fetching or parsing countdown_config.json:', error);
    }
}

export async function settingsFunction() {
    const configFilePath = path.join(__dirname, 'countdown', 'countdown_config.json');

    // 讀取 countdown_config.json 文件
    try {
        const response = await fetch('./countdown/countdown_config.json');
        if (!response.ok) {
            throw new Error('Network response was not ok');
        }
        const configData = await response.json();  // 解析 JSON 文件內容

        // 檢查讀取的數據結構是否正確
        if (!Array.isArray(configData.days_off)) {
            configData.days_off = [];
        }
        if (!Array.isArray(configData.work_weekends)) {
            configData.work_weekends = [];
        }

        // 獲取 DOM 元素
        const contentDiv = document.getElementById('content');
        contentDiv.style.border = 'none';

        // 使用 configData 中的數據設置內容
        contentDiv.innerHTML = `

            <h1 style="text-align: center; width: 100%;">设置</h1>
            
            <div style="display: flex; flex-direction: column; align-items: center;">

                <div style="width: 50%; margin-bottom: 20px;">
                    <label for="daysOff">休息的工作日:</label><br>
                    <textarea id="daysOff" rows="5" style="width: 100%;">${configData.days_off.join(',')}</textarea>
                </div>
                
                <div style="width: 50%; margin-bottom: 20px;">
                    <label for="workWeekends">工作的休息日:</label><br>
                    <textarea id="workWeekends" rows="5" style="width: 100%;">${configData.work_weekends.join(',')}</textarea>
                </div>

                <div style="text-align: center; width: 50%;">
                    <button id="saveBtn" style="padding: 10px 20px; width: 100%;">保存</button>
                </div>
            </div>
        `;

        // 保存按鈕點擊事件
        document.getElementById('saveBtn').addEventListener('click', function() {
            // 獲取輸入框中的數據，並拆分為數組
            const daysOffInput = document.getElementById('daysOff').value.split(',').map(day => day.trim());
            const workWeekendsInput = document.getElementById('workWeekends').value.split(',').map(day => day.trim());
            
            // 更新 configData
            configData.days_off = daysOffInput.filter(day => day !== '');  // 去除空行
            configData.work_weekends = workWeekendsInput.filter(day => day !== '');

           // 將更新後的 configData 寫入本地文件
            fs.writeFile(configFilePath, JSON.stringify(configData, null, 2), (err) => {
                if (err) {
                    console.error('Error saving file:', err);
                    alert('保存失敗！');
                } else {
                    alert('保存成功！');
                }
            });
        });

    } catch (error) {
        console.error('Error fetching or parsing countdown_config.json:', error);
    }
}
// text_comparison.js

const { ipcRenderer } = require('electron');

export async function text_comparisonFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>文本对比</h1>
        
        <div class="import">

            <div>
                <label>文本1：</label>
            </div>
            <div>
                <textarea id="text_1" rows="10"></textarea>
            </div>

            <div>
                <label>文本2：</label>
            </div>
            <div>
                <textarea id="text_2" rows="10"></textarea>
            </div>

            <div>
                <button id="comparisonButton">开始对比</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>对比结果：</label>
            </div>
            <div>
                <textarea id="result_output" rows="10"></textarea>
            </div>
        </div>
    `;

    const text_1 = document.getElementById('text_1');
    const text_2 = document.getElementById('text_2');

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    // 选择文件夹按钮js代码
    document.getElementById('comparisonButton').addEventListener('click', async () => {

        const data = {
            'text_1': text_1.value,
            'text_2': text_2.value
        }

        ipcRenderer.send('asynchronous-message', { command: 'text_comparison', data: data});
    });

    ipcRenderer.removeAllListeners('asynchronous-reply');

    ipcRenderer.on('asynchronous-reply', (event, result) => {

        if (result[0] === 'text_comparison') {

            document.getElementById(`result_output`).value = result[1];

        };
    });
}
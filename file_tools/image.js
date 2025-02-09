// image.js

const { ipcRenderer } = require('electron');

let num = 0;
let imagelist = [];

export async function imageFunction() {

    // 獲取 DOM 元素
    const contentDiv = document.getElementById('content');

    contentDiv.innerHTML = `

        <h1>自动修改文件名</h1>

        <div class="import">
            <div>
                <canvas id="imageCanvas" width="860" height="540"></canvas>
            </div>

            <div>
                <label style="width: 10%;">文件夹路径</label>
                <input id="file_path" type="text">
                <button id="selectButton">选择文件夹</button>
            </div>

            <div>
                <button id="previousButton">⇤</button>
                <button id="leftButton">⟲</button>
                <button id="zoominButton">+</button>
                <button id="zoomoutButton">−</button>
                <button id="rightButton">⟳</button>
                <button id="nextButton">⇥</button>
            </div>

            <div>
                <input id="name_1" type="text">
                <label style="width: 5px; margin-right: 5px; margin-left: 5px;">-</label>
                <input id="name_2" type="text">
                <label style="width: 5px; margin-right: 5px; margin-left: 5px;">-</label>
                <input id="name_3" type="text">
                <label style="width: 5px; margin-right: 5px; margin-left: 5px;">-</label>
                <input id="name_4" type="text">
                <label style="width: 5px; margin-right: 5px; margin-left: 5px;">-</label>
                <input id="name_5" type="text">
                <button id="renameButton">修改文件名</button>
            </div>
        </div>

        <div class="export">
            <div>
                <label>查找/修改结果：</label>
            </div>
            <div>
                <textarea id="result_output" rows="10"></textarea>
            </div>
        </div>
        `;

    var input = document.getElementById('file_path');
    input.classList.add('readonly');
    input.readOnly = true;

    var result_output = document.getElementById('result_output');
    result_output.classList.add('readonly');
    result_output.readOnly = true;

    document.getElementById('selectButton').addEventListener('click', async () => {
        // 向主進程發送請求，打開文件選擇對話框
        const fileFilters = [{ name: 'Image Files', extensions: ['png', 'jpg', 'jpeg'] }];
        const filePath = await ipcRenderer.invoke('dialog:openFile', fileFilters);

        if (!filePath) {
            console.log('File selection was canceled.');
            return;
        }

        document.getElementById(`file_path`).value = filePath;

        const data = {
            select_function: 'select',
            filePath: filePath
        };

        ipcRenderer.send('asynchronous-message', { command: 'filename_image', data: data });
    });

    const canvas = document.getElementById('imageCanvas');
    const ctx = canvas.getContext('2d');
    let img = new Image();
    let scale = 1.0;
    let rotation = 0;
    let offsetX = 0, offsetY = 0; // 圖片偏移位置
    let isDragging = false;
    let startX = 0, startY = 0;
    let lastX = 0, lastY = 0;
    let velocityX = 0, velocityY = 0;
    let friction = 0.9;  // 慣性滾動摩擦係數

    // 加載初始圖片
    img.src = 'filename_tools/image.jpeg';
    img.onload = () => {
        setInitialScaleAndDrawImage();
    };

    // 根據 canvas 尺寸設置圖片初始縮放比例
    function setInitialScaleAndDrawImage() {
        // 獲取圖片和 canvas 的寬高比
        const imgAspectRatio = img.width / img.height;
        const canvasAspectRatio = canvas.width / canvas.height;

        // 根據 canvas 與圖片的寬高比調整縮放比例
        if (imgAspectRatio > canvasAspectRatio) {
            // 當圖片寬高比大於 canvas，將縮放比例設為 canvas 寬度與圖片寬度的比例
            scale = canvas.width / img.width;
        } else {
            // 當圖片寬高比小於或等於 canvas，將縮放比例設為 canvas 高度與圖片高度的比例
            scale = canvas.height / img.height;
        }

        drawImage();
    }

    // 渲染圖片
    function drawImage() {
        ctx.clearRect(0, 0, canvas.width, canvas.height);  // 清除畫布
        ctx.save(); // 保存當前繪製狀態
        ctx.translate(canvas.width / 2 + offsetX, canvas.height / 2 + offsetY);
        ctx.rotate(rotation * Math.PI / 180);  // 旋轉圖片
        ctx.scale(scale, scale);  // 縮放圖片
        ctx.drawImage(img, -img.width / 2, -img.height / 2);  // 繪製圖片
        ctx.restore();  // 還原繪製狀態
    }

    // 按鈕事件綁定
    document.getElementById('zoominButton').onclick = () => {
        scale += 0.1;
        drawImage();
    };

    document.getElementById('zoomoutButton').onclick = () => {
        scale = Math.max(scale - 0.1, 0.1);  // 確保縮放比例不低於 0.1
        drawImage();
    };

    document.getElementById('leftButton').onclick = () => {
        rotation -= 90;
        drawImage();
    };

    document.getElementById('rightButton').onclick = () => {
        rotation += 90;
        drawImage();
    };

    canvas.addEventListener('mousedown', (e) => {
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        lastX = startX;
        lastY = startY;
    });

    canvas.addEventListener('mousemove', (e) => {
        if (isDragging) {
            let dx = e.clientX - lastX;
            let dy = e.clientY - lastY;
            offsetX += dx;
            offsetY += dy;
            lastX = e.clientX;
            lastY = e.clientY;
            drawImage();
            velocityX = dx;  // 記錄速度，用於慣性滾動
            velocityY = dy;
        }
    });

    canvas.addEventListener('mouseup', () => {
        isDragging = false;
        requestAnimationFrame(inertiaScroll);
    });

    canvas.addEventListener('mouseleave', () => {
        isDragging = false;
    });

    // 慣性滾動效果
    function inertiaScroll() {
        if (Math.abs(velocityX) > 0.1 || Math.abs(velocityY) > 0.1) {
            offsetX += velocityX;
            offsetY += velocityY;
            velocityX *= friction;  // 減少速度
            velocityY *= friction;
            drawImage();
            requestAnimationFrame(inertiaScroll);
        }
    }

    document.getElementById('previousButton').addEventListener('click', async () => {
        if (!imagelist) {
            return;
        }

        num = (num - 1 + imagelist.length) % imagelist.length;

        document.getElementById(`file_path`).value = imagelist[num];
        img.src = imagelist[num];
        img.onload = () => {
            setInitialScaleAndDrawImage();
        };
    });

    document.getElementById('nextButton').addEventListener('click', async () =>{
        if (!imagelist) {
            return;
        }

        num = (num + 1) % imagelist.length;

        document.getElementById(`file_path`).value = imagelist[num];
        img.src = imagelist[num];
        img.onload = () => {
            setInitialScaleAndDrawImage();
        };
    });

    document.getElementById('renameButton').addEventListener('click', async () => {    
        if (!imagelist) {
            console.log('File selection was canceled.');
            return;
        }

        const data = {
            select_function: 'rename',
            num: num,
            imagelist: imagelist,
            name_1: document.getElementById('name_1').value,
            name_2: document.getElementById('name_2').value,
            name_3: document.getElementById('name_3').value,
            name_4: document.getElementById('name_4').value,
            name_5: document.getElementById('name_5').value,
            result_text: document.getElementById('result_output').value
        };

        ipcRenderer.send('asynchronous-message', { command: 'filename_image', data: data });
    });

    ipcRenderer.on('asynchronous-reply', (event, result) => {
        if (result[0] === 'filename_image') {

            if (result[1][0] == 'select'){
                num = result[1][1][0];
                imagelist = result[1][1][1];

                img.src = imagelist[num];
                img.onload = () => {
                    setInitialScaleAndDrawImage();
                };
            };

            if (result[1][0] == 'no_new'){
                document.getElementById(`result_output`).value = '';
                document.getElementById(`result_output`).value = result[1][1]
            }

            if (result[1][0] == 'rename'){
                num = result[1][1][0];
                imagelist = result[1][1][1];

                document.getElementById(`file_path`).value = imagelist[num];
                img.src = imagelist[num];
                img.onload = () => {
                    setInitialScaleAndDrawImage();
                };

                document.getElementById(`result_output`).value = '';
                document.getElementById(`result_output`).value = result[1][1][2]
            }
        };
    });
}
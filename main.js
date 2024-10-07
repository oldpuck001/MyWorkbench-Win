// main.js

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { spawn } = require('child_process');
const path = require('path');

// 禁用硬件加速，在編寫程序用的MacBook pro2018上，使用硬件加速會報錯，所以禁用
app.disableHardwareAcceleration();

function createWindow() {

  // 創建瀏覽器窗口
  const win = new BrowserWindow({
    width: 1400,
    height: 820,
    webPreferences: {
      nodeIntegration: true,        // 禁用直接使用 Node.js 模块
      contextIsolation: false,      // 禁用上下文隔离，以便访问 window.myAPI
    }
  });

  // 加載應用的 index.html
  win.loadFile('index.html');
}

// 當 Electron 完成初始化並準備創建瀏覽器窗口時調用
app.whenReady().then(createWindow);

// IPC 處理文件選擇窗口
ipcMain.handle('dialog:openFile', async (event, filters) => {
  // 打開文件選擇對話框
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: filters
  });
  // 如果取消選擇文件，返回 null
  if (canceled) {
    return null;
  } else {
    return filePaths[0];  // 返回選擇的文件路徑
  }
});

// IPC 處理文件夾選擇窗口
ipcMain.handle('dialog:openDirectory', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openDirectory']
  });
  // 如果取消選擇，返回 null
  if (canceled) {
    return null;
  }
  return filePaths[0];
});

// IPC 處理文件保存窗口
ipcMain.handle('dialog:saveFile', async (event, filters, defaultFileName) => {
  // 打開保存文件對話框，讓用戶選擇保存文件的位置
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: 'Save File',
    defaultPath: defaultFileName,       // 設置預設文件名
    filters: filters                    // 使用前端傳遞的文件類型過濾器
  });
  // 如果取消保存文件，返回 null
  if (canceled) {
    return null;
  } else {
    return filePath;
  }
});

// IPC 處理與 Python 後端的通信
ipcMain.on('asynchronous-message', (event, arg) => {

  // 通過 spawn 函數創建的 Python 子進程對象，允許你與該子進程進行交互，例如監聽它的輸出、錯誤信息，或向其發送輸入數據。
  const pythonProcess = spawn('python3', [path.join(__dirname, 'backend.py')]);

  // 發送數據到 Python
  pythonProcess.stdin.write(JSON.stringify(arg));
  pythonProcess.stdin.end();

  // 讀取 Python 的輸出
  pythonProcess.stdout.on('data', (data) => {
    const response = JSON.parse(data.toString());
    event.reply('asynchronous-reply', response.result);
  });

  // pythonProcess.stderr 是 Python 子進程的標準錯誤輸出流（stderr）。
  pythonProcess.stderr.on('data', (data) => {
    console.error(`stderr: ${data}`);
  });

  // pythonProcess 是通過 spawn 創建的子進程對象，用於為子進程對象添加事件監聽器，監聽特定的事件並在事件觸發時執行回調函數。
  pythonProcess.on('close', (code) => {
    console.log(`Python process exited with code ${code}`);
  });
});

// 在 macOS 上的應用通常在關閉所有窗口後，應用本身仍然會保持活躍，圖標會繼續留在 Dock 中，因此檢查當前運行平台，避免所有窗口關閉後立即退出。
// darwin 是 Node.js 中用來標識 macOS 的操作系統平台名稱。名稱來自於 macOS 的核心內核 Darwin。
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// 在 macOS 上，所有窗口關閉後，可以點擊 Dock 中的應用圖標來重新打開應用。
// activate 事件允許應用在這種情況下重新創建窗口（如果沒有窗口打開）。
// 如果所有窗口都關閉了，則調用 createWindow() 函數來重新創建窗口。
app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
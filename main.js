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

// IPC 處理與 Python 後端的通信和文件選擇窗口
ipcMain.handle('dialog:openFile', async (event, filters) => {
  // 打開文件選擇對話框，讓用戶選擇一個 XLSX 文件
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

// 处理文件夹选择请求
ipcMain.handle('dialog:openDirectory', async () => {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openDirectory']
  });
  
  if (!canceled && filePaths.length > 0) {
    return filePaths[0]; // 返回选择的文件夹路径
  }
  return null; // 用户取消选择，返回 null
});

ipcMain.handle('dialog:saveFile', async (event, filters, defaultFileName) => {
  // 打開保存文件對話框，讓用戶選擇保存文件的位置
  const { canceled, filePath } = await dialog.showSaveDialog({
    title: 'Save File',
    defaultPath: defaultFileName,  // 設置預設文件名
    filters: filters               // 使用前端傳遞的文件類型過濾器
  });
  // 如果取消選擇文件，返回 null
  if (canceled) {
    return null;
  } else {
    return filePath;  // 返回選擇的文件保存路徑
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

  // pythonProcess 是通過 spawn 創建的子進程對象，代表 Python 腳本的執行過程。
  pythonProcess.on('close', (code) => {
    console.log(`Python process exited with code ${code}`);
  });
});

// 當所有窗口都關閉時退出應用。
// darwin 是 Node.js 中用來標識 macOS 的操作系統平台名稱。名稱來自於 macOS 的核心內核 Darwin。
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// 在 macOS 上，所有窗口關閉後，可以點擊 Dock 中的應用圖標來重新打開應用。
app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
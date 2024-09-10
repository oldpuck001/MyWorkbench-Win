// main.js

//從 Electron 中導入核心模塊，用來創建應用窗口和處理應用程序生命週期及進程間通信。
//app：Electron 應用程序的主進程對象。
//BrowserWindow：用來創建和管理應用的窗口。
//ipcMain：用於處理主進程和渲染進程之間的進程間通信（IPC, Inter-Process Communication）。
const { app, BrowserWindow, ipcMain } = require('electron');

//從 Node.js 的 child_process 模塊中導入 spawn，用於創建子進程，這通常用來運行外部程序（如 Python 腳本）。
const { spawn } = require('child_process');

//從 Node.js 的 path 模塊中導入，用來處理文件路徑，使其在不同的操作系統中都能正確運行。
const path = require('path');

// 禁用硬件加速，在編寫程序用的MacBook pro2018上，使用硬件加速會報錯，所以禁用
app.disableHardwareAcceleration();

function createWindow() {
  // 創建瀏覽器窗口
  // webPreferences 是 Electron 中 BrowserWindow 的一個選項對象，用來配置新窗口（即渲染進程）的各種設置。
  // nodeIntegration: true 在渲染進程中啟用 Node.js，這樣網頁腳本可以使用 Node.js 提供的所有功能（例如可以使用 require 加載模塊）。
  // 默認情況下，Electron 會禁用 Node.js 集成，這樣可以提高安全性，防止惡意腳本通過 Node.js API 操作文件系統或進行其他敏感操作。
  // contextIsolation 是一個控制是否將渲染進程中的 JavaScript 上下文與 Electron 提供的上下文進行隔離的設置。
  // false 禁用上下文隔離。這種情況下，安全性會降低，但可以更方便地集成前端和後端功能。
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

// IPC 處理與 Python 後端的通信
// ipcMain 是 Electron 中的一個模塊，用來在主進程中處理渲染進程發送的 IPC（進程間通信）消息。它用於監聽和處理來自渲染進程的請求。
// on() 是 ipcMain 模塊的方法，用於監聽特定類型的事件。
// 'asynchronous-message' 是事件名稱，表示主進程正在監聽來自渲染進程發送的這種類型的消息。
// 這個名稱可以是任意字符串，具體取決於你在應用中的命名規範。通常，這個名稱是渲染進程中通過 ipcRenderer.send() 發送消息時指定的。
// event 是 IPC 消息事件對象，包含了一些關於這次通信的信息。可以通過 event.reply() 方法回覆渲染進程，也可以用其他屬性來檢查它的詳細情況。
// arg 是從渲染進程發送過來的數據。渲染進程發送的第二個參數會作為 arg 傳遞到這裡。它可以是任何類型的數據，如字符串、數字、對象或 JSON 數據。
ipcMain.on('asynchronous-message', (event, arg) => {

  // 啟動 Python 進程，spawn 是 Node.js 的 child_process 模塊中的一個方法，它允許你創建一個子進程來運行一個外部命令。
  // 通過 spawn 函數創建的 Python 子進程對象，允許你與該子進程進行交互，例如監聽它的輸出、錯誤信息，或向其發送輸入數據。
  const pythonProcess = spawn('python3', [path.join(__dirname, 'backend.py')]);

  // 發送數據到 Python
  // stdin 代表的是 Python 子進程的標準輸入流。write() 是一個方法，用來將數據寫入標準輸入流。
  // JSON.stringify(arg)：將 JavaScript 對象 arg 轉換為 JSON 字符串。因為子進程（Python）通常無法直接理解 JavaScript 對象。
  // stdin.end() 方法用來結束輸入流。如果不調用 stdin.end()，Python 子進程可能會持續等待更多的輸入數據，導致應用無法正常繼續執行。
  pythonProcess.stdin.write(JSON.stringify(arg));
  pythonProcess.stdin.end();

  // 讀取 Python 的輸出
  // pythonProcess.stdout 是 Python 子進程的標準輸出流。
  // on('data', ...) 是一個事件監聽器，用來監聽標準輸出中是否有新的數據。
  // 'data' 是從Python子進程輸出的數據，以二進制或字符串的形式接收。
  // data 通常是以二進制格式（Buffer）傳入的，因此需要通過 toString() 方法將其轉換為可讀的字符串。
  // JSON.parse() 是 JavaScript 的方法，用來將 JSON 字符串轉換為 JavaScript 對象。
  // event.reply() 是 Electron 中的 ipcMain 模塊的一個方法，用來將數據從主進程發送回渲染進程。
  // 這裡的 event 通常是之前從渲染進程收到消息的事件對象。
  // 'asynchronous-reply' 是發送回渲染進程的消息標識符。這個標識符可以用來在渲染進程中識別這條消息。
  // response.result 是從 Python 腳本中獲取的處理結果。
  pythonProcess.stdout.on('data', (data) => {
    const response = JSON.parse(data.toString());
    event.reply('asynchronous-reply', response.result);
  });

  // pythonProcess.stderr 是 Python 子進程的標準錯誤輸出流（stderr）。
  // on('data', ...) 是一個事件監聽器，用來監聽子進程的 stderr 流中是否有數據。
  // 當有錯誤數據輸出時，'data' 事件會被觸發，並將錯誤數據作為回調函數的參數傳遞進來。
  // data 是Python腳本中輸出的錯誤信息，通常以二進制數據格式（Buffer）傳遞。需要將其轉換為可讀的字符串格式來顯示。
  // console.error() 是 Node.js 的一個方法，用於在控制台中打印錯誤信息。
  // ${data} 是一個變量插值語法，將 data（即 Python 腳本中輸出的錯誤信息）轉換為字符串後插入到模板字符串中。
  pythonProcess.stderr.on('data', (data) => {
    console.error(`stderr: ${data}`);
  });

  // pythonProcess 是通過 spawn 創建的子進程對象，代表 Python 腳本的執行過程。
  // .on() 方法用於為子進程對象添加事件監聽器，監聽特定的事件並在事件觸發時執行回調函數。
  // 'close' 是要監聽的事件名稱，表示子進程已經結束並關閉了所有標準輸入輸出流（stdin、stdout、stderr）。
  // 當 'close' 事件被觸發時要執行的回調函數，code 是子進程的退出代碼。
  // (code) => { console.log(...); } 是一個箭頭函數，當子進程觸發 'close' 事件時執行。
  // code 是子進程的退出代碼，是一個整數值。根據慣例，0 表示正常退出，非零值表示異常退出或錯誤。
  // console.log()：Node.js 的方法，用於在控制台輸出信息。
  // ${code} 將變量 code 的值插入到字符串中，從而顯示子進程的退出代碼。
  pythonProcess.on('close', (code) => {
    console.log(`Python process exited with code ${code}`);
  });
});

// 當所有窗口都關閉時退出應用。
// 在 Windows 和 Linux 系統上，當窗口全部關閉時，調用 app.quit() 來退出應用。
// 在 macOS 上的應用通常在關閉所有窗口後，應用本身仍然會保持活躍，圖標會繼續留在 Dock 中，因此檢查當前運行平台，避免所有窗口關閉後立即退出。
// process 是 Node.js 中的一個全局對象，提供了關於當前 Node.js 進程的信息和控制功能。
// platform 屬性，它返回當前運行應用的操作系統平台的字符串標識符。
// darwin 是 Node.js 中用來標識 macOS 的操作系統平台名稱。名稱來自於 macOS 的核心內核 Darwin。
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// 在 macOS 上，所有窗口關閉後，可以點擊 Dock 中的應用圖標來重新打開應用。
// activate 事件允許應用在這種情況下重新創建窗口（如果沒有窗口打開）。如果所有窗口都關閉了，則調用 createWindow() 函數來重新創建窗口。
app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});
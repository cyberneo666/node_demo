import { app, BrowserWindow, nativeTheme } from 'electron';
import path from 'path';
import os from 'os';
import { dialog, ipcMain } from 'electron';
import { transform2json } from './excel2json';
let excelPath = '';
let jsonDir = '';
async function handleFileOpen() {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
  });
  if (!canceled) {
    excelPath = filePaths[0];
    return filePaths[0];
  }
}
async function handleDirSave() {
  const { canceled, filePaths } = await dialog.showOpenDialog({
    properties: ['openDirectory'],
  });
  if (!canceled) {
    jsonDir = filePaths[0];
    return filePaths[0];
  }
}

async function handleTransform2Json() {
  if (excelPath == '' && excelPath == undefined) {
    // alert('Excel file is null!');
    return;
  }
  if (jsonDir == '' && jsonDir == undefined) {
    // alert('Target json directory is null!');
    return;
  }
  console.log(
    '[handleTransform2Json]excelPath:' + excelPath + ',jsonDir:' + jsonDir
  );
  await transform2json(excelPath, jsonDir);
}
// needed in case process is undefined under Linux
const platform = process.platform || os.platform();

try {
  if (platform === 'win32' && nativeTheme.shouldUseDarkColors === true) {
    require('fs').unlinkSync(
      path.join(app.getPath('userData'), 'DevTools Extensions')
    );
  }
} catch (_) {}

let mainWindow: BrowserWindow | undefined;
console.log(process.env.QUASAR_ELECTRON_PRELOAD);
function createWindow() {
  /**
   * Initial window options
   */
  mainWindow = new BrowserWindow({
    icon: path.resolve(__filename, 'icons/icon.png'), // tray icon
    width: 700,

    height: 400,
    useContentSize: true,
    webPreferences: {
      sandbox: false,
      contextIsolation: true,
      nodeIntegration: true,
      // More info: https://v2.quasar.dev/quasar-cli-vite/developing-electron-apps/electron-preload-script
      preload: path.resolve(__filename, process.env.QUASAR_ELECTRON_PRELOAD),
      // preload: path.resolve(__filename,'/Users/miya/Desktop/node_demo/electron-demo/src-electron/electron-preload.ts'),
    },
  });

  //
  ipcMain.on('open-file-dialog-for-xlsx', (event) => {
    dialog
      .showOpenDialog({
        properties: ['openFile'],
        filters: [{ name: 'Excel', extensions: ['xlsx', 'xls'] }],
      })
      .then((result) => {
        if (!result.canceled && result.filePaths.length > 0) {
          event.sender.send('selected-file', result.filePaths[0]);
        }
      })
      .catch((err) => {
        console.log(err);
      });
  });

  mainWindow.loadURL(process.env.APP_URL);

  if (process.env.DEBUGGING) {
    // if on DEV or Production with debug enabled
    mainWindow.webContents.openDevTools();
  } else {
    // we're on production; no access to devtools pls
    mainWindow.webContents.on('devtools-opened', () => {
      //mainWindow?.webContents.closeDevTools();
    });
  }

  mainWindow.on('closed', () => {
    mainWindow = undefined;
  });
  ipcMain.handle('load-prefs', () => {
    return {
      // 包含 preferences 的对象
    };
  });
}
ipcMain.handle('dialog:openFile', handleFileOpen);
ipcMain.handle('dialog:saveDir', handleDirSave);
ipcMain.handle('transform:save2json', handleTransform2Json);
app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
  if (platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (mainWindow === undefined) {
    createWindow();
  }
});

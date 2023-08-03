import { app, BrowserWindow, nativeTheme } from 'electron';
// import * as electronPath from 'path';
import os from 'os';
import {dialog,ipcMain} from 'electron'


// needed in case process is undefined under Linux
const platform = process.platform || os.platform();
const electronPath=require('path')
try {
  if (platform === 'win32' && nativeTheme.shouldUseDarkColors === true) {
    require('fs').unlinkSync(
      electronPath.join(app.getPath('userData'), 'DevTools Extensions')
    );
  }
} catch (_) {}

let mainWindow: BrowserWindow | undefined;

function createWindow() {
  /**
   * Initial window options
   */
  mainWindow = new BrowserWindow({
    icon: electronPath.resolve(__filename, 'icons/icon.png'), // tray icon
    width: 700,
    height: 400,
    useContentSize: true,
    webPreferences: {
      contextIsolation: false,
      nodeIntegration:true,
      // More info: https://v2.quasar.dev/quasar-cli-vite/developing-electron-apps/electron-preload-script
      preload: electronPath.resolve(__filename, process.env.QUASAR_ELECTRON_PRELOAD),
    },
  });

  mainWindow.loadURL(process.env.APP_URL);

  if (process.env.DEBUGGING) {
    // if on DEV or Production with debug enabled
    mainWindow.webContents.openDevTools();
  } else {
    // we're on production; no access to devtools pls
    mainWindow.webContents.on('devtools-opened', () => {
      mainWindow?.webContents.closeDevTools();
    });
  }
  
  mainWindow.on('closed', () => {
    mainWindow = undefined;
  });

  ipcMain.on('open-file-dialog-for-xlsx',(event)=>{
    dialog.showOpenDialog({
      properties:['openFile'],
      filters:[
        {name:'Excel',extensions:['xlsx','xls']}
      ]
    }).then(result=>{
      if(!result.canceled&&result.filePaths.length>0){
        event.sender.send('selected-file',result.filePaths[0]);
      }
    }).catch(err=>{
      console.log(err)
    });
  });
}

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

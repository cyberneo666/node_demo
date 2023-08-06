/**
 * This file is used specifically for security reasons.
 * Here you can access Nodejs stuff and inject functionality into
 * the renderer thread (accessible there through the "window" object)
 *
 * WARNING!
 * If you import anything from node_modules, then make sure that the package is specified
 * in package.json > dependencies and NOT in devDependencies
 *
 * Example (injects window.myAPI.doAThing() into renderer thread):
 *
 *   import { contextBridge } from 'electron'
 *
 *   contextBridge.exposeInMainWorld('myAPI', {
 *     doAThing: () => {}
 *   })
 *
 * WARNING!
 * If accessing Node functionality (like importing @electron/remote) then in your
 * electron-main.ts you will need to set the following when you instantiate BrowserWindow:
 *
 * mainWindow = new BrowserWindow({
 *   // ...
 *   webPreferences: {
 *     // ...
 *     sandbox: false // <-- to be able to import @electron/remote in preload script
 *   }
 * }
 */
// 注入 window.myAPI.doAThing() 到渲染进程的示例



//import { ipcRenderer } from "electron/renderer"
const { contextBridge,ipcRenderer } = require('electron')


contextBridge.exposeInMainWorld('api',
 {
    selectXlsxFile: ()  => {
            ipcRenderer.send('open-file-dialog-for-xlsx')
            ipcRenderer.on('selected-file',(event,args)=>{
                console.log(args)
                return args
            })
            console.log("selectXlsxFile");
            },
    // send:(channel: string,data: any)=>{
    //     ipcRenderer.invoke(channel,data).catch(e=>console.log(e))
    // },
    // receive:(channel: string,func: (arg0: string) => void)  => {
    //     console.log('preload:receive from '+ channel )
    //     ipcRenderer.on(channel,(event,args)=>func(args))   
    // }
 })
 contextBridge.exposeInMainWorld('electronAPI', {
    openFile: () => ipcRenderer.invoke('dialog:openFile')
  })
 